"""
This script updates an Excel file with ratings of usernames from lichess to
create a leaderboard. The Excel file will contain the names of the players.
The ratings and activity will be updated in the Excel file.
"""


import lichess
import openpyxl


def get_rating(username, control):
    # This function returns the rating of the username
    client = lichess.Client()
    user = client.get_data(f'{username}')
    rating = user['perfs'][f'{control}']['rating']
    return rating


def get_player_list():
    # This function returns a list of usernames from the Excel file
    wb = openpyxl.load_workbook("players.xlsx")
    sheet = wb.active
    player_list = [cell.value for cell in sheet['A']]
    return player_list


def write_rating(rating_list):
    # This function prints the ratings to the Excel file
    wb = openpyxl.load_workbook("players.xlsx")
    sheet = wb.active
    for index, value in enumerate(rating_list, start=1):
        sheet.cell(row=index, column=2, value=value)
    wb.save("players.xlsx")


def check_key(dictionary, target_key):
    # Recursive function. Returns a boolean value for a single instance of an activity
    if isinstance(dictionary, dict):
        if target_key in dictionary:
            return True
        else:
            for key, value in dictionary.items():
                if isinstance(value, dict):
                    if check_key(value, target_key):
                        return True
    return False


def check_activity(username, control):
    # This function checks for the weekly activity of the player
    client = lichess.Client()
    user = client.get_activity(f'{username}')
    # Create a list with boolean values of daily activity of time control
    activity = []
    for item in user:
        value = check_key(item, f'{control}')
        activity.append(value)
    status = "Inactive"
    if True in activity:
        status = "Active"
    return status


def write_activity(status_list):
    wb = openpyxl.load_workbook("players.xlsx")
    sheet = wb.active
    for index, value in enumerate(status_list, start=1):
        sheet.cell(row=index, column=3, value=value)
    wb.save("players.xlsx")


def main():
    print("Specify the time control")
    print("Blitz, Bullet, Classical, Rapid, etc")
    control = input(">> ").lower()

    # Fetch the list of players
    player_list = get_player_list()

    # Update the ratings of players
    rating_list = []
    for username in player_list:
        rating = get_rating(username, control)
        rating_list.append(rating)
    write_rating(rating_list)
    print("Ratings updated!")

    # Update status activity of players
    status_list = []
    for username in player_list:
        status = check_activity(username, control)
        status_list.append(status)
    write_activity(status_list)
    print("Status updated!")


if __name__ == '__main__':
    main()


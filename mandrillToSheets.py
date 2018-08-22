import mandrill
import argparse
import gspread
import traceback
import pandas as pd
from datetime import datetime, date, timedelta
from oauth2client.service_account import ServiceAccountCredentials

from config import GoogleSheetsConfig, MandrillConfig

config = GoogleSheetsConfig()
mandrill_config = MandrillConfig()


def find_element_in_list(element, list_element):
    """This is to handle the index search of the dates to remove
    list.index(item) throws a huge error if the value is not
    present.

    Args:
        element: Object to search for in the list
        list_element: List to search through

    Returns:
        Either the index of the item if it is present
        in the list, else None.
    """
    try:
        index_element = list_element.index(element)
        return index_element
    except ValueError:
        return None


# Command line argument parser
parser = argparse.ArgumentParser(
                    description="Grab stats from Mandrill, upload to Google Sheets")

parser.add_argument("--today",
                    help="Today's stats only. Else previous 28 days.",
                    action="store_false")

parser.add_argument("--campaign",
                    help="Campaign to grab the stats for.",
                    type=str,
                    required=True)

args = parser.parse_args()

if args:
    end_date = date.today()
    # if arg = True, then only todays stats
    # False * 28 = 0, True * 28 = 28
    days = 28 * args.today
    start_date = end_date - timedelta(days)

    API_KEY = mandrill_config.config["API_KEY"]
    campaign_config = mandrill_config.config["campaigns"][args.campaign]

    tags = campaign_config["tags"]
    SHEET_KEY = campaign_config["sheet"]

    # Initialise DataFrame, loop through the previous 28 days,
    # and add an empty row for sent, opens, clicks
    # Quicker than adding in the main loop, and allows for empty days

    report = pd.DataFrame(columns=['Date', 'Tag', 'Sent', 'Opens', 'Clicks'])

    # Range(days+1) will do the previous 28 days if days = 28
    # Or if days = 0 then just today

    for day in range(days+1):
        for tag in tags:
            report = report.append(pd.DataFrame({
                                'Date': (start_date + timedelta(days=day)).strftime("%Y-%m-%d"),
                                'Tag': tag,
                                'Sent': 0,
                                'Opens': 0,
                                'Clicks': 0
                                }, index=[0]
                                ), ignore_index=True)

    # Re-order report DataFrame as append
    # changes the column order to alphabetical
    report = report[['Date', 'Tag', 'Sent', 'Opens', 'Clicks']]

    # Connect to Mandrill
    mandrill_client = mandrill.Mandrill(API_KEY)
    date_fmt = "%Y-%m-%d %H:%M:%S"

    # Loop through the email tags, grab the stats from the daterange
    # Update each date in the report DataFrame

    for tag in tags:
        try:
            result = mandrill_client.messages.search_time_series(
                date_from=start_date.strftime(date_fmt),
                date_to=end_date.strftime(date_fmt),
                tags=[tag]
            )
            if result:
                for line in result:
                    # Mandrill response is in %Y-%m-%d %H:%M"%S format
                    # and given in ~2hr increments. So for each date we need
                    #  to sum the three stats.
                    mandrill_date = datetime.strptime(line['time'], date_fmt).strftime("%Y-%m-%d")
                    for index, row in report.iterrows():
                        if ((row['Date'] == mandrill_date) & (row['Tag'] == tag)):
                            report.loc[index, 'Sent'] += line['sent']
                            report.loc[index, 'Opens'] += line['unique_opens']
                            report.loc[index, 'Clicks'] += line['unique_clicks']

        except Exception as e:
            print(e)
            continue

    try:
        scope = config.scope
        cred_json = config.cred_json
        credentials = ServiceAccountCredentials.from_json_keyfile_dict(cred_json, scope)
        client = gspread.authorize(credentials)
        sheet = client.open_by_key(SHEET_KEY)
        worksheets = sheet.worksheets()

        for i, tag in enumerate(tags):
            # for the worksheet, grab the first column's values (Dates)
            dates_to_ignore = [x for x in worksheets[i].col_values(1) if x is not '']
            number_of_rows = worksheets[i].row_count
            number_of_free_rows = number_of_rows - len(dates_to_ignore)
            # If we are looking at just today instead of 28 days previous, only update today.
            if days == 0:
                dates_to_remove = [date.today().strftime("%Y-%m-%d")]
            else:
                dates_to_remove = [date.today().strftime("%Y-%m-%d"), (date.today()-timedelta(1)).strftime("%Y-%m-%d")]

            # Take the values we need to update, check if they exist in the column.
            # If it doesn't exist, it would break a normal list comprehension statement
            # So this function find_element_in_list is required
            indexs_to_remove = [find_element_in_list(x, dates_to_ignore) for x in dates_to_remove]

            # Remove the dates we need to update from our ignore list.
            [dates_to_ignore.remove(x) for x in dates_to_remove if find_element_in_list(x, dates_to_ignore) is not None]

            # Remove the rows containing the dates required to update.
            # Index + 1 because Google Sheets starts at 1, and Python at 0.
            [worksheets[i].delete_row(index+1) for index in indexs_to_remove if index is not None]

            # If there is data to push to Google Sheets, push it. Else don't.
            if report[(report.Tag == tag) & (~report.Date.isin(dates_to_ignore))].values.tolist():
                if (number_of_free_rows < len(report[(report.Tag == tag) & (~report.Date.isin(dates_to_ignore))])):
                    worksheets[i].add_rows(len(report[(report.Tag == tag) & (~report.Date.isin(dates_to_ignore))]))
                for row in report[(report.Tag == tag) & (~report.Date.isin(dates_to_ignore))].values.tolist():
                    worksheets[i].append_row(row)

    except Exception as e:
        print(e)
        traceback.print_exc()

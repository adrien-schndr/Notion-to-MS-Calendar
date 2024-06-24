from notion_client import Client
from O365 import Account, MSGraphProtocol
from dotenv import dotenv_values
import datetime as dt


config = dotenv_values(".env")

notion = Client(auth=config["NOTION_TOKEN"])

calendar_dict = notion.databases.query(
    **{
        "database_id": config["NOTION_DB_ID"],
        "filter": {
            "property": "Date",
            "date": {
                "on_or_after": "2024-06-20"
            },
        },
    }
)

clean_cal_list = []
liste_events_noms = []
for evenement in calendar_dict["results"]:
    temp_dict = {
        "ID": evenement["id"],
        "Name": evenement["icon"]["emoji"] + " | " + evenement["properties"]["Name"]["title"][0]["text"]["content"],
        "Type": evenement["properties"]["Type"]["select"]["name"]
    }
    liste_events_noms.append(evenement["properties"]["Name"]["title"][0]["text"]["content"])
    date_dict = evenement["properties"]["Date"]["date"]


    def date_format(date_dict, date_type):
        starting_date = date_dict[date_type][:10]
        if len(date_dict[date_type]) > 10:
            starting_time = date_dict[date_type][11:16]
        else:
            starting_time = None
        return starting_date, starting_time


    if date_dict["end"]:
        temp_dict["Date début"] = date_format(date_dict, "start")
        temp_dict["Date fin"] = date_format(date_dict, "end")
    else:
        temp_dict["Date"] = date_format(date_dict, "start")

    # location_dict = evenement["properties"]["Location"]["rich_text"]
    # if location_dict:
    #     temp_dict["Location"] = location_dict["name"]
    # else:
    #     temp_dict["Location"] = None

    clean_cal_list.append(temp_dict)


def afficher_dict(events_dict):
    for key, val in events_dict.items():
        print(key, ":", val)
    print("----------------------------------")


for evenement in clean_cal_list:
    afficher_dict(evenement)

credentials = (config["OUTLOOK_CLIENT"], config["OUTLOOK_SECRET"])

protocol = MSGraphProtocol()

scopes = ['Calendars.ReadWrite', 'User.Read']
account = Account(credentials, protocol=protocol)

if account.authenticate(scopes=scopes):
    print('Authenticated!')
    schedule = account.schedule()
    calendar = schedule.get_default_calendar()
    query = calendar.new_query('start').greater_equal(dt.date.today())
    events = calendar.get_events(query=query, include_recurring=False)
    for event in events:
        titre = event.subject[2:]
        emoji = {
            "type": "emoji",
            "emoji": event.subject[0]
        }
        start = event.start.isoformat()
        end = event.end.isoformat()
        location = event.location['displayName']
        if titre not in liste_events_noms:
            properties = {
                "Name": {"title": [{"text": {"content": titre}}]},
                "Date": {"date": {"start": start}},
                "Location": {"rich_text": [{"text": {"content": location}}]}
            }
            if end:
                properties["Date"]["date"]["end"] = end

            page = notion.pages.create(parent={"database_id": config["NOTION_DB_ID"]}, icon=emoji, properties=properties)

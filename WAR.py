import win32com.client
from tabulate import tabulate
from datetime import datetime, timedelta
import next_week_planned

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
calendar = outlook.GetDefaultFolder(9)
appointments = calendar.Items
appointments.Sort("[Start]")
appointments.IncludeRecurrences = "True"

begin = datetime.today() - timedelta(days=4)
end = datetime.today() + timedelta(days=1)
print(f"Activities from: {begin}, to: {end}")
restriction = "[Start] >= '" + begin.strftime("%m/%d/%Y") + "' AND [End] <= '" +end.strftime("%m/%d/%Y") + "'"
print("Restriction:", restriction)
restrictedItems = appointments.Restrict(restriction)

calcTableHeader = ['Date', 'Organizer', 'Subject']
calcTableBody_red = []
calcTableBody_yellow = []
calcTableBody_misc = []
calcTableBody_canceled = []

for appointmentItem in restrictedItems:
    red = []
    yellow = []
    misc = []
    canceled = []
    if "Canceled" not in appointmentItem.Subject:
        if appointmentItem.Categories == "Red Category":
            red.append(appointmentItem.Start.strftime("%m/%d"))
            red.append(appointmentItem.Organizer)
            red.append(appointmentItem.Subject)
            calcTableBody_red.append(red)
        elif appointmentItem.Categories == "Yellow Category":
            yellow.append(appointmentItem.Start.strftime("%m/%d"))
            yellow.append(appointmentItem.Organizer)
            yellow.append(appointmentItem.Subject)
            calcTableBody_yellow.append(yellow)
        elif appointmentItem.Categories == "Maroon Category" or appointmentItem.Categories == "Blue Category" or appointmentItem.Categories == "Orange Category" or appointmentItem.Categories == "Green Category":
            misc.append(appointmentItem.Start.strftime("%m/%d"))
            if appointmentItem.Organizer != "":
                misc.append(appointmentItem.Organizer)
            else:
                misc.append("N/A")
            misc.append(appointmentItem.Subject)
            calcTableBody_misc.append(misc)
    else:
        canceled.append(appointmentItem.Start.strftime("%m/%d"))
        canceled.append(appointmentItem.Organizer)
        canceled.append(appointmentItem.Subject)
        calcTableBody_canceled.append(canceled)

#struct: [['date', 'organizer', 'subject'], [#2]]

war_end = end - timedelta(days=1)

with open("war.html", "w") as f:
    print("<strong>Weekly Report –  Zach Denney</strong><br>", file=f)
    print("<strong>Dates: {}-{}</strong><br>".format(begin.strftime("%m/%d"), war_end.strftime("%m/%d")), file=f)
    print("<br><strong>On-Site Meetings:</strong><br><br>", file=f)
    print(tabulate(calcTableBody_red, headers=calcTableHeader, tablefmt="html"), file=f) #tablefmt="fancy_grid"
    print("<br><strong>Virtual Meetings:</strong> <br><br>", file=f)
    print(tabulate(calcTableBody_yellow, headers=calcTableHeader, tablefmt="html"), file=f)
    print("<br><strong>Misc. Meetings:</strong><br><br>", file =f)
    print(tabulate(calcTableBody_misc, headers=calcTableHeader, tablefmt="html"), file=f)
    print(f"<br><strong>Canceled Meetings ({len(calcTableBody_canceled)}):</strong><br><br>", file =f)
    print(tabulate(calcTableBody_canceled, headers=calcTableHeader, tablefmt="html"), file=f)
    print("<br><strong>Next Week Planned:</strong><br><br>", file=f)
    print("<br><img src=\"cropped.jpg\"><br><br>", file=f)
    f.close()

next_week_planned.findOutlook()
time.sleep(5)
next_week_planned.next_week_planned()


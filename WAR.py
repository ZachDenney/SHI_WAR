import win32com.client
from tabulate import tabulate
from datetime import datetime, timedelta
import next_week_planned
import time
import winsound

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
calendar = outlook.GetDefaultFolder(9)
appointments = calendar.Items
appointments.Sort("[Start]")
appointments.IncludeRecurrences = "True"

#end = datetime.today() + timedelta(days=1) #assumes you run on friday, if else statement below makes running day dynamic
if datetime.weekday(datetime.today()) != 4:
    fridayDelta = 4 - datetime.weekday(datetime.today())
else:
    fridayDelta = 1

begin = datetime.today() - timedelta(days=fridayDelta+1)
end = datetime.today() + timedelta(days=fridayDelta+1)

print(f"Activities from: {begin}, to: {end}")
restriction = "[Start] >= '" + begin.strftime("%m/%d/%Y") + "' AND [End] <= '" +end.strftime("%m/%d/%Y") + "'"
print("Restriction:", restriction)
restrictedItems = appointments.Restrict(restriction)

# Calendar Color/Category Mapping
# Red --> In Situ
# 	Red = External In Situ
# 	Maroon = Internal In Situ
#
# Yellow --> Virtual
# 	Yellow = Virtual
#
# Misc --> Misc
# 	Blue = Optional Misc
# 	Green = Optional Relationship Building
# 	Orange = Priority-Business
#
#
# Don't show:
#
# Purple = Padding
# Green = Incentive
# Green = Holiday
# Orange = Priority-Personal
# Dark Blue = Tracking


calcTableHeader = ['Date', 'Organizer', 'Subject']
calcTableBody_InSitu = []
calcTableBody_Virtual = []
calcTableBody_misc = []
calcTableBody_canceled = []

for appointmentItem in restrictedItems:
    InSitu = []
    Virtual = []
    misc = []
    canceled = []
    if "Canceled" not in appointmentItem.Subject:
        if appointmentItem.Categories == "External In Situ":
            InSitu.append(appointmentItem.Start.strftime("%m/%d"))
            InSitu.append(appointmentItem.Organizer)
            InSitu.append(appointmentItem.Subject)
            calcTableBody_InSitu.append(InSitu)
        elif appointmentItem.Categories == "Virtual":
            Virtual.append(appointmentItem.Start.strftime("%m/%d"))
            Virtual.append(appointmentItem.Organizer)
            Virtual.append(appointmentItem.Subject)
            calcTableBody_Virtual.append(Virtual)
        elif appointmentItem.Categories == "Internal In Situ" or appointmentItem.Categories == "Optional Misc" or appointmentItem.Categories == "Priority Business" or appointmentItem.Categories == "Optional Relationship Building":
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

war_begin = begin + timedelta(days=1)
war_end = end - timedelta(days=1)

with open("war.html", "w") as f:
    print("<strong>Weekly Report –  Zach Denney</strong><br>", file=f)
    print("<strong>Dates: {}-{}</strong><br>".format(war_begin.strftime("%m/%d"), war_end.strftime("%m/%d")), file=f)
    print(f"<br><strong>On-Site Meetings ({len(calcTableBody_InSitu)}):</strong><br><br>", file=f)
    print(tabulate(calcTableBody_InSitu, headers=calcTableHeader, tablefmt="html"), file=f) #tablefmt="fancy_grid"
    print(f"<br><strong>Virtual Meetings ({len(calcTableBody_Virtual)}):</strong> <br><br>", file=f)
    print(tabulate(calcTableBody_Virtual, headers=calcTableHeader, tablefmt="html"), file=f)
    print(f"<br><strong>Misc. Meetings ({len(calcTableBody_misc)}):</strong><br><br>", file =f)
    print(tabulate(calcTableBody_misc, headers=calcTableHeader, tablefmt="html"), file=f)
    print(f"<br><strong>Canceled Meetings ({len(calcTableBody_canceled)}):</strong><br><br>", file =f)
    print(tabulate(calcTableBody_canceled, headers=calcTableHeader, tablefmt="html"), file=f)
    print("<br><strong>Next Week Planned:</strong><br><br>", file=f)
    print("<br><img src=\"cropped.jpg\"><br><br>", file=f)
    f.close()

next_week_planned.findOutlook()
time.sleep(3)
next_week_planned.next_week_planned()
frequency = 2500  # Set Frequency To 2500 Hertz
duration = 500  # Set Duration To 1000 ms == 1 second
winsound.Beep(frequency, duration)

# SHI_WAR
Automate WAR process.

WAR.py == main method. uses win api to grab outlook calendar for current week, and dump data by category. effectivley i have harded a 1:1 linking between category type, and calendar classification... so, setting up category types is required. ex: red = onsite(customer), blue = misc meetings, etc. 

next_week_planned == grab outlook window, bring front, send keypress to find calendar, snapshot and crop calendar image for main war report.

Don't judge the messy code-- it was built for speed, not efficiency (and definitely not elegance).

Additional work intended:
1. streamline/modify into proper init/methods.
2. automate CRM interaction.
3. auto generate email (ask if additional input required/commentary).

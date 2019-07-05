# SHI_WAR
Automate WAR process.

WAR.py == main method. uses win api to grab outlook calendar for current week, and dump data by category.

next_week_planned == grab outlook window, bring front, send keypress to find calendar, snapshot and crop calendar image for main war report.

Don't judge the messy code-- it was built for speed, not efficiency (and definitely not elegance).

Additional work intended:
1. streamline/modify into proper init/methods.
2. automate CRM interaction.
3. auto generate email (ask if additional input required/commentary).
4. work on window wildcarding as sometimes Outlook doesn't detect properly.
5. timedelta still a bit sensitive to +/- 1 day (timestamp related). work on this variable calculation.
6. breakout customer name. should be able to take appointment and seperate by ":" (ex: subject line = company: topic of meeting), and add to own category.

Calendar Color/Category Mapping (change these in outlook to link to category breakout in WAR.py)

-Red --> In Situ   
   -Red = External In Situ  
   -Maroon = Internal In Situ  
-Yellow --> Virtual  
    -Yellow = Virtual  
-Misc --> Misc  
    -Blue = Optional Misc  
    -Green = Optional Relationship Building  
    -Orange = Priority Business  
-Don't Present:  
    -Purple = Padding  
    -Green = Incentive  
    -Green = Holiday  
    -Orange = Priority Personal  
    -Dark Blue = Tracking  
  


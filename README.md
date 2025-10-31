# google-cal2google-sheet

# Requirements

Have a Google calendar.


# Installation

1. Create a blank google sheet.
2. Go to Extensions > Apps Script
3. Dump this script in
4. Adjust the calendar year
5. Review the code so you're cool with letting it see your calendar
6. Save it
7. Run the `etl_cal` function
8. Allow permissions
9. You should see an execution log find your calendar and process events -- it may take a while if your calendar is crazy
10. After it's done, go to your sheet and make some graphs


# Future Items

- Color code things by their color in Google Calendar
- Exclude vacation and out-of-office better
- Option to exclude "meetings" without any other attendees
- If it's the current year, stop reporting after the current week b/c it will only show lots of free time, some planned meetings, and recurring meetings
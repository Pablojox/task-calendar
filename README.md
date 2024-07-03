# Create Calendar project (VBA)
## 🗺️ Context
I was working for a company that creates an annual calendar of tasks to plan their work according to project management principles. Typically, they generate the calendar for the upcoming year using an Excel file. This involves creating formulas to add one year to the previous calendar and checking if the new dates fall on weekends or holidays to make necessary adjustments.

The company wanted to make this process faster and more efficient.

## 🎯 Objectives
Develop a VBA script to increment each date by one year and adjust for weekends and holidays.

## ✅ Solution steps
The steps for the script are as follows:
- Set the list of holidays for the next year.
- Add one year to each date.
- Create a loop to iterate through each date and:
  - Check if the date falls on a holiday and subtract one day if it does.
  - Check if the date falls on a Saturday (subtract one day) or a Sunday (subtract two days).
- Repeat the loop three times to account for scenarios where two holidays fall in the same week, such as a Thursday and Friday. If a date falls on the weekend of that week, it needs to go through the loop three times to ensure all adjustments are made correctly.

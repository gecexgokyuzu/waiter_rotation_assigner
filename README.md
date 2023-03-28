Restaurant Waiter Scheduler

This is a Python script for scheduling restaurant waiters based on their history and table sections. The script uses the pandas and openpyxl libraries for data manipulation and Excel file handling.

Setup------------

Install the required libraries with pip install pandas openpyxl.

How to use------------

1. Open the RotationCopy.xlsx file in the Data folder and update the sheet named Cycle with the date and the sections where the waiters were assigned for the day.

2. Run the script by executing python scheduler.py in the terminal.

3. The script will assign waiters to sections according to the rules in the CheckAvailableWaiter() function and update the section_assignments dictionary.

4. The updated section_assignments dictionary will be logged to the RotationCopy.xlsx file under the Cycle sheet along with the date.

5. The cells in the RotationCopy.xlsx file will be colored to indicate the location of the section assigned to the waiter.

Dependencies------------

This code has the following dependencies:

pandas
openpyxl

This code is licensed under the MIT License.
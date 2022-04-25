# Timesheet Butler

This is a little utility for those who, like me, have a hard time filling out your timesheet every day and don't remember the details of the week by the time Friday rolls around.

Once it is installed, you will get a pop up once an hour that asks you what you have been working on. It then saves this information to an Excel file. At the end of the week, you'll have a detailed record of how the week went that should improve the reliability of your timesheet's data :)

## Requirements

This script requires the following to be installed on your computer:

* Python 3
* the python xlwings library
* Excel

## Installation

This is a quick script so that installation is a manual process.

### Install dependencies

If you do not have python installed, install "Anaconda 3" from ArupApps.

### Save the files

Clone or download the repo to your hard drive.

### Update the file paths

Copy the "record of time.xlsx" file to whereever you want to store it.

In "butler.bat", change to python path, and the path to the butler.py script to be correct for your system.

In "butler.py", change the path to 'record of time.xlsx' to be correct for your system.

### Set up Scheduled Tasks

Search for "Scheduled Tasks" in the Start Menu. Add a new task with "butler.bat" as the target. Have a look at "Timesheet Butler.xml" for the triggers I use.

## Enjoy!

Each time you answer Timesheet Butler's question, it will be added to the top of the Excel file - you can then refer to it when you do your timesheet at the end of the week.
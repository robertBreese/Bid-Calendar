# Bid-Calendar

A VBA macro for Excel Calendar template that auto creates and finds folders within a SharePoint (or local) library based on calendar event name. Also uses Outlook VBA to send calendar event invites to a 365 group. Macro has the ability to search two separate folders for the project currently. 

## Installation

Import .bas modules into one of the calendar templates found on this page:
https://www.vertex42.com/calendars/2022.html

Can easily be modified to work with other templates.

## Usage

Create a settings sheet with the following information (examples shown, place into next cell):

```
Folder A | \Test Folder\Customers - Documents\Test\Jan to Dec 2022 Bids
Folder B | \Test Folder\Customers - Documents\Test\
Recepients | <email> - either a single email or group can be used. 
Shared Mailbox | <group name> - either a single email or group can be used. 
Save Path | \Test Folder\Customers - Documents\Test\ - Where new projects that don't already exist are created.
Search Steps | <value> (1, 2, 3, etc) how many nested folders deep beyond the parent path the macro will search.
Share Point Folder | \Test Folder\ - typically the name of \Test Folder\te, how it displays as a path within Windows.
```
Type a project name into any date on the calendar, then click the project name once. The macro will search for a folder by that name, open it in the file explorer and create an event within Outlook with reminder one day before (Outlook desktop must be installed as this uses VBA). If the macro cannot find a folder by that name, it will create a new folder. You can add project ID's, for example, in front of project names and the macro will still find the correct folder without issue even if that ID doesn't exist in the folder name. 

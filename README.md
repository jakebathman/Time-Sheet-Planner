### Time-Sheet-Planner

This is a tool used internally at Collin County HLS to manage PeopleSoft time for employees.

**The current version is v7.9 which can be downloaded here: http://jakebathman.com/work/timesheetplanner/**

========

### Changelog

**v7.9 (released 3/3/2015)**
* Added Office Closure time code to dropdown on first sheet
* Updates to staff name list

**v7.8 (unreleased)**
 
**v7.7 (released 4/14/2014)**
* Minor changes to formatting and formulas on first sheet
* Updates to staff name list

**v7.6 (released 1/21/2014)**
* Minor bug and formatting fixed
* Removed button to import from PeopleSoft, since it's broken
* Added & removed some employee names
* Fixed error in email times email address list

**v7.5 (released 11/25/2013)**
* Support for accrued time off and generating a form when it happens

**v7.4 (released 10/21/2013)**
* Revisions and updates to Time Off Sheet formatting and back-end code
* Other minor formatting and code changes

**v7.3.1 (released 9/23/2013)**
* Minor bug fixes

**v7.3 (unreleased)**
* Fully functional Time Off Form
* New button on main sheet when time off exists
* Minor formatting changes and bug fixes

**v7.2 (unreleased)**
* Added drop-down for time off type selection
* Minor reformatting of main sheet

**v7.1.2 (released 9/4/2013)**
* Minor bug fix with entering time into Time Off Form sheet

**v7.1.2 (unreleased)**
* Calendar control wasn't going to work on other computers, so it was replaced with a stock control only version
* Added back the time off code column on main sheet, by request
* Other minor bug fixes, code cleanup, and sheet formatting

**v7.1 (released 9/4/2013)**
* In User Preferences sheet, now a "No Prompt" can be toggled to decrease prompts (right now only suppresses Clear Sheet prompt).
* Fixed major bug in Email Times button code
* Added Time Off Form sheet with helper code
* Minor main-sheet formatting changes (Issues #4, #5)
* Fixed bug with string logic for something like 30742 (3:07:42 pm). Logic wasn't dividing string correctly to calculate time. Changed mid(str,1,2) to mid(str,2,2) (was outputting min as 30 instead of 07).

**v7.0.1 (released 8/5/2013)**
* Removed "Example" sheet
* Changed name of prefs sheet to "User Preferences" 
* Removed peoplesoft prompt from Email button, if no times existed (now just ends with warning)
* Added Oscar to Email list, and revised the order of people to email (now, direct supervisors are first)
* Revision to punch calculation, where only two punches in the first and second columns will produce a time (fixed from prior releases)


**v7.0 (unreleased)**
* Added Supervisor sheet for calculating by pay period
* Added code to GitHub, which will track all issues and the changelog from now on: https://github.com/jakebathman/Time-Sheet-Planner
* Re-formatted main sheet
* Commented out backup prompt/code on Clear Sheet (untested)
* Added Instructions form to eliminate clunky text on the sheet
* Removed hidden rows & relevant code ("step 2" stuff)
* Removed or commented out backup code and sheets, since they're pretty much worthless and only slow everything down
* Slightly revised the auto-PM code in main sheet change event


**v6.5 Beta (unreleased)**
* To-Do: Fix PeopleSoft importing (maybe add web browser or refreshable web query?)
* To-Do: Fix last punch estimator (not showing up anymore, because formulas point to blank hidden rows)
* To-Do: Remove references to hidden rows in code, and delete those rows
* Fixed lingering progress bar when cancelling the clearing of a sheet
* Changed how the Clear Times sub works, since we don't care about the hidden rows anymore (simply clears a range of cells now, instead of slowly looping)
* Added Sunday to the timesheet (even though it's rarely used)
* To-Do: Automatically email about time off needed to be entered, if any exists (only for full-weeks? Figure out how to calculate)
* To-Do: If entering afternoon (last punch) in a row, make it PM if it looks like it should be
* Times with seconds can now be entered without punctuation (e.g. 12:22:46 can be entered as 122246 and will format correctly)
* Times entered in OUT column are assumed to be PM and a final (night) punch


**v6.4 RC (released 1/7/2013)**
* Added prompt on PeopleSoft Importer notifying user that it's currently broken

**v6.3 RC (released 12/10/2012)**
* Adjusted Clear Times sub to remove comments, which could incorrectly report previously imported PeopleSoft times
* Re-wrote adding code, which is all done in the background and doesn't use (breakable) sheet formulas
* Slightly revised Clear Times sub to run faster and with less screen flicker

**v6.2 RC (released 05/29/2012)**
* Added "working" form to Clear Times button
* Fixed bug where some subs wouldn't heed command to not show application alerts, which resulted in user confusion
* Hidden working sub now included, which will eventually process PeopleSoft HTML code for timesheet import
* After clearing times on main page, active cell now returns to Monday's first punch 

**v6.2 beta (released 02/10/2012)**
* Revised pre-populated list of names/emails for emailling times based on staff turnover
* Fixed bug in Email sub that halted when no Outlook signature file existed (based on font and font size)
* Fixed bug in VBA references when running on older versions of Office (removed superfluous references)
* Added error catching code in Email sub, prompting with a more descriptive (and user-friendly) error message

**v6.1 beta (released 11/03/2011)**
* Fixed Bug: Importing from PeopleSoft was not possible when using any browser other than Chrome
* Fixed Bug: Various input revision bugs when inputting directly to sheet (not using military time)
* Fixed: (v.6) New Bug: adding an OUT time should check to see if it's after the IN time, and adjust to military accordingly (e.g. 5-->1700)
* Fixed Bug: Inputting times directly to sheet using two digits over 24 rolls over to future days (fix: disallow and blank >24 inputs)
* Added: highlight (faint) those times that were either new or confirmed with PeopleSoft imports (cleared as well). Should retain between edits.
* Fixed Bug: adding/deleting additional names and emails inside email times routine is now more robust
* Added: when running PeopleSoft importer, now checks for open window with timesheet. If not, prompts to open timesheet URL
* Fixed: (v.6) New Bug: Highlighting times in email script does not work, sometimes highlights header values
* New Bug: Too much screen flicker (probably because of poorly-placed Application.ScreenUpdating commands)

**v6.0 beta (released 10/24/2011)**
* Beta release of PeopleSoft importer
* Allows inputting of times in sheet in military time without : seperator (e.g., 1300, 13, 13:00, 13.00 all convert to 1:00:00 PM)
* Beta release of button to email times to someone
* New Bug: Midnight times are not converted correctly if : seperator is not used (fix: look for hrs>24)
* New Bug: Highlighting times in email script does not work, sometimes highlights header values
* New Bug: adding an OUT time should check to see if it's after the IN time, and adjust to military accordingly (e.g. 5-->1700)
* Fixed Bug: selecting multiple options in user prefs now disallowed; part time employees added, new comp time forumla in sheet

**v5.0 (released 8/15/2011)**  
* Updated example sheet to reflect new planner
* Added button on main sheet to clear previous times (uses macro)

**v4.0 (released 4/29/2011)**  
* Began changelog
* Fixed bug that could prevent Friday punch helper from appearing
* Added formula in Friday punch helper that will make automatic 1-hour lunch more obvious to user
* Added time code for holiday hours
* Fixed bug that could incorrectly report the number of incorrect or missing timecodes for time off

========

### Task list for future development

* Test and refine PeopleSoft importer
* PS Importer currently discards (ignores) time off values
* Update conflict resolution to include ability to move punches (for missing IN PS punches)
* Decrease number of prompts to user; become more autonomous
* ~~Eventually fully automate calculations, instead of using hidden cells~~
* ~~Re-work design and layout to be easier to understand (less big blocks of text)~~
* Adjust last-punch calculator to be VBA, include all contingencies and proper trigger for showing
* Make auto-round of PeopleSoft times a user-selectable option (default rounded)


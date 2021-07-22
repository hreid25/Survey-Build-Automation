# Survey-Build-Automation
Using Selenium and Python to automate Engagement Survey construction
at McLean & Company. Primary goal is to reduce low effort (copy-paste), 
repetitive work (button-clicking), while reduce human error.

This program interacts with the Sergeant platform to automatically construct engagement 
surveys (pulse surveys, PEP surveys and full engagement). 
Process:

1. Returned Survey Build Doc (from client) is read into the program which we then use to construct a 2d matrix. 

2. Adding Hovers to Question Items in Array
The program parses 'hovers' from one of the excel worksheets and replaces text strings in the array with the newly
created hover, for every language if more than one is present.
Example: {{"Executive Team (The CEO,CIO,CFO,COO)" | hover}}

3. Creates Chrome driver and gets us the Sergeant address (needs to first call gets against the info tech login page). 
Note the user must be logged into Connect or Sergeant already in order for the program to work. The user will need to 
authenticate using the secure app installed on your phone. Chrome Driver is loaded using the default Chrome user data. Whatever
profile is loaded needs to be logged in already and have credentials saved on their browser.

4. The program will check all 'pretty names' and 'slug names' for drivers on a given survey. These are grouped as key, value pairs
in order to return a match when we go to add questions (click the appropriate Add question toggle).

5. The program will then scan all three pages of the full engagement survey and build a list of those for comparison. We should be able to
skip questions already added to the survey, or skip deletions if a question has already been deleted.

6. Program then checks against Page 2 to ensure that we are ungrouping (selecting delete toggles) the senior management relationships
driver. This element is tricky for a few reasons, in that delete and edit commits wont be saved the first time over while the questions
are still in a group. Reduces chances of error in future blocks.

7. Program goes onto delete questions on page 2 and 3 (EXM question is never removed for PEP, Full Eng or Pulse). This 
works by checking if a question id is present and its adjacent cell is NoneType. Based on this it then toggles the delete in Sergeant.

8. We then add questions to the survey based on the key, value pairs corresponding to our excelrowlistobject's value. If we have a match,
we return the slug's name, pass that through to the XPATH to find the right Add Question button and send the text and click save.

9. Editing questions is last, starting on page 1 and cycling through all on page 2 and 3. Saving on each page, it will then return
us to the last page and will now change languages and repeat this loop until all languages have been completed. Note here:
questions added and deleted in English, will be cause them to be removed from all other languages present.

Rationale:
The amount of time spent by Project Coordinators copying and pasting for multilingual and single language 
surveys is substantial. The primary goal of this program is to effectively eliminate or substantially reduce 
low-effort, redundant work. Secondary goals are to eliminate the incidence of human error in question deletion, 
pasting incorrectly, or other inconsistencies which places effort on the development team to re-add those questions
back into Sergeant.

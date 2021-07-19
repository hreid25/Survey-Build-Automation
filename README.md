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
authenticate using the secure app installed on your phone.

4. Program then checks against Page 2 to ensure that we are ungrouping (selecting delete toggles) the senior management relationships
driver. This element is tricky for a few reasons, in that delete and edit commits wont be saved the first time over.
Removing this was deemed necessary as it was critical if we were to then delete, add and edit questions more easily.

5. Program goes onto delete questions on page 2 and 3 (EXM question is never removed for PEP, Full Eng or Pulse). This 
works by checking if a question id is present and its adjacent cell is NoneType. Then toggles delete in Sergeant.

6. We then begin adding questions, matching the question ids from the array to the driver and running a few checks
to grab the question text from the QIL. We then need to grab the newly created custom question id and reinsert that 
into our array.

7. Editing questions is last, starting on page 1 and cycling through all on page 2 and 3. Saving on each page, it will then return
us to the last page and will now change languages and repeat this loop until all languages have been completed. Note here:
questions added and deleted in English, will be cause them to be removed from all other languages present.

Rationale:
The amount of time spent by Project Coordinators copying and pasting for multilingual and single language 
surveys is substantial. The primary goal of this program is to effectively eliminate or substantially reduce 
low-effort, redundant work. Secondary goals are to eliminate the incidence of human error in question deletion, 
pasting incorrectly, or other inconsistencies which places effort on the development team to re-add those questions
back into the web platform Sergeant.

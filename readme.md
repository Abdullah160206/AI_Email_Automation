
1.Python package -> Anaconda

2.Python Libraries -> 
	1.os -> Help script to Interact Directly with Operating System.
	2.re -> allowing you to search, split, alter, and validate text patterns.
	3.datetime -> Use for date & time related operations.
	4.openpyxl -> Use for write excel file.
	5.win32com -> Allowing you to Interact with Windows component Object Model (COM).

3.Modules
	python 3.13.14
	file ->	email_automation.py

4.VARIABLES:
	1.INTERNAL MAIL DOMAINS -> domains whose emails you want in Internal_emails folder.
	2.ENABLE FOLDER CLEANUP -> set it to true if you want to delete day old folder from DELETE_OLDER_THAN_DAYS your outlook.
	3.DELETE_OLDER_THAN_DAYS -> set days in this for which last day folder you want to delete(15).
	4.NO_OF_DAYS -> no of days excel file you want.
	5.HOD_EMAIL -> define email on which escalation and top up emails send to HOD.
	6.ESCALATION_THRESHOLD -> define the threshold for which after that number consecutive unread send escalation email to HOD & after that number to & fro convo send escalation email to HOD.
	7.FOLLOWUP_DAYS_THRESHOLD -> define no of follow up days threshold after which the system send the follow up email to the customer for the email we send to it regarding to any document or any number like chases, motor, or Mobile number.
	8.IS_CUSTOMER_CARE -> make it true if the system is use by customer care department,    (MAKE It FALSE if system is use by normal user on normal email); When you make it false it don't send HOD_EMAIL, no follow up, and no draft.
	9.DAYS_OFFSET -> if you run system on today set it to 0 and on yesterday set it to 1.
	10.promtoions -> define the array of promotions(which emails you want in promotion folder on outlook).
	11.HOD_POPUP_BUSINESS_DAY -> After read customer email and after that no of days still no reply, Then a pop up msg is send to the HOD, that the xyz email is not entertain
	
5.FOLDER PATH:
    PROJECT_DIRECTORY -> set directory in which you want all excel files.
    FOLDER PATH -> C:\Users\Abdullahh\OneDrive\Desktop\Anaconda_Project\firstproject


6.What the system does
    Reads every email currently in Outlook's Inbox (and previously filed date folders), across every date present — not just today.
    Classifies each complaint by category (safety, mechanical, warranty, service delay, general) and by urgency, based on keyword rules.
    Tracks conversations, not individual emails — a customer's back-and-forth thread is treated as one unit, with reply timing and unanswered streaks computed across the whole history.
    Files each conversation into a date folder matching when it started (or its original folder, if it's a continuing thread), unless it belongs to an internal domain (→ Internal_Mails) or matches a promotion keyword (→ Promotions).
    Drafts replies automatically for common cases — asking for a missing contact number, sending a category-specific first response, or confirming resolution — when IS_CUSTOMER_CARE is enabled.
    Sends follow-up nudges to customers who've gone quiet after we asked for missing information.
    Escalates to the HOD in two ways:
    When a conversation has too many consecutive unanswered messages or too many total exchanges (ESCALATION_THRESHOLD).
    When a customer's message has gone unanswered for HOD_POPUP_BUSINESS_DAYS+ working days.
    Exports Excel reports — one row per conversation, color-coded by priority and status, for both the current run and a rolling multi-day window (NO_OF_DAYS).
    Cleans up old folders (optional) once conversations inside them are old enough (DELETE_OLDER_THAN_DAYS).
    Date folders and required subfolders (Promotions, Internal_Mails) are created automatically under Inbox on first run — no manual setup needed.


7.ARCHITECTURE DIAGRAM:

   Outlook mailbox -------> |Classifier|  |State Tracker|  |Folder manager| --------------> |Excel Reports|  &   |Email actions|    ----------> Stakeholders
(inbox, sent, draft)         (Type)         (escalation)   (Dates & Cleanup)                (Daily & Weekly)  (Draft, nudges, alert)           (customers, Team, HOD)



8.Functions
	1.get_target_date() -> Give excat date w.r.t day_offset today or yesterday.
	2.
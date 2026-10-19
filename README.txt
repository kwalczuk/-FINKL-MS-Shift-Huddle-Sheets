SST & SHIFT HUDDLE SHEETS
=========================

SETUP (once)
  1. Before unzipping: right-click the .zip > Properties > tick "Unblock" > OK.
  2. Unzip the folder anywhere (e.g. J:\Departments\Machine\4M\KACPER\Shift Huddles\HuddleSheets).
  Needs Excel (sign-in sheets). For the topic sheets: Word / PowerPoint for .docx / .pptx topics.

EVERY WEEK
  1. Get the email out of Outlook, any of these work:
       - New Outlook / Outlook on the web: open the email > ... > Save as / Download  (gives an .eml)
       - Classic Outlook: drag the email onto your desktop  (gives a .msg)
       - Or "Save all attachments" / "Download all" (folder or .zip)
  2. Drag that file (or folder / zip / the attachments) onto "Huddle Sheets.bat".
     Or double-click the .bat, press Enter, and pick the email from the Downloads list.
  3. Answer the prompts:
       topics by day  -> week (P / T / N, a week # like 43, or a date like 10/19)
       -> shops (A = all, or 1 / HMS / 1,3)  -> include topic sheets? how many copies?
  4. The PDF opens when it's done. Saved in PDFs\Week 43 - 2026-10-19\ next to the script
     (change with: set OutputFolder J:\some\folder).

  Every attachment is a topic (Word, PowerPoint, PDF, Excel, pictures), in the order of the
  email / zip: 1st -> Mon, 2nd -> Tue ... 6th -> Sat. A 7th+ file is printed after the others
  but isn't on the sheet. Change with 'sat 3', 'swap mon tue', 'skip 7' (leave a file out).

  The 3rd shift (setting NightShift) prints dates one day earlier: Mon column = Sunday's date.
  The topics + day mapping are saved with each week, so "Reuse topics" can reprint later.

EDITING PEOPLE / SHOPS
  Either open HuddleConfig.xlsx and edit the Employees / Shops / Settings sheets
  (keep the header row; close Excel before using the script's edit commands),
  or use the commands in the script window. Type  help  for the full list:

    list [shop] [shift]      list hms 3
    add                      asks shop, shift, last/first name, assignment
    remove [name]            remove torres
    edit [name]              change name / assignment / shop / shift
    find <text>
    shops | add-shop | remove-shop | shifts mms 3 | title hms
    settings | set ExtraBlankRows 4
    weeks                    week numbers around today
    open config | open template | open output
    cancel (at any prompt)   back out

  Row order on the sheet = row order in the Employees sheet. Supervisors added
  with "add" go to the top of their shift.

SHEET LAYOUT
  SignSheetTemplate.xlsx is the look of one sheet - change fonts, widths, the
  Huddle Reminders text, etc. Keep the same cell positions (title A1, shift F1,
  date I1, days C2:H2, topics C3:H3, names from A5, I5 merged down the name rows).

FILES
  HuddleSheets.ps1        the script
  Huddle Sheets.bat       launcher
  HuddleConfig.xlsx       shops, shifts, people, settings
  SignSheetTemplate.xlsx  sheet layout
  lib\PdfSharp.dll        combines the PDFs (open source, MIT)
  HuddleSheets-error.log  created if something fails - send it to whoever fixes the script

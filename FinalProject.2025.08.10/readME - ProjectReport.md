# - Introduction - #

I worked as a program administrative assistant for five years. In one of those positions, I 
manually maintained a program team list in Excel for the 170-ish people on our team. I used it 
for event planning, coordinating catering, determining meeting attendees, creating email mailing 
lists, and so forth. Doing all the data management manually was a pretty significant time-sink.

The premise of this assignment is that, when someone joins the program team, the admin 
provides them with a specific table template to fill out and return to the admin. This template 
would be provided as an excel file, but the application also handles data from Word documents 
and PDF files, in case the team member decides to return the file as a Word or PDF file instead.

For the project’s dataset, we have five files, submitted by team members (three Word 
docs, one Excel file, and one PDF file). The template provided to team members is as follows:

If any of these fields do not apply, please leave them blank. 
First Name: 
Last Name:
Company:
Department:
Supervisor:
Start Date:
End Date:
Email:
Landline:
Mobile:
Food Allergies:

# - Application Description - #

This project isn’t intended for an external user; it’s intended for an admin who would be 
in charge of utilizing and modifying it as needed. The script, at this point, does five main things:

1. It creates an Admin Workbook, to hold all the team information.
2. It reads from three possible file types (Word, Excel, and PDF).
3. Once data has been read, the data is input into the Admin Workbook.
4. There is a function that clears some extraneous text, such as “n/a”, “not applicable”, etc.
5. And finally, there is a process for generating reports in the form of more specific 
worksheets. In this case, there is one worksheet for food allergies, and one for emails.

# - Results - #

With the code as it’s currently written, we end up with an Admin Workbook that has three sheets: 
‘Program Team’ for all the team info, ‘Allergies’ for just people’s food allergies, and ‘Emails’ for 
a list of all emails. The code can be added to and modified as-needed to include new files as the 
admin receives them. I’ve included comments indicating where code can be modified in-future.

# - Conclusion - #

The application overall works well to compile data and provide useful sheet reports. 
Something like this would have been very helpful in my previous employment position. Trying 
to keep that particular Excel file current and accurate was pretty arduous, and there was some 
risk involved, too, in the case of missing’s someone’s food allergies (particularly if they were 
serious), or in the case of sending emails to the wrong recipient(s). Having one standard template 
for everyone to fill out, and then automating that data input instead of the admin typing it all 
manually, could significantly cut down on time, and could help mitigate some of the risk factors.

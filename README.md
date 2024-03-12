# Outlook_Email_Scheduler
Send emails from MS Outlook with times specified

Origin Story

The ideation of this began at the back end of 2015 when I was working at an agency where we were asked to send 400 emails daily from multiple inboxes for client accounts. I’m not sure what you’d understand by that so allow me to mansplain: we had email templates created by the CMO, and data filled into spreadsheets by the data team, where we had validated email addresses and prospect card information. A messaging person’s task was to:

Open a new email object on Outlook
Paste the template
Change the variable fields like - first name, company name, prospect category, company category, etc.

Repeat that for like 400 times.
The Realization

After going through the process and seeing myself and everyone around me get rattled and seeking a better purpose in life, and watching the founders do nothing about solving problems, I realized it was better to nudge my brain into the matter. 

There was a six-month period in 2016, when I quit the job, sat and learned Python, and learned how to code in VBA. Lots of chasing videos on YouTube, reading Stackoverflow, and learning on Codeacademy.

The Solution
I realized that learning how to represent human actions in code would solve the problem. The only thing in the way was an understanding of the Outlook object model. There was a lot of reading the documentation and trial and error.

How to use the Code


Create an Excel file with the following column in Sheet1 (maintain the order of the columns column names may change)

Company Name
State
Time Zone
First Name
Designation
Email Address
Time (date-time at which the email needs to be sent)
Personalization
Template Name
Template Path (Path of the oft template file)

Create Oft Template files

Following is a sample oft template that has seen great results for the sender.

Subject: Can I have your Friday evening?

Hi %FirstName,

%Personalization

This is a cold email. I’d like to help %Company grow by 8000%.

Do you have 2 hours on your calendar for a discussion this Friday evening at 6 PM?

Regards,
Jan Theman
Director Inside Sales

Go to the VBA editor on the excel file and save the code in a new module

Run the module/macro


If you are so inclined, you may also create a macro button for this in the ribbon and run it from there, instead of having to run the macro from the VBA editor every time. 

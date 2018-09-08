# OutlookAutoGreeting
VBA Outlook macro that generates reply-alls, prepended with an automatically-generated greeting. I wrote this to never again misspell someone's name in an email (on the first line at least), and so I'd have to write emails less manually.

Sample output -- "Good morning Jonathan,"

This macro will look at your local time, and generate a time-appropriate salutation (Good morning/Good afternoon/Good evening). It will look at the first word of the sender display name, adjust the case, and append it that the salutation. It will then produce a reply-all.

You can add this macro to your Quick Access toolbar/Ribbon in Outlook, and use this macro in place of the reply button. 

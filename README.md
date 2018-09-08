# OutlookAutoGreeting
VBA Outlook macro that generates replies, prepended with an automatically-generated greeting. I wrote this to never again misspell someone's name in an email (on the first line at least), and so I'd have to write emails less manually.

Sample output -- "Hello Mr. Smith,"

This macro will look at your local time, and generate a time-appropriate salutation (Good morning/Good afternoon/Good evening). It will look at the first word of the sender display name, adjust the case, and append it that the greeting.

You can add this macro to your quick access toolbar in Outlook, and use this macro in place of the reply button. 

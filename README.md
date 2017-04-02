# ![ommplogo](https://cloud.githubusercontent.com/assets/15161168/24591436/4a2e9204-17be-11e7-8ae7-83b682617bd6.png) Outlook Mail Merge Plus

## Synopsis 

A Visual Basic script that does two impossible things. 

1. Add attachments to a mail merge.
2. Change the sender on an Exchange account. 

Mail merge is a useful tool, but it neglects several features that are often necessary when sneding a mail merge. Those features are the ability to add attachments, as well as to allow delegates to send the FROM option. 

## Usage 

1 Put Outlook in OFFLINE mode. Send/Receive Tab "Press Offline Mode".
2 Run through your mail merge completely, all messages should stick in the Outbox. 
3 Run the .VBS Script. 
4 Add Attachments. 
5 Change the Sender if desired. 
6 Check that the items in the Outboox look like what you want to send. 

## Compatibility 

Outlook 2013 and onward (probably backwards compatible all the way until 2003). 

## Error Suppression 

In Outlook 2013, it has been noted that some security settings may prevent the script from running correctly. To fix this try the following: launch Outlook as admin see here: [http://www.msoutlook.info/question/353]
Shift right-click should allow you to run as admin. Then you can go to File > Options> Trust Center > Trust Center Settings > Programmatic Access > Never Warm me about suspicious activity. That same dialog will tell you if anti-virus is valid, and if it is we’re fine there. 

Make sure Windows firewall is turned ON. 

## Contributors 

- Wooter Westerveld, Lead Developer[OMMA]{http://omma.sourceforge.net}
- Hexadecimalz

## Issues 

This code needs cleanup. I've attempted to clean-up most of the big issues such as spelling and grammar, but the code itself is not pretty.  

It would be nice of some of the user-intervention, such as putting the program in Offline mode were scripted, such that it reduces the possibility for error while utilizing. 


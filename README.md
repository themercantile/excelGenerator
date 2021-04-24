# excelGenerator
This JS/HTML generates excel formulas useful for the creation of hyperlinks in Excel. 

Generate useful Excel formulas from user input URLs.

For the purposes of this project, a useful Excel formula is one we can use within the NDIA office environment. We're using tools that are available on the locked down computers. This includes the MS Office suite and 
front-end programming tools that can run in a web browser.

A useful Excel formula can be used within NDIA Jira ticket resolution process (to be finalised). Examples of useful formulas are below.

Create hyperlinks of URLs:

	E.g. https://zoomies.pupper.org/MYRECORD122/displayme.jizz <-- MYRECORD122 is the part of the URL that changes when you want to see another record.
	The Excel formula to load this URL as a hyperlink would be =HYPERLINK("https://zoomies.pupper.org/MYRECOD122/displayme.jizz", "CLICK") <-- "CLICK" is the text displayed in the cell which obfuscates the link.
	
	What if we want to load MYRECORD999 instead (or any other record)? We could use the Excel concatenate function as follows:
	=CONCATENATE("https://zoomies.pupper.org/", A1, "/displayme.jizz") <-- Where A1 is a cell reference to the record we want to load. Note that this formula won't create a useful hyperlink with display text.
	
	To do that, we'd need to combine the two: =HYPERLINK(CONCATENATE("https://zoomies.pupper.org/", A1, "/displayme.jizz"), "CLICK") <-- This like creates the nice hyperlink and concatenates at the same time.

Create batch file commands to open a browser at a certain page:

	Say we had 10 tickets in Jira and we wanted to add the exact same note to all of them. We could manually open each ticket one at a time and add the note OR we could open all the tickets at once and CTRL-TAB through each of the tabs and paste in the text required.
	
	The command to open a web page from the command line we would enter: start https://zoomies.pupper.org/MYRECOD122/displayme.jizz <-- This would open the page using the default web browser in Windows.
	
	If we wanted to use a specific browser like Chrome, the command is: start chrome https://zoomies.pupper.org/MYRECOD122/displayme.jizz. 
	To produce this outcome in Excel, we'd use =CONCATENATE("start chrome ", B1) <-- Where B1 is the cell reference pointing to =CONCATENATE("https://zoomies.pupper.org/", A1, "/displayme.jizz")
	
	It would be really handy to be able to select 20 rows of these commands and export them to an automatically named batch file. Perhaps using the date/time to name it. Will need to check that out.
	
	The format of the batch file will be:
	
		@echo off
		start chrome https://zoomies.pupper.org/MYRECOD123/displayme.jizz
		start chrome https://zoomies.pupper.org/MYRECOD124/displayme.jizz
		start chrome https://zoomies.pupper.org/MYRECOD125/displayme.jizz
		start chrome https://zoomies.pupper.org/MYRECOD127/displayme.jizz
		...
		etc
		
	When run, each of these lines will open the specified record on its own tab. Ready for workifying!

So - Um - What is the JavaScript for?

	We're going to create a one-page application using front-end tools available on a massively locked down system. The workflow will be:
	
	1) User identifies a resource that can be easily changed by modifying the record number. 
		A real world example of this would be google maps. https://www.google.com/maps/place/5+King+St+Maidenhead+United+Kingdom/ will take us to the Kebab Elite take away in England. https://www.google.com/maps/place/17A+Wanneroo+Rd+Joondanna+WA+6060/ will take us to Marco's Pizza in Western Australia. You can see in the URL we're just changing the address.
		
	2) User copies the full URL from the web page and pastes it in a text field in our app. 
		Using an example from above, we might paste in https://www.google.com/maps/place/17A+Wanneroo+Rd+Joondanna+WA+6060/.
	3) User tells us which part of the URL can be manipulated by pasting in only that part to another text input field. This is the RECORD part.
		Using the same example, they would paste in "17A+Wanneroo+Rd+Joondanna+WA+6060".
	4) Program now asks for user input - Which cell needs to be concatenated into our URL? We're assuming the user has an excel file already that has a list of RECORDSs that can be concatenated. 
		Let's say the user wants cell C3 to replace part of the URL. Cell C3 contains a correctly formatted address.
	5) The program replaces the RECORD part of the URL from the full URL to create the base URL we'll use. 
		"https://www.google.com/maps/place/17A+Wanneroo+Rd+Joondanna+WA+6060/" becomes ("https://www.google.com/maps/place/", C3, "/"). NB - Both JS and Python both have a .replace() method.
	5) Now we can use the JS concatenate function to create the Excel formula.
		=HYPERLINK(CONCATENATE("https://www.google.com/maps/place/", C3, "/"), "CLICK")
	6) This Excel formula is displayed in our application and can be copied to the computer clipboard.
	7) Paste into Excel and just drag down to create formulas for each row of the Excel file.

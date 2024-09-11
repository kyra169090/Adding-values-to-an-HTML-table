# Automated Web Input Process - Adding values to an HTML table from an Excel Spreadsheet
You have an Excel file with values and a webpage with an HTML table? And you need to migrate the data from one to the other? No CSV import or anything? How annoying! This script can help.

This AutoHotkey script, combined with JavaScript (for DOM manipulation), automates the process of copying selected cells from Excel and inputting the data into a web page (populating its DOM elements), using G33kDude's Chrome.ahk library. It's designed to work with a specific web page identified by its title.

# How it works
* Install AutoHotkey
* You will need this Chrome.ahk file to be able to automate Google Chrome (many thanks to G33kDude, he saved my life with that!): https://github.com/G33kDude/Chrome.ahk
* You need to edit the shortcut you usually use for opening Google Chrome. Right click and go to „Properties”. Then you need to edit the „Target" section. Don’t delete the path in here, only paste this text after the original path (after a whitespace): „ --profile-directory=Default --remote-debugging-port=9222 --remote-allow-origins=*”
* In the script, ensure you have the correct path to Chrome.ahk included.
* Adjust the web page selectors and event types in the JavaScript code as needed for your specific use case.
* Copy Selected Cells to Arrays: Press Alt+1 to copy the selected cells from the first column in Excel to an array (Array1), and Alt+2 to copy the selected cells from the second column to another array (Array2)
* Press Alt+P to start the process of inputting the copied data into the web page.
* Any other things are commented in the code

# Just a tip
Use this function in Chrome DevTools console, if you would like to know all event listeners attached to a specified DOM element:
```javascript  
  getEventListeners(document.querySelector('your-element-selector'));




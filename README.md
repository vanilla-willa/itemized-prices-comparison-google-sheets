# Compare Price Options on Google Sheets


## Background
Google Sheets is used to organize data collaboratively and privately. Many times, people use Google Sheets or Excel sheets to do financial calculations and evaluations. Usually, basic techniques would be to use built-in formulas, such as `COUNT()` or `SUM()`, while more advanced techniques would involve things like pivot tables, Vlookups, Index Match, conditional formatting. But at some point, the task may be more complex, formulas can become a pain to maintain, or the user may just be doing the same thing repetitively (such as doing different variations of `= SUM(A1, A2, A5...)` multiple times throughout the sheet).

## Purpose
I wanted to a create a multi-purpose solution to simplify the interface of comparing prices and options by generating a new sheet with checkboxes that automatically sums up the prices.  
Using this list of items and the items' corresponding prices as the database, a new sheet with checkboxes is generated. Checkmarking a box will retrieve the price from the database; checkmarking multiple boxes will calculate the total of all the checked boxes. The total price of all boxes checked will be outputted in the last cell in that row or column.
<br /> <br />
There are two main scripts that either generate a script with the categorized items in rows or with the categorized items into columns. The categorized items in rows is generally better for vertically-oriented monitors, while the categorized items in columns is generally better for horizontally-oriented monitors.
<br /> </br>
Use cases I had in mind were applied to collaborative or personal budgeting (ie. planning an itinerary and comparing total costs for different variations) or summing up IOUs to charge friend(s) one lump sum on Venmo.  
Note: This is by no means an elegant solution. Instead, it was a project fueled by personal curiosity of Apps Script (with some level in development + having done VBA and a ton of Excel in previous roles) and being struck by the idea as I watched my partner try to guesstimate expenses for an upcoming trip.

## Features
Note: This was designed to be published as a private or unlisted add-on. But after development, I found that because Google has moved the process of publishing add-ons from the Chrome Web Store to G Suite Marketplace SDK, publishing a private or unlisted add-on (for anyone in any domain to use) is pretty much impossible. To publish, there are additional things that need to be submitted, such as a Terms of Service, to go through their approval process.
<br />
- The user can only have one list source to generate sheet.
- List sources have to follow the format of:
  - Category in a colored cell
  - Listed under the category with the colored cell:
    - Item is in the first column
    - Pricing formatted in $XXX.XX (using the built-in Number format tool) is in the second column
- When the user loads the spreadsheet, a shortcut menu named 'Surprised Pikachu' shows up in the menu bar to access the scripts.
- When the user uses the menu dropdown, they can run the main scripts directly or open a Pokemon-themed sidebar panel to either run the main scripts or embed a shortcut to their source data for quick access.
- If the user embeds a shortcut, 11 rows of colored cells will be added to the top of the sheet, and shortcut images with scripts will be added to those rows.
- The user can only use scripts on the currently active sheet.
- After running a script, user will receive a sheet automatically named the current day and time (in PST) in the format of yyyy-MM-dd HH:mm:ss (ie. 2021-01-01 00:00:00).
- The user can rename their sheets without breaking the scripts.
- The user will be able to expand and collapse categories in these script-generated sheets to view the items and use the checkboxes.
- The user can check multiple boxes to get the total of the items checked, without needing to look at the original list.

Some basic error handling:
- When the user tries to run a script on a sheet formatted incorrectly, user will get a error popup.
- To prevent accidental clicks on the images, any time the images are clicked, user will get a popup to confirm whether they meant to click on the script.

Future ToDos:
- [ ] The user can have multiple list sources to generate sheet.

## Screenshots & Gifs

![Image of Add Dialog](https://github.com/vanilla-willa/itemized-prices-comparison-google-sheets/blob/main/assets/dialog-clicktoadd.png)
<br />
![Image of Remove Dialog](https://github.com/vanilla-willa/itemized-prices-comparison-google-sheets/blob/main/assets/dialog-clicktoremove.png)
<br />
<img src="https://github.com/vanilla-willa/itemized-prices-comparison-google-sheets/blob/main/assets/expandingrowspika-510x450.png" width="227" height="200">
<br />
<img src="https://github.com/vanilla-willa/itemized-prices-comparison-google-sheets/blob/main/assets/expandingcolumnspika-600x450.png" width="266" height="200">

## Live Demo
https://docs.google.com/spreadsheets/d/1T2nYpsmoaYgHwHFjTZtoi5HJ944XGn5mbtnZspb2YxA/edit?usp=sharing

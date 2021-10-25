# Course Reserves Auto-fill Template
Google Sheets and Google Apps Script documents to automate course reserves inventory


## Format
The Google Sheet contains columns for:
  * **Bibliographic Information**: Title, Author, Edition, ISBN
  * **Course Information**: Course, Section, Instructor
  * **Inventory**: Print Copy, E-Access, Notes
  * **Liasons**: Purchase Print?, Purchase Ebook?, Notes
  * **Acquisitions**: Purchase Date	Received Date	Ebook Activated
  * **Reserves Processing**: Added to Reading List
  * **Query**: Query made?

The associated Google Script contains code that auto-fills relevant fields in the sheet, given an ISBN.

## Function

* Upon opening, the script will check the 'ISBN' and 'Query Made?' columns for each row of the spreadsheet.
* In a given row, if 'Query Made?' is empty, and 'ISBN' is filled, it will query WorldCat for alternate ISBNs and metadata, and then query Primo to check our inventory for each ISBN.
* If 'Title' or 'Author' are empty, they will be filled using metadata from WorldCat
* 'Print Copy' and 'E-Access' will be filled to reflect our inventory
* 'Notes' will display the status of each ISBN queried in our inventory (via Primo)
* 'Query Made?' will be changed to 'done' when the row is finished


## Usage

* **Triggers have to be enabled manually**
  * With the spreadsheet open, navigate to the *Tools* tab and click on *<> Script editor*
  * In the script editor, click on the *Triggers* button on the left side (the clock icon)
  * In triggers, add a new trigger
    * select triggerOnOpen from the menu, and for Event select *On open*
  * Repeat for triggerOn Edit, for Event select *On Edit*

* The only *necessary* data for the spreadsheet to function are ISBNs.
* Copy/paste ISBN, and any available metadata into the appropriate columns.
* Reopen/refresh the sheet, and it will begin to auto-fill.
* Each query takes a second or two; as long as the spreadsheet remains open, it will continue until finished.

## Issues

* Changing column order will cause the script to malfunction, as it indexes by the column letter/number position.
* Columns added after the 'Query' column on the far right will not affect the functioning of the script, so any new columns should go there
* Presently, the script generally doesn't find alternate editions. We may find a solution with WorldCat soon.

## Setup

In Google Drive, simply copy the template, and fill in available data.


To report any issues or for general help, contact Ethan Davis (edavis@lclark.edu, Digital Initiatives)




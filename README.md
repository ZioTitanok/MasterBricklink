# MasterBricklink
<b><i>MasterBricklink, AFOL Tools for Bricklink on Google Spreadsheet</b></i>, was conceived by <b>GianCann</b> (between 2018 and 2019) to simplify the management of LEGO brick inventories. After inheriting the project, I improved existing features and introduced new ones, aiming to make the spreadsheet as standalone and user-friendly as possible.<br></br>

## Installation
The "installation" process, or rather the "generation" of the spreadsheet, requires completing a few steps:

1. Create a Google Sheets workbook.
2. Import (by copying) various scripts using the Script Editor/App Script.
3. After reloading the sheet, in the Bricklink Tool menu, choose "Regenerate" and then "Regenerate Settings".
4. In the Settings sheet, insert TurboBricksManager API key, needed for generating the databases.
5. Optionally, in Settings, insert Bricklink API keys (recommended for utilizing the full potential of the spreadsheet).
6. Proceed with generating the other sheets, moving from top to bottom in the Bricklink Tool menu.

## Introduction to Features
Here's a brief overview of the purpose and functionalities of the different sheets:

* <i>Settings</i>: Hosts parameters needed by the spreadsheet and some configurations for Lab.
* <i>DBs</i>: Databases needed by the spreadsheet.
* <i>Inventory</i>: Downloads inventory from your Bricklink account based on the filters set in the sheet.
* <i>PartOut</i>: Downloads the part-out of the selected set based on the filters set in the sheet.
* <i>XML</i>: Generates XML files from Lab data for manual export to Bricklink.
* <i>Lab</i>: Designed for element analysis, data can be automatically imported from Inventory and/or PartOut based on sheet filters, or manually input.

## To-Do List
As with any respectable project, there are still many ideas to be realized:

* Error handling and user experience improvements.
* Multi-part-out for sets and related filters.
* Multi-platform support: synchronization with Brickowl, Rebrickable.
* Introduction of automatic export to Bricklink without the need for XML.
* Write explanations for various functionalities of the spreadsheet.

## Changelog
From 1.0.0 to present (GitHub)
v1.2.3: General code improvement and Lab enhancement.<br>
v1.2.2: PartOut improvement.<br>
v1.2.1: Colors and Category databases no longer depend on Bricklink APIs.<br>
v1.2.0: Introduced automatic update of Parts, Minifigures, and Sets Databases through TurboBricksManager.<br>
v1.1.3: Introduced Lab function for price suggestions and updates.<br>
v1.1.2: Minor Lab update.<br>
v1.1.1: Improved import performance. Minor PartOut update.<br>
v1.1.0: Enhanced performance in XML generation and Price Guide download.<br>
v1.0.1: ReadMe and minor fixes.<br>
v1.0.0: Introduced creation/restoration functionalities for the spreadsheet sheets.<br>

From 0.0.1 to 0.9.0 (pre-GitHub)
v0.9.0: Automatic update of Categories and Colors databases.<br>
v0.8.0: Introduced filters in Inventory, PartOut, and Lab.<br>
v0.7.0: Introduced Settings and other minor user experience functions.<br>
v0.6.0: Beyond Parts: Management of Minifigures, Sets, and more.<br>
v0.5.0: Introduced Import functionalities between Inventory, PartOut, and Lab.<br>
v0.4.0: Introduced PartOut for downloading Sets' part-outs.<br>
v0.3.1: Introduced XML Export (Upload/Upgrade) for manual inventory synchronization with Bricklink.<br>
v0.3.0: Introduced XML Export (Wanted) for manual creation of WantedLists on Bricklink (GianCann).<br>
v0.2.0: Lab improvement and introduction of Parts, Minifigures, and Sets Databases.<br>
v0.1.0: Introduced OAuth1 in scripts (eliminated external PHP). (GianCann)<br>
v0.0.2: Introduced Inventory with inventory download from Bricklink (via external PHP). (GianCann)<br>
v0.0.1: Introduced Lab with Price Guide download for Parts (via external PHP). (GianCann)<br>

## Dedication
To the memory of <b>GianCann</b>, who undoubtedly would have done a better job than me in advancing this project.
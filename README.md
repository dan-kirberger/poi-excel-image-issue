# poi-excel-image-issue
Run this program with `./gradlew run` (or however you like to run apps in your IDE)
 
It will generate two files (generated with different options).

These options each demonstrate an issue with embedding images in a worksheet

#####example-MOVE_AND_RESIZE.xlsx
Open this file in Excel. 

Sort the worksheet using the "Name" or "Type" column filter options.

Observe that the thumbnails do not get sorted along with the rest of the row data, they remain fixed at their previous row numbers.


#####example-MOVE_DONT_RESIZE.xlsx
Open this file in Excel. 

Filter the "Type" column to just show "Text only"

Observe that the "One" and "Three" rows disappear but their images remain, hovering over the document
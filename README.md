# wordrubric
VBA script to create, highlight and save individual feedback forms with word/excel

## How it works
This works by taking a set of results in a formatted excel spreadsheet and merging them with a pre-defined template in word. The VBA hijacks the mail to individual documents process to highlight cells in a rubric table and then save the results to individual PDF's.

The use case for this is quite specific but the rubric template could easily be modified. 

There are two folders here Rubric Win and Rubric Mac. There's no difference in the VBA or the files. The only difference is the user guides. 

## How to use it.
* Make a duplicate of the Rubric folder, including the sub folder. All the files should be in the same folder and there should be a folder called Feedback
* Put all of your marks in the spreadsheet.
  * Rubric columns should follow the naming format `Rubric n`
* Open the word document. 
* Change the Rubric as required. The number of rubric criteria columns should match the columns in the spreadsheet.
* Add the spreadsheet as a recipient list.
* Do a Mail Merge > Edit individual documents. 
* The individual PDF feedback forms should be in the Feedback folder. 

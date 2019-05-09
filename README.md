# HelloQUEST

## TO DO:
* pack these scripts into executable using 
  * pyinstaller
  * tkinter

## 5-10-2019 Updates
* everyting was linked correctly, now need to create user interface portion
* ideas for interface:
  * check boxes for which sheets to be exported
  * fields for month, timezone, average wage, file names

## 5-9-2019 Updates
* sheets are all linked, but need to check accuracy of numbers 

## 5-4-2019 Updates
### Notes on datalinkage/datalink.py
* added a month string field to differentiate separate month file names
* changed from inner to full outer join on humanity and bob
* standardized the order of colummns
* creates a general joined file, nosale excel file, paid events file, and free events file
* TO DO: 
  * create aggregate sale file by event id, then merge with SEL
  * clean the code into methods
  * create an executable / GUI
  * question: should I create a file that aggregates by employee id?
    * this would be more about employee performance than event performance.

## 4-29-2019 Updates
### Notes on datalinkage/datalink.py
* currently able to [inner] join on humanity and bob
  * joins on employee id, office name, and date from both sheets
  * in joined.xlsx in output folder
* takes from input folder --> humanity_eventid.xlsx and mastersales.xlsx
* humanity_eventid.xlsx is generated from datalinkage/eventid_parse.py script
* need to create a script that generates a 1-sheet "master sheet" from raw bob data

### Notes on datalinkage/eventid_parse.py
* takes raw humanity data in excel form from the input folder
* parses the event id from shift_title into a separate event_id column
* essentially prepares the humanity script for datalink.py script

### Notes on eventIDGenerator.py:
* Currently asks for the row number in the excel sheet and the abbreviation of the office code. Still need to look into the autmatic generation of the office code, and we'll be set.
* The last four digits are just randomly generated. Will deal with pulling the pin from existing data once I figure out a source.

## Getting Started (Windows)
### Download Github Desktop
* https://desktop.github.com/

## Getting Started (MAC OSX)
### Here is a list of commands to run in **Terminal** before running this code:
* xcode-select --install
* /usr/bin/ruby -e "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/master/install)"
* brew install python3
* pip install xlrd

### Now, clone the repository onto your computer. Type these commands into Terminal:
* git --version
* cd
* cd Desktop
* mkdir HelloFresh
* git clone https://github.com/laurayuzhengUMD/HelloQUEST.git

Check on your Desktop to see that the files have been added. 

## To run the script, type in Terminal:
* cd 
* cd Desktop
* cd HelloFresh
* python eventIDGenerator.py

## If updates are made and you would like to see them locally, type in Terminal:
* cd
* cd Desktop
* cd HelloFresh
* git pull

## Resources
* [Reading from an excel sheet](https://www.geeksforgeeks.org/reading-excel-file-using-python/)


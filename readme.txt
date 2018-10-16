This tool is built to work with the provided "nametagTemplate.png". Follow these instructions exactly, and it will generate the PSR nametags.

TA pictures need to be named as "FirstLast.jpg".

The Excel sheet provided by the department is messy. Fix it by doing the following:

#1. Removing non-PSR TAs:
	
		1. Delete all reader names (beige highlight)
		
		2. Delete all CCS names (yellow highlight)
		
		3. Delete all PSR fellow names (at bottom)
		
#2. Fix Course Numbers:
	
		The default organization only labels Course Number on the first TA (sometimes first 2+) for a particular course.
		Additional TAs for the same course will be placed beneath the labeled TA, and will have their Course Numbers blank.
		
		1. For each labeled TA with unlabled TAs below them, copy their course assignment downward into the blank spots immediately below them.
		
#3. Generate info.xlsx (holds TA data for tag generation):

	1. Create new excel document, info.xlsx, in the folder.
	
	2. Copy column A (TA names) and column J (Course Numbers), paste to columns C and D of new sheet.
	
	3. Resolve Multi-class Assignments:
	
		* The tool cannot handle duplicate entries -- where a TA is serving for multiple classes.
		
		* Alphabetize by name to identify duplicates. If a TA is serving two courses, combine entries to one row with Course Number as "PHYS #, #"
	
	4. Delete row 1.
	
	5. Reformat Course Numbers: After each step, fill the column with the formula by selecting the top entry and dragging the black square downward over the active rows.
		
		* Go to cell B1. Paste in formula as: ="Phys "&D1
		
	6. Reformat Names:
		
		* Go to cell E1. Paste in formula as: =MID(C1&" "&C1,FIND(" ",C1)+1,LEN(C1))
		
		* Go to cell A1. Past in formula as: =IF(RIGHT(E1,1)=",",LEFT(E1,LEN(E1)-1),E1)

	7. The info.xlsx file should now be formatted as:

		* First column: TA Names
			Format: "First Last"
		
		* Second column: Course Number
			Format: "Phys #"
			
			(If serving multiple courses, Format: "Phys #, #")

#To use the tool:

	1. Place the BadgeMake.py, nametagTemplate.png, and the info.xlsx files into a folder.

	2. Make the subfolders "Output" and "Pictures".

	3. Place all the TA pictures in the Pictures folder.

	4. Run BadgeMake.py using run.bat. (will take ~10 seconds)

	5. Completed badges will save to the Output folder.
	
		*Names which cause errors (largely, through missing pictures) will be written to errors.txt in the Output folder.

#To print multiple badges per page:

	1. Navigate in Explorer to the Output folder.
	
	2. Select all pictures (Ctrl + A)
	
	3. Right click and choose "Print"
	
	4. From print window, select appropriate printing option along right edge
'''
AdvisorTrac Report Extractor
at-report-extract.py

Author: Jake Sherwood (MothsInTheMachine on GitHub)
Last Modified: Jan. 9, 2024

Purpose:
The purpose of this program is to take report text data from AdvisorTrac and 
convert it into a CSV file for use with Excel.

Reason:
Values from these reports are delivered by either HTML (screen) or TEXT (.4sp). 
However, to use these values in Excel requires a trained data-entry eye and 
deep attention, making it hard to import the data. It requires extreme Excel 
know-how or manual entry.

I wrote this code to help me with this and make data reporting from AdvisorTrac 
easier. It's written for a specific department in a specific institution, so it 
may not work for all.

Requires:
	- Python3 installed

Process:
1.	Get the data from AdvisorTrac delivered in TEXT (.4sp) form and save it to 
	the computer. Remember the location of this file.
2.	Open a command line interface. Suggestions are:
	- Command Prompt (Windows)
	- Powershell (Windows)
	- Terminal (Windows or *nix, Mac included)
3.	Use python to run the script.

	Syntax:

	> python at-report-extract.py <source (required)> <destination (optional)>

4.	Profit.

Revisions:

01/10/24 - Renamed directories as A and B (as in 'how to get from A to B'). 
	Outputs by default to B directory as 'output.csv'.

01/09/24 - Initial script and commit to GitHub private repo
'''

import sys
import os, os.path

print()
print('/*******************************************/')
print('/* AdvisorTrac Report Extractor in process */')
print('/*******************************************/')
print()

def main():
	
	# Define departments (labs) to report on
	labList = ['SV Math Lab','SV Writing Lab','DC Math Lab','DC Writing Lab']
	# Define reasons that the department uses for visits
	mathList = ['Chemistry','CIS/Programming','Math','Physics','Other Course Tutoring']
	writingList = ['Languages','MS Office','Reading','Social Sciences','Workshop','Writing']
	# Extend the array of keywords with previous lists and additional keywords to search for
	keywords = []
	keywords.extend(labList)
	keywords.extend(mathList)
	keywords.extend(writingList)
	keywords.extend(['Other','Grand Total'])

	# Check if input file and path are existent
	infname = 'A\\reportOut.4sp' # Default path and name
	try:
		inputPath = sys.argv[1]
	except:
		quitProcess('you did not define an input file path')
		exit()
	
	# Default output file name and path
	outfname = 'B\\output.csv' # Formatted for Windows directories only
	try:
		outputPath = sys.argv[2]
		if (len(outputPath.split('.')) > 1):
			os.makedirs(os.path.dirname(outputPath), exists_ok=True)
			outfname = outputPath[0] + ".csv"
		else:
			outfname = outputPath + ".csv"
	except:
		print(f'Output file will be named {outfname} and stored in same folder as this script.')

	if (os.path.exists(infname)):
		f = open(infname,'r')
	else:
		quitProcess('the file path does not exist')
		exit()
	
	if (os.path.exists(outfname)):
		choice = input('The output file name already exists. Do you want to overwrite it? [y/n]: ')
		choice = choice.lower()
		choice = choice[0]
		while (choice != 'n') and (choice != 'y'):
			print("Invalid choice. Please enter Y for 'Yes' and N for 'No'...")
			choice = input('The output file name already exists. Do you want to overwrite it? [y/n]: ')
			choice = choice.lower()
		if (choice == 'n'):
			print("Cancelling process. Please run command again later.")
			exit()
		if (choice == 'y'):
			print(f'Confirmed overwriting previous file {outfname}.')
			# Remove the old file
			os.remove(outfname)
	
	nf = open(outfname,'w')

	# Reading and writing files here
	nf.write('AdvisorTrac Lab and Subject Report,,,\n')

	for x in f:
		found = x.find('from')
		if (found != -1):
			# Write the dates of the report and skip to next line
			nf.write('Dates ' + x[found:found+29] + ',Visits,Hours,Students\n\n')
			continue

		prevList = [] # This is used for detecting duplicates

		for word in keywords:
			found = x.find(word)
			if (found != -1):
				y = x[found:-1].split('\t')
				# Remove null values
				while '' in y:
					y.remove('')
				# Skip header line
				if ('Visits' in y):
					continue
				# R-strip the ':' from the first elements
				y[0] = y[0].rstrip(':')
				# Remove 'Students'
				if ('Students' in y):
					y.remove('Students')
				# Skip lines less than 2 elements in length
				if (len(y) < 2):
					continue
				# Skip lines where 'Math Lab' is the first element (Duplicates)
				if (y[0] == 'Math Lab') or (y[0] == 'Math Lab Total') or (y[0] == 'Writing Lab') or (y[0] == 'Writing Lab Total'):
					continue
				# Remove HTML from Grand Total line
				if (y[0] == 'Grand Total'):
					y[1] = (y[1].replace('<b>','')).replace('</b>','')
					y[3] = (y[3].replace('<b>','')).replace('</b>','')
				# Compare array to previous array and skip the duplicate
				if (prevList != []) and (prevList == y):
					continue
				prevList = y
				
				z = ','.join(y) + '\n'
				# Add a newline character to separate lab data
				for lab in labList:
					w = y[0].find(lab + " Total")
					if (w != -1):
						z = z + '\n'
				# Strip newline from very last line
				if (y[0] == 'Grand Total'):
					z = z.rstrip()
				# Write to file
				nf.write(z)
	# Process completed
	f.close()
	nf.close()
	print(f'... Done. Outputted file is {outfname}.')

def quitProcess(reason='unknown'):
	print(f'Sorry, I could not understand that command. The reason is {reason}.')
	print('Please enter the path and file name to use extractor.')
	print('Syntax:\n\n\t> python at-report-extract.py <PATH TO 4sp FILE> [OPTIONAL OUTPUT PATH]\n')
	print('/***** Quitting process. *****/\n')
				
# Run the script
main()
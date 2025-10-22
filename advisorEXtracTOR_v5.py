"""
advisorEXtracTOR -- The AdvisorTrac Report Extractor 
Version 5
Author: Jake Sherwood (MothsInTheMachine on GitHub)
Last Modified: 22 October 2025
"""
import sys
import os, os.path
from datetime import datetime
ERROR_FLAG = "[Error] "

def main():
	centers_filename = 'centers.txt'
	reasons_filename = 'reasons.txt'
	output_filename  = 'output.csv'
	input_foldername = 'input'

	print_header()

	# Read in centers and reasons names from external files.
	try:
		centers = read_centers_from_file(centers_filename)
		print("Centers:")
		for c in centers:
			print('\t' + c)
	except:
		print(f"{ERROR_FLAG}There was an issue reading {centers_filename}. Quitting program.")
		return
	
	try:
		reasons = read_reasons_from_file(reasons_filename)
		print("Reasons:")
		for r in reasons:
			print('\t' + r)
	except:
		print(f"{ERROR_FLAG}There was an issue reading {reasons_filename}. Quitting program.")
		return
	
	# Verify overwriting of existing output file.
	if os.path.exists(output_filename):
		choice = ""
		while (choice != 'n') and (choice != 'y'):
			print(f"The output file {output_filename} already exists.")
			choice = input("Do you want to overwrite it? [y/n]: ")
			choice = choice[0].lower()
			# Gets to this line if invalid choice is made
			if (choice != 'n' and choice != 'y'):
				print("\nInvalid choice. Please enter Y for 'Yes' and N for 'No'.")
		
		if (choice == 'n'):
			print(f"\nThe previous {output_filename} will not be overwritten. Quitting program.")
			return
		
		if (choice == 'y'):
			print(f"\nConfirmed overwriting of {output_filename}.")
	
	# Verify there is an input file folder in the directory.
	if not os.path.exists(input_foldername):
		print(f"{ERROR_FLAG}There is no input folder in {os.getcwd()}. Please create a folder named {input_foldername} in the directory. Quitting program.")
		return

	# Prepare the output file for writing.
	dt = datetime.now().strftime('%m/%d/%Y %I:%M:%S %p')
	out_file = open(output_filename, 'w')
	out_file.write(f"AdvisorEXtracTOR Report Output \ngenerated at {dt}\n\n")
	out_file.write(','.join(['',"Visits","Hours","Students"]) + "\n\n")

	# Read data from input directory.
	success_count, failure_count = extract_data(out_file, centers, reasons, input_foldername)

	# Close files and quit the program.
	out_file.close()
	print(f"The program has ended with {success_count} successes and {failure_count} failures.")
	print_footer()

	
def extract_data(out_file, centers, reasons, input_foldername):

	# Develop keywords to search for.
	keywords = []
	for c in centers:
		keywords.append(c + " Total:")
	for r in reasons:
		keywords.append(r + " Total:")
	keywords.append(" Total:") # For the strange 'no reason' numbers.
	# This was an error at times where students could chose the blank
	# space below the list of Reasons.
	keywords.append("Grand Total:") # Look for the Grand Total as well.

	success_count = 0
	failure_count = 0

	for filename in os.listdir(input_foldername):
		# Fix this to work in Windows and Linux
		filepath = input_foldername + "\\" + filename # TODO
		print(f"Extracting data from {filepath}...")

		try:
			in_file = open(filepath, 'r')
			success_count += 1
		except:
			print(f"{ERROR_FLAG}Cannot extract data from {filepath}. Moving to next file.")
			failure_count += 1
			continue

		# Append totals to a new list to prepare to write to file.
		totals = []

		for line in in_file:

			# Record the date of the AdvisorTrac report
			# by splicing it directly from the report.
			found = line.find("from")
			if (found != -1):
				dateline = ("Dates " + line[found:found+29]).split()
				out_file.write(' '.join(dateline) + '\n\n')
				continue # Nothing else worthwhile on this line
			
			# Skip line that appears at beginning of report.
			if (line.find("<!--# vis at beginning of report:") != -1):
				continue

			# Skip lines that are newlines only.
			if (line == '\n'):
				continue

			# Skip lines that begin with series of &nbsp;
			if (line.find("&nbsp;&nbsp;&nbsp;&nbsp;") != -1):
				continue
			
			# Additional criteria for skipping lines
			if (line.find("<!--step3") != -1): continue

			if (line.find("<!--studline") != -1): continue
			
			if (line.find("-->\n") != -1): continue

			if (line.find("#name##addlStud#") != -1): continue
			
			if (line.find("-->Name") != -1): continue

			# Split the line by tab characters if the string
			# is not empty nor equal to just "Students".
			a = list(filter(lambda str: str != "" and str != "Students\n", line.split('\t')))

			# Some of the lines with this tag includes totals.
			# Pop this tag from the array for later code.
			if (a[0] == "<!--Group1 Dif-->"):
				a.pop(0)
			
			# Only look for 1st item to be in the keywords.
			if (a[0] not in keywords):
				continue
			
			# Clean html tags from numbers in Grand Total.
			if (a[0] == "Grand Total:"):
				a[1] = a[1].replace("<b>", '').replace("</b>", '')
				a[3] = a[3].replace("<b>", '').replace("</b>", '')
			
			totals.append(a)

		# Close the file when done.
		in_file.close()
		print(f"Extraction from {filename} complete.")

		for t in totals:
			if (t[0][:-7] in centers):
				out_file.write("\n" + ','.join(t) + "\n\n")
			elif (t[0] == "Grand Total:"):
				out_file.write(','.join(t) + "\n\n")
			else:
				out_file.write(','.join(t) + "\n")
	
	return success_count,failure_count


def read_centers_from_file(filename):
	centers = []
	with open(filename, 'r') as file:
		for line in file:
			centers.append(line.rstrip())
	file.close()
	return centers


def read_reasons_from_file(filename):
	reasons = []
	with open(filename, 'r') as file:
		for line in file:
			reasons.append(line.rstrip())
	file.close()
	return reasons


def print_header():
	print()
	print('/*******************************************/')
	print('/* AdvisorTrac Report Extractor in process */')
	print('/*******************************************/')
	print()


def print_footer():
	print()
	print('/*******************************************/')
	print('/*  Qutting AdvisorTrac Report Extractor.  */')
	print('/*******************************************/')
	print()


# Invoke main()
main()
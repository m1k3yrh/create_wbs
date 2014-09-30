#!/bin/python

#This script has been crafted to do the following:
#We need to read in an csv file and create an object for each row
#We need to clean up Rank and Parent fields.
#We need to sort on Planned For, Priority, Rank, ID to get Ranked List.
#Then we need to recursively take and consume the top item and search the list for items which have it as parent_column. And output to xlsx.

#This requires python 2.7.8 
#This also requires xlswriter. To install, see http://xlsxwriter.readthedocs.org/getting_started.html#getting-started

import csv     
import sys  
import os    
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from operator import itemgetter
import re	
import locale

# Load configuration from local config.py file
sys.path.append(os.getcwd())
from config import configured_data

# read the input parameters
try:
	rtc_export_file_name = sys.argv[1]	
	output_excel_file_name = sys.argv[2]
except:
	print("Usage: python create_wbs.py <input_csv_file> <output_xlsx_file>")
	print("Where <input_csv_file is export from RTC query")
	print("      <output_csv_file> will contain report from script")
	sys.exit(0)

workbook = xlsxwriter.Workbook(output_excel_file_name)	#creates output excel file 
grouped_worksheet = workbook.add_worksheet('Work Breakdown')	#creates a worksheet for grouping
grouped_worksheet.outline_settings(True, False, False, True)	# displays the grouping summary above
ranked_worksheet = workbook.add_worksheet('Ranked')	#creates a worksheet for ranking
input_worksheet = workbook.add_worksheet('Input')	#creates a worksheet for regurgitating input
error_worksheet = workbook.add_worksheet('Errors') # workbook for all suspicious data found from input
error_worksheet.set_tab_color('green')

# format variables
percent_complete_format = workbook.add_format()
percent_complete_format.set_num_format(0x09) # Predefined Excel format for %age with no decimals
hyperlink_format=workbook.add_format()
hyperlink_format.set_font_color('blue')
hyperlink_format.set_underline()
formats={} # Create empty dictionary
filters={} # Create empty dictionary
hidden_columns=[] # Create empty list

#hyperlink
hyperlink_prefix=None

x_grouped_sheet = 0 # row index for grouped worksheet
x_ranked_sheet = 0	# row index for ranked worksheet
x_raw_sheet = 0	# row index for input worksheet
x_error_sheet =0
type_column = 0
percent_complete_column=0
parent_list={}


def load_cell_formats():
	global formats
	configs=configured_data['Format']
	for config_key in configs.keys():
		config_value=configs[config_key]
		try:
			formats[config_key]=workbook.add_format(config_value)
		except:
			print("json format error:  Couldn't parse:",config_value)
			sys.exit(-1)
	
def set_cell_format(cell_value):
	try:
		return formats[cell_value]
	except:
		return None
	
# Test the Filters to see if any of the columns match
def check_filters(arow):
	for key,filter in filters.items():
		y=headers[key]
		if arow[y] in filter:
			return True
	return False
		
# function to write a row to the output file
# function returns the index for the next available row
def print_a_row(ax,arow,aworksheet,depth=0): 
	for i,y in enumerate(arow):
		if i==id_column and hyperlink_prefix!=None:
			aworksheet.write_url(ax,i,hyperlink_prefix + y ,None, y )
		elif i==percent_complete_column :
			aworksheet.write(ax,i,y,percent_complete_format)
		else:
			cell_format=set_cell_format(y)
			aworksheet.write(ax,i,y,cell_format)	
			
	aworksheet.set_row(ax, None, None, {'level': depth})		# sets the grouping level for this row
	ax +=1
	return ax

# function to write a row to the output file
# function returns the index for the next available row
def print_header(ax,arow,aworksheet,hidden=False): 
	for y,col in enumerate(arow):	
		aworksheet.write(ax,y,col,header_format)
		if col in hidden_columns:
			aworksheet.set_column(y,y,None,None,{'hidden':True})
	ax +=1
	return ax

#Creates a dictionary of lists of rows with a specific parent.
def create_parent_list_dictionary():
	global parent_list
	global id_dictionary
	
	for row in data:
		if not check_filters(row): # Ignore everything which filters say to ignore
			parent=row[parent_column]
			if parent in parent_list:
				parent_list[parent].append(row) # add to list against key parent
			else:
				parent_list[parent]=[row] # Create a new list against key parent


# function to recursively group children in work break down structure and write to output file
def search_children(acurrent,depth): 
	global x_grouped_sheet
	global data
	
	mypts = 0
	my_completed_points=0
	childpts = 0
	completed_child_points=0
	
	if acurrent==None:
		current_id=""		# initial recursion. find roots.
	else:
		myx = x_grouped_sheet
		x_grouped_sheet+=1 # reserve a row for me
		current_id = acurrent[id_column]
		mypts = acurrent[storypts_column]
		status=acurrent[status_column]
		try:
			my_completed_points=progress_table[status]*mypts
		except:
			my_completed_points=0

	try:
		children=parent_list[current_id]
	except:
		pass # no children so move on
	else:
		for item in children:
#			if not check_filters(item): # Not needed as we filter the parent_list already
				temp = search_children(item,depth if acurrent==None else depth+1)	# recursively add the story point for the child
				childpts+=temp['child_pts']
				completed_child_points+=temp['completed_child_points']
	
	if acurrent!=None:
		acurrent.append(my_completed_points)
		acurrent.append(mypts+childpts)
		acurrent.append(my_completed_points+completed_child_points)
		try:
			percent_complete=(my_completed_points+completed_child_points)/(mypts+childpts)
		except ZeroDivisionError:
			percent_complete=""
		acurrent.append(percent_complete)
		print_a_row(myx,acurrent,grouped_worksheet,depth)				
	return {'child_pts':mypts+childpts,'completed_child_points':my_completed_points+completed_child_points}	# base call to return story point for current

# Check that all items in data have parents in data as well
def missing_parents_report():
	global x_error_sheet
	
	error=False
	for id in parent_list:
		if not id in id_dictionary:
			if not error:
				error_worksheet.set_tab_color('red')
				error_worksheet.write(x_error_sheet,0,"Fatal: The follow parents are missing from the input file.  Data will be missing from WBS")
				x_error_sheet+=1
				error=True
			error_worksheet.write_url(x_error_sheet,0,hyperlink_prefix + id ,None, id)			

	return error

######################## Main starts here #####################
#read the configuration
load_cell_formats()
header_format=formats['Header']
planned_for_order={value:key for key,value in enumerate(configured_data['Planned For'])}
priority_order={value:key for key,value in enumerate(configured_data['Priority'])}
try:
	progress_table=configured_data['Progress']
except:
	progress_table=None
try:
	filters=configured_data['Filters']
except:
	filters=None
try:
	hidden_columns=configured_data['Hidden']
except:
	hidden_columns=None
try:
	hyperlink_prefix=configured_data['Hyperlink']
except KeyError:
	hyperlink_prefix=None

try:
#	log_file = open('debug.txt', 'w')	#open debug output file to dump trace
	print("Opening File...")
	input_file = open(rtc_export_file_name, encoding='utf-16') # opens the csv file
except IOError:
	print ('Error. Cannot open', rtc_export_file_name)
	sys.exit(0)

print("Reading File...")
reader = csv.reader(input_file,delimiter='\t')  # creates the reader 

input = list()
for r in reader:	# creates a list of row items
	input.append(r)

# Split input into header and data
header = input[0]	#extract the header from the top row
data = input[1:]	#extract all the work items

# Add columns for calculated values to header
header.append("Earned Story Points")
earned_story_points_column=len(header)-1
header.append("Accumulated Story Points")
accumulated_story_points_column=len(header)-1
header.append("Accumulated Earnt Points")
accumulated_earnt_points_column=len(header)-1
header.append("Percent Complete")
percent_complete_column=len(header)-1
#create enumerated dictionary of header to help finding columns
headers={value:key for key,value in enumerate(header)}
#Find the special columns needed for processing
try:
	planned_for_column = headers['Planned For']
except:
	print("Error: csv file must contain 'Planned For' attribute")
	sys.exit(0)
try:	
	priority_column = headers['Priority']
except:
	print("Error: csv file must contain 'Priority' attribute")
	sys.exit(0)
try:	
	rank_column = headers['Rank (relative to Priority)']
except:
	print("Error: csv file must contain 'Rank (relative to Priority)' attribute")
	sys.exit(0)
try:
	id_column = headers['Id']
except:
	print("Error: csv file must contain 'Id' attribute")
	sys.exit(0)
try:
	parent_column = headers['Parent']
except:
	print("Error: csv file must contain 'Parent' attribute")
	sys.exit(0)
try:
	type_column = headers['Type']
except:
	print("Error: csv file must contain 'Type' attribute")
	sys.exit(0)
try:
	status_column = headers['Status']
except:
	print("Error: csv file must contain 'Status' attribute")
	sys.exit(0)
try:
	storypts_column = headers['Story Points']
except:
	print("Error: csv file must contain 'Story Points' attribute")
	sys.exit(0)

print("Create Raw Sheet...")
x_raw_sheet = print_header(x_raw_sheet,header, input_worksheet)
for i,row in enumerate(data,1):
	print_a_row(i,row, input_worksheet)	# write to input worksheet

# create dictionary for the header names and indices
print("Cleaning Data...")
for row in data:
	try:
		rank = row[rank_column].split(' ')[1]
	except:
		rank = 'z' # no rank assigned so give is very high ranking
	priority=priority_order[row[priority_column]]
	id=int(row[id_column])
	row[rank_column]='%06x.%s.%08x'%(priority,rank,id) # dots used as separators as they come before 0 in ascii.  short string comes before longer one.
	row[parent_column] = row[parent_column].lstrip('#')	# clean up parent_column id_column
	points = row[storypts_column].split(' ')[0]	# clean up story points
	if points is '':
		points = 0
	row[storypts_column] = int(points)
	
print("Sorting...")
data.sort(key= itemgetter(rank_column)) #creates ranked list


print("Creating Worksheets...")


x_grouped_sheet = print_header(x_grouped_sheet,header, grouped_worksheet,True)
	
id_dictionary={row[id_column]:row  for row in data if not check_filters(row)}
create_parent_list_dictionary()

missing_parents_report()
search_children(None,0)		# creates grouping and writes to grouped worksheet

x_ranked_sheet = print_header(x_ranked_sheet,header, ranked_worksheet,True)
for row in data:
	if not check_filters(row):
		x_ranked_sheet = print_a_row(x_ranked_sheet,row, ranked_worksheet)	# writes to ranked worksheet

print("Closing...")
try:
	workbook.close()	# close output file	
except IOError:
	print ('Error. Please close', output_excel_file_name, 'and rerun the program')
	sys.exit(0)
	
input_file.close()      # close input file				
print("Done")


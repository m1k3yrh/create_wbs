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
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from operator import itemgetter
import re	
import json
import locale

# read the input parameters
rtc_export_file_name = sys.argv[1]	
output_excel_file_name = sys.argv[2]
config_file_name = 'config.json'

workbook = xlsxwriter.Workbook(output_excel_file_name)	#creates output excel file 
grouped_worksheet = workbook.add_worksheet('Grouped')	#creates a worksheet for grouping
grouped_worksheet.outline_settings(True, False, False, True)	# displays the grouping summary above
ranked_worksheet = workbook.add_worksheet('Ranked')	#creates a worksheet for ranking
input_worksheet = workbook.add_worksheet('Input')	#creates a worksheet for ranking

# format variables
header_format = workbook.add_format()	# Add a bold format to use to highlight cells.
percent_complete_format = workbook.add_format()
percent_complete_format.set_num_format(0x09)
hyperlink_format=workbook.add_format()
hyperlink_format.set_font_color('blue')
hyperlink_format.set_underline()
formats={} # Create empty dictionary

#hyperlink
hyperlink_prefix=None

x_grouped_sheet = 0 # row index for grouped worksheet
x_ranked_sheet = 0	# row index for ranked worksheet
x_raw_sheet = 0	# row index for input worksheet
type_column = 0
percent_complete_column=0

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
	global formats
	try:
		return formats[cell_value]
	except:
		return None
	
		
# function to write a row to the output file
# function returns the index for the next available row
def print_a_row(ax,arow,aworksheet,depth=0): 
	global type_column
	global percent_complete_column
	
	for y in range(0,len(arow)):
		cell_format=set_cell_format(arow[y])
		if y==id_column and hyperlink_prefix!=None:
			aworksheet.write(ax,y,'=hyperlink("' + hyperlink_prefix + arow[y] +'",' + arow[y] + ')',hyperlink_format)
		elif y==percent_complete_column :
			aworksheet.write(ax,y,arow[y],percent_complete_format)
		else:
			aworksheet.write(ax,y,arow[y],cell_format)	
			
	aworksheet.set_row(ax, None, None, {'level': depth})		# sets the grouping level for this row
	ax +=1
	return ax

# function to write a row to the output file
# function returns the index for the next available row
def print_header(ax,arow,aworksheet): 
	y=0
	for col in arow:	
		if y <= len(arow):
			aworksheet.write(ax,y,col,header_format)	
		y +=1
	ax +=1
	return ax


# function to recursively group children in work break down structure and write to output file
def search_children(acurrent,depth): 
	global x_grouped_sheet
	myx = x_grouped_sheet-1
	mypts = 0
	my_completed_points=0
	childpts = 0
	completed_child_points=0
	children=False
	
	if acurrent==None:
		current_id=""		# orphan
	else:
		current_id = acurrent[id_column]
		mypts = acurrent[storypts_column]
		status=acurrent[status_column]
		if status=="In Progress" :
			my_completed_points=0.25*mypts
		elif status=="Implemented" :
			my_completed_points=0.75*mypts
		elif status=="Done" :
			my_completed_points=1.0*mypts
	
	for item in data:
		if current_id == item[parent_column]:
			x_grouped_sheet = print_a_row(x_grouped_sheet,item,grouped_worksheet,depth)	
			temp = search_children(item,depth+1)	# recursively add the story point for the child
			childpts+=temp['child_pts']
			completed_child_points+=temp['completed_child_points']
			children=True
			
	if myx >= 1:
		grouped_worksheet.write(myx,storypts_column,mypts)
		acurrent.append(my_completed_points)
		grouped_worksheet.write(myx,earned_story_points_column,my_completed_points)
		if children:
			acurrent.append(mypts+childpts)
			grouped_worksheet.write(myx,accumulated_story_points_column,mypts+childpts)
#			grouped_worksheet.write(myx,len(acurrent),"=sum("+xl_rowcol_to_cell(myx,storypts_column)+":"+xl_rowcol_to_cell(x_grouped_sheet-1,storypts_column)+")")
			acurrent.append(my_completed_points+completed_child_points)
			grouped_worksheet.write(myx,accumulated_earnt_points_column,my_completed_points+completed_child_points)
		try:
			percent_complete=(my_completed_points+completed_child_points)/(mypts+childpts)
		except ZeroDivisionError:
			percent_complete=""
		acurrent.append(percent_complete)
		grouped_worksheet.write(myx,percent_complete_column,percent_complete,percent_complete_format)

			
	return {'child_pts':mypts+childpts,'completed_child_points':my_completed_points+completed_child_points}	# base call to return story point for current


######################## Main starts here #####################
#read the config file
try:
	config_file = open(config_file_name)	#open the config file

except IOError:
    print ('Error. Cannot open', config_file_name)
    sys.exit(0)

configured_data = json.loads(config_file.read())	#read the config file
load_cell_formats()
header_format=formats['Header']

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

header = input[0]	#extract the header from the top row
	
# create dictionary for the header names and indices
print("Cleaning Data...")
headers = dict()	
for i in range(0, len(header)):
	temp = {header[i]:i}
	headers.update(temp)

planned_for_column = headers['Planned For']
priority_column = headers['Priority']
rank_column = headers['Rank (relative to Priority)']
id_column = headers['Id']
parent_column = headers['Parent']
type_column = headers['Type']
status_column = headers['Status']
storypts_column = headers['Story Points']

x_raw_sheet = print_header(x_raw_sheet,header, input_worksheet)

# Clean up the input data and do the sorting	
data = input[1:]	#extract all the work items
for row in data:
	x_raw_sheet = print_a_row(x_raw_sheet,row, input_worksheet)	# write to input worksheet
	rank_split = row[rank_column].split(' ')
	row[rank_column]= rank_split.pop()	# clean up rank_column
	row[parent_column] = row[parent_column].lstrip('#')	# clean up parent_column id_column
	points = row[storypts_column].split(' ')[0]	# clean up story points
	if points is '':
		points = 0
	row[storypts_column] = int(points)
	
print("Sorting...")
data.sort(key= itemgetter(planned_for_column,priority_column,rank_column,id_column)) #creates ranked list


print("Creating Worksheets...")
# add extra calculation rows
header.append("Earned Story Points")
earned_story_points_column=len(header)-1
header.append("Accumulated Story Points")
accumulated_story_points_column=len(header)-1
header.append("Accumulated Earnt Points")
accumulated_earnt_points_column=len(header)-1
header.append("Percent Complete")
percent_complete_column=len(header)-1

x_grouped_sheet = print_header(x_grouped_sheet,header, grouped_worksheet)
search_children(None,0)		# creates grouping and writes to grouped worksheet

x_ranked_sheet = print_header(x_ranked_sheet,header, ranked_worksheet)
for row in data:
	x_ranked_sheet = print_a_row(x_ranked_sheet,row, ranked_worksheet)	# writes to ranked worksheet

print("Closing...")
try:
	workbook.close()	# close output file	
except IOError:
	print ('Error. Please close', output_excel_file_name, 'and rerun the program')
	sys.exit(0)
	
input_file.close()      # close input file				
config_file.close()	# close config file
print("Done")


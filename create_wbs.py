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
iteration_report_worksheet = workbook.add_worksheet('Iteration Report')	#creates a iteration_report_worksheet for grouping
iteration_report_worksheet.outline_settings(True, False, False, True)	# displays the grouping summary above
ranked_worksheet = workbook.add_worksheet('Ranked')	#creates a worksheet for ranking
input_worksheet = workbook.add_worksheet('Input')	#creates a worksheet for regurgitating input
error_worksheet = workbook.add_worksheet('Errors') # workbook for all suspicious data found from input
error_sheet_color='green'


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
	try:
		configs=load_config('Format')
	except:
		formats=None
	else:
		try:
			formats={config_key:workbook.add_format(config_value) for config_key,config_value in configs.items()}
		except:
			print("config format error:  Couldn't parse:",config_value)
			sys.exit(-1)
	
def set_cell_format(cell_value):
	try:
		return formats[cell_value]
	except:
		return None
	
# Test the Filters to see if any of the columns match
def check_filters(arow):
	if filters:
		for key,filter in filters.items():
			y=headers[key]
			if arow[y] in filter:
				return True
		return False
		
# function to write a row to the output file
# function returns the index for the next available row
def print_a_row(ax,arow,aworksheet,depth=0,format=True,options={}): 
	for i,y in enumerate(arow):
		if format and ( i==id_column or i==parent_column ) and hyperlink_prefix!=None:
			aworksheet.write_url(ax,i,hyperlink_prefix + y ,None, y )
		elif format and i==percent_complete_column :
			aworksheet.write(ax,i,y,percent_complete_format)
		else:
			if format: 
				cell_format=set_cell_format(y)
			else:
				cell_format=None
			aworksheet.write(ax,i,y,cell_format)	
	
	options['level']=depth	
	aworksheet.set_row(ax, None, None, options)		# sets the grouping level for this row
	ax +=1
	return ax

# print a list of rows
def print_list(ax,alist,aworksheet,depth=0,format=True,options={}):
	for row in alist:
		ax=print_a_row(ax,row,aworksheet,depth,format,options={})
	return ax

# function to write a row to the output file
# function returns the index for the next available row
def print_header(ax,arow,aworksheet,hidden=False): 
	for y,col in enumerate(arow):	
		aworksheet.write(ax,y,col,header_format)
		if hidden and hidden_columns and col in hidden_columns:
			aworksheet.set_column(y,y,None,None,{'hidden':True})
	aworksheet.freeze_panes(ax+1, 0)
	ax +=1
	return ax

#Creates a dictionary of lists of rows with a specific parent.
def create_parent_list_dictionary():
	global parent_list
	global id_dictionary
	
	for row in filtered_data:
		parent=row[parent_column]
		if parent in parent_list:
			parent_list[parent].append(row) # add to list against key parent
		else:
			parent_list[parent]=[row] # Create a new list against key parent


# function to recursively group children in work break down structure and write to output file
def search_children(acurrent,depth): 
	global x_grouped_sheet
	
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


# Checks if the status has risen (green,orange, red)
def set_error_sheet_color(status):
	global error_sheet_color
	
	if error_sheet_color=='red':
		return # already at max level.
	else:
		error_sheet_color=status
		error_worksheet.set_tab_color(error_sheet_color)
	return	

# Check that all items in data have parents in data as well
def missing_parents_report():
	global x_error_sheet
	
	is_error=False
	for id in parent_list:
		if not id=="" and not id in id_dictionary:
#Note: (not id=="") above is to ignore roots (items with no parent.
			if not is_error:
				set_error_sheet_color('red')
				error_worksheet.write(x_error_sheet,0,
									"Fatal: The follow items don't have parents in the input file.  Data will be missing from WBS",error_format)
				x_error_sheet+=1
				is_error=True
			x_error_sheet=print_list(x_error_sheet,parent_list[id],error_worksheet)
	return is_error

def parents_before_children_report():
	global x_error_sheet
	planned_for_order={value:key for key,value in enumerate(planned_for_list)}
	
	is_error=False
	for row in filtered_data:
		try:
			parent=id_dictionary[row[parent_column]]
		except KeyError:
			continue # Item has no parent so move on and check next one
		parent_planned_for=planned_for_order[parent[planned_for_column]] # Find ranking of parent
		child_planned_for=planned_for_order[row[planned_for_column]]
		if parent_planned_for<child_planned_for:
			if not is_error:
				set_error_sheet_color('orange')
				error_worksheet.write(x_error_sheet,0,
						"Warning: The follow items are in Iterations after their parent. (A parent can't be completed until all children are completed.  Suggest you move the Parent to Backlog or the same iteration as child)",error_format)
				x_error_sheet+=1
				is_error=True
			x_error_sheet=print_a_row(x_error_sheet,row,error_worksheet)
	return is_error
	
	
	
# Optional checks done on data to check that items are in correct location and state based on progress etc.
def wrong_state_report():
	global x_error_sheet

	if new_states:
		list=[row for row in filtered_data
				if row[status_column] in new_states and len(row)>accumulated_earnt_points_column and row[accumulated_earnt_points_column]>0]
		if list:
			set_error_sheet_color('orange')
			error_worksheet.write(x_error_sheet,0,
								"Warning: The follow items are marked new but have children with progress",error_format)
			x_error_sheet+=1
			x_error_sheet=print_list(x_error_sheet,list,error_worksheet)

	if completed_states:
		list=[row for row in filtered_data if not row[status_column] in completed_states and len(row)>percent_complete_column \
													and row[percent_complete_column]!='' and row[percent_complete_column]>0.99]
		if list:
			set_error_sheet_color('orange')
			error_worksheet.write(x_error_sheet,0,
								"Warning: The follow items are have all children complete but are still in early progress",error_format)
			x_error_sheet+=1
			x_error_sheet=print_list(x_error_sheet,list,error_worksheet)
	
	if impeded_states:
		is_warning=False
		for row in filtered_data:
			try:
				if row[status_column] in impeded_states:
					unimpeded_children=[row for row in parent_list[row[id_column]] if not row[status_column] in impeded_states]
					if not unimpeded_children==[]:
						if is_warning==False:
							set_error_sheet_color('orange')
							error_worksheet.write(x_error_sheet,0,
												"Warning: The follow items have open children even though the parent is Impeded",error_format)
							x_error_sheet+=1
							is_warning=True
						x_error_sheet=print_a_row(x_error_sheet,row,error_worksheet)
			except (KeyError):
				continue # KeyError occurs if item has no children

		is_warning=False
		for row in filtered_data:
			try:
				if not row[status_column] in impeded_states:
					unimpeded_children=[row for row in parent_list[row[id_column]] if not row[status_column] in impeded_states]
					if unimpeded_children==[]:
						if is_warning==False:
							set_error_sheet_color('orange')
							error_worksheet.write(x_error_sheet,0,
												"Warning: The follow items are open even though all children are Impeded",error_format)
							x_error_sheet+=1
							is_warning=True
						x_error_sheet=print_a_row(x_error_sheet,row,error_worksheet)
			except (KeyError):
				continue # KeyError occurs if item has no children
	
	if product_work_items and product_backlogs:
		list=[row for row in filtered_data if row[type_column] in product_work_items and not row[planned_for_column] in product_backlogs]
		if list:
			set_error_sheet_color('orange')
			error_worksheet.write(x_error_sheet,0,
								"Warning: The follow product items are not plannedFor Product Backlog",error_format)
			x_error_sheet+=1
			x_error_sheet=print_list(x_error_sheet,list,error_worksheet)
	
	if product_work_items and product_categories:
		is_warning=False
		list=[row for row in filtered_data if row[type_column] in product_work_items and not row[filed_against_column] in product_categories]
		if list:
			set_error_sheet_color('orange')
			error_worksheet.write(x_error_sheet,0,
								"Warning: The follow Product Items are not FiledAgainst Product Categories",error_format)
			x_error_sheet+=1
			x_error_sheet=print_list(x_error_sheet,list,error_worksheet)
	
	if team_work_items and team_categories:
		list=[row for row in filtered_data if row[type_column] in team_work_items and not row[filed_against_column] in team_categories]
		if list:
			set_error_sheet_color('orange')
			error_worksheet.write(x_error_sheet,0,
								"Warning: The follow Team Items are not FiledAgainst Team Categories",error_format)
			x_error_sheet+=1
			x_error_sheet=print_list(x_error_sheet,list,error_worksheet)
	
	return

def load_config(key):
	try:
		return configured_data[key]
	except:
		return None
	
def find_column(key):
	try:
		return headers[key]
	except:
		print("Error: csv file must contain '%s' attribute"%(key))
		sys.exit(0)
		
def create_iteration_team_report():
	x=0
	x=print_header(x,header,iteration_report_worksheet,hidden=True)

# Build a dictionary of everything in each Sprint/iteration	
	d={}
	if team_categories:
		tc=team_categories
	else:
		tc=['all'] # Create a dummy team_categories
	for row in filtered_data:
		if team_categories:
			t=(row[planned_for_column],row[filed_against_column])
		else:
			t=(row[planned_for_column],tc[0]) # If team_categories not defined, stuff everything in all
		if t in d:
			d[t].append(row)
		else:
			d[t]=[row]
			
	for p in planned_for_list:
		iteration_report_worksheet.write(x,0,p)
		px=x # Remember the Iteration Row (aka PlannedFor) so can add data later
		x+=1
		iteration_total_pts=0.0
		iteration_earned_pts=0.0
		for f in tc:
			if team_categories: # Only create a Category row if team_categories is defined
				iteration_report_worksheet.write(x,0,f)
				iteration_report_worksheet.set_row(x, None, None, {'level': 1,'collapsed':True})		# sets the grouping level for this row
				fx=x # remember the FiledAgainst row so can add data later
				x+=1
				depth=2
			else:
				depth=1
			category_total_pts=0.0
			category_earned_pts=0.0
			try:
				l=d[(p,f)]
			except KeyError:
				pass # Nothing to print for this combination
			else:
				for row in l: # Add up points from all the children
					try:
						category_total_pts+=row[storypts_column] 
						category_earned_pts+=row[earned_story_points_column]
# Note: We only add points belonging to this object (not accumulated points which includes children) to prevent double counting.
					except IndexError:
						pass # Ignore index error.  Occurs when an item has parents which are not in data set.  A Fatal error is reported to Error Sheet
					x=print_a_row(x,row,iteration_report_worksheet,depth=depth,options={'hidden':True})
			if team_categories: # Only update the Category row if team_categories is defined
				iteration_report_worksheet.write(fx,storypts_column,category_total_pts)
				iteration_report_worksheet.write(fx,earned_story_points_column,category_earned_pts)
				if category_total_pts: # Avoid div_zero.  Don't calculate percent if there is no points.
					iteration_report_worksheet.write(fx,percent_complete_column,category_earned_pts/category_total_pts,percent_complete_format)
			iteration_total_pts+=category_total_pts
			iteration_earned_pts+=category_earned_pts
		iteration_report_worksheet.write(px,storypts_column,iteration_total_pts)
		iteration_report_worksheet.write(px,earned_story_points_column,iteration_earned_pts)
		if iteration_total_pts: # Avoid div_zero.  Don't calculate percent if there is no points.
			iteration_report_worksheet.write(px,percent_complete_column,iteration_earned_pts/iteration_total_pts,percent_complete_format)

# Calculate ranking and clean up Parent and Story Points columns
def clean_data():
	priority_order={value:key for key,value in enumerate(priority_list)}
	for row in filtered_data:
		try:
			rank = row[rank_column].split(' ')[1]
		except:
			rank = 'z' # no rank assigned so give is very high ranking (i.e bottom of list)
		priority=priority_order[row[priority_column]]
		id=int(row[id_column])
		row[rank_column]='%06x.%s.%08x'%(priority,rank,id) # dots used as separators as they come before 0 in ascii.  short string comes before longer one.
		row[parent_column] = row[parent_column].lstrip('#')	# clean up parent_column id_column
		points = row[storypts_column].split(' ')[0]	# clean up story points
		if points is '':
			points = 0
		row[storypts_column] = int(points)
		

######################## Main starts here #####################
#read the configuration

load_cell_formats()
try:
	header_format=formats['Header']
except:
	header_format=None
try:
	error_format=formats['Error_Item']
except:
	error_format=None
planned_for_list=load_config('Planned For')
priority_list=load_config('Priority')
if not planned_for_list or not priority_list:
	print("Error 'Planned For' and 'Priority' lists must be defined in Configuration file")
	sys.exit(-1)
progress_table=load_config('Progress')
filters=load_config('Filters')
hidden_columns=load_config('Hidden')
hyperlink_prefix=load_config('Hyperlink')
team_categories=load_config('Team Categories')
team_work_items=load_config('Team Work Items')
product_work_items=load_config('Product Work Items')
product_categories=load_config('Product Categories')
product_backlogs=load_config('Product Backlogs')
impeded_states=load_config('Impeded States')
completed_states=load_config('Completed States')
new_states=load_config('New States')

try:
#	log_file = open('debug.txt', 'w')	#open debug output file to dump trace
	print("Opening File...")
	input_file = open(rtc_export_file_name, encoding='utf-16') # opens the csv file
except IOError:
	print ('Error. Cannot open', rtc_export_file_name)
	sys.exit(0)

print("Reading File...")
reader = csv.reader(input_file,delimiter='\t')  # creates the reader 
input = [r for r in reader]

# Split input into header and data
header = input[0]	#extract the header from the top row
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
planned_for_column = find_column('Planned For')
filed_against_column = find_column('Filed Against')
priority_column = find_column('Priority')
rank_column = find_column('Rank (relative to Priority)')
id_column = find_column('Id')
parent_column = find_column('Parent')
type_column = find_column('Type')
status_column = find_column('Status')
storypts_column = find_column('Story Points')

#
data = input[1:]	#extract all the work items
filtered_data=[row for row in data if not check_filters(row)]


try:
	print("Create Raw Sheet...")
	x_raw_sheet = print_header(x_raw_sheet,header, input_worksheet)
	print_list(1,data, input_worksheet,format=False)	# write to input worksheet
	
	# Prep error sheet
	x_error_sheet = print_header(x_error_sheet,header, error_worksheet)
	set_error_sheet_color(error_sheet_color)
	
	print("Cleaning Data...")
	clean_data()
	print("Sorting...")
	filtered_data.sort(key= itemgetter(rank_column)) #creates ranked list
	
	
	print("Creating Worksheets...")
	x_grouped_sheet = print_header(x_grouped_sheet,header, grouped_worksheet,hidden=True)
		
	id_dictionary={row[id_column]:row  for row in filtered_data}
	create_parent_list_dictionary()
	
	search_children(None,0)		# creates grouping and writes to grouped worksheet
	
	x_ranked_sheet = print_header(x_ranked_sheet,header, ranked_worksheet,hidden=True)
	x_ranked_sheet = print_list(x_ranked_sheet,filtered_data, ranked_worksheet)	# writes to ranked worksheet
	
	# check data consistency and report to Error tab
	missing_parents_report()
	wrong_state_report()
	parents_before_children_report()
	
	# create iteration report
	create_iteration_team_report()
	
except:
	workbook.close() # close what we've achieved
	raise # re-throw the error

print("Closing...")
try:
	workbook.close()	# close output file	
except IOError:
	print ('Error. Please close', output_excel_file_name, 'and rerun the program')
	sys.exit(0)
	
input_file.close()      # close input file				
print("Done")


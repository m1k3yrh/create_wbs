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





class output_worksheet_class:
	def __init__(self,name):
		self.worksheet=workbook.add_worksheet(name)
		self.x=0
	
	# function to write a row to the output file
	# function returns the index for the next available row
	def append_a_row(self,arow,depth=0,format=True,options={}): 
		aworksheet=self.worksheet
		ax=self.x
		for i,y in enumerate(arow):
			if format and ( i==header.id_column or i==header.parent_column ) and config.hyperlink_prefix!=None:
				aworksheet.write_url(ax,i,config.hyperlink_prefix + y ,None, y )
			elif format and i==header.percent_complete_column:
				aworksheet.write(ax,i,y,config.percent_complete_format)
			else:
				if format: 
					cell_format=config.set_cell_format(y)
				else:
					cell_format=None
				aworksheet.write(ax,i,y,cell_format)	
		
		options['level']=depth	# Add Depth to options string.
		aworksheet.set_row(ax, None, None, options)		# sets the grouping level for this row
		self.x+=1 # Move to next row.

	# print a list of rows
	def append_list(self,alist,depth=0,format=True,options={},raw=False):
		for row in alist:
			if raw:
				r=row.raw_row
			else:
				r=row.processed_row
			self.append_a_row(r,depth=depth,format=format,options=options)
		
	# function to write a row to the output file
	# function returns the index for the next available row
	def append_header(self,header,hidden=False): 
		aworksheet=self.worksheet
		for y,col in enumerate(header):	
			aworksheet.write(self.x,y,col,config.header_format)
			if hidden and config.hidden_columns and col in config.hidden_columns:
				aworksheet.set_column(y,y,None,None,{'hidden':True})
		aworksheet.freeze_panes(self.x+1, 0)
		self.x +=1
	

class iteration_report_worksheet_class(output_worksheet_class):
	def append_iteration(self,s):
		self.worksheet.write(self.x,0,s,config.iteration_summary_format)
		x=self.x
		self.x+=1 # move x to next row.
		return x # return row of iteration data
	
	def append_category(self,s):
		self.worksheet.write(self.x,0,s,config.category_summary_format)
		self.worksheet.set_row(self.x, None, None, {'level': 1,'collapsed':True})		# sets the grouping level for this row
		x=self.x
		self.x+=1 # move x to next row.
		return x # return row of iteration data

# Add total points, earned points and percent complete to a row that was previously written to.
	def update_row(self,fx,total_pts,earned_pts):
		self.worksheet.write(fx,header.storypts_column,total_pts)
		if config.progress_table:
			self.worksheet.write(fx,header.earned_story_points_column,earned_pts)
			if total_pts: # Avoid div_zero.  Don't calculate percent if there is no points.
				self.worksheet.write(fx,header.percent_complete_column,earned_pts/total_pts,config.percent_complete_format)

		
class output_worksheet_error_class(output_worksheet_class):	
	def __init__(self,name):
		super().__init__(name)
		self.error_sheet_color='green'
	
	def append_warning(self,s):	
		if not self.error_sheet_color=='red':
			self.error_sheet_color='orange'
			self.worksheet.set_tab_color(self.error_sheet_color)
		self.worksheet.write(self.x,0,s,config.error_format)
		self.x+=1 # Move to next row
		
	def append_error(self,s):	
		self.error_sheet_color='red'
		self.worksheet.set_tab_color(self.error_sheet_color)
		self.worksheet.write(self.x,0,s,config.error_format)
		self.x+=1 # Move to next row
		
# Class for configuration
class config_class:
	def __init__(self):
# Load configuration from local config.py file
		sys.path.append(os.getcwd())
		from config import configured_data
		self.configured_data=configured_data
		
		self.formats=self.load_cell_formats()
		try:
			self.header_format=self.formats['Header']
		except:
			self.header_format=None
		try:
			self.error_format=self.formats['Error_Item']
		except:
			self.error_format=None
		try:
			self.iteration_summary_format=self.formats['Iteration_Summary']
		except:
			self.iteration_summary_format=None
		try:
			self.category_summary_format=self.formats['Category_Summary']
		except:
			self.category_summary_format=None
		self.percent_complete_format = workbook.add_format()
		self.percent_complete_format.set_num_format(0x09) # Predefined Excel format for %age with no decimals
	
		self.planned_for_list=self.load_config('Planned For')
		self.priority_list=self.load_config('Priority')
		if not self.planned_for_list or not self.priority_list:
			print("Error 'Planned For' and 'Priority' lists must be defined in Configuration file")
			sys.exit(-1)
			
		self.priority_order={value:key for key,value in enumerate(self.priority_list)}

		self.progress_table=self.load_config('Progress')
		self.filters=self.load_config('Filters')
		self.hidden_columns=self.load_config('Hidden')
		self.hyperlink_prefix=self.load_config('Hyperlink')
		self.team_categories=self.load_config('Team Categories')
		self.team_work_items=self.load_config('Team Work Items')
		self.product_work_items=self.load_config('Product Work Items')
		self.product_categories=self.load_config('Product Categories')
		self.product_backlogs=self.load_config('Product Backlogs')
		self.impeded_states=self.load_config('Impeded States')
		self.completed_states=self.load_config('Completed States')
		self.new_states=self.load_config('New States')
		self.average_sizes=self.load_config('Average Sizes')

	def load_cell_formats(self):
		configs=self.load_config('Format')
		if not configs:
			formats=None
		else:
			formats={config_key:workbook.add_format(config_value) for config_key,config_value in configs.items()}
		return formats
		
	def set_cell_format(self,cell_value):
		try:
			return self.formats[cell_value]
		except:
			return None

	def load_config(self,key):
		try:
			return self.configured_data[key]
		except:
			return None


# Class to for header row.  Interprets the header and calculates column for specific attributes used in processing.
class header_class:
	
	def __init__(self,r):
		self.raw_header=r

		self.full_header=list(self.raw_header) # Clone list
		if config.progress_table:
			self.full_header.append("Earned Story Points")
		self.full_header.append("Accumulated Story Points")
		if config.progress_table:
			self.full_header.append("Accumulated Earnt Points")
			self.full_header.append("Percent Complete")		
		if config.average_sizes:
			self.full_header.append("Based on Shirt Size")

		#create enumerated dictionary of header to help finding columns
		self.headers={value:key for key,value in enumerate(self.full_header)}

		#Find the special columns needed for processing
		self.planned_for_column = self.find_column('Planned For')
		self.filed_against_column = self.find_column('Filed Against')
		self.priority_column = self.find_column('Priority')
		self.rank_column = self.find_column('Rank (relative to Priority)')
		self.id_column = self.find_column('Id')
		self.parent_column = self.find_column('Parent')
		self.type_column = self.find_column('Type')
		self.status_column = self.find_column('Status')
		self.storypts_column = self.find_column('Story Points')	
		if config.progress_table:
			self.earned_story_points_column=self.find_column("Earned Story Points")
		else:
			self.earned_story_points_column=None
		self.accumulated_story_points_column=self.find_column("Accumulated Story Points")
		if config.progress_table:
			self.accumulated_earnt_points_column=self.find_column("Accumulated Earnt Points")
			self.percent_complete_column=self.find_column("Percent Complete")
		else:
			self.accumulated_earnt_points_column=None
			self.percent_complete_column=None
		if config.average_sizes:
			self.shirt_size_column=self.find_column("SSPoints")	
			self.based_on_shirt_size_column=self.find_column("Based on Shirt Size")
		else:
			self.based_on_shirt_size_column=None
			self.shirt_size_column=None

	def find_column(self,key,optional=False):
		try:
			return self.headers[key]
		except:
			if not optional:
				print("Error: csv file must contain '%s' attribute"%(key))
				sys.exit(0)	
			else:
				return None


# Class to store each row from the input file.
class spreadsheet_row:
	def __init__(self,r):
		self.raw_row=[] # where input raw_row is stored
		self.processed_row=[] # row with cleaned up and calculated values added.
		self.parent=None # where parent is stored after decode
		self.children=[]
		self.rank=None
		self.id=None
		self.points=None
		self.status=None
		self.earned_points=None
		self.accumulated_points=None
		self.accumulated_earned_points=None
		self.percent_complete=None
		self.based_on_shirt_size=False
		self.planned_for=None
		self.filed_against=None
		self.type=None
		self.shirt_size=None
	
		self.raw_row=r
		self.processed_row=list(self.raw_row) # clone the row so changes to processed row do not impact raw row.
		try:
			rank = self.raw_row[header.rank_column].split(' ')[1]
		except:
			rank = 'z' # no rank assigned so give is very high ranking (i.e bottom of list)
		priority=config.priority_order[self.raw_row[header.priority_column]]
		self.id=self.raw_row[header.id_column]
		self.status=self.raw_row[header.status_column]
		self.processed_row[header.rank_column]=self.rank='%06x.%s.%08x'%(priority,rank,int(self.id)) # dots used as separators as they come before 0 in ascii.  short string comes before longer one.
		
		self.processed_row[header.parent_column]=self.parent = self.raw_row[header.parent_column].lstrip('#')	# clean up parent_column id_column
		pts = self.raw_row[header.storypts_column].split(' ')[0]	# clean up story points
		if pts is '':
			pts=None
		else:
			pts=int(pts)
		self.processed_row[header.storypts_column]=self.points = pts
		self.processed_row.extend(['']*(len(header.full_header)-len(header.raw_header))) # Add extra columns to make to length of full header
		
		if header.planned_for_column!=None:
			self.planned_for=self.raw_row[header.planned_for_column]
		if header.filed_against_column!=None:
			self.filed_against=self.raw_row[header.filed_against_column]
		if header.type_column!=None:
			self.type=self.raw_row[header.type_column]
		if header.shirt_size_column!=None:
			self.shirt_size=self.raw_row[header.shirt_size_column]
			if self.shirt_size=='Unassigned' or len(self.shirt_size)==0:
				self.shirt_size=None

	# Test the Filters to see if any of the columns match
	def check_filters(self):
		if config.filters:
			for key,filter in config.filters.items():
				y=header.find_column(key)
				if self.processed_row[y] in filter:
					return True
		return False

	def __lt__(self, other):
		return self.rank < other.rank	
	
	def add_child(self, child):
		self.children.append(child)	
		
	def set_accumulated_points(self,pts):
		self.accumulated_points=pts
		self.processed_row[header.accumulated_story_points_column]=pts
		
	def set_earned_points(self,pts):
		self.earned_points=pts
		self.processed_row[header.earned_story_points_column]=pts
		
	def set_acc_earned_points(self,pts):
		self.accumulated_earned_points=pts
		self.processed_row[header.accumulated_earnt_points_column]=pts
		
	def set_percent_complete(self,pts):
		self.percent_complete=pts
		self.processed_row[header.percent_complete_column]=pts
		
	def set_based_on_shirt_size(self,b):
		self.based_on_shirt_size=b
		if b:
			self.processed_row[header.based_on_shirt_size_column]='E'

# Takes a list of spreadsheet_rows and builds a dictionary using the id of each item as key.
class id_dictionary_class(dict):
	def __init__(self,data):
		r={row.id:row  for row in data}
		self.update(r)
	

# function to recursively group children in work break down structure and write to output file
class work_breakdown_class:
	def __init__(self,data,id_dictionary):
		self.roots=[]
		self.items_with_missing_parents=[]
		for n in data:
			if n.parent=='':
				self.roots.append(n) # Add to list of roots
			else:
				try:
					parent=id_dictionary[n.parent]
				except:
					self.items_with_missing_parents.append(n) # If no parent add to list of items with no parents in data provided
				else:
					parent.add_child(n)
		if len(self.items_with_missing_parents):
			self.missing_parents_report(self.items_with_missing_parents)
		
	def calculate_points(self,acurrent=None):
		if acurrent==None:
			for row in self.roots:
				self.calculate_points(row)
		else:
			mypts=acurrent.points
			status=acurrent.status
			try:
				my_earned_pts=config.progress_table[status]*mypts
			except:
				my_earned_pts=0
			acc_earned_pts=0

			based_on_shirt_size=False
			if len(acurrent.children):
				acc_pts=0
				for row in acurrent.children:
					[p,q,e]=self.calculate_points(row) # recursively add points from children
					if p==None:
						acc_pts=None # If any child has no points we cannot estimate points for parent
					elif not acc_pts==None:
						acc_pts+=p
					acc_earned_pts+=q
					based_on_shirt_size|=e
			else:
				acc_pts=None
			if config.average_sizes and (acc_pts==None or acc_pts==0):  # If average sizes are defined for Shirt Sizes then try to use these when Story points don't exist
				shirt_size=acurrent.shirt_size
				try:
					ss_value=config.average_sizes[shirt_size]
				except:
					based_on_shirt_size=False
					acc_pts=None
					pass # if we fail to match just press on
				else:
					acc_pts=ss_value
					based_on_shirt_size=True
			if acc_pts==None and mypts!=None:
				acc_pts=mypts #  don't have points but we do so our points are the points.
			elif acc_pts!=None and mypts!=None:
				acc_pts+=mypts # If I have points and so does my child, add them together.
			acc_earned_pts+=my_earned_pts
			acurrent.set_accumulated_points(acc_pts)
			if config.progress_table:
				acurrent.set_earned_points(my_earned_pts)
				acurrent.set_acc_earned_points(acc_earned_pts)
			acurrent.set_based_on_shirt_size(based_on_shirt_size)
			if acc_pts and acc_pts>0 and header.percent_complete_column:
				percent_complete=acc_earned_pts/acc_pts
				acurrent.set_percent_complete(percent_complete)
			return [acc_pts,acc_earned_pts,based_on_shirt_size]
		
	def write_to_spreadsheet(self,worksheet,acurrent=None,depth=0):
		if acurrent==None:
			for row in self.roots:
				self.write_to_spreadsheet(worksheet,row,depth)
		else:
			worksheet.append_a_row(acurrent.processed_row,depth=depth)
			if len(acurrent.children):
				for row in acurrent.children:
					self.write_to_spreadsheet(worksheet,row,depth+1)

	# Check that all items in data have parents in data as well
	def missing_parents_report(self,missing_parent_list):
		error_worksheet.append_error("Fatal: The follow items don't have parents in the input file.  Data will be missing from WBS")
		error_worksheet.append_list(missing_parent_list)


class iteration_team_report_class:
	def __init__(self,data):
# Build a dictionary of everything in each Sprint/iteration	
		self.data={}
		if config.team_categories:
			tc=config.team_categories
		else:
			tc=['all'] # Create a dummy team_categories
		for row in data:
			if config.team_categories:
				t=(row.planned_for,row.filed_against)
			else:
				t=(row.planned_for,tc[0]) # If team_categories not defined, stuff everything in all
			if t in self.data:
				self.data[t].append(row)
			else:
				self.data[t]=[row]

	def create_iteration_team_report(self):
		iteration_report_worksheet.append_header(header.full_header,hidden=True)

		if config.team_categories:
			tc=config.team_categories
		else:
			tc=['all'] # Use the dummy team_categories					
		for p in config.planned_for_list:
			px=iteration_report_worksheet.append_iteration(p)
			iteration_total_pts=0.0
			iteration_earned_pts=0.0
			for f in tc:
				if config.team_categories: # Only create a Category row if team_categories is defined
					fx=iteration_report_worksheet.append_category(f)
					depth=2
				else:
					depth=1
				category_total_pts=0.0
				category_earned_pts=0.0
				try:
					l=self.data[(p,f)]
				except KeyError:
					pass # Nothing to print for this combination
				else:
					for row in l: # Add up points from all the children
						try:
							category_total_pts+=row.points 
							category_earned_pts+=row.earned_points
	# Note: We only add points belonging to this object (not accumulated points which includes children) to prevent double counting.
						except (IndexError,TypeError):
							pass # Ignore index error.  Occurs when an item has parents which are not in data set.  A Fatal error is reported to Error Sheet
								# Ignore TypeError.  Occurs when item doesn't have points.
						x=iteration_report_worksheet.append_a_row(row.processed_row,depth=depth,options={'hidden':True})
				if config.team_categories: # Only update the Category row if team_categories is defined
					iteration_report_worksheet.update_row(fx,category_total_pts,category_earned_pts)
				iteration_total_pts+=category_total_pts
				iteration_earned_pts+=category_earned_pts
			iteration_report_worksheet.update_row(px,iteration_total_pts,iteration_earned_pts)


def parents_before_children_report():
	planned_for_order={value:key for key,value in enumerate(config.planned_for_list)}
	
	is_error=False
	for row in filtered_data:
		try:
			parent=id_dictionary[row.parent]
		except KeyError:
			continue # Item has no parent so move on and check next one
		parent_planned_for=planned_for_order[parent.planned_for] # Find ranking of parent
		child_planned_for=planned_for_order[row.planned_for]
		if parent_planned_for<child_planned_for:
			if not is_error:
				error_worksheet.append_warning("Warning: The follow items are in Iterations after their parent. (A parent can't be completed until all children are completed.  Suggest you move the Parent to Backlog or the same iteration as child)")
				is_error=True
			error_worksheet.append_a_row(row.processed_row)
	return is_error
	
	
	
# Optional checks done on data to check that items are in correct location and state based on progress etc.
def wrong_state_report():
	if config.new_states:
		list=[row for row in filtered_data if row.status in config.new_states and row.accumulated_earned_points!=None and row.accumulated_earned_points>0]
		if list:
			error_worksheet.append_error("Warning: The follow items are marked new but have children with progress")
			error_worksheet.append_list(list)

	if config.completed_states:
		list=[row for row in filtered_data if not row.status in config.completed_states and row.percent_complete!=None and row.percent_complete>0.99]
		if list:
			error_worksheet.append_warning("Warning: The follow items are have all children complete but are still in early progress")
			error_worksheet.append_list(list)
	
	if config.impeded_states:
		is_warning=False
		for row in filtered_data:
			if row.status in config.impeded_states and not row.children==[]:
				unimpeded_children=[l for l in row.children if not l.status in config.impeded_states]
				if not unimpeded_children==[]:
					if is_warning==False:
						error_worksheet.append_warning("Warning: The follow items have open children even though the parent is Impeded")
						is_warning=True
					error_worksheet.append_a_row(row.processed_row)

		is_warning=False
		for row in filtered_data:
			if not row.status in config.impeded_states and not row.children==[]:
				impeded_children=[l for l in row.children if not l.status in config.impeded_states]
				if impeded_children==[]:
					if is_warning==False:
						error_worksheet.append_warning("Warning: The follow items are open even though all children are Impeded")
						is_warning=True
					error_worksheet.append_a_row(row.processed_row)
	
	if config.product_work_items and config.product_backlogs:
		list=[row for row in filtered_data if row.type in config.product_work_items and not row.planned_for in config.product_backlogs]
		if list:
			error_worksheet.append_warning("Warning: The follow product items are not plannedFor Product Backlog")
			error_worksheet.append_list(list)
	
	if config.product_work_items and config.product_categories:
		is_warning=False
		list=[row for row in filtered_data if row.type in config.product_work_items and not row.filed_against in config.product_categories]
		if list:
			error_worksheet.append_warning("Warning: The follow Product Items are not FiledAgainst Product Categories")
			error_worksheet.append_list(list)
	
	if config.team_work_items and config.team_categories:
		list=[row for row in filtered_data if row.type in config.team_work_items and not row.filed_against in config.team_categories]
		if list:
			error_worksheet.append_warning("Warning: The follow Team Items are not FiledAgainst Team Categories")
			error_worksheet.append_list(list)
	
	return

def shirtsizechecks():
	parent_with_shirt_size=[]
	leaf_without_shirt_size=[]
	undefined_shirt_size=[]
	item_started_but_no_children=[]
	
# Find Work items which are e2e (product items).  Leaves of Product item tree should contain ShirtSizes
	product_work_items=[row for row in filtered_data if row.type in config.product_work_items]
	
	for row in product_work_items:
		child_product_work_items=[r for r in product_work_items if r in row.children]
		if row.shirt_size==None and child_product_work_items==[]:
			leaf_without_shirt_size.append(row)
		elif row.shirt_size!=None and child_product_work_items!=[]:
			parent_with_shirt_size.append(row)
		if row.shirt_size!=None and row.shirt_size not in config.average_sizes:
			undefined_shirt_size.append(row)
		if row.children==[] and row.status not in config.new_states:
			item_started_but_no_children.append(row)
			
	if leaf_without_shirt_size!=[]:
		error_worksheet.append_warning("Warning: The following Product Items have no Shirt Size defined.  (Total backlog sizing may be inaccurate)")
		error_worksheet.append_list(leaf_without_shirt_size)
	if parent_with_shirt_size!=[]:
		error_worksheet.append_warning("Warning: The following Product Items have ShirtSizes even though they have children.  (Only leaf end-2-end work items should have ShirtSizes otherwise double counting may occur)")
		error_worksheet.append_list(parent_with_shirt_size)
	if undefined_shirt_size!=[]:
		error_worksheet.append_error("FATAL: The following Product Items have ShirtSizes which are not defined")
		error_worksheet.append_list(undefined_shirt_size)
	if item_started_but_no_children!=[]:
		error_worksheet.append_warning("Warning: The following Product Items are not New but haven't got any Implementation Items associated with them.  (Velocity of current or past sprints will likely be under-estimated)")
		error_worksheet.append_list(item_started_but_no_children)
				
				

######################## Main starts here #####################
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

#read the configuration
config=config_class()

# Load the input file
try:
	print("Opening File...")
	input_file = open(rtc_export_file_name, encoding='utf-16') # opens the csv file
except IOError:
	print ('Error. Cannot open', rtc_export_file_name)
	sys.exit(0)

print("Reading File...")
reader = csv.reader(input_file,delimiter='\t')  # creates the reader 
input = [r for r in reader]

# split of header and data.
header=header_class(input[0])
data=[spreadsheet_row(row) for row in input[1:]]
filtered_data=[row for row in data if not row.check_filters()]

# Create the output workbooks
grouped_worksheet = output_worksheet_class('Work Breakdown')	#creates a worksheet for grouping
grouped_worksheet.worksheet.outline_settings(True, False, False, True)	# displays the grouping summary above
iteration_report_worksheet = iteration_report_worksheet_class('Iteration Report')	#creates a iteration_report_worksheet for grouping
iteration_report_worksheet.worksheet.outline_settings(True, False, False, True)	# displays the grouping summary above
ranked_worksheet = output_worksheet_class('Ranked')	#creates a worksheet for ranking
input_worksheet = output_worksheet_class('Input')	#creates a worksheet for regurgitating input
error_worksheet = output_worksheet_error_class('Errors') # workbook for all suspicious data found from input


try:
	print("Create Raw Sheet...")
	input_worksheet.append_header(header.raw_header)
	input_worksheet.append_list(data,format=False,raw=True)	# write to input worksheet
	
	# Prep error sheet
	error_worksheet.append_header(header.full_header)
	
	print("Sorting...")
	filtered_data.sort() #creates ranked list
	
	
	print("Creating Worksheets...")
	grouped_worksheet.append_header(header.full_header,hidden=True)
		
	id_dictionary=id_dictionary_class(filtered_data)
	work_breakdown=work_breakdown_class(filtered_data,id_dictionary)		# creates grouping and writes to grouped worksheet
	work_breakdown.calculate_points()
	work_breakdown.write_to_spreadsheet(grouped_worksheet)
	
	ranked_worksheet.append_header(header.full_header,hidden=True)
	ranked_worksheet.append_list(filtered_data)	# writes to ranked worksheet

	# create iteration report
	iteration_team_report=iteration_team_report_class(filtered_data)
	iteration_team_report.create_iteration_team_report()
	
# Comment out warning reports for now as not refactored to new code.  Will work on adding shirtsize story point mappings.
	# check data consistency and report to Error tab
	wrong_state_report()
	parents_before_children_report()
	if config.product_work_items:
		shirtsizechecks()
	
	
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


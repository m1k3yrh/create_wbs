#
# This Python file contains all the configuration for create_wbs.
# Edit this file to match the configuration of your project.
#
# All configuration is defined within the configured_data dictionary.  Each key defines specific configuration.
# Ensure you use Python syntax when defining content.
# In particular ensure that each item in lists etc has a comma at the end unless it's the last item in the list


configured_data={ # Do not change this line.  All configuration is kept within the configured_data dictionary

#
# Mandatory Configuration.
# -----------------------
#
# The following configuration must be completed.

# Set hyperlink string to point to prefix required to open a work item in your RTC instance.
# (Easiest way to get this URL prefix is to open a work item and copy the URL removing the work item ID from the end
	"Hyperlink":"https://clm.bdc6.aess.accenture.com/ccm1/web/projects/Managed%20Services%20-%20Asset#action=com.ibm.team.workitem.viewWorkItem&id=",

# List all Sprints below in chronological order.
    "Planned For":[
		"FY14Q4 Sprint 2",
		"FY14Q4 Sprint 3",
		"FY14Q4 Sprint 4",
		"FY14-Q4",
		"FY15-Q1-Sprint1",
		"FY15-Q1-Sprint2",
		"FY15-Q1",
		"Backlog"
    ],
    
# List all Priority values here.  Highest priority down to lowest.  RTC uses Priority, then Rank, then ID to determine order of items in backlogs.
	"Priority":[
		"High",
		"Medium",
		"Low",
		"Unassigned"
	],
	
#
# Optional Configuration.
# -----------------------
#
# The following configuration enables additional features.  It can be either modified or removed as desired.
	
# List all formatting below.  "Header" and "Error Title" are special and used to format Header rows and Error descriptions respectively.  
# The rest are used for string matching and will format cells containing the text specified.
# Wildcards are not supported at this stage.
# See http://xlsxwriter.readthedocs.org/en/latest/working_with_formats.html for description how to set formats for below
	"Format":{
# Special row formatting rules
		"Header":{"font":"Times New Roman","size":15,"bold":1,"color":"blue"},
		"Error_Item":{"font":"Times New Roman","size":20,"bold":1},
		"Iteration_Summary":{"font":"Times New Roman","size":16,"bold":1},
		"Category_Summary":{"font":"Times New Roman","size":14,"italic":1},

# String match formatting rules
		"Capability":{"font":"Times New Roman","size":10},
		"Defect":{"font":"Times New Roman","size":10},
		"Epic":{"font":"Times New Roman","size":16,"italic":1},
		"Feature":{"font":"Times New Roman","size":14},
		"Story":{"font":"Times New Roman","size":12},
		"Impeded":{"color":"red"},
		"More Information":{"color":"orange"},
		"Done":{"color":"green"},
		"In Progress":{"color":"#00ff77"},
		"Implementing":{"color":"#00ff77"},
		"Implemented":{"color":"#00ffbb"}
	},
	
# Set the rules for how much value is earned for particular states.  Values should be in range 0.0-1.0 and represent the fraction of total 
# story points for a work item which are considered done when the item reaches this state.  Insert multiple rows with the same multiplier if required
# e.g. Implemented and "Ready for Integration" might both be 0.75 for example.
# Any states not listed are assumed to have a multiplier of 0.0 (i.e. no points are earned).
	"Progress":{
		"In Progress" : 0.25,
		"Implemented" : 0.75,
		"Done" : 1.0
	},


# Average Sizes for unestimated work
	"Average Sizes":{
		"S" : 13,
		"M"	: 30,
		"L" : 100
	},
	
# List the columns to be searched and the values that should be ignored. e.g. '"Status":["Invalid","Rejected"]' means all items with status "Invalid"
# or "Rejected" will be ignored.  
# Ignored items are neither output nor processed (so any points in these items is not counted.
	"Filters": {
		"Status":["Invalid","Rejected"],
		"Type":["Defect","Task"]
	},
	
# List the columns to be hidden in the output excel file.
# This configuration is mainly for tidiness.  The items are output to the file and can be unhidden via the context menus.
	"Hidden": [
		"Parent",
		"Rank (relative to Priority)",
		"Modified Date",
		"Owned By",
		"Filed Against"
	],
	
# List all completed states (Those where all children should have earned 100% of their points.
# Used to check that State of Story, Epic etc corresponds to earned values recorded from children work items.
	"Completed States":["Done","Implemented"],

# List all states in which progress should be at 0%.  If not, then a warning will be added to Errors worksheet.
	"New States":["New"],
	
#
########################################################################################
# Below are extensions designed for Scaled Agile teams (with mWallet as client).
# Comment out or delete to disable these features.
########################################################################################

# Team Categories is list of all Categories that are linked to teams.  All Capabilities should be assigned to one of these	
	"Team Categories":[	'Managed Services - Asset/mWallet/MLV_BL',
						'Managed Services - Asset/mWallet/MLV_CR_Emulators',
						'Managed Services - Asset/mWallet/MLV_GUI',
						'Managed Services - Asset/mWallet/MLV_MobileApps-Android',
						'Managed Services - Asset/mWallet/MLV_MobileApps-iOS',
						'Managed Services - Asset/mWallet/mWallet Design',
						'Managed Services - Asset/mWallet/RTAE'],

# Team Work Items are kinds of Work Items that should be assigned to teams (and not to Product Categories)
	"Team Work Items":['Capability','Task'],
	
# Product Work Items.  These should not be assigned to teams but to Product Categories
	"Product Work Items":['Epic','Feature','Story'],
	
# Product Categories: These are where Product Work Items should be assigned.
	"Product Categories":['Products/mWallet Product'],
	
# Backlog(s) into which a Product Work Items should be assigned.  (Commented out because mWallet prefer to set Story to PlannedFor when last Capabilities are assigned to a Sprint).
#	"Product Backlogs":['Backlog'],
	
# List impeded states here. Used to check that if a parent is impeded all it's children are impeded and vica versa.
	"Impeded States":["Impeded","More Information"],

# Items that should have valid Story Points assigned
	"Story Pointed Work Items":['Capability'],

	
	
} # DO NOT DELETE OR COMMENT THIS LINE.  IT IS END OF CONFIGURATION DATA DEFINITION



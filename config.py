configured_data={
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
	"Priority":[
		"High",
		"Medium",
		"Low",
		"Unassigned"
	],
	
# See http://xlsxwriter.readthedocs.org/en/latest/working_with_formats.html for description how to set formats for below
	"Format":{
		"Header":{"font":"Times New Roman","size":15,"bold":1,"color":"blue"},
		"Error_Item":{"font":"Times New Roman","size":20,"bold":1},
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
	"Progress":{
		"In Progress" : 0.25,
		"Implemented" : 0.75,
		"Done" : 1.0
	},
	"Filters": {
		"Status":["Invalid","Rejected"],
		"Type":["Defect","Task"]
	},
	"Hidden": [
		"Parent",
		"Rank (relative to Priority)",
		"Modified Date",
		"Owned By",
		"Filed Against"
	],
	"Hyperlink":"https://clm.bdc6.aess.accenture.com/ccm1/web/projects/Managed%20Services%20-%20Asset#action=com.ibm.team.workitem.viewWorkItem&id=",

# Team Categories is list of all Categories that are linked to teams.  All Capabilities should be assigned to one of these	
	"Team Categories":['Managed Services - Asset/mWallet/MLV_BL',
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
	"Product Categories":['Products/mWallet Product']
}



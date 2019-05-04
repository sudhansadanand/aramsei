#!/usr/bin/python


import sys
sys.path.append("/Users/susadana/anaconda/lib/python2.7/site-packages")
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import prettytable
from prettytable import PrettyTable

#Google sheets with spark 2019 registrations
#SPARK_SHEET = 'https://docs.google.com/spreadsheets/d/1VSyTMtqdWP43aPEOnjTvj38UcxW1uuhbkfr4ffTlKqI'
SPARK_SHEET = 'https://docs.google.com/spreadsheets/d/1ltvAR-kqmusP3VsyXd3p0iFc97oKTrTk_u6xwUBIMFc'

SPARK_CREDENTIALS_FILE = 'spark-236622-15fe0499d9a3.json'

SparkTeams = []
p_e_table=[] # participant name, [[event,team name]]
EventsList = ["Quiz","World Dance","Spark Tank", "World Music", "Mini-Makers"]
Event_StartTimes = { "Quiz": "14:00",
                     "World Dance": "15:30",
                     "Spark Tank": "13:00",
                     "Mini-Makers": "9:30",
                     "World Music": "9:30"
                    }

class SparkOrder:
    order_num = 0
    participant_names = []
    contact_email = ""
    def __init__(self, o_num, c_email):
        self.order_num = o_num
        self.participant_names = [] 
        self.contact_email = c_email
    def addParticipants(self,Participants):
        self.participant_names.append(Participants)

class SparkTeam:
    TeamName = ""
    Events = []
    p_num = 0
    Category = ""
    time_slot = "" 
    orders = [] #List of each order containing {Order Number, Participant Name List, Contact Email}

    def __init__(self, STeamName, SEvent, SP_num, SCategory):
        self.TeamName = STeamName
        self.Events = [] 
        self.orders = []
        self.p_num = int(SP_num)
        self.Category = SCategory
        self.time_slot = ":" 

    def getevents(self):
        return ",".join(self.Events)

    def addorder(self,Spark_order):
        self.orders.append(Spark_order)

    def addEvent(self,Spark_event):
        #Add the event, only if its not already in here.
        for event in self.Events:
            if(str(Spark_event.strip()) == str(event.strip())):
                return 
        self.Events.append(Spark_event)

    def addp_num(self,p_num):
        self.p_num +=int(p_num)

def generate_spark_teams():
    global SparkTeams
    scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive'] 
    credentials = ServiceAccountCredentials.from_json_keyfile_name(SPARK_CREDENTIALS_FILE, scope)
    gc = gspread.authorize(credentials)
    sheet = gc.open_by_url(SPARK_SHEET)
    spark_sheet = sheet.get_worksheet(1)
    spark_data = spark_sheet.get_all_values()
    hdr_count = 2
    entry_i = 1
    for each_order in spark_data:
        print("order: " + str(each_order))
        p_list=[]
        if entry_i <= hdr_count:
	    entry_i+=1
            continue
     
        #Get order number
        order_num = each_order[0]
        print("order number: " + str(order_num))

        #Get Participant List
        p_list = str(each_order[2].encode("utf-8"))

        #Get contact email list
        email_id = each_order[-3]

        #Create the order Object
        new_order = SparkOrder(order_num,email_id)
        new_order.addParticipants(p_list)

        #Get team name
        team_name = each_order[3]
        print "team name: "+str(team_name)

        #Get Event
        team_event = each_order[-2]
        print "event: "+str(team_event)

        #Get grade
        team_category = each_order[-4]
        print "grade: "+str(team_category)

        #Get num_participants
        print "p_num: ->"+str(each_order[4])+"<-"
        p_num = int(each_order[4])

        print (""+str(order_num)+": {"+str(team_name)+", "+str(team_event)+", "
                +str(p_num)+", "+str(team_category)+", "+str(p_list)+", "+str(email_id)+"}")

        #Check if Team already exists in spark team list
        #If There are no teams or of team doesn't exist, create a new team
        #If team already exists, just add this order to the existing team
        if not SparkTeams:
            new_team = SparkTeam(team_name,team_event,p_num,team_category)
            new_team.addorder(new_order)
            new_team.addEvent(str(team_event))
            SparkTeams.append(new_team)
        else:
            team_exists = 0
            for this_team in SparkTeams:
                if(team_name == this_team.TeamName):
                    team_exists = 1
                    this_team.addorder(new_order)
                    this_team.addEvent(str(team_event))
                    this_team.addp_num(p_num)
                    break
            if(team_exists == 0):
                new_team = SparkTeam(team_name,team_event,p_num,team_category)
                new_team.addorder(new_order)
                new_team.addEvent(str(team_event))
                SparkTeams.append(new_team)

def getParticipatingEvents(p_name):
    e_list = [] 
    category = ""
    for this_team in SparkTeams:
	for each_order in this_team.orders:
	    for this_participant in each_order.participant_names[0].split(","):
                #print "Checking "+str(p_name)+" with "+str(this_participant)
		if(p_name == this_participant.strip()):
                    category = str(this_team.Category)
                    #print "Matched"
		    #e_list.append(""+str(this_team.Events)+","+str(this_team.TeamName)+"")
		    e_list.append(""+this_team.getevents()+","+str(this_team.TeamName)+"")
    return [str(p_name),e_list,str(category)]

def is_participant_already_in_table(p_name):
    global p_e_table # participant name, [[event,team name]]
    if not p_e_table:
        return False
    else:
        for p_i in p_e_table:
            if p_name == p_i[0]:
                return True


#Generate Participant - Event data
#for certificates
def generate_participant_event_data():
    global p_e_table
    participating_events_info = getParticipatingEvents("Kanira Venkat".strip())
    p_e_table.append(participating_events_info)

    for this_team in SparkTeams:
	for each_order in this_team.orders:
	    for this_participant in each_order.participant_names[0].split(","):
		if(is_participant_already_in_table(this_participant.strip())):
		    continue
		#print "This participant: "+str(this_participant)          
		participating_events_info = getParticipatingEvents(this_participant.strip()) 
		p_e_table.append(participating_events_info)

def display_kids_with_multiple_events(): 
    p_i = 1
    for each_participant in p_e_table:
	#print len(each_participant[1])
	if(len(each_participant[1]) > 1):
	    print "S.No: "+str(p_i)+", Category: "+str(each_participant[2])+", Name: "+str(each_participant[0])+", {"
	    p_events = each_participant[1]
	    for event in p_events: 
		print event
	    print "}"
	    print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
	    p_i+=1



#1. Form Spark Teams from Orders
generate_spark_teams()
generate_participant_event_data()

#2.Write orders to google spreadsheet, under each team for each event
print "Displaying order for each team, event wise"

#Open google sheets
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name(SPARK_CREDENTIALS_FILE, scope)
gc = gspread.authorize(credentials)
sheet = gc.open_by_url(SPARK_SHEET)

for this_event in EventsList:
    order_table = PrettyTable()
    order_table.title = str(this_event)+" ---- [Teams - Orders]"
    order_table.hrules = prettytable.ALL
    order_table.field_names = ["Team Name", " Order#", "Participant names", "contact email"]
    for this_team in SparkTeams:
        if(this_event in this_team.Events):
	    order_no = 1
	    for each_order in this_team.orders:
                order_table.add_row([str(this_team.TeamName),str(each_order.order_num),str(each_order.participant_names), str(each_order.contact_email)])
		order_no+=1
    print(order_table)

    #Write it back to google sheets
    # '''
    result = []
    spark_sheet = sheet.worksheet("orders")
    sheet.del_worksheet(spark_sheet)
    spark_sheet = sheet.add_worksheet("orders",100,30)
    spark_sheet = sheet.worksheet("orders")

    for line in str(order_table).splitlines():
        splitdata = line.split("|")
        if len(splitdata) == 1:
            continue  # skip lines with no separators
        linedata = []
        for field in splitdata:
            field = field.strip()
            if field:
                linedata.append(field)
        result.append(linedata)
    #print result
    #Open google sheets
    for each_row in result:
        spark_sheet.append_row(each_row,value_input_option='RAW')
    # '''
    
#3.Spark Teams for Each Event
#Write this data into appropriate spreadsheets in google spreadsheet
print "Displaying Spark Team data, event wise"
s_no = 1
# Print event list
print ("Events: " + str(EventsList))

Eventslist = ["Quiz"]
#Eventslist = ["World Dance"]
#Eventslist = ["Spark Tank"]
#Eventslist = ["World Music"]
#Eventslist = ["Mini-Makers"]

for this_event in EventsList:
    event_table = PrettyTable()
    event_table.field_names = ["Team Order", "Time Slot", "Team Name", "Category", "participant Num", "Participant Names"]
    s_no = 1
    print "Event: "+str(this_event)
    print "========================"
    for this_team in SparkTeams:
        #print this_team.TeamName
        if(this_event in this_team.Events):
            p_list = []
            for each_order in this_team.orders:
                p_list.append(','.join(str (eo) for eo in each_order.participant_names))
                 
            event_table.add_row([str(s_no), str(this_team.time_slot), str(this_team.TeamName), str(this_team.Category), str(this_team.p_num), ','.join(str(p) for p in p_list)])
            #print str(s_no)+" <"+str(this_team.TeamName)+">"
            s_no +=1
    print(event_table)

    #Write it back to google sheets
    # '''
    result = []

    for line in str(event_table).splitlines():
        splitdata = line.split("|")
        if len(splitdata) == 1:
            continue  # skip lines with no separators
        linedata = []
        for field in splitdata:
            field = field.strip()
            if field:
                linedata.append(field)
        result.append(linedata)
    #print result
    #Open google sheets
    scope = ['https://spreadsheets.google.com/feeds']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(SPARK_CREDENTIALS_FILE, scope)
    gc = gspread.authorize(credentials)
    sheet = gc.open_by_url(SPARK_SHEET)
    spark_sheet = sheet.worksheet(str(this_event))
    sheet.del_worksheet(spark_sheet)
    spark_sheet = sheet.add_worksheet(str(this_event),100,30)
    spark_sheet = sheet.worksheet(str(this_event))
    for each_row in result:
        spark_sheet.append_row(each_row,value_input_option='RAW')      
    # '''

#Print Participant Events data
participant_table = PrettyTable()
participant_table.hrules = prettytable.ALL
participant_table.field_names = ["SNo","Name","Category","{Event,Team} list"]
p_i = 1
for each_participant in p_e_table:
    #print len(each_participant[1])
    #print "S.No: "+str(p_i)+", Category: "+str(each_participant[2])+", Name: "+str(each_participant[0])+", {"
    p_events = each_participant[1]
    participant_table.add_row([str(p_i),str(each_participant[0]),str(each_participant[2]),p_events])
    #for event in p_events: 
    #    print event
    #	print "}"
    #print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    p_i+=1
print(participant_table)
#Write it back to google sheets
# '''
result = []

for line in str(participant_table).splitlines():
    splitdata = line.split("|")
    if len(splitdata) == 1:
        continue  # skip lines with no separators
    linedata = []
    for field in splitdata:
        field = field.strip()
        if field:
            linedata.append(field)
    result.append(linedata)
#print result
#Open google sheets
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name(SPARK_CREDENTIALS_FILE, scope)
gc = gspread.authorize(credentials)
sheet = gc.open_by_url(SPARK_SHEET)
spark_sheet = sheet.worksheet("participant-events")
sheet.del_worksheet(spark_sheet)
spark_sheet = sheet.add_worksheet("participant-events",100,30)
spark_sheet = sheet.worksheet("participant-events")
for each_row in result:
    spark_sheet.append_row(each_row,value_input_option='RAW')
# '''

#5. Display Participants with Multiple Events
display_kids_with_multiple_events()
#Fill up time slots for each team per event 





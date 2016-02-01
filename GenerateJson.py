# GenerateJson.py: Create a json file from agile project server, Rally at the moment
#
# Usage:
#   python.exe GenerateJson.py --config=user.cfg sprint_name 
#
#   e.g. 
#       python.exe GenerateJson.py --config=user.cfg "Sprint 40[FY16]
#
# History:
#       Init                2/1/2016
#


import sys
import json

from array import array
from pyral import Rally, rallyWorkset

options = [arg for arg in sys.argv[1:] if arg.startswith('--')]
args    = [arg for arg in sys.argv[1:] if arg not in options]

server, user, password, apikey, workspace, project = rallyWorkset(options)

print "%s %s" % (server, apikey)
print "%s %s" % (workspace, project)
print ""

if apikey:
    rally = Rally(server, apikey=apikey, workspace=workspace, project=project)
else:    
    rally = Rally(server, user=user, password=password, workspace=workspace, project=project)
rally.enableLogging('mypyral.log')

#projects = ['Install: Civil 3D']
projects = ['Install: Civil 3D', 'Install: Plant 3D', 'Continuous Delivery', 'Team Vela']
projects_dict = {'Install: Civil 3D': 'Civil', 'Install: Plant 3D': 'Plant', 'Continuous Delivery': 'CI/CD', 'Team Vela': 'Install Automation'}

pre_us_hl = "https://rally1.rallydev.com/#/26211768400d/detail/userstory/"
pre_defect_hl = "https://rally1.rallydev.com/#/26211768400d/detail/defect/"

items_ist = []
items_coverspy = []

# create query 
def query_US(iteration_name):

    for project in projects:
        rally.setProject(project)
        print "project: %s" % (project)

        project_abbr = projects_dict[project]

        if project == 'Continuous Delivery':
            data[project_abbr] = []
        else:
            data['Install'][project_abbr] = []

        items = []
        # User Story
        for sty in rally.get('UserStory', query='Iteration.Name = ' + '"' + iteration_name + '"'):
            if sty.Project.Name == project:
                print "%s %s" % (sty.FormattedID, sty.Name)
                link = pre_us_hl + sty.oid
                # Special cases for IST and CoverSpy
                if sty.Name.lower().startswith('[ist]'):
                    items_ist.append([sty.FormattedID, sty.Name, link])
                elif sty.Name.lower().startswith('[coverspy]'):
                    items_coverspy.append([sty.FormattedID, sty.Name, link])
                else:
                    items.append([sty.FormattedID, sty.Name, link])

        # Defect
        for defect in rally.get('Defect', query='Iteration.Name = ' + '"' + iteration_name + '"'):
            if defect.Project.Name == project:
                print "%s %s" % (defect.FormattedID , defect.Name)
                #print "%s%s" %(pre_defect_hl, defect.oid)
                link = pre_defect_hl + defect.oid
                # Special cases for IST and CoverSpy
                if defect.Name.lower().startswith('[ist]'):
                    items_ist.append([defect.FormattedID, defect.Name, link])
                elif defect.Name.lower().startswith('[coverspy]'):
                    items_coverspy.append([defect.FormattedID, defect.Name, link])
                else:
                    items.append([defect.FormattedID, defect.Name, link])

        if project == 'Continuous Delivery':
            data[project_abbr] = items
        else:
            data['Install'][project_abbr] = items

        if len(items_ist) > 0:
            data['Install']['IST'] = items_ist

        if len(items_coverspy) > 0:
            data['CoverSpy'] = items_coverspy

        print ""

data = {}
data['Install'] = {}

# Fetch user story and defect in a sprint
query_US(args[0])

#print data

# write to json file
with open('data.json', 'w') as fp:
    json.dump(data, fp)

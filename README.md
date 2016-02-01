# SprintReport
A tool to create Sprint report table

2 Python script files and 1 bat file are included:
### GenerateJson.py
Create a json file from agile project server, Rally at the moment
'''
Usage:
  python.exe GenerateJson.py --config=user.cfg sprint_name 

  e.g. 
      python.exe GenerateJson.py --config=user.cfg "Sprint 40[FY16]
'''

### Convert.py: 
convert json file to a html file with a table in sprint report
Table example: Template1.orig.htm
'''
Usage:
  python.exe convert.py jsonfile.json
Output:
  A htm file wiht the table
'''

### CreateSpirntTable.bat
Call above 2 Python scripts and generate the sprint table
'''
Usage:
CreateSprintTable.bat Sprint_Name
e.g.
     CreateSprintTable.bat "Sprint 40[FY16]" 
Output:
     data.htm
'''

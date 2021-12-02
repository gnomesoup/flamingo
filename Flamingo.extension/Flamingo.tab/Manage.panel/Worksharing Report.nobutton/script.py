import codecs
from datetime import datetime, timedelta
from flamingo.excel import OpenWorkbook, GetWorksheetData
from glob import glob
from os import path
from pyrevit import forms, script
import re
import shutil
import tempfile
import uuid

def TimestampAsDateTime(string):
    return datetime.strptime(
        string[10:29],
        format="%Y-%m-%d %H:%M:%S"
    )

def OutputMD(output, string, printReport=True):
    if printReport:
        output.print_md(string)

def GetWorksharingReport(output, printReport):
    fd, tempPath = tempfile.mkstemp()
    shutil.copy2(slogPath, tempPath)

    sessions = {}
    currentSessionId = None
    minStart = datetime.max
    maxEnd = datetime.now()
    with codecs.open(tempPath, "r", encoding="UTF-16") as f:
        for line in f:
            if line.startswith(" ") and currentSessionId:
                m = re.match(r" (.+)=.*\"(.*)\"", line)
                if m:
                    sessions[currentSessionId][m.group(1)] = m.group(2)
            elif re.findall(r">Session", line):
                currentSessionId = line[0:9]
                startDateTime = TimestampAsDateTime(line)
                sessions[currentSessionId] = {"start": startDateTime}
                minStart = startDateTime if startDateTime < minStart \
                    else minStart
            elif re.findall(r"<Session", line) and currentSessionId:
                currentSessionId = line[0:9]
                endDateTime = TimestampAsDateTime(line)
                sessions[currentSessionId]['end'] = endDateTime
            elif re.findall(r"<STC[^:]", line):
                sessions[currentSessionId]['lastSync'] = TimestampAsDateTime(
                    line
                )
            elif line[0:9] in sessions:
                sessions[currentSessionId]['lastActive'] = TimestampAsDateTime(
                    line
                )

    for session in sessions:
        if "end" in sessions[session]:
            endDateTime = sessions[session]['end']
        else:
            endDateTime = datetime.now()
            if "lastSync" in sessions[session]:
                sessions[session]['timeSinceLastSync'] = \
                    datetime.now() - sessions[session]['lastSync']
        sessions[session]['sessionLength'] = \
            endDateTime - sessions[session]['start']

    activeSessions = [
        session for session, sessionInfo in sessions.items()
        if "end" not in sessionInfo
    ]
    sortedSessionKeys = sorted(
        sessions,
        key=lambda x: sessions[x]['start']
    )
    OutputMD(output, "# Worksharing Log Info", printReport)
    OutputMD(output, "Active: {}".format(maxEnd - minStart), printReport)
    OutputMD(output, "central: {}".format(list(sessions.values())[0]['central']), printReport)
    OutputMD(output, "Total Sessions: {}".format(len(sessions)), printReport)
    OutputMD(output, "Active Sessions: {}".format(len(activeSessions)), printReport)
    OutputMD(output, "", printReport)
    OutputMD(output, "# Sessions", printReport)
    keysToPrint = {
        "build": "Revit Version",
        "host": "Computer Name",
        "end": "Session End",
        "sessionLength": "Session Length",
        "lastActive": "Last Activity",
        "lastSync": "Last Sync To Central",
        "timeSinceLastSync": "Time Since Last Sync",
    }
    
    for session in sortedSessionKeys:
        values = sessions[session]
        icon = ":white_heavy_check_mark:" if "end" in values else ":cross_mark:"
        output.print_md(
            "## {}: {} {}".format(values['user'], values['start'], icon)
        )
        OutputMD(output, "id: {}".format(session), printReport)
        for key, title in keysToPrint.items():
            if key in values:
                OutputMD(output, "{}: {}".format(title, values[key]), printReport)
        OutputMD(output, "", printReport)

    return sessions

if __name__ == "__main__":
    output = script.get_output()

    workbook = OpenWorkbook("C:/Users/michael_p/Downloads/HiveProjectsList-8f6b78ba-9307-4bba-b93f-0907e0d6e30b.xlsx")
    excelData = GetWorksheetData(workbook.Sheets[1])
    projectData = []
    headers = None
    for i in range(excelData['rowCount']):
        row = excelData[i+1]['data']
        if not headers:
            headers = row
            continue
        data = { header: value for header, value in zip(headers, row) }
        projectData.append(data)

        
    models = {}
    projects = {}
    for project in projectData:
        projectId = uuid.uuid4()
        projectNumber = project['ProjectNumber']
        if projectNumber not in projects:
            projects[projectNumber] = {
                "projectId": projectId,
                "projectNumber": projectNumber,
                "projectName": project["Project Name"],
            }
        modelPaths = set(
            [
                modelPath.replace("\\\\Server\\Revit", "N:\\")
                for modelPath in (project['Filepaths']).split(",")
            ]
        )
        for modelPath in modelPaths:
            print(modelPath)
            projects[projectNumber]["nPath"] = modelPath
            models[modelPath] = {
                "projectId": projectId,
                "modelId": uuid.uuid4(),
                "projectNumber": projectNumber,
                "modelPath": modelPath
            }

    print("models:")
    for path, projectInfo in models.items():
        print("  - modelId: {}".format(uuid.uuid4()))
        print("    modelPath: {}".format(projectInfo["modelPath"]))
        print("    projectId: {}".format(
            projects[projectInfo["projectNumber"]]["projectId"]
        ))
    
    print("projects:")
    for projectNumber, project in projects.items():
        print("  - projectId: {}".format(project["projectId"]))
        print("    projectName: {}".format(project["projectName"]))
        print("    projectNumber: {}".format(project["projectNumber"]))
        print("    nPath: {}".format(project["nPath"]))
        
    
    script.exit()
    backupFolderPath = forms.SelectFromList.show(
        models,
        title="Select model to report",
        button_name="Run Report",
        multiselect=False,
        info_panel=True,
        checked_only=True,
        width=400,
        height=750,
        item_template=forms.utils.load_ctrl_template(
            script.get_bundle_file("WorksharingReportTemplate.xml")
        )
    )
    script.exit()
    # backupFolderPath = "N:/2006 GSA Silvio Mollo Federal Building/01_Central Files/01_Javits/NY0282ZZ_A_37-38_r2021_backup"
    revitModelPath = forms.pick_file(
        file_ext="rvt",
        init_dir="N:\\",
        title="Pick a revit file to report"
    )
    revitModel, x = path.splitext(revitModelPath)
    print(revitModel)
    backupFolderPath = revitModel + "_backup"
    if path.exists(backupFolderPath):
    #     output.print_md(listdir(backupFolderPath, ".slog"))
        slogPaths = glob(backupFolderPath + "/*.slog")
        if len(slogPaths) > 1:
            forms.alert(
                "Multiple slog files found. Createing report off of the fist file"
                " This model should be recentralized to cleanup the backup folder."
            )
        elif not slogPaths:
            forms.alert(
                "No worksharing log file could be found in the specified folder",
                exitscript=True
            )
        slogPath = slogPaths[0]
    else:
        forms.alert(
            "The workshare folder for the specified model does not exist.",
            exitscript=True
        )

    GetWorksharingReport(output, printReport=True)
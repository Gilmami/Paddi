import word
import os
import re
import shutil
from pdga_survey import get_pdga_data_survey

# gets list of schools from the completed reports location on google drive
# (windows only)


def get_schools(country, state):
    """Enumerates the existing names of directories in a path and adds them to
    a list, appends New to that list."""
    school_list = list()
    try:
        path = os.path.normpath(r"G:\Shared Drives\EDU Consulting - Completed Reports"
                                + "\\" + country + "\\" + state)
        school_list = os.listdir(path)
        school_list.sort()
        school_list
        return(school_list)
    except FileNotFoundError:
        school_list.append("New")
        return(school_list)


def make_proj_tree(country, state, school, year, proj, campuses=None):
    """Creates a path based on country, state, school year and proj variables,
    and populates subdirectories underneath based on proj value."""
    print(campuses)
    if campuses is None:
        path = os.path.normpath("G:\\Shared Drives\\EDU Consulting - Completed Reports"
                                + "\\" + country + "\\" + state + "\\" + school + "\\" + proj + "\\" + year)
        if proj in ["EIS", "PSEIS", "VoD", "Capital"]:
            fileNames = [r"\Drafts\Spode", r"\Finals\Spode", r"\Finals\PressPacket", "\Admin", "\ForDesign",
                         r"\Data\Spode"]
            for i in fileNames:
                os.makedirs(path + i, exist_ok=True)

        elif proj == "PDGA":
            fileNames = [r"\Drafts\Spode",
                         r"\Finals\Spode", "\Admin", r"\Data\Spode"]
            for i in fileNames:
                os.makedirs(path + i, exist_ok=True)
    else:
        for campus in campuses:
            path = os.path.normpath("G:\\Shared Drives\\EDU Consulting - Completed Reports"
                                    + "\\" + country + "\\" + state + "\\" + school + "\\" + proj + "\\" + year)
            if proj in ["EIS", "PSEIS", "VoD", "Capital"]:
                fileNames = [f"\\{campus}\\Drafts\\Spode",
                             f"\\{campus}\\Finals\\Spode",
                             f"\\{campus}\\Finals\\PressPacket",
                             f"\\{campus}\\Admin", f"\\{campus}\\ForDesign",
                             f"\\{campus}\\Data\\Spode"]
                for i in fileNames:
                    os.makedirs(path + i, exist_ok=True)

            elif proj == "PDGA":
                fileNames = [f"\\{campus}\\Drafts\\Spode"
                             f"\\{campus}\\Finals\\Spode", f"\\{campus}\\Admin",
                             f"\\{campus}\\Data\\Spode"]
                for i in fileNames:
                    os.makedirs(path + i, exist_ok=True)

# make sure to add in .agg folder in agg parent directory so we have a place to
# put all the agg things.


def make_agg_tree(country, state, aggName, schoolList, year, proj):
    """Creates a path and directory tree for aggregate studies that follows a
    similar format to make_proj_tree."""
    path = os.path.normpath("G:\\Shared Drives\\EDU Consulting - Completed Reports"
                            + "\\" + country + "\\" + state + "\\" + aggName)
    os.makedirs(path+r"\Admin", exist_ok=True)
    os.makedirs(path+r"\.Agg", exist_ok=True)

    for school in schoolList:
        path = os.path.normpath("G:\\Shared Drives\\EDU Consulting - Completed Reports"
                                + "\\" + country + "\\" + state + "\\" + aggName +
                                "\\" + school + "\\" + proj + "\\" + year)

        if proj == "EIS":
            fileNames = [r"\Drafts\Spode", r"\Finals\Spode", "\ForDesign",
                         r"\Data\Spode", r"\Finals\PressPacket"]
            for i in fileNames:
                os.makedirs(path + i, exist_ok=True)

        elif proj == "PDGA":
            fileNames = [r"\Drafts\Spode",
                         r"\Finals\Spode", "\Admin", r"\Data\Spode"]
            for i in fileNames:
                os.makedirs(path + i, exist_ok=True)


def import_admin(country, state, school, year, acronym, projType, schoolType,
                 schoolID=None, dataRun=None, completionsBaseYear=None,
                 campuses=None):
    """Imports data survey, timeline, and utility for each unique project type."""
    projList = ["EIS", "PSEIS", "VoD", "Capital"]

    make_proj_tree(country, state, school, year, projType, campuses)

    USUtilityRegex = re.compile(r"EIS_Utility.*.xlsx")
    USUtilityLoc = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates"

    timeOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\Project_timeline.xlsx"

    CANUtilityRegex = re.compile(r"CAN_Utility.*.xlsx")
    CANUtilityLoc = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\Canada"

    timePath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + \
        school + "\\" + projType + "\\" + year + "\\" + "Admin" + \
        "\\" + acronym + "_Timeline_" + year + ".xlsx"

    if not os.path.exists(timePath):
        os.makedirs(os.path.split(timePath)[0], exist_ok=True)
        shutil.copy(timeOrig, timePath)
    else:
        pass

    if campuses is None:
        if projType == "EIS" and country == "US":
            USUtilityPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + \
                "\\" + school + "\\" + projType + "\\" + year + "\\" + \
                "Data" + "\\" + acronym + "_Utility_" + year + ".xlsx"

            if not os.path.exists(USUtilityPath):
                os.chdir(USUtilityLoc)
                matchReCopy("Utility", USUtilityRegex, '.xlsx', 'Data', country, state, school,
                            projType, acronym, year)
            else:
                pass

        elif projType == "VoD":
            VoDUtilityOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\Value of Degree\VoD_Utility.xlsx"
            VoDSurvOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\Value of Degree\VoD_DataSurvey.xlsx"
            VoDUtilityPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + \
                "\\" + school + "\\" + projType + "\\" + year + "\\" + \
                "Data" + "\\" + acronym + "_Utility_" + year + ".xlsx"
            VoDSurvPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + \
                school + "\\" + projType + "\\" + year + "\\" + "Data" + \
                "\\" + acronym + "_Data_Survey_" + year + ".xlsx"

            if not os.path.exists(VoDUtilityPath):
                shutil.copy(VoDUtilityOrig, VoDUtilityPath)
            else:
                pass
            if not os.path.exists(VoDUtilityPath):
                shutil.copy(VoDSurvOrig, VoDUtilityPath)
            else:
                pass

        elif projType == "Capital":
            CapUtilityOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\Capital\Capital_Utility.xlsx"
            CapSurvOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\Capital\Capital_DataSurvey_Lightcast.xlsx"
            CapUtilityPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + \
                "\\" + school + "\\" + projType + "\\" + year + "\\" + \
                "Data" + "\\" + acronym + "_Utility_" + year + ".xlsx"
            CapSurvPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + \
                school + "\\" + projType + "\\" + year + "\\" + "Data" + \
                "\\" + acronym + "_Data_Survey_" + year + ".xlsx"

            if not os.path.exists(CapUtilityPath):
                shutil.copy(CapUtilityOrig, CapUtilityPath)
            else:
                pass
            if not os.path.exists(CapSurvPath):
                shutil.copy(CapSurvOrig, CapSurvPath)
            else:
                pass

        elif projType == "PSEIS":
            PSUtilityOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\PSEIS\PSEIS_Utility_WorkingVersion.xlsx"

            PSSurvOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\PSEIS\PSEIS_DataSurvey_Lightcast.xlsx"

            PSUtilityPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + \
                "\\" + school + "\\" + projType + "\\" + year + "\\" + \
                "Data" + "\\" + acronym + "_Utility_" + year + ".xlsx"

            PSSurvPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + \
                school + "\\" + projType + "\\" + year + "\\" + "Data" + \
                "\\" + acronym + "_Data_Survey_" + year + ".xlsx"

            if not os.path.exists(PSUtilityPath):
                shutil.copy(PSUtilityOrig, PSUtilityPath)
            else:
                pass

            if not os.path.exists(PSSurvPath):
                shutil.copy(PSSurvOrig, PSSurvPath)
            else:
                pass

        elif projType == "EIS" and country == "CAN":
            CANUtilityPath = r"G:\Shared Drives\EDU Consulting - Private\Completed Reports" + "\\" + country + "\\" + state + \
                "\\" + school + "\\" + projType + "\\" + year + "\\" + \
                "Data" + "\\" + acronym + "_Utility_" + year + ".xlsx"

            if not os.path.exists(CANUtilityPath):
                os.chdir(CANUtilityLoc)
                matchReCopy("Utility", CANUtilityRegex, '.xlsx', 'Data', country, state, school,
                            projType, acronym, year)
            else:
                pass

        if projType == "EIS" and schoolType == "College" and country == "US":
            colSurvLoc = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\College"
            colSurvOrigRegex = re.compile(r"College_DataSurvey.*")

            colSurvPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + \
                school + "\\" + projType + "\\" + year + "\\" + "Data" + \
                "\\" + acronym + "_DataSurvey_" + year + ".xlsx"

            if not os.path.exists(colSurvPath):
                os.chdir(colSurvLoc)
                matchReCopy("Data_Survey", colSurvOrigRegex, '.xlsx',
                            'Data', country, state, school, projType, acronym, year)
            else:
                pass

        elif projType == "EIS" and schoolType == "University-Public" and country == "US":
            uniSurvLoc = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\University"
            uniSurvOrigRegex = re.compile(r"University_DataSurvey.*")

            uniSurvPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + \
                school + "\\" + projType + "\\" + year + "\\" + "Data" + \
                "\\" + acronym + "_DataSurvey_" + year + ".xlsx"

            if not os.path.exists(uniSurvPath):
                os.chdir(uniSurvLoc)
                matchReCopy("Data_Survey", uniSurvOrigRegex, '.xlsx', 'Data', country, state, school,
                            projType, acronym, year)
            else:
                pass

            uniSurvMedOrigRegex = r"University_DataSurvey_MedicalTab_VignetteTab.*"
            uniSurvMedLoc = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\University"

            uniSurvMedPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + \
                school + "\\" + projType + "\\" + year + "\\" + "Data" + \
                "\\" + acronym + "_DataSurvey_MedicalTab_" + year + ".xlsx"

            if not os.path.exists(uniSurvMedPath):
                os.chdir(uniSurvMedLoc)
                matchReCopy("Data_Survey", uniSurvMedOrigRegex, '.xlsx', 'Data', country, state, school,
                            projType, acronym, year, aggName=aggName, campuses=campuses)
            else:
                pass

        elif projType == "EIS" and schoolType == "University-Private" and country == "US":
            uniSurvOrigRegex = r"University_DataSurvey_Private.*"
            uniSurvLoc = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\University"

            uniSurvPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + \
                school + "\\" + projType + "\\" + year + "\\" + "Data" + \
                "\\" + acronym + "_DataSurvey_" + year + ".xlsx"

            if not os.path.exists(uniSurvOrig):
                os.chdir(uniSurvLoc)
                matchReCopy("Data_Survey", uniSurvOrigRegex, '.xlsx', 'Data', country, state, school,
                            projType, acronym, year, aggName=aggName, campuses=campuses)
            else:
                pass

        elif projType == "EIS" and country == "CAN" and schoolType == "College":
            colSurvOrigRegex = r"CanadaCollege_DataSurvey_Lightcast.*"
            colSurvLoc = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\Canada"

            colSurvOrigIntRegex = r"CanadaCollege_DataSurvey_International_Lightcast.*.xlsx"
            colSurvIntLoc = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\Canada"

            colSurvPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + \
                school + "\\" + projType + "\\" + year + "\\" + "Data" + \
                "\\" + acronym + "_DataSurvey_" + year + ".xlsx"

            colSurvIntPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + \
                school + "\\" + projType + "\\" + year + "\\" + "Data" + "\\" + \
                acronym + "_DataSurvey_International_" + year + ".xlsx"

            if not os.path.exists(colSurvPath):
                os.chdir(colSurvLoc)
                matchReCopy("Data_Survey", colSurvOrigRegex, '.xlsx', 'Data', country, state, school,
                            projType, acronym, year, aggName=aggName, campuses=campuses)
            else:
                pass

            if not os.path.exists(colSurvIntPath):
                os.chdir(colSurvIntLoc)
                matchReCopy("Data_Survey", colSurvOrigIntRegex, '.xlsx', 'Data', country, state, school,
                            projType, acronym, year, aggName=aggName, campuses=campuses)
            else:
                pass

        elif projType == "EIS" and country == "CAN" and schoolType == "University":
            uniSurvOrigRegex = r"CanadaUniversity_DataSurvey.*"
            uniSurvLoc = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\Canada"

            uniSurvOrigIntRegex = r"CanadaUniversity_DataSurvey_International.*"
            uniSurvIntLoc = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\Canada"

            uniSurvPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + \
                school + "\\" + projType + "\\" + year + "\\" + "Data" + \
                "\\" + acronym + "_DataSurvey_" + year + ".xlsx"

            uniSurvIntPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + \
                school + "\\" + projType + "\\" + year + "\\" + "Data" + "\\" + \
                acronym + "_DataSurvey_International_" + year + ".xlsx"

            if not os.path.exists(uniSurvPath):
                os.chdir(uniSurvLoc)
                matchReCopy("Data_Survey", uniSurvOrigRegex, '.xlsx', 'Data', country, state, school,
                            projType, acronym, year, aggName=aggName, campuses=campuses)
            else:
                pass

            if not os.path.exists(uniSurvIntPath):
                os.chdir(uniSurvIntLoc)
                matchReCopy("Data_Survey", uniSurvOrigIntRegex, '.xlsx', 'Data', country, state, school,
                            projType, acronym, year)
            else:
                pass

        elif projType == "PDGA":
            pdgaDataPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + \
                school + "\\" + projType + "\\" + year + "\\" + "Data" + \
                "\\" + acronym + "_DataSurvey_" + year + ".xlsx"

            if not os.path.exists(pdgaDataPath):
                get_pdga_data_survey(schoolID, dataRun, completionsBaseYear,
                                     pdgaDataPath)
            else:
                pass

        else:
            pass

    else:
        for campus in campuses:
            if projType == "EIS" and country == "US":
                USUtilityPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + \
                    school + "\\" + projType + "\\" + year + "\\" + campus + \
                    "\\" + "Data" + "\\" + campus + "_Utility_" + year + ".xlsx"

                if not os.path.exists(USUtilityPath):
                    shutil.copy(USUtilityOrig, USUtilityPath)
                else:
                    pass

            elif projType == "VoD":
                VoDUtilityOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\Value of Degree\VoD_Utility.xlsx"
                VoDSurvOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\Value of Degree\VoD_DataSurvey.xlsx"
                VoDUtilityPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + \
                    school + "\\" + projType + "\\" + year + "\\" + campus + \
                    "\\" + "Data" + "\\" + campus + "_Utility_" + year + ".xlsx"
                VoDSurvPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + school + \
                    "\\" + projType + "\\" + year + "\\" + campus + "\\" + \
                    "Data" + "\\" + campus + "_Data_Survey_" + year + ".xlsx"

                if not os.path.exists(VoDUtilityPath):
                    shutil.copy(VoDUtilityOrig, VoDUtilityPath)
                else:
                    pass
                if not os.path.exists(VoDUtilityPath):
                    shutil.copy(VoDSurvOrig, VoDUtilityPath)
                else:
                    pass

            elif projType == "Capital":
                CapUtilityOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\Capital\Capital_Utility.xlsx"
                CapSurvOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\Capital\Capital_DataSurvey_Lightcast.xlsx"
                CapUtilityPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + \
                    school + "\\" + projType + "\\" + year + "\\" + campus + \
                    "\\" + "Data" + "\\" + campus + "_Utility_" + year + ".xlsx"
                CapSurvPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + school + \
                    "\\" + projType + "\\" + year + "\\" + campus + "\\" + \
                    "Data" + "\\" + campus + "_Data_Survey_" + year + ".xlsx"

                if not os.path.exists(CapUtilityPath):
                    shutil.copy(CapUtilityOrig, CapUtilityPath)
                else:
                    pass
                if not os.path.exists(CapSurvPath):
                    shutil.copy(CapSurvOrig, CapSurvPath)
                else:
                    pass

            elif projType == "PSEIS":
                PSUtilityOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\PSEIS\PSEIS_Utility_WorkingVersion.xlsx"

                PSSurvOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\PSEIS\PSEIS_DataSurvey_Lightcast.xlsx"

                PSUtilityPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + \
                    school + "\\" + projType + "\\" + year + "\\" + campus + \
                    "\\" + "Data" + "\\" + campus + "_Utility_" + year + ".xlsx"

                PSSurvPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + school + \
                    "\\" + projType + "\\" + year + "\\" + campus + "\\" + \
                    "Data" + "\\" + campus + "_Data_Survey_" + year + ".xlsx"

                if not os.path.exists(PSUtilityPath):
                    shutil.copy(PSUtilityOrig, PSUtilityPath)
                else:
                    pass

                if not os.path.exists(PSSurvPath):
                    shutil.copy(PSSurvOrig, PSSurvPath)
                else:
                    pass

            elif projType == "EIS" and country == "CAN":
                CANUtilityPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + \
                    school + "\\" + projType + "\\" + year + "\\" + campus + \
                    "\\" + "Data" + "\\" + campus + "_Utility_" + year + ".xlsx"

                if not os.path.exists(CANUtilityPath):
                    shutil.copy(CANUtilityOrig, CANUtilityPath)
                else:
                    pass

            if projType == "EIS" and schoolType == "College" and country == "US":
                colSurvOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\College\College_DataSurvey_w Instructions.xlsx"

                colSurvPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + school + \
                    "\\" + projType + "\\" + year + "\\" + campus + "\\" + \
                    "Data" + "\\" + campus + "_DataSurvey_" + year + ".xlsx"

                if not os.path.exists(colSurvPath):
                    shutil.copy(colSurvOrig, colSurvPath)
                else:
                    pass

            elif projType == "EIS" and schoolType == "University-Public" and country == "US":
                uniSurvOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\University\University_DataSurvey.xlsx"

                uniSurvPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + school + \
                    "\\" + projType + "\\" + year + "\\" + campus + "\\" + \
                    "Data" + "\\" + campus + "_DataSurvey_" + year + ".xlsx"

                if not os.path.exists(uniSurvPath):
                    shutil.copy(uniSurvOrig, uniSurvPath)
                else:
                    pass

                uniSurvMedOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\University\University_DataSurvey_Private.xlsx"

                uniSurvMedPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + school + \
                    "\\" + projType + "\\" + year + "\\" + campus + "\\" + "Data" + \
                    "\\" + campus + "_DataSurvey_MedicalTab_" + year + ".xlsx"

                if not os.path.exists(uniSurvMedPath):
                    shutil.copy(uniSurvMedOrig, uniSurvMedPath)
                else:
                    pass

            elif projType == "EIS" and schoolType == "University-Private" and country == "US":
                uniSurvOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\University\University_DataSurvey_Private.xlsx"

                uniSurvPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + school + \
                    "\\" + projType + "\\" + year + "\\" + campus + "\\" + \
                    "Data" + "\\" + campus + "_DataSurvey_" + year + ".xlsx"

                uniSurvMedOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\University\University_DataSurvey_Private.xlsx"

                uniSurvMedPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + school + \
                    "\\" + projType + "\\" + year + "\\" + campus + "\\" + "Data" + \
                    "\\" + campus + "_DataSurvey_MedicalTab_" + year + ".xlsx"

                if not os.path.exists(uniSurvMedPath):
                    shutil.copy(uniSurvMedOrig, uniSurvMedPath)
                else:
                    pass

            elif projType == "EIS" and country == "CAN" and schoolType == "College":
                colSurvOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\Canada\CanadaCollege_DataSurvey_Lightcast.xlsx"

                colSurvOrigInt = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\Canada\CanadaCollege_DataSurvey_International_Lightcast.xlsx"

                colSurvPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + school + \
                    "\\" + projType + "\\" + year + "\\" + campus + "\\" + \
                    "Data" + "\\" + campus + "_DataSurvey_" + year + ".xlsx"

                colSurvIntPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + school + \
                    "\\" + projType + "\\" + year + "\\" + campus + "\\" + "Data" + \
                    "\\" + campus + "_DataSurvey_International_" + year + ".xlsx"

                if not os.path.exists(colSurvPath):
                    shutil.copy(colSurvOrig, colSurvPath)
                else:
                    pass

                if not os.path.exists(colSurvIntPath):
                    shutil.copy(colSurvOrigInt, colSurvIntPath)
                else:
                    pass

            elif projType == "EIS" and country == "CAN" and schoolType == "University":
                uniSurvOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\Canada\CanadaUniversity_DataSurvey_Lightcast.xlsx"

                uniSurvOrigInt = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\Canada\CanadaUniversity_DataSurvey_International_Lightcast.xlsx"

                uniSurvPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + school + \
                    "\\" + projType + "\\" + year + "\\" + campus + "\\" + \
                    "Data" + "\\" + campus + "_DataSurvey_" + year + ".xlsx"

                uniSurvIntPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + school + \
                    "\\" + projType + "\\" + year + "\\" + campus + "\\" + "Data" + \
                    "\\" + campus + "_DataSurvey_International_" + year + ".xlsx"

                if not os.path.exists(uniSurvPath):
                    shutil.copy(uniSurvOrig, uniSurvPath)
                else:
                    pass

                if not os.path.exists(uniSurvIntPath):
                    shutil.copy(uniSurvOrigInt, uniSurvIntPath)
                else:
                    pass

            elif projType == "PDGA":
                pdgaDataPath = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" + state + "\\" + school + \
                    "\\" + projType + "\\" + year + "\\" + campus + "\\" + \
                    "Data" + "\\" + campus + "_DataSurvey_" + year + ".xlsx"

                if not os.path.exists(pdgaDataPath):
                    get_pdga_data_survey(schoolID, dataRun, completionsBaseYear,
                                         pdgaDataPath)
                else:
                    pass

            else:
                pass


def import_agg_admin(country, state, aggName, year):
    """import_admin but for aggs. imports timeline only."""
    USUtilityOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\EIS_Utility.xlsx"
    timeOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\US\Project_timeline.xlsx"
    CANUtilityOrig = r"G:\Shared drives\EDU Consulting - Private\EIS survey templates\Canada\CAN_Utility.xlsx"

    adminPath = os.path.normpath("G:\\Shared Drives\\EDU Consulting - Completed Reports" +
                                 "\\" + country + "\\" + state + "\\" + aggName + "\\" + "Admin")
    if not os.path.exists(adminPath + "\\" + aggName + "_Timeline_" + year + ".xlsx"):
        shutil.copy(timeOrig, adminPath + "\\" +
                    aggName + "_Timeline_" + year + ".xlsx")
    else:
        pass


def disaggregate_agg(country, state, aggName, year, proj):
    path = os.path.normpath(
        f"G:\\Shared Drives\\EDU Consulting - Completed Reports\\{country}\\{state}\\{aggName}")
    for dirname in os.path.listdir(path):
        # do some iterating on the list of directories, check if they're
        # directories, and then if so, move them to the parent directory outside
        # of aggName.
        pass
        # the following takes a regex, matches with a filename in current working directory,
        # then copies that file to the location in drafts, and appends the correct
        # filetype, docx or ppt, or xlsx.


def matchReCopy(report, regex, filetype, location, country, state, school,
                projType, acronym="", year="2021", aggName="", campuses=""):
    """Matches a regex and finds that in the working directory, and copies that
    file to a location specified by the pathing for the project."""
    if campuses is not None:
        for campus in campuses:
            if type(school) == type("string"):
                for file in os.listdir():
                    pathToNew = r"G:\Shared Drives\EDU Consulting - Completed Reports" \
                        + "\\" + country + "\\" + state + "\\" + aggName +\
                        "\\" + school + "\\" + projType + "\\" + year + "\\" +\
                        campus + "\\" + location + "\\" + acronym + "_" + report +\
                        "_" + year + "_Draft." + filetype

                    if regex.match(file) and not os.path.exists(pathToNew):
                        shutil.copy(file, pathToNew)
                        return(pathToNew)
                    else:
                        pass
            elif type(school) == type([]):
                pathList = []
                for file in os.listdir():
                    for schoolname in school:
                        aggPathToNew = r"G:\Shared Drives\EDU Consulting - Completed Reports" \
                            + "\\" + country + "\\" + state + "\\" + aggName + \
                            "\\" + schoolname + "\\" + projType + "\\" + year + \
                            "\\" + location + "\\" + schoolname + "_" + report + \
                            "_" + year + "_Draft." + filetype
                        if regex.match(file) and not os.path.exists(aggPathToNew):
                            shutil.copy(file, aggPathToNew)
                            pathList.append(aggPathToNew)
                        else:
                            pass
                return(pathList)
            else:
                pass
    else:
        if type(school) == type("string"):
            for file in os.listdir():
                pathToNew = r"G:\Shared Drives\EDU Consulting - Completed Reports" \
                    + "\\" + country + "\\" + state + "\\" + aggName +\
                    "\\" + school + "\\" + projType + "\\" + year + \
                    "\\" + location + "\\" + acronym + "_" + report +\
                    "_" + year + "_Draft." + filetype

                if regex.match(file) and not os.path.exists(pathToNew):
                    shutil.copy(file, pathToNew)
                    return(pathToNew)
                else:
                    pass
        elif type(school) == type([]):
            pathList = []
            for file in os.listdir():
                for schoolname in school:
                    aggPathToNew = r"G:\Shared Drives\EDU Consulting - Completed Reports" \
                        + "\\" + country + "\\" + state + "\\" + aggName + \
                        "\\" + schoolname + "\\" + projType + "\\" + year + \
                        "\\" + location + "\\" + schoolname + "_" + report + \
                        "_" + year + "_Draft." + filetype
                    if regex.match(file) and not os.path.exists(aggPathToNew):
                        shutil.copy(file, aggPathToNew)
                        pathList.append(aggPathToNew)
                    else:
                        pass
            return(pathList)
        else:
            pass


def import_model(country, state, schoolList, projType, acronym="", year="2021",
                 aggName="", campuses=None):
    """Looks for model in a fixed location based on model name, for eis,
    US_EIS.xlsm, for pdga, GapModel2.2.xlsm, Currently located in C:\EIS
    folder, and places it in the working directory of the current project as
    defined by the users inputs."""
    pathDict = {}
    if campuses is None:
        if aggName == "":
            for school in schoolList:
                if projType in ["EIS", "VoD", "PSEIS"]:
                    pathToExistingModel = os.path.normpath(
                        r'C:\EIS\US_EIS.xlsm')
                if projType == "PDGA":
                    pathToExistingModel = os.path.normpath(
                        r"C:\EIS\GapModel2.2.xlsm")
                pathToNew = r"G:\Shared Drives\EDU Consulting - Completed Reports"\
                    + "\\" + country + "\\" + state + "\\" + school\
                    + "\\" + projType + "\\" + year + "\\" + "Drafts" + "\\"\
                    + acronym + "_" + "Model" + "_" + year + "_Draft.xlsm"
                if not os.path.exists(pathToNew):
                    shutil.copy(pathToExistingModel, pathToNew)
                pathDict[school] = [pathToExistingModel, pathToNew]
        else:
            for school in schoolList:
                if projType in ["EIS", "VoD", "PSEIS"]:
                    pathToExistingModel = os.path.normpath(
                        r'C:\EIS\US_EIS.xlsm')
                if projType == "PDGA":
                    pathToExistingModel = os.path.normpath(
                        r"C:\EIS\GapModel2.2.xlsm")
                pathToNew = r"G:\Shared Drives\EDU Consulting - Completed Reports"\
                    + "\\" + country + "\\" + state + "\\" + aggName + "\\" + school\
                    + "\\" + projType + "\\" + year + "\\" + "Drafts" + "\\"\
                    + acronym + "_" + "Model" + "_" + year + "_Draft.xlsm"
                if not os.path.exists(pathToNew):
                    shutil.copy(pathToExistingModel, pathToNew)
                pathDict[school] = [pathToExistingModel, pathToNew]
    else:
        if aggName == "":
            for campus in campuses:
                for school in schoolList:
                    if projType in ["EIS", "VoD", "PSEIS"]:
                        pathToExistingModel = os.path.normpath(
                            r'C:\EIS\US_EIS.xlsm')
                    if projType == "PDGA":
                        pathToExistingModel = os.path.normpath(
                            r"C:\EIS\GapModel2.2.xlsm")
                    pathToNew = r"G:\Shared Drives\EDU Consulting - Completed Reports"\
                        + "\\" + country + "\\" + state + "\\" + school\
                        + "\\" + projType + "\\" + year + "\\" + campus + "\\" + "Drafts" + "\\"\
                        + campus + "_" + "Model" + "_" + year + "_Draft.xlsm"
                    if not os.path.exists(pathToNew):
                        shutil.copy(pathToExistingModel, pathToNew)
                    pathDict[campus] = [pathToExistingModel, pathToNew]
            else:
                for campus in campuses:
                    for school in schoolList:
                        if projType in ["EIS", "VoD", "PSEIS"]:
                            pathToExistingModel = os.path.normpath(
                                r'C:\EIS\US_EIS.xlsm')
                        if projType == "PDGA":
                            pathToExistingModel = os.path.normpath(
                                r"C:\EIS\GapModel2.2.xlsm")
                        pathToNew = r"G:\Shared Drives\EDU Consulting - Completed Reports"\
                            + "\\" + country + "\\" + state + "\\" + aggName + "\\" + school\
                            + "\\" + projType + "\\" + year + "\\" + campus + "\\" + "Drafts" + "\\"\
                            + campus + "_" + "Model" + "_" + year + "_Draft.xlsm"
                        if not os.path.exists(pathToNew):
                            shutil.copy(pathToExistingModel, pathToNew)
                        pathDict[school] = [pathToExistingModel, pathToNew]
    return(pathDict)


def find_model(country, state, schoolList, projType, year, aggName="",
               campuses=None):
    modelPathDict = {}
    if campuses is None:
        for school in schoolList:
            if projType in ["EIS", "VoD", "PSEIS"]:
                pathToExistingModel = os.path.normpath(r'C:\EIS\US_EIS.xlsm')
            if projType == "PDGA":
                pathToExistingModel = os.path.normpath(
                    r"C:\EIS\GapModel2.2.xlsm")

            os.chdir(os.path.normpath(r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country +
                     "\\" + state + "\\" + aggName + "\\" + school + "\\" + projType + "\\" + year + "\\" + "Drafts"))
            matchDict = {}
            mostRecentMod = 0
            mostRecentNonPlayModel = None
            for file in os.listdir():
                modelRegex = re.compile(".*[mM][oO][dD][eE][lL].*")
                playRegex = re.compile(".*[pP][lL][aA][yY].*")
                if modelRegex.search(file) is not None and playRegex.search(file) is None:
                    matchDict[os.path.abspath(file)] = os.path.getmtime(
                        os.path.abspath(file))
            for file in matchDict:
                if ".xlsm" in file:
                    if mostRecentMod < matchDict[file]:
                        mostRecentNonPlayModel = file
                    else:
                        continue
            modelPathDict[school] = [
                pathToExistingModel, mostRecentNonPlayModel]
    else:
        for campus in campuses:
            for school in schoolList:
                if projType in ["EIS", "VoD", "PSEIS"]:
                    pathToExistingModel = os.path.normpath(
                        r'C:\EIS\US_EIS.xlsm')
                if projType == "PDGA":
                    pathToExistingModel = os.path.normpath(
                        r"C:\EIS\GapModel2.2.xlsm")

                os.chdir(os.path.normpath(r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + country + "\\" +
                         state + "\\" + aggName + "\\" + school + "\\" + projType + "\\" + year + "\\" + campus + "\\" + "Drafts"))
                matchDict = {}
                mostRecentMod = 0
                mostRecentNonPlayModel = None
                for file in os.listdir():
                    modelRegex = re.compile(".*[mM][oO][dD][eE][lL].*")
                    playRegex = re.compile(".*[pP][lL][aA][yY].*")
                    if modelRegex.search(file) is not None and playRegex.search(file) is None:
                        matchDict[os.path.abspath(file)] = os.path.getmtime(
                            os.path.abspath(file))
                for file in matchDict:
                    if ".xlsm" in file:
                        if mostRecentMod < matchDict[file]:
                            mostRecentNonPlayModel = file
                        else:
                            continue
                modelPathDict[school] = [
                    pathToExistingModel, mostRecentNonPlayModel]

    return(modelPathDict)


def import_templates(projType, schoolType, country, state, school, year,
                     acronym, aggName="", campuses=None):
    """Imports templates from their location on the saved drive into the project
    directory.
    Also imports all admin related documents as well and should create the project
    directory if it does not exist. WARNING: This will overwrite if the filenames
    are not changed, so be sure to change the filenames when editing them to not run that risk."""
    pathDict = {}
    if aggName == "":
        import_admin(country, state, school, year,
                     acronym, projType, schoolType)

    else:
        import_agg_admin(country, state, aggName, year)

    if campuses is not None:
        for campus in campuses:
            if projType == "EIS" and schoolType == "College" and country == "US":
                # Specific shared drive locations for each document. and regex to match
                # the template, regardless of what the date at the end says.
                os.chdir(
                    r"G:\\Shared drives\\EDU Consulting - Private\\EIS report masters\\US\\EIS_Individual\\College Template")
                colExecRe = re.compile(r"College_ExecSum.*")
                colFactRe = re.compile(r"College_FactSheet.*")
                colMainRe = re.compile(r"College_MainReport.*")

                pathToModel = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + \
                    country + "\\" + state + "\\" + aggName + "\\" + school + \
                    "\\" + projType + "\\" + year + "\\" + campus + "\\" + "Drafts" + "\\" + \
                    acronym + "_" + "Model" + "_" + year + "_Draft.xlsm"

                filePath = matchReCopy("ExecSum", colExecRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["ExecSum"] = filePath

                filePath = matchReCopy("FactSheet", colFactRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["FactSheet"] = filePath

                filePath = matchReCopy("MainReport", colMainRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["MainReport"] = filePath

            elif projType == "Capital" and country == "US":
                os.chdir(
                    r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\Capital")
                colExecRe = re.compile(r"Capital_ExecSum.*")
                colFactRe = re.compile(r"Capital_FactSheet.*")
                colMainRe = re.compile(r"Capital_MainReport.*")

                filePath = matchReCopy("ExecSum", colExecRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["ExecSum"] = filePath

                filePath = matchReCopy("FactSheet", colFactRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["FactSheet"] = filePath

                filePath = matchReCopy("MainReport", colMainRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["MainReport"] = filePath

            elif projType == "PSEIS" and schoolType == "College" and country == "US":
                os.chdir(
                    r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\EIS_Individual\College Template")
                colExecRe = re.compile(r"College_ExecSum.*")
                colFactRe = re.compile(r"College_FactSheet.*")
                colMainRe = re.compile(r"PSEIS_MainReport.*")

                filePath = matchReCopy("ExecSum", colExecRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["ExecSum"] = filePath

                filePath = matchReCopy("FactSheet", colFactRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["FactSheet"] = filePath

                filePath = matchReCopy("MainReport", colMainRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["MainReport"] = filePath

            elif projType == "PSEIS" and schoolType == "University" and country == "US":
                os.chdir(
                    r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\EIS_Individual\University Template")
                colExecRe = re.compile(r"University_ExecSum.*")
                colFactRe = re.compile(r"University_FactSheet.*")

                filePath = matchReCopy("ExecSum", colExecRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["ExecSum"] = filePath
                filePath = matchReCopy("FactSheet", colFactRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["FactSheet"] = filePath

                os.chdir(
                    r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\EIS_Individual\College Template")

                colMainRe = re.compile(r"PSEIS_MainReport.*")
                filePath = matchReCopy("MainReport", colMainRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["MainReport"] = filePath

            elif projType == "VoD" and country == "US":
                os.chdir(
                    r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\VoD")
                colFactRe = re.compile(r"VoD_FactSheet.*")
                colMainRe = re.compile(r"VoD_MainReport.*")

                filePath = matchReCopy("FactSheet", colFactRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["FactSheet"] = filePath

                filePath = matchReCopy("MainReport", colMainRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["MainReport"] = filePath

            elif projType in "EIS" and schoolType == "University" and country == "US":
                os.chdir(
                    r"G:\\Shared drives\\EDU Consulting - Private\\EIS report masters\\US\\EIS_Individual\\University Template")

                uniExecRe = re.compile(r"University_ExecSum.*")
                uniFactRe = re.compile(r"University_FactSheet.*")
                uniMainRe = re.compile(r"University_MainReport.*")

                filePath = matchReCopy("ExecSum", uniExecRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["ExecSum"] = filePath

                filePath = matchReCopy("FactSheet", uniFactRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["FactSheet"] = filePath

                filePath = matchReCopy("MainReport", uniMainRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["MainReport"] = filePath

            elif projType == "EIS" and country == "CAN":
                os.chdir(
                    r"G:\\Shared drives\\EDU Consulting - Private\\EIS report masters\\CAN\\EIS Individual")
                canExecRe = re.compile(r"CAN_ExecSum.*")
                canFactRe = re.compile(r"CAN_FactSheet.*")
                canMainRe = re.compile(r"CAN_MainReport.*")

                filePath = matchReCopy("ExecSum", canExecRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["ExecSum"] = filePath

                filePath = matchReCopy("FactSheet", canFactRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["FactSheet"] = filePath

                filePath = matchReCopy("MainReport", canMainRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["MainReport"] = filePath

            elif projType == "PDGA" and country == "US":
                os.chdir(
                    r"G:\\Shared drives\\EDU Consulting - Private\\PDGA report masters and surveys\\US\\PDGA Template Reports")
                usPdgaMainEnvRe = re.compile(r".*MainReport_Template_wEnv.*")
                usPdgaMainNoEnvRe = re.compile(
                    r".*MainReport_Template_[0-9].*")
                usPdgaAppendix = re.compile(r".*Appendix.*")

                filePath = matchReCopy("MainReport_wEnvScan", usPdgaMainEnvRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["MainReportEnv"] = filePath
                filePath = matchReCopy("MainReport", usPdgaMainNoEnvRe, "docx", "Drafts", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["MainReportNoEnv"] = filePath

                filePath = matchReCopy("Appendix", usPdgaAppendix, "docx", "Drafts", country, state,
                                       school, projType, acronym, year, aggName)
                pathDict["Appendix"] = filePath
    else:
        if projType == "EIS" and schoolType == "College" and country == "US":
            # Specific shared drive locations for each document. and regex to match
            # the template, regardless of what the date at the end says.
            os.chdir(
                r"G:\\Shared drives\\EDU Consulting - Private\\EIS report masters\\US\\EIS_Individual\\College Template")
            colExecRe = re.compile(r"College_ExecSum.*")
            colFactRe = re.compile(r"College_FactSheet.*")
            colMainRe = re.compile(r"College_MainReport.*")

            pathToModel = r"G:\Shared Drives\EDU Consulting - Completed Reports" + "\\" + \
                country + "\\" + state + "\\" + aggName + "\\" + school + \
                "\\" + projType + "\\" + year + "\\" + "Drafts" + "\\" + \
                acronym + "_" + "Model" + "_" + year + "_Draft.xlsm"

            filePath = matchReCopy("ExecSum", colExecRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["ExecSum"] = filePath

            filePath = matchReCopy("FactSheet", colFactRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["FactSheet"] = filePath

            filePath = matchReCopy("MainReport", colMainRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["MainReport"] = filePath

        elif projType == "Capital" and country == "US":
            os.chdir(
                r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\Capital")
            colExecRe = re.compile(r"Capital_ExecSum.*")
            colFactRe = re.compile(r"Capital_FactSheet.*")
            colMainRe = re.compile(r"Capital_MainReport.*")

            filePath = matchReCopy("ExecSum", colExecRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["ExecSum"] = filePath

            filePath = matchReCopy("FactSheet", colFactRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["FactSheet"] = filePath

            filePath = matchReCopy("MainReport", colMainRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["MainReport"] = filePath

        elif projType == "PSEIS" and schoolType == "College" and country == "US":
            os.chdir(
                r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\EIS_Individual\College Template")
            colExecRe = re.compile(r"College_ExecSum.*")
            colFactRe = re.compile(r"College_FactSheet.*")
            colMainRe = re.compile(r"PSEIS_MainReport.*")

            filePath = matchReCopy("ExecSum", colExecRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["ExecSum"] = filePath

            filePath = matchReCopy("FactSheet", colFactRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["FactSheet"] = filePath

            filePath = matchReCopy("MainReport", colMainRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["MainReport"] = filePath

        elif projType == "PSEIS" and schoolType == "University" and country == "US":
            os.chdir(
                r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\EIS_Individual\University Template")
            colExecRe = re.compile(r"University_ExecSum.*")
            colFactRe = re.compile(r"University_FactSheet.*")

            filePath = matchReCopy("ExecSum", colExecRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["ExecSum"] = filePath
            filePath = matchReCopy("FactSheet", colFactRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["FactSheet"] = filePath

            os.chdir(
                r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\EIS_Individual\College Template")

            colMainRe = re.compile(r"PSEIS_MainReport.*")
            filePath = matchReCopy("MainReport", colMainRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["MainReport"] = filePath

        elif projType == "VoD" and country == "US":
            os.chdir(
                r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\VoD")
            colFactRe = re.compile(r"VoD_FactSheet.*")
            colMainRe = re.compile(r"VoD_MainReport.*")

            filePath = matchReCopy("FactSheet", colFactRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["FactSheet"] = filePath

            filePath = matchReCopy("MainReport", colMainRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["MainReport"] = filePath

        elif projType in "EIS" and schoolType == "University" and country == "US":
            os.chdir(
                r"G:\\Shared drives\\EDU Consulting - Private\\EIS report masters\\US\\EIS_Individual\\University Template")

            uniExecRe = re.compile(r"University_ExecSum.*")
            uniFactRe = re.compile(r"University_FactSheet.*")
            uniMainRe = re.compile(r"University_MainReport.*")

            filePath = matchReCopy("ExecSum", uniExecRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["ExecSum"] = filePath

            filePath = matchReCopy("FactSheet", uniFactRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["FactSheet"] = filePath

            filePath = matchReCopy("MainReport", uniMainRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["MainReport"] = filePath

        elif projType == "EIS" and country == "CAN":
            os.chdir(
                r"G:\\Shared drives\\EDU Consulting - Private\\EIS report masters\\CAN\\EIS Individual")
            canExecRe = re.compile(r"CAN_ExecSum.*")
            canFactRe = re.compile(r"CAN_FactSheet.*")
            canMainRe = re.compile(r"CAN_MainReport.*")

            filePath = matchReCopy("ExecSum", canExecRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["ExecSum"] = filePath

            filePath = matchReCopy("FactSheet", canFactRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["FactSheet"] = filePath

            filePath = matchReCopy("MainReport", canMainRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["MainReport"] = filePath

        elif projType == "PDGA" and country == "US":
            os.chdir(
                r"G:\\Shared drives\\EDU Consulting - Private\\PDGA report masters and surveys\\US\\PDGA Template Reports")
            usPdgaMainEnvRe = re.compile(r".*MainReport_Template_wEnv.*")
            usPdgaMainNoEnvRe = re.compile(r".*MainReport_Template_[0-9].*")

            filePath = matchReCopy("MainReport_wEnvScan", usPdgaMainEnvRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["MainReportEnv"] = filePath
            filePath = matchReCopy("MainReport", usPdgaMainNoEnvRe, "docx", "Drafts", country, state,
                                   school, projType, acronym, year,
                                   aggName)
            pathDict["MainReportNoEnv"] = filePath
    return(pathDict)


def import_finals(projType, schoolType, country, state, school, year, acronym,
                  aggName="", campuses=None):
    """Imports Press Packet and PPTs into the existing project directory for
    EIS, for PDGA, imports PPT and Data Tables."""
    pathDict = {}
    if campuses is not None:
        for campus in campuses:
            if projType == "EIS" and schoolType == "College" and country == "US":
                os.chdir(
                    r"G:\\Shared drives\\EDU Consulting - Private\\EIS report masters\\US\\EIS_Individual\\College Template")
                colPptNoCon = re.compile(r".*_PowerPoint_[0-9].*")
                colPptCon = re.compile(r"College_PowerPoint_w_Con.*")

                filePath = matchReCopy("PowerPoint", colPptNoCon, "pptm", "Finals", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["PowerPoint_NoCon"] = filePath

                filePath = matchReCopy("PowerPoint_w_Construction", colPptCon, "pptm", "Finals", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["PowerPoint_Con"] = filePath

                os.chdir(
                    r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\EIS_Individual\College Template\Press Packet - College")
                mktEx = re.compile(r".*Marketing Examples.docx")
                method = re.compile(r".*Methodology.pdf")
                takeaways = re.compile(r".*Takeaways.*")

                filePath = matchReCopy("Marketing_Examples", mktEx, "docx", r"Finals\PressPacket", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)

                filePath = matchReCopy("Methodology", method, "pdf", r"Finals\PressPacket", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)

                filePath = matchReCopy("Takeaways", takeaways, "docx", r"Finals\PressPacket", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["Takeaways"] = filePath

            elif projType == "Capital" and schoolType == "College" and country == "US":
                os.chdir(
                    r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\Capital")
                colPpt = re.compile(r"College_Capital_PowerPoint.*")
                filePath = matchReCopy("PowerPoint", colPpt, "pptm", "Finals", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["PowerPoint"] = filePath

                os.chdir(
                    r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\EIS_Individual\College Template\Press Packet - College")
                mktEx = re.compile(r".*Marketing Examples.docx")
                method = re.compile(r".*Methodology.pdf")
                takeaways = re.compile(r".*Takeaways.*")

                filePath = matchReCopy("Marketing_Examples", mktEx, "docx", r"Finals\PressPacket", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)

                filePath = matchReCopy("Methodology", method, "pdf", r"Finals\PressPacket", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)

                filePath = matchReCopy("Takeaways", takeaways, "docx", r"Finals\PressPacket", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["Takeaways"] = filePath

            elif projType == "Capital" and schoolType == "University" and country == "US":
                os.chdir(
                    r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\Capital")
                uniPpt = re.compile(r"Uni_Capital_PowerPoint.*")
                filePath = matchReCopy("PowerPoint", uniPpt, "pptm", "Finals", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["PowerPoint"] = filePath

                os.chdir(
                    r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\EIS_Individual\University Template\Press Packet - University")
                mktEx = re.compile(r".*Marketing Examples_Lightcast.pdf")
                method = re.compile(r".*Methodology_Lightcast.pdf")
                takeaways = re.compile(r".*Takeaways_Lightcast\..*")
                takeawaysPriv = re.compile(r".*aways_PRIVATE.*_Lightcast.*")

                filePath = matchReCopy("Marketing_Examples", mktEx, "pdf", r"Finals\PressPacket", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)

                filePath = matchReCopy("Methodology", method, "pdf", "Finals", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)

                filePath = matchReCopy("Takeaways", takeaways, "docx", r"Finals\PressPacket", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["Takeaways_Public"] = filePath

                filePath = matchReCopy("Takeaways_Private", takeawaysPriv, "docx",
                                       r"Finals\PressPacket", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["Takeaways_Private"] = filePath

            elif projType == "PSEIS" and schoolType == "College" and country == "US":
                # os.chdir(
                #     r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\EIS_Individual\College Template")
                # matchReCopy("PowerPoint", colPptNoCon, "pptm", "Finals", country, state,
                #             school, projType, acronym, year, aggName)
                # matchReCopy("PowerPoint_w_Construction", colPptCon, "pptm", "Finals", country, state,
                #             school, projType, acronym, year, aggName)
                pass

            elif projType == "EIS" and schoolType == "University" and country == "US":
                os.chdir(
                    r"G:\\Shared drives\\EDU Consulting - Private\\EIS report masters\\US\\EIS_Individual\\University Template")
                uniPptEven = re.compile(r".*_PowerPoint_Even.*")
                uniPptOdd = re.compile(r".*_PowerPoint_Hospital.*")
                filePath = matchReCopy("PowerPoint_Even", uniPptEven, "pptm", "Finals", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["PowerPoint_Even"] = filePath

                filePath = matchReCopy("PowerPoint_Hospital_Odd", uniPptOdd, "pptm", "Finals", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["PowerPoint_Odd"] = filePath

                os.chdir(
                    r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\EIS_Individual\University Template\Press Packet - University")
                mktEx = re.compile(r".*Marketing Examples_Lightcast.pdf")
                method = re.compile(r".*Methodology_Lightcast.pdf")
                takeaways = re.compile(r".*Takeaways_Lightcast\..*")
                takeawaysPriv = re.compile(r".*aways_PRIVATE.*_Lightcast.*")

                filePath = matchReCopy("Marketing_Examples", mktEx, "pdf", r"Finals\PressPacket", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)

                filePath = matchReCopy("Methodology", method, "pdf", "Finals", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)

                filePath = matchReCopy("Takeaways", takeaways, "docx", r"Finals\PressPacket", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["Takeaways"] = filePath

                filePath = matchReCopy("Takeaways_Private", takeawaysPriv, "docx",
                                       r"Finals\PressPacket", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["Takeaways_Private"] = filePath

            elif projType in ["EIS", "Capital", "PSEIS", "VoD"] and country == "CAN":
                os.chdir(
                    r"G:\Shared drives\EDU Consulting - Private\EIS report masters\CAN\Press Packet")
                mktEx = re.compile(r".*MarketingExamples.*")
                method = re.compile(r".*Methodology.*")
                takeaways = re.compile(r".*Takeaways.*")

                filePath = matchReCopy("Marketing_Examples", mktEx, "indd", r"Finals\PressPacket", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)

                filePath = matchReCopy("Methodology", method, "indd", r"Finals\PressPacket", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)

                filePath = matchReCopy("Takeaways", takeaways, "indd", r"Finals\PressPacket", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)

                if schoolType == "University":
                    os.chdir(
                        r"G:\Shared drives\EDU Consulting - Private\EIS report masters\CAN")
                    ppt = re.compile(r".*_PPT_Uni.*")
                    filePath = matchReCopy("Ppt_New", ppt, "pptx", r"Finals", country, state,
                                           school, projType, acronym, year,
                                           aggName, campus)
                    pathDict["PowerPoint"] = filePath

                elif schoolType == "College":
                    os.chdir(
                        r"G:\Shared drives\EDU Consulting - Private\EIS report masters\CAN")
                    pptCon = re.compile(r".*_PPT_Const.*")
                    pptNew = re.compile(r".*_PPT_New.*")
                    filePath = matchReCopy("Ppt_Construction", pptCon, "pptx", r"Finals", country, state,
                                           school, projType, acronym, year,
                                           aggName, campus)
                    pathDict["PowerPoint_Construction"] = filePath

                    filePath = matchReCopy("Ppt_New", pptNew, "pptx", r"Finals", country, state,
                                           school, projType, acronym, year,
                                           aggName, campus)
                    pathDict["PowerPoint_New"] = filePath

            elif projType == "PDGA" and country == "US":
                os.chdir(
                    r"G:\\Shared drives\\EDU Consulting - Private\\PDGA report masters and surveys\\US\\PDGA Template Reports")

                usPdgaPptRe = re.compile(r".*PowerPoint.*Lightcast.*")
                filePath = matchReCopy("Ppt", usPdgaPptRe, "pptx", "Finals", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)
                pathDict["PowerPoint"] = filePath

                usPdgaDataTablesRe = re.compile(r".*DataTables.*")
                filePath = matchReCopy("DataTables", usPdgaDataTablesRe, "xlsx", "Finals", country, state,
                                       school, projType, acronym, year,
                                       aggName, campus)

            else:
                pass
    else:
        if projType == "EIS" and schoolType == "College" and country == "US":
            os.chdir(
                r"G:\\Shared drives\\EDU Consulting - Private\\EIS report masters\\US\\EIS_Individual\\College Template")
            colPptNoCon = re.compile(r".*_PowerPoint_[0-9].*")
            colPptCon = re.compile(r"College_PowerPoint_w_Con.*")

            filePath = matchReCopy("PowerPoint", colPptNoCon, "pptm", "Finals", country, state,
                                   school, projType, acronym, year, aggName)
            pathDict["PowerPoint_NoCon"] = filePath

            filePath = matchReCopy("PowerPoint_w_Construction", colPptCon, "pptm", "Finals", country, state,
                                   school, projType, acronym, year, aggName)
            pathDict["PowerPoint_Con"] = filePath

            os.chdir(
                r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\EIS_Individual\College Template\Press Packet - College")
            mktEx = re.compile(r".*Marketing Examples.docx")
            method = re.compile(r".*Methodology.pdf")
            takeaways = re.compile(r".*Takeaways.*")

            filePath = matchReCopy("Marketing_Examples", mktEx, "docx", r"Finals\PressPacket", country, state,
                                   school, projType, acronym, year, aggName)

            filePath = matchReCopy("Methodology", method, "pdf", r"Finals\PressPacket", country, state,
                                   school, projType, acronym, year, aggName)

            filePath = matchReCopy("Takeaways", takeaways, "docx", r"Finals\PressPacket", country, state,
                                   school, projType, acronym, year, aggName)
            pathDict["Takeaways"] = filePath

        elif projType == "Capital" and schoolType == "College" and country == "US":
            os.chdir(
                r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\Capital")
            colPpt = re.compile(r"College_Capital_PowerPoint.*")
            filePath = matchReCopy("PowerPoint", colPpt, "pptm", "Finals", country, state,
                                   school, projType, acronym, year, aggName)
            pathDict["PowerPoint"] = filePath

            os.chdir(
                r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\EIS_Individual\College Template\Press Packet - College")
            mktEx = re.compile(r".*Marketing Examples.docx")
            method = re.compile(r".*Methodology.pdf")
            takeaways = re.compile(r".*Takeaways.*")

            filePath = matchReCopy("Marketing_Examples", mktEx, "docx", r"Finals\PressPacket", country, state,
                                   school, projType, acronym, year, aggName)

            filePath = matchReCopy("Methodology", method, "pdf", r"Finals\PressPacket", country, state,
                                   school, projType, acronym, year, aggName)

            filePath = matchReCopy("Takeaways", takeaways, "docx", r"Finals\PressPacket", country, state,
                                   school, projType, acronym, year, aggName)
            pathDict["Takeaways"] = filePath

        elif projType == "Capital" and schoolType == "University" and country == "US":
            os.chdir(
                r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\Capital")
            uniPpt = re.compile(r"Uni_Capital_PowerPoint.*")
            filePath = matchReCopy("PowerPoint", uniPpt, "pptm", "Finals", country, state,
                                   school, projType, acronym, year, aggName)
            pathDict["PowerPoint"] = filePath

            os.chdir(
                r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\EIS_Individual\University Template\Press Packet - University")
            mktEx = re.compile(r".*Marketing Examples_Lightcast.pdf")
            method = re.compile(r".*Methodology_Lightcast.pdf")
            takeaways = re.compile(r".*Takeaways_Lightcast\..*")
            takeawaysPriv = re.compile(r".*aways_PRIVATE.*_Lightcast.*")

            filePath = matchReCopy("Marketing_Examples", mktEx, "pdf", r"Finals\PressPacket", country, state,
                                   school, projType, acronym, year, aggName)

            filePath = matchReCopy("Methodology", method, "pdf", "Finals", country, state,
                                   school, projType, acronym, year, aggName)

            filePath = matchReCopy("Takeaways", takeaways, "docx", r"Finals\PressPacket", country, state,
                                   school, projType, acronym, year, aggName)
            pathDict["Takeaways_Public"] = filePath

            filePath = matchReCopy("Takeaways_Private", takeawaysPriv, "docx",
                                   r"Finals\PressPacket", country, state,
                                   school, projType, acronym, year, aggName)
            pathDict["Takeaways_Private"] = filePath

        elif projType == "PSEIS" and schoolType == "College" and country == "US":
            # os.chdir(
            #     r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\EIS_Individual\College Template")
            # matchReCopy("PowerPoint", colPptNoCon, "pptm", "Finals", country, state,
            #             school, projType, acronym, year, aggName)
            # matchReCopy("PowerPoint_w_Construction", colPptCon, "pptm", "Finals", country, state,
            #             school, projType, acronym, year, aggName)
            pass

        elif projType == "EIS" and schoolType == "University" and country == "US":
            os.chdir(
                r"G:\\Shared drives\\EDU Consulting - Private\\EIS report masters\\US\\EIS_Individual\\University Template")
            uniPptEven = re.compile(r".*_PowerPoint_Even.*")
            uniPptOdd = re.compile(r".*_PowerPoint_Hospital.*")
            filePath = matchReCopy("PowerPoint_Even", uniPptEven, "pptm", "Finals", country, state,
                                   school, projType, acronym, year, aggName)
            pathDict["PowerPoint_Even"] = filePath

            filePath = matchReCopy("PowerPoint_Hospital_Odd", uniPptOdd, "pptm", "Finals", country, state,
                                   school, projType, acronym, year, aggName)
            pathDict["PowerPoint_Odd"] = filePath

            os.chdir(
                r"G:\Shared drives\EDU Consulting - Private\EIS report masters\US\EIS_Individual\University Template\Press Packet - University")
            mktEx = re.compile(r".*Marketing Examples_Lightcast.pdf")
            method = re.compile(r".*Methodology_Lightcast.pdf")
            takeaways = re.compile(r".*Takeaways_Lightcast\..*")
            takeawaysPriv = re.compile(r".*aways_PRIVATE.*_Lightcast.*")

            filePath = matchReCopy("Marketing_Examples", mktEx, "pdf", r"Finals\PressPacket", country, state,
                                   school, projType, acronym, year, aggName)

            filePath = matchReCopy("Methodology", method, "pdf", "Finals", country, state,
                                   school, projType, acronym, year, aggName)

            filePath = matchReCopy("Takeaways", takeaways, "docx", r"Finals\PressPacket", country, state,
                                   school, projType, acronym, year, aggName)
            pathDict["Takeaways"] = filePath

            filePath = matchReCopy("Takeaways_Private", takeawaysPriv, "docx",
                                   r"Finals\PressPacket", country, state,
                                   school, projType, acronym, year, aggName)
            pathDict["Takeaways_Private"] = filePath

        elif projType in ["EIS", "Capital", "PSEIS", "VoD"] and country == "CAN":
            os.chdir(
                r"G:\Shared drives\EDU Consulting - Private\EIS report masters\CAN\Press Packet")
            mktEx = re.compile(r".*MarketingExamples.*")
            method = re.compile(r".*Methodology.*")
            takeaways = re.compile(r".*Takeaways.*")

            filePath = matchReCopy("Marketing_Examples", mktEx, "indd", r"Finals\PressPacket", country, state,
                                   school, projType, acronym, year, aggName)

            filePath = matchReCopy("Methodology", method, "indd", r"Finals\PressPacket", country, state,
                                   school, projType, acronym, year, aggName)

            filePath = matchReCopy("Takeaways", takeaways, "indd", r"Finals\PressPacket", country, state,
                                   school, projType, acronym, year, aggName)

            if schoolType == "University":
                os.chdir(
                    r"G:\Shared drives\EDU Consulting - Private\EIS report masters\CAN")
                ppt = re.compile(r".*_PPT_Uni.*")
                filePath = matchReCopy("Ppt_New", ppt, "pptx", r"Finals", country, state,
                                       school, projType, acronym, year, aggName)
                pathDict["PowerPoint"] = filePath

            elif schoolType == "College":
                os.chdir(
                    r"G:\Shared drives\EDU Consulting - Private\EIS report masters\CAN")
                pptCon = re.compile(r".*_PPT_Const.*")
                pptNew = re.compile(r".*_PPT_New.*")
                filePath = matchReCopy("Ppt_Construction", pptCon, "pptx", r"Finals", country, state,
                                       school, projType, acronym, year, aggName)
                pathDict["PowerPoint_Construction"] = filePath

                filePath = matchReCopy("Ppt_New", pptNew, "pptx", r"Finals", country, state,
                                       school, projType, acronym, year, aggName)
                pathDict["PowerPoint_New"] = filePath

        elif projType == "PDGA" and country == "US":
            os.chdir(
                r"G:\\Shared drives\\EDU Consulting - Private\\PDGA report masters and surveys\\US\\PDGA Template Reports")

            usPdgaPptRe = re.compile(r".*PowerPoint.*Lightcast.*")
            filePath = matchReCopy("Ppt", usPdgaPptRe, "pptx", "Finals", country, state,
                                   school, projType, acronym, year, aggName)
            pathDict["PowerPoint"] = filePath

            usPdgaDataTablesRe = re.compile(r".*DataTables.*")
            filePath = matchReCopy("DataTables", usPdgaDataTablesRe, "xlsx", "Finals", country, state,
                                   school, projType, acronym, year, aggName)

        else:
            pass

    return(pathDict)

#                   ⣰⣦⣄           
#                ⢀⣴⣿⡿⠃          
#               ⠭⠚⠿⣋            
#             ⡜ ⠱⡀             
#            ⡐   ⠑⡀            
#   ⣄      ⡰⣠⠼⠚⠛⢦⣜⣆      ⢠⠧⢴⣶⡆
#⠰⠿⠛⠹    ⡰⡹⠃ ⢠  ⠙⢭⣧⡀    ⠳⡀⠈ 
#  ⡠⠃ ⢀⣀⡴⠙⠷⢄ ⢸⠆ ⣠⠾⠉⠹⣶⠦⢤⣀⡇  
#  ⠉⠉⠉⠉⡰⠧  ⡵⣯⠿⠿⣭⠯⠤⠤⠤⠬⣆ ⠈   
#     ⣰⣁⣁⣀⣀⣄⣛⠉⠉⢟⠈⡄   ⢈⣆    
#    ⡰⡇⡘   ⣰⣀⣀⣀⣀⣀⣇⣀⣀⣀⣆⣋⡂     
#          ⢸    ⡇                  
#          ⠈⢆   ⠘⠒⢲⠆       
#            ⢹⡀   ⠸        
#            ⠈⡇   ⠁        

VERSION = "2.0.0r11"

#TODO:
#DONE: requestSample() - Add ability to add Text Comments to Add. Notes field; [008180] in any rules refers to that field and $ is "text contains"

import configparser
import datetime
import logging
import os
import pyautogui
#import pydirectinput
import pyperclip
import subprocess
import time
import win32gui #install pywin32
import re

DEFAULT_CLINICIAN = "UNK"
DEFAULT_SOURCE_IP = "RSW100"
DEFAULT_SOURCE_OP = "RSOP"
DEFAULT_SOURCE_GP = "M82046"
DEFAULT_LASTNAME  = "UKASBIO-LKB"
ROOTDIR = "W:/Pathology/Biochem/LKBecker/Projects/autoTester_CP/"
LOGFORMAT = '%(asctime)s: %(name)-18s:%(levelname)-7s:%(message)s'

SHORT_DELAY     = 0.4
MED_DELAY       = 0.55
LONG_DELAY      = 0.8
LONGEST_DELAY   = 1.50

WORKING_FROM_HOME = False

if WORKING_FROM_HOME:
    SHORT_DELAY     = 0.75
    MED_DELAY       = 1.00
    LONG_DELAY      = 1.30
    LONGEST_DELAY   = 2.50


RESULT_STATIC_X = 900
RESULT_Y_PER_ANALYTE = 20
AUTH_QUEUE_PER_Y= 22.5 

SCENARIOS = []
pyautogui.PAUSE = 0.25


class TargetResult():
    def __init__(self, profile:str, disciplineIdx:int, profileIdx:int, analyteIdx:int, analyte:str, result:float, totalOffset:int) -> None:
        self.Profile = profile
        self.DisciplineIndex = disciplineIdx
        self.ProfileIndex = profileIdx
        self.AnalyteIndex = analyteIdx
        self.Analyte = analyte
        self.Value = result
        self.TotalOffset = totalOffset

    def __repr__(self):
        return f"TargetResults({self.Profile}:{self.Analyte}({self.DisciplineIndex},{self.ProfileIndex},{self.AnalyteIndex}),{self.Value},TotalY={self.TotalOffset})"

    def __str__(self):
        return f"[{self.Profile}] {self.Analyte}={self.Value};"


class TestingScenario():
    def __init__(self, textLine) -> None:
        #TODO: rewrite
        self.LabNumber = None
        self.AuthQueue = "[Not Retrieved]"
        self.targetResults = []

        textLine            = textLine.split("\t")
        
        assert len(textLine) == 12 #new File format!

        self.ID             = textLine[0]
        self.Phase          = int(textLine[1])
        self.ScenarioSex    = textLine[2]
        self.Location       = textLine[3]
        self.SampleTaken    = datetime.datetime.strptime(textLine[4], "%Y-%m-%d %H:%M")
        self.SampleReceived = datetime.datetime.strptime(textLine[5], "%Y-%m-%d %H:%M")
        self.PatientDOB     = datetime.datetime.strptime(textLine[6], "%Y-%m-%d")
        self.SubScenarioID  = None

        scenarioStrSplit = textLine[7].split(";")
        for analyteStr in scenarioStrSplit:
            subStrSplit = analyteStr.strip().split("|")
            testSet     = subStrSplit[0]
            analyte     = subStrSplit[1]
            result      = subStrSplit[2].strip()
            
            disciplineIdx   = int(subStrSplit[3].strip())
            profileIdx      = int(subStrSplit[4].strip())
            analyteIdx      = int(subStrSplit[5].strip())
            totalOffset     = int(subStrSplit[6].strip())

            try:
                result = float(result)
            except ValueError:
                result = result

            self.targetResults.append(TargetResult(testSet, disciplineIdx, profileIdx, analyteIdx, analyte, result, totalOffset))

        self.ClinicalDetails    = textLine[8].strip()
        self.PatientTags        = textLine[9].strip().split(";")
        if len(self.PatientTags) == 0:
            self.PatientTags = None
        self.SubScenarioID      = textLine[10].strip()
        self.ClinNotes          = textLine[11].strip()
        if not self.ClinNotes:
            self.ClinNotes = None

        self.requiredTestSets = set()

        self.targetResults.sort(key = lambda x: (x.DisciplineIndex, x.ProfileIndex, x.AnalyteIndex))
        self.requiredTestSets = list(dict.fromkeys(map(lambda x: (x.Profile, x.TotalOffset), self.targetResults))) 

        self.targetResultStr = ''.join(map(lambda x: str(x), self.targetResults))

    def __str__(self):
        if self.SubScenarioID:
            return f"Scenario({self.SubScenarioID},{self.ScenarioSex},{self.PatientDOB.strftime('%d/%m/%Y')},{self.targetResultStr})"
        return f"Scenario({self.ID},{self.ScenarioSex},{self.PatientDOB.strftime('%d/%m/%Y')},{self.targetResultStr})"

    def __repr__(self):
        if self.SubScenarioID:
            return f"Scenario({self.SubScenarioID},{self.ScenarioSex},{self.PatientDOB.strftime('%d/%m/%Y')},{self.targetResultStr})"
        return f"Scenario({self.ID},{self.ScenarioSex},{self.PatientDOB.strftime('%d/%m/%Y')},{self.targetResultStr})"

    def clearAll():
        global SCENARIOS
        SCENARIOS = []

    def parseAll(fileName:str):
        with open(os.path.join(ROOTDIR, fileName), 'r') as ScenarioDataFile:
            ScenarioLines = ScenarioDataFile.readlines()
            ScenarioLines = ScenarioLines[1:] #Skip header
            for line in ScenarioLines:
                SCENARIOS.append( TestingScenario(line) )
        logging.info(f"processScenarios(): {len(SCENARIOS)} scenarios loaded from file.")


class WindowMgr:
    def __init__ (self):
        self._handle = None

    def find_window(self, class_name, window_name=None):
        self._handle = win32gui.FindWindow(class_name, window_name)

    def _window_enum_callback(self, hwnd, wildcard):
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) is not None:
            self._handle = hwnd

    def find_window_wildcard(self, wildcard):
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)

    def set_foreground(self):
        win32gui.SetForegroundWindow(self._handle)


config = configparser.ConfigParser()
config.read(os.path.join(ROOTDIR, 'config.ini'))

logging.basicConfig(filename=os.path.join(ROOTDIR, 'Output/debug.log'), filemode='w', level=logging.DEBUG, format=LOGFORMAT)
console = logging.StreamHandler()
console.setLevel(logging.DEBUG)
console.setFormatter(logging.Formatter(LOGFORMAT))
logging.getLogger().addHandler(console)
logging.info("Starting program...")

WindowManager = WindowMgr()

#These are HARDCODED values assuming a 1920 * 1080 screen (my screen, to be precise). 
# #It _should_ work elsewhere, but no promises...
WinPath = dict()
WinPath['QuitButton']                       = (1890,  10)
WinPath['Btn_RequestEntry']                 = (  40, 120)
WinPath['Btn_ResultEntry']                  = (  40, 185)
WinPath['Btn_Search']                       = (  40, 250)
WinPath['Btn_Authorisation']                = (  40, 435)

RequestEntry = dict()
RequestEntry['Btn_UseNextFreeLabNo']        = ( 380, 135)
RequestEntry['Btn_Go']		                = ( 300, 130)
RequestEntry['Field_Labno_Start']           = ( 183, 130)
RequestEntry['Field_Labno_End']             = ( 281, 130)
RequestEntry['Btn_SwitchToManualEntry']     = (1840, 130)
RequestEntry['Field_Surname']               = ( 185, 260)
RequestEntry['Field_Forename']              = ( 185, 290)
RequestEntry['Area_NoPtFound']              = ( 100, 220,  200, 250)
RequestEntry['Btn_NewPatient']              = (1830, 270)
RequestEntry['Field_NewPtDOB']              = ( 270, 300)
RequestEntry['Field_NewPtSex']              = ( 270, 270)
RequestEntry['Btn_ConfirmNewPatient']       = (1790, 990)
RequestEntry['Btn_AcceptPatient']           = (1695, 250)
RequestEntry['Field_Clinician']             = ( 190, 335)
RequestEntry['Field_Source']                = ( 190, 360)
RequestEntry['Field_ClinicalDetails1']      = ( 190, 455)
RequestEntry['Field_AddNotes']              = ( 190, 505)
RequestEntry['Field_FirstTest']             = ( 190, 585)
RequestEntry['Field_SampleDate']            = ( 190, 545)
RequestEntry['Field_SampleTime']            = ( 295, 545)
RequestEntry['Field_ReceivedDate']          = ( 455, 545)
RequestEntry['Field_ReceivedTime']          = ( 575, 545)
RequestEntry['Btn_SaveRequest']             = (1800, 990)
RequestEntry['Btn_AmendPatient']            = (1695, 220)
RequestEntry['Field_PatientDOB_Day']        = ( 190, 225)
RequestEntry['Field_PatientDOB_Month']      = ( 210, 225)
RequestEntry['Field_PatientDOB_Year']       = ( 230, 225)
RequestEntry['Btn_SaveAndReturn']           = (1610, 990)

Results = dict()
Results['Field_LabID']                      = ( 194, 154)
Results['Btn_Go']                           = ( 307, 159)
Results['Results_Field_Start']              = ( 393, 285)
Results['Btn_Save']                         = (1795, 990)
Results['Btn_Queue']                        = (1710, 990)
Results['Btn_ExpandDetails']                = (1910, 200)
Results['Area_Report']                      = (  90,  80, 1830, 955)
Results['Area_Auth']                        = ( 760,  955, 150,  30)

Search = dict()
Search['Btn_RequestSearch']                 = (1845, 195)
Search['Btn_RequestSearch']                 = (1845, 195)
Search['Field_LabNoRange']                  = ( 575, 330)
Search['Btn_Search']                        = ( 932, 140)
Search['Btn_ExpandDetails']                 = (1910, 200)
Search['Area_Report']                       = (  90,  80, 1830, 955)

Authorisation = dict()
Authorisation['Btn_Search']                 = ( 240, 965)
Authorisation['Btn_List_PASS']              = (1110, 140)
Authorisation['Btn_List_FAIL']              = ( 210, 140)
Authorisation['PASS_Queue_Start']           = (1080, 185)
Authorisation['FAIL_Queue_Start']           = ( 180, 185)
Authorisation['TopOfAuthQueueList']         = ( 155, 140)
Authorisation['Btn_Authorise_List']         = ( 140, 970)
Authorisation['Btn_Authorise_Report']       = ( 130, 990)
Authorisation['Btn_AuthLists_OK']           = ( 959, 694)
Authorisation['TopOfResults']               = ( 190, 275)
Authorisation['Area_Auth_Rules']            = (  90, 900,  580, 75)
Authorisation['Btn_Auth_Cancel']            = (1875, 990)
Authorisation['Area_Scrn_Auth']             = (  90,  80, 1830, 955)
Authorisation['Area_Queue_Popup']           = ( 720, 450,  490, 160)


PASSQueues = [
    "URGENT HAEMATOLOGY", "URGENT COAGULATION", "DAWN COAGULATION", "URGENT GENERAL BIOCHEMISTRY", "ENDOCRINOLOGY", 
    "PROTEINS", "BIOCHEM DUPLICATE REVIEW", "ROUTINE HAEMATOLOGY", "ROUTINE COAGULATIONS", "MANUAL BIOCHEMISTRY", 
    "SPECIALIST MANUAL BIOCHEM", "NON-URGENT BIOCHEMISTRY", "FILMS", "SPECIAL COAGUALTION", "FLOW CYTOMETRY", 
    "AUTOIMMUNITY (AUTO)", "AUTOIMMUNITY (MANUAL)", "AUTOIMMUNITY IIF", "ALLERGY", "BIOCHEMISTRY REFERRALS", 
    "HAEMATOLOGY REJECTIONS", "IMMUNOLOGY REFERRALS", "BIOCHEMISTRY CLINICAL REVIEW"
]

FAILQueues = [
    "(DEFAULT)", "URGENT COAGULATION", "URGENT GENERAL BIOCHEMISTRY", "ENDOCRINOLOGY", "PROTEINS", "BIOCHEM DUPLICATE REVIEW", 
    "ROUTINE HAEMATOLOGY", "ROUTINE COAGULATION", "DFT", "MANUAL BIOCHEMISTRY", "SPECIALIST MANUAL BIOCHEM", 
    "MOLECULAR GENETIC TESTING", "NON-URGENT BIOCHEMISTRY", "FILMS", "URGENT FILMS", "SPECIAL COAGULATION", 
    "PD FLUID", "FLOW CYTOMETRY", "SPECIAL HAEMATOLOGY", "GF", "IMMUNOLOGY REJECTIONS", 
    "AUTOIMMUNITY (AUTO)", "AUTOIMMUNITY (MANUAL)", "AUTOIMMUNITY IIF", "IMMUNODEFICIENCY", "URGENT IMMUNOLOGY CONS.", 
    "NON URGENT IMMUNOLOGY CONS.", "BIOCHEMISTRY REFERRALS", "HAEMATOLOGY REJECTIONS", "IMMUNOLOGY REFERRALS", 
    "BIOCHEMISTRY CLINICAL REVIEW"
]

DOBVerified = []

def timestamp(forFile:bool=False):
    if forFile:
        return datetime.datetime.now().strftime("%y%m%d-%H%M%S")
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def waitForWindow(windowTitle:str, maxWaitSeconds:int=10) -> None:
    passedTime = 0
    while True:
        time.sleep(SHORT_DELAY)
        passedTime = passedTime + 1
        if passedTime >= maxWaitSeconds:
            raise Exception(f"Could not locate Window '{windowTitle}' within {maxWaitSeconds} seconds.")

        BeakerWindow = pyautogui.getWindowsWithTitle(windowTitle) #Does partial matches, the closer the better
        logging.debug(f"waitForWindow(): {len(BeakerWindow)} windows located matching query.")
        if BeakerWindow:
            BeakerWindow = BeakerWindow[0]
            BeakerWindow.restore()
            break

def focusWindowIfExists(windowTitle:str) -> bool:
    WindowManager.find_window_wildcard(f".*{windowTitle}.*")
    if WindowManager._handle:
        WindowManager.set_foreground()
        return True
    else:
        return False

def captureScreenshot(name:str="UNKNOWN", area:tuple=None, addTS:bool=False) -> None:
    if not os.path.isdir(f"{ROOTDIR}/Screenshots/"):
        os.mkdir(f"{ROOTDIR}/Screenshots/")
    if area:
        if len(area) != 4:
            raise Exception("captureScreenshot(): area must be a tuple of 4 items, (top, left, height, width)")
    if addTS:
        ts = datetime.datetime.now().strftime("%y%m%d-%H%M%S")
        fileName = f"{ROOTDIR}/Screenshots/{name}_{ts}.png"
    else:
        fileName = f"{ROOTDIR}/Screenshots/{name}.png"
    if not area:
        pyautogui.screenshot(imageFilename=fileName)
        logging.info(f"Saving screenshot '{fileName}'...")
    else:
        pyautogui.screenshot(imageFilename=fileName, region=area)
        logging.info(f"Saving screenshot '{fileName}', area {area}...")

def retrieveViaClipboard(startCoordTuple:tuple, endCoordTuple:tuple) -> str:
    pyautogui.click(*startCoordTuple)
    pyautogui.dragTo(*endCoordTuple, button="left")
    pyautogui.hotkey('ctrl', 'c') #bootlegHotkey('ctrl', 'c')
    time.sleep(.2)
    return pyperclip.paste()

def openWinPath():
    logging.info("Opening Remote Desktop endpoint...")
    subprocess.call(f"mstsc {os.path.join(ROOTDIR, 'cpub-Winpath-Collection1-CmsRdsh.rdp')}")
    time.sleep(5)
    WinPath = focusWindowIfExists("(REMOTEUHNM.WINPATH.CO.UK)")
    if not WinPath:
        time.sleep(15)
        pyautogui.write(config['Default']['Password'])
        pyautogui.press('enter') # pydirectinput.press('enter')
        time.sleep(15)
    WinPath = focusWindowIfExists("(REMOTEUHNM.WINPATH.CO.UK)")
    pyautogui.press('enter') # pydirectinput.press('enter') #Select Stoke Location
    if WinPath:
        logging.info("Logged into WinPath.")
 
def closeWinPath():
    WinPath = focusWindowIfExists("(REMOTEUHNM.WINPATH.CO.UK)")
    if WinPath:
        pyautogui.hotkey('alt', 'f4') #bootlegHotkey('alt', 'f4') #Launches Logout And Close? prompt
        time.sleep(SHORT_DELAY)

        pyautogui.press('enter') # pydirectinput.press('enter')

def requestSample(Scenario:TestingScenario, modeMRI:bool=False, useOffset:bool=False):
    def checkForLabNoAllocError():
        focusWindowIfExists("WinPath") #TODO Check
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(SHORT_DELAY)
        pyautogui.hotkey('ctrl', 'c')
        windowVal = pyperclip.paste()
        if windowVal:
            windowVal = [x for x in windowVal.split("\n") if x]
            if len(windowVal)>3:
                if windowVal[3] == "This lab no. has been entered on another terminal since you started keying it in.":
                    logging.info(f"Detected Lab No allocation error for scenario {Scenario.SubScenarioID}-{Scenario.Phase}. Looping...")
                    return True
        return False
        

    focusWindowIfExists("WinPath Enterprise : ")
    logging.debug("requestSample(): Start")
    #pyautogui.click(*WinPath['Btn_RequestEntry'])
    pyautogui.hotkey('shift', 'F1')
    time.sleep(MED_DELAY)
    logging.debug("Request entry opened.")

    #TODO: recode for Not Actually Free Lab No

    LabNoAllocError = True
    while LabNoAllocError:
    
        if useOffset:
            logging.info(f"Processing Scenario {Scenario.SubScenarioID}-{Scenario.Phase} via -offset- sample no [{Scenario.LabNumber}]...")
            #request entry opens with cursor in Lab No field
            pyautogui.write(Scenario.LabNumber)

        else:
            pyautogui.click(*RequestEntry['Btn_UseNextFreeLabNo'])
            time.sleep(MED_DELAY)
            Scenario.LabNumber = retrieveViaClipboard(RequestEntry['Field_Labno_Start'], RequestEntry['Field_Labno_End'])
            assert Scenario.LabNumber != ""
            logging.info(f"Processing Scenario {Scenario.SubScenarioID}-{Scenario.Phase} via -acquired- sample no [{Scenario.LabNumber}]...")

        pyautogui.click(*RequestEntry['Btn_Go'])
        time.sleep(SHORT_DELAY)
        pyautogui.click(*RequestEntry['Btn_SwitchToManualEntry'])
        time.sleep(MED_DELAY)
        
        pyautogui.click(*RequestEntry['Field_Surname'])
        pyautogui.write(DEFAULT_LASTNAME)
        pyautogui.press('tab')
        pyautogui.write(Scenario.ID)
        pyautogui.press('enter')
        time.sleep(SHORT_DELAY)

        if WORKING_FROM_HOME:
            time.sleep(30)

        try:
            PatientCreatedHere = False
            _PatientDoenstExist = pyautogui.locateOnScreen(image=os.path.join(ROOTDIR, "noPatientFound.png"), region = RequestEntry['Area_NoPtFound'])
            if _PatientDoenstExist:
                pyautogui.click(*RequestEntry['Btn_NewPatient'])   
                PatientCreatedHere = True
                time.sleep(MED_DELAY)
                pyautogui.click(*RequestEntry['Field_NewPtDOB'])
                pyautogui.write(Scenario.PatientDOB.strftime("%d/%m/%Y"))

                pyautogui.click(*RequestEntry['Field_NewPtSex'])
                pyautogui.write(Scenario.ScenarioSex)

                time.sleep(MED_DELAY)
                pyautogui.click(*RequestEntry['Btn_ConfirmNewPatient'])
                time.sleep(MED_DELAY)
            else:
                if modeMRI == False and not (Scenario.ID in DOBVerified):
                    pyautogui.click(*RequestEntry['Btn_AmendPatient'])
                    time.sleep(SHORT_DELAY)
                    pyautogui.click(*RequestEntry['Field_PatientDOB_Day'])
                    time.sleep(SHORT_DELAY)
                    pyautogui.hotkey('ctrl', 'c')
                    time.sleep(SHORT_DELAY)
                    DOBDay = pyperclip.paste()

                    pyautogui.click(*RequestEntry['Field_PatientDOB_Month'])
                    time.sleep(SHORT_DELAY)
                    pyautogui.hotkey('ctrl', 'c')
                    time.sleep(SHORT_DELAY)
                    DOBMonth = pyperclip.paste()
                    time.sleep(SHORT_DELAY)
                    
                    pyautogui.click(*RequestEntry['Field_PatientDOB_Year'])
                    time.sleep(SHORT_DELAY)
                    pyautogui.hotkey('ctrl', 'c')
                    time.sleep(SHORT_DELAY)
                    DOBYear = pyperclip.paste()
                    CurrentDOB = datetime.datetime.strptime(f"{DOBDay}-{DOBMonth}-{DOBYear}", "%d-%b-%Y")
                    del(DOBDay, DOBMonth, DOBYear)

                    if not (Scenario.PatientDOB.date() == CurrentDOB.date()):
                        pyautogui.click(*RequestEntry['Field_PatientDOB_Day'])
                        pyautogui.keyDown('delete')
                        time.sleep(MED_DELAY)
                        pyautogui.keyUp('delete')
                        pyautogui.write(Scenario.PatientDOB.strftime("%d/%m/%Y"))
                        logging.info(f"Amending DOB for patient from {CurrentDOB} to {Scenario.PatientDOB}...")
                        DOBVerified.append(Scenario.ID)

                    pyautogui.click(*RequestEntry['Btn_SaveAndReturn'])
                    time.sleep(SHORT_DELAY)
                pyautogui.click(*RequestEntry['Btn_AcceptPatient'])   
                time.sleep(MED_DELAY)

        except pyautogui.ImageNotFoundException:
                logging.debug("ImageNotFoundException triggered; patient appears to exist. Continuing...")
                pyautogui.click(*RequestEntry['Btn_AcceptPatient'])
            #TODO: Reimplement?
            # SetDOB = retrieveViaClipboard(RequestEntry['Field_DOB_Start'], RequestEntry['Field_DOB_End'])
            # SetDOB = datetime.datetime.strptime(SetDOB, "%d/%m/%Y")
            # if not (Scenario.PatientDOB.date() == SetDOB.date()):
            #     pyautogui.click(*RequestEntry['DOBFieldStart'], clicks=2)
            #     pyautogui.write(Scenario.PatientDOB.strftime("%d/%m/%Y"))
            #     changedDOB = True

        if PatientCreatedHere:
            if Scenario.PatientTags:
                if Scenario.PatientTags != "":
                    logging.info("requestSample(): Patient was newly created and requires Flags to be applied. Opening flagging interface...")
                    pyautogui.hotkey('shift', 'f9') #Open Patient Flag Panel, could also click at 2 coords
                    time.sleep(SHORT_DELAY)
                    pyautogui.press('tab') #Select lab number field
                    time.sleep(SHORT_DELAY)
                    pyautogui.write(Scenario.LabNumber)
                    pyautogui.press('enter') #Confirm current lab number
                    time.sleep(MED_DELAY)
                    pyautogui.press('tab') #Select first flag field
                    for _Flag in Scenario.PatientTags:
                        logging.info(f"requestSample(): Applying flag {_Flag} to patient.")
                        pyautogui.write(_Flag)
                        pyautogui.press('enter')
                        #TODO: Check for error window
                    pyautogui.hotkey('alt', 'o') #Presses OK button
                    time.sleep(MED_DELAY)
        
        pyautogui.click(*RequestEntry['Field_Clinician'])  
        pyautogui.write(DEFAULT_CLINICIAN)

        pyautogui.click(*RequestEntry['Field_Source'])
        if Scenario.Location == "GP":
            pyautogui.write(DEFAULT_SOURCE_GP)
        elif Scenario.Location == "IP":
            pyautogui.write(DEFAULT_SOURCE_IP)
        elif Scenario.Location == "OP":
            pyautogui.write(DEFAULT_SOURCE_OP)
        else:
            pyautogui.write(Scenario.Location)
    
        pyautogui.click(*RequestEntry['Field_SampleDate'])
        pyautogui.hotkey('shift', 'end')
        pyautogui.press('delete')
        pyautogui.write(Scenario.SampleTaken.strftime("%d/%m/%Y"))
        pyautogui.press('tab', presses=2)
        pyautogui.write(Scenario.SampleTaken.strftime("%H%M"))

        pyautogui.press('tab')
        pyautogui.write(Scenario.SampleReceived.strftime("%d/%m/%Y"))
        pyautogui.press('tab', presses=2)
        pyautogui.write(Scenario.SampleReceived.strftime("%H%M"))

        if Scenario.ClinicalDetails:
            logging.info(f"Scenario {Scenario.SubScenarioID}-{Scenario.Phase} requires Clinical Detail [{Scenario.ClinicalDetails}]. Entering code.")    
            pyautogui.click(*RequestEntry['Field_ClinicalDetails1'])
            pyautogui.write(Scenario.ClinicalDetails)

        if Scenario.ClinNotes:
            logging.info(f"Scenario {Scenario.SubScenarioID}-{Scenario.Phase} specifies Add. Notes [{Scenario.ClinNotes}]. Entering into designated field...")
            pyautogui.click(*RequestEntry['Field_AddNotes'])
            pyautogui.write(Scenario.ClinNotes)
        
        #Now to create the sample and request...
        pyautogui.click(*RequestEntry['Field_FirstTest'])
        logging.info(f"Submitting test requests for Scenario {Scenario.SubScenarioID}-{Scenario.Phase}...")
        for requiredSet in Scenario.requiredTestSets:
            logging.info(f"Entering required set [{requiredSet}].")
            pyautogui.write(requiredSet[0])
            pyautogui.press('tab')
        pyautogui.click(*RequestEntry['Btn_SaveRequest'])
        LabNoAllocError = checkForLabNoAllocError()

        time.sleep(SHORT_DELAY)
        pyautogui.press('enter')

        time.sleep(MED_DELAY)
    
    pyautogui.press('esc', presses=5, interval=0.1)
    time.sleep(LONGEST_DELAY)

def enterResults(Scenario:TestingScenario, modeMRI:bool=False):
    logging.debug("enterResults(): Start")
    focusWindowIfExists("WinPath Enterprise : ")
    RESULT_Y_START  = 305   
    logging.info(f"Proceeding to Results Entry for specimen {Scenario.LabNumber}")
    #pyautogui.click(*WinPath['Btn_ResultEntry'])
    pyautogui.hotkey('shift', 'F2')
    time.sleep(MED_DELAY)
    # pyautogui.click(*Results['Field_LabID'])
    # pyautogui.hotkey('ctrl', 'a')
    # pyautogui.press('delete')
    pyautogui.write(Scenario.LabNumber)
    pyautogui.press('enter')
    time.sleep(MED_DELAY)

    totalYOffset = 0
    for testSet in Scenario.requiredTestSets:
        _tmpResults = list(filter(lambda x: x.Profile == testSet[0], Scenario.targetResults))
        logging.debug(f"Processing {len(_tmpResults)} results for Profile/TLC [{testSet[0]}]...")
        for result in _tmpResults:
            yPos = RESULT_Y_START + totalYOffset + (RESULT_Y_PER_ANALYTE * result.AnalyteIndex)
            logging.info(f"Attempting to write value [{result.Value}] at y={yPos}, for analyte [{result.Analyte}] in profile [{result.Profile}].")
            pyautogui.click(x=RESULT_STATIC_X, y=yPos)
            time.sleep(SHORT_DELAY)
            pyautogui.write(str(result.Value))
            time.sleep(SHORT_DELAY)
        totalYOffset = totalYOffset + testSet[1]
        logging.debug(f"All results for Profile {testSet[0]} processed, total Y Offset is now {totalYOffset}(+{testSet[1]}).")
    time.sleep(MED_DELAY)
    
    pyautogui.click(*Results['Btn_Save'])
    time.sleep(MED_DELAY)
    pyautogui.click(*Results['Btn_Go'] )
    time.sleep(LONG_DELAY)

    pyautogui.click(*Results['Btn_Queue'])  # Force-queue sample
    logging.debug("Force-Queued sample.")
    time.sleep(SHORT_DELAY)

    pyautogui.press('enter')                # Dismiss any error message / #TODO: detect warning window via Win32API...
    pyautogui.press('esc', presses= 4)       # And if there wasn't an error popup and instead the sample entry was re-opened,  close it
    
def processSampleAuthQueue(Scenario:TestingScenario, takeScreenshot:bool = False) -> str:
    logging.info("processSampleAuthQueue(): Start")
    focusWindowIfExists("WinPath Enterprise : ")
    time.sleep(LONGEST_DELAY)               # Allow time for processing in background...
    #pyautogui.click(*WinPath['Btn_Authorisation'])
    pyautogui.hotkey('shift', 'F6')
    time.sleep(SHORT_DELAY)
    pyautogui.click(*Authorisation['Btn_Search'])
    time.sleep(SHORT_DELAY)
    pyautogui.write(Scenario.LabNumber)
    pyautogui.press('enter')
    time.sleep(MED_DELAY)
    #pyautogui.click(*Search['Area_Report'])
    
    logging.debug("Retrieving auth queue...")
    focusWindowIfExists("Search results")
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(SHORT_DELAY)
    pyautogui.hotkey('ctrl', 'c')
    windowVal = pyperclip.paste()
    Scenario.AuthQueue = [x.strip() for x in windowVal.split('\r\n') if x and "queue" in x]
    if takeScreenshot==True:
        logging.debug("Screenshotting auth queue...")
        captureScreenshot(f"{Scenario.SubScenarioID}_{Scenario.Phase}_0Queue", area=Authorisation['Area_Queue_Popup'])
    
    pyautogui.press('enter')
    pyautogui.press('esc', presses= 5)

def documentSampleReport(Scenario:TestingScenario, modeMRI:bool=False):
    logging.debug("documentSampleReport(): Start")
    focusWindowIfExists("WinPath Enterprise : ")
    #pyautogui.click(*WinPath['Btn_Search'])
    pyautogui.hotkey('shift', 'F3')
    time.sleep(MED_DELAY)
    pyautogui.click(*Search['Btn_RequestSearch'])
    time.sleep(SHORT_DELAY)
    pyautogui.click(*Search['Field_LabNoRange'])
    pyautogui.write(Scenario.LabNumber)
    pyautogui.press('tab')
    pyautogui.write(Scenario.LabNumber)
    pyautogui.click(*Search['Btn_Search'])
    time.sleep(MED_DELAY)
    pyautogui.press('enter')
    time.sleep(MED_DELAY)
    pyautogui.click(*Search['Btn_ExpandDetails'])
    time.sleep(SHORT_DELAY)

    if Scenario.SubScenarioID:
        captureScreenshot(f"{Scenario.SubScenarioID}_{Scenario.Phase}", area=Results['Area_Report'])
    
    else:
        captureScreenshot(f"{Scenario.ID}_{Scenario.Phase}_1Report", area=Search['Area_Report'])

def authoriseLastSampleOfQueue(Scenario:TestingScenario, TakeAuthScreeshot:bool=True, AuthIndex:int=1, SILLY_MODE:bool=False) -> None:
    FirstAuthQueue = Scenario.AuthQueue[0].split(" - ")
    FirstAuthQueue = list(map(lambda x: x.strip(), FirstAuthQueue))
    logging.debug(f"authoriseLastSampleOfQueue(): Aiming to authorise sample {Scenario.LabNumber} in queue {Scenario.AuthQueue}")
    focusWindowIfExists("WinPath Enterprise : ")
    pyautogui.click(*WinPath['Btn_Authorisation'])
    time.sleep(SHORT_DELAY)
    if FirstAuthQueue[1]=="PASS queue":
        X_Pos, Y_Pos = Authorisation['PASS_Queue_Start']
        Y_Pos = Y_Pos + int(AUTH_QUEUE_PER_Y * PASSQueues.index(FirstAuthQueue[2]))
        if SILLY_MODE:
            Y_Pos = Y_Pos + int(AUTH_QUEUE_PER_Y)
        Btn_List = Authorisation['Btn_List_PASS']

    elif FirstAuthQueue[1]=="FAIL queue":
        X_Pos, Y_Pos = Authorisation['FAIL_Queue_Start']
        Y_Pos = Y_Pos + int(AUTH_QUEUE_PER_Y * FAILQueues.index(FirstAuthQueue[2]))
        if SILLY_MODE:
            Y_Pos = Y_Pos + int(AUTH_QUEUE_PER_Y)
        Btn_List = Authorisation['Btn_List_FAIL']
    
    else:
        raise Exception(f"TestQueue[1] is neither 'PASS queue' nor 'FAIL queue'. TestQueue: {Scenario.AuthQueue}")

    #logging.debug(f"authoriseLastSampleOfQueue(): Attempting to click at x={X_Pos}, y={Y_Pos} to hit queue {Scenario.AuthQueue}.")
    pyautogui.click(x=X_Pos, y=Y_Pos)
    pyautogui.click(*Btn_List)
    time.sleep(MED_DELAY)
    pyautogui.click(*Authorisation['Btn_AuthLists_OK']) #TODO: Optional? 

    time.sleep(MED_DELAY)
    pyautogui.click(*Authorisation['TopOfAuthQueueList'])
    pyautogui.hotkey('ctrl', 'end')
    pyautogui.click(*Authorisation['Btn_Authorise_List'])
    time.sleep(MED_DELAY)
    pyautogui.click(*Authorisation['Btn_AuthLists_OK'])
    time.sleep(LONG_DELAY)

    if TakeAuthScreeshot==True:
        pyautogui.click(*Authorisation['TopOfResults'])
        for i in range(1, AuthIndex+1):
            pyautogui.press("down", presses=1)
            time.sleep(SHORT_DELAY)
            logging.debug(f"Screenshotting Authorisation logic status, iteration {i} of {AuthIndex}...")
            captureScreenshot(f"{Scenario.SubScenarioID}_{Scenario.Phase}_AuthLogic_{i}", area=Authorisation['Area_Auth_Rules'])

    pyautogui.click(*Authorisation['Btn_Authorise_Report'])
    time.sleep(MED_DELAY)
    logging.debug("authoriseLastSampleOfQueue(): Sample should be authorised.")
    pyautogui.press('esc', presses= 5)

#TODO: def assignOffsetLabNo()? in case Next Free Lab No > current pre-assigned Lab No during testing?

def processScenarios(fileName:str, modeMRITest:bool=False, useLabNoOffset:bool=False):
    logFileExists = os.path.isfile(os.path.join(ROOTDIR, "Output/AutoTestingSession.log"))
    logFile = open(os.path.join(ROOTDIR, "Output/AutoTestingSession.log"), 'a')
    logging.info(f"Opening output file '{ROOTDIR}/Output/AutoTestingSession.log'...")
    if not logFileExists:
        logFile.write("Timestamp\tPatient ID\tPhase\tPatient Name\tSex\tSample Origin\tPatient Flags\t"
        "Clinical Details\tAdditional Notes\tSample ID\tSample Collected\tSample Received\tTarget Result\tAuth Queue\tScenario ID\n")
        logFile.flush()

    logging.info(f" ==== Auto-Resulting Tool v{VERSION} - CliniSys Branch")
    logging.info("processScenarios(): Loading Scenarios from R-Script generated file...")
    TestingScenario.parseAll(fileName)

    logging.info(f"processScenarios(): {len(SCENARIOS)} scenarios loaded from file.")
    logging.info("processScenarios(): Start processing Scenarios...")

    if useLabNoOffset:
        #TODO: Check length of scenarios to process. Create offset. 
        TESTINGOFFSET = 0
        pyautogui.click(*WinPath['Btn_RequestEntry'])
        time.sleep(MED_DELAY)
        logging.debug("Request entry opened.")
        pyautogui.click(*RequestEntry['Btn_UseNextFreeLabNo'])
        time.sleep(SHORT_DELAY)
        NextLabNo = retrieveViaClipboard(RequestEntry['Field_Labno_Start'], RequestEntry['Field_Labno_End'])
        assert NextLabNo != ""
       
        TESTINGOFFSET = int(NextLabNo[3:])
        TESTINGOFFSET = round(TESTINGOFFSET + len(SCENARIOS), -2)
        CUR_YEAR = datetime.datetime.now().strftime("%y")

        logging.info(f"Preparing lab number offset: Next lab no {NextLabNo} => Offset to {CUR_YEAR}B{TESTINGOFFSET:08d}- {CUR_YEAR}B{(TESTINGOFFSET+len(SCENARIOS)):08d}.")
        pyautogui.press('esc', presses= 5)

        for x in range(0, len(SCENARIOS)):
            SCENARIOS[x].LabNumber = f"{CUR_YEAR}B{(TESTINGOFFSET + x):08d}"

    for _TestingScenario in SCENARIOS: #Replace with num counter, get next scenario to clear list only when needed
        logging.info(f"processScenarios(): Processing Scenario {_TestingScenario.SubScenarioID}-{_TestingScenario.Phase} from {_TestingScenario.Location}")
        
        requestSample(Scenario=_TestingScenario, modeMRI= modeMRITest, useOffset=useLabNoOffset)
        time.sleep(SHORT_DELAY)
        enterResults(Scenario=_TestingScenario, modeMRI= modeMRITest)
        time.sleep(SHORT_DELAY)
        processSampleAuthQueue(Scenario=_TestingScenario, takeScreenshot=False)
        time.sleep(SHORT_DELAY)
        authoriseLastSampleOfQueue(_TestingScenario, SILLY_MODE=False)
        time.sleep(SHORT_DELAY)
        documentSampleReport(_TestingScenario, modeMRI=modeMRITest)     
        
        outStr = (  f"{timestamp(forFile=False)}\t{_TestingScenario.ID}\t{_TestingScenario.Phase}\t"
                    f"{DEFAULT_LASTNAME}, {_TestingScenario.ID}\t{_TestingScenario.ScenarioSex}\t{_TestingScenario.Location}\t"
                    f"{';'.join(_TestingScenario.PatientTags)}\t{_TestingScenario.ClinicalDetails}\t{_TestingScenario.ClinNotes}\t"
                    f"{_TestingScenario.LabNumber}\t{_TestingScenario.SampleTaken.strftime('%d/%m/%Y %H:%M')}\t"
                    f"{_TestingScenario.SampleReceived.strftime('%d/%m/%Y %H:%M')}\t"
                    f"{_TestingScenario.targetResultStr}\t{';'.join(_TestingScenario.AuthQueue)}\t{_TestingScenario.SubScenarioID}"
        )
        logging.info(outStr)
        logFile.write(outStr + "\n")
        logFile.flush()
        
        pyautogui.press('esc', presses=6)
        time.sleep(SHORT_DELAY)

    logging.info(f"All scenarios processed. Have a nice day!")

#pyautogui.displayMousePosition()
#openWinPath()
#processScenarios(fileName="Output/ScriptDigest_MRI.tsv", modeMRITest=True)
#processScenarios(fileName="Output/ScriptDigest_CMRI-UHNM.tsv", modeMRITest=False)
#processScenarios(fileName="Output/ScriptDigest_VB12.tsv", modeMRITest=False)
#processScenarios(fileName="Output/ScriptDigest_XAN.tsv", modeMRITest=False)


#LOG IN AS MACCLESFIELD / LEIGHTON FOR:
#processScenarios(fileName="Output/ScriptDigest_B12M.tsv", modeMRITest=False)
#processScenarios(fileName="Output/ScriptDigest-MRI-MCHT.tsv", modeMRITest=True)
processScenarios(fileName="Output/ScriptDigest_CMRI-MCHT.tsv", modeMRITest=False)
#closeWinPath()

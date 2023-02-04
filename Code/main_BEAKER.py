PROGRAM = "autoTester v1.2"

import configparser
import datetime
import logging
from math import floor
import os
import pyautogui
import pydirectinput
import pyperclip
import subprocess
import time
import win32gui
import re

DEFAULT_SUBMITTER= "GP Surgery"
DEFAULT_LASTNAME = "UKASBIO18"
ROOTDIR = "S:/CS-Pathology/BIOCHEMISTRY/Z Personal/LKBecker/Code/Python/autoTester/"
LOGFORMAT = '%(asctime)s: %(name)-18s:%(levelname)-7s:%(message)s'
RESULT_FIELD_X_POS = 850
SHORTEST_DELAY = 0.5
SHORT_DELAY = 1.0
MED_DELAY   = 1.5
LONG_DELAY  = 2.25
screenWidth, screenHeight = pyautogui.size()
pyautogui.PAUSE = 0.15

config = configparser.ConfigParser()
config.read('./config.ini')
LOCATION = config['Default']['Context']
LOCATION = LOCATION[:LOCATION.find('[')-1]

logging.basicConfig(filename=os.path.join(ROOTDIR, 'debug.log'), filemode='w', level=logging.DEBUG, format=LOGFORMAT)

class TargetResult():
    def __init__(self, analyte:str, result:float, status:str) -> None:
        self.Analyte = analyte
        self.Value = result
        self.Status = status

    def __repr__(self):
        return f"TargetResults({self.Analyte},{self.Value},{self.Status})"


class TestingScenario():
    def __init__(self, textLine) -> None:
        self.targetResults = []
        textLine = textLine.split("\t")
        self.ScriptID = textLine[0]
        self.Phase    = int(textLine[1])
        self.PatientID = textLine[2]
        self.Scenario = textLine[3]
        self.Test = textLine[4]
        self.TestCode = textLine[5]
        self.PatientSex = textLine[6]
        self.PatientDTBirth = datetime.datetime.strptime(textLine[7], "%Y-%m-%d %H:%M")
        targetResults = textLine[8].split(";")
        self.extraInfo = textLine[9].strip()
        self.specimenID = None
        self.Intrument = None
        self.Requisition = None
        for targetResult in targetResults:
            analyte, result = targetResult.split("=")
            analyte = analyte.strip()
            result = result.strip()
            status = "normal"
            if result[-6:] == "(calc)":
                status = "calculated"
                result = result[:-6].strip()
            if result == "LLN":
                raise Exception(f"LLN value received for {analyte} - R script must be rewritten to handle this case! Aborting progam.")
            try:
                result = float(result)
            except ValueError:
                result = result
            self.targetResults.append(TargetResult(analyte, result, status))

    def __str__(self):
        return f"Scenario({self.ScriptID},{self.TestCode},{self.PatientID},{self.PatientSex},{self.PatientDTBirth},{len(self.targetResults)} target results)"

    def __repr__(self):
        return f"Scenario({self.ScriptID},{self.TestCode},{self.PatientID},{self.PatientSex},{self.PatientDTBirth},{len(self.targetResults)} target results)"


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


class TestYCoordinate():
    Analytes = {}
    AnalyteCounter = 0
    def __init__(self, dataLine) -> None:
        dataLine = dataLine.split("\t")
        dataLine = list(map(lambda x: x.strip(), dataLine))
        self.TestCode = dataLine[1]
        self.Analyte  = dataLine[2].strip()
        self.YOffset  = float(dataLine[5])
        self.PostScroll = dataLine[4] == "TRUE"
        self.CombinedSetIndex = int(dataLine[6])

        if not dataLine[1] in TestYCoordinate.Analytes.keys():
            TestYCoordinate.Analytes[self.TestCode] = dict()
        TestYCoordinate.Analytes[self.TestCode][self.Analyte] = self
        TestYCoordinate.AnalyteCounter = TestYCoordinate.AnalyteCounter + 1

    def calcYPos(self) -> int:
        assert(self.YOffset>0)
        return floor(302+(45.5*(self.YOffset-1)))


console = logging.StreamHandler()
console.setLevel(logging.INFO)
console.setFormatter(logging.Formatter(LOGFORMAT))
logging.getLogger().addHandler(console)
logging.info("Starting program...")

WindowManager = WindowMgr()

#assert(screenWidth == 1920)
#assert(screenWidth == 1080)

#These are HARDCODED values assuming a 1920 * 1080 screen. My screen, to be precise. It _should_ work elsewhere, but no promises...
RequisitionEntry = dict()
RequisitionEntry['DOBFieldStart']           = (295,  245)
RequisitionEntry['DOBFieldEnd']             = (500,  245)
RequisitionEntry['TOBField']                = (633,  245)
RequisitionEntry['SexField']                = (677,  217)
RequisitionEntry['ProcedureField']          = (162,  485)
RequisitionEntry['CreateSpecimenButton']    = (211,  610)
RequisitionEntry['ReceiveButton']           = (285,  604)
RequisitionEntry['LabReqCommentField']      = (1293, 668)
RequisitionEntry['AcceptAndNewButton']      = (221,  120)
RequisitionEntry['CollectionDateField']     = (780,  658)
RequisitionEntry['CollectionTimeField']     = (873,  658)
RequisitionEntry['ExternalIDField']         = (1044, 658)
RequisitionEntry['SpecimenIDStart']         = (152,  654)
RequisitionEntry['SpecimenIDStop']          = (722,  654)

PatientLookup = dict()
PatientLookup['SurnameBox']                 = (663,  512)
PatientLookup['ForenameBox']                = (1032, 512)
PatientLookup['FindPatientBtn']             = (710,  648)
PatientLookup['NoPatientMatchConfirmBtn']   = (1049, 562)
PatientLookup['NewPatientBtn']              = (592,  650)
PatientLookup['BirthDateField']             = (1032, 543)
PatientLookup['SexField']                   = (665,  546)

Beaker = dict()
Beaker['InBasket']                          = (1720, 13)
Beaker['Search']                            = (1900, 60)
Beaker['MyDashboard']                       = (  20, 60)
Beaker['LogoutButton']                      = ( 880, 35)

Results = dict()
Results['InstrumentStart']                  = (1588, 170)
Results['InstrumentEnd']                    = (1742, 170)
Results['RequisitionStart']                 = (1752, 170)
Results['RequisitionEnd']                   = (1900, 170)
Results['AcknowledgeWarnings']              = (1814, 971)
Results['SummaryButton']                    = (1375, 290)
Results['OpenSnapshotButton']               = (1900, 600)
Results['SnapshotScrollDownButton']         = (1905, 965)
Results['TrackingPaneButton']               = (1450, 290)
Results['TrackingPaneContent']              = (1450, 350)
Results['AcknowledgeWarnings']              = (1814, 971)
Results['ConfirmFinal']                     = (1860, 1020)
Results['Verify']                           = (1860, 1020)
Results['TestScrollDownNoSnapshot']         = (1875, 965)
Results['TestScrollDownSnapshot']           = (1275, 950)
Results['TestScrollUpNoSnapshot']           = (1875, 290)
Results['TestScrollUpSnapshot']             = (1275, 300)
Results['ClearClinicalHoldButton']          = (1720, 330)
Results['AcceptClearHoldButton']            = (1325, 735)

def timestamp(forFile:bool=False):
    if forFile:
        return datetime.datetime.now().strftime("%y%m%d-%H%M%S")
    return datetime.datetime.now().strftime("%y-%m-%d %H:%M:%S")

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

def openBeaker():
    loginWindow = focusWindowIfExists("Hyperspace - Testing")
    beakerLoggedIn = False
    if not loginWindow:
        beakerLoggedIn = focusWindowIfExists(f"Hyperspace - {LOCATION} - Testing")
    
    if (loginWindow == False) and (beakerLoggedIn == False):
        logging.info("Opening CITRIX endpoint...")
        subprocess.call(f"\"{config['Default']['AppPath']}\"" + f" -qlaunch \"{config['Default']['AppName']}\"", shell=True)
        time.sleep(2)
        waitForWindow(f"Hyperspace - Testing", 10)
        loginWindow = True
    
    if (loginWindow == True):
        focusWindowIfExists("Hyperspace - Testing")
        pyautogui.click(x=794, y=456)
        pyautogui.typewrite(config['Default']['Username'])
        pydirectinput.press('tab') #PyAutoGUI uses older keypresses, which Citrix does not recognise.
        pyautogui.typewrite(config['Default']['Password'])
        pydirectinput.press('enter')
        time.sleep(1)
        pyautogui.typewrite(config['Default']['Context'])
        pydirectinput.press('enter')
        pydirectinput.press('enter')
    
    time.sleep(5)
    logging.info("Logged into Beaker.")

def useSearch(query):
    pyautogui.click(*Beaker['Search'])
    time.sleep(SHORT_DELAY)
    pyautogui.typewrite(query)
    pydirectinput.press('enter')
    time.sleep(MED_DELAY)

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
    bootlegHotkey('ctrl', 'c')
    time.sleep(.2)

    return pyperclip.paste()

def bootlegHotkey(key1:str, key2:str, delay:float=0.25):
    pydirectinput.keyDown(key1)
    pydirectinput.keyDown(key2)
    time.sleep(delay)
    pydirectinput.keyUp(key2)
    pydirectinput.keyUp(key1)

def closeBeaker():
    pyautogui.click(*Beaker['LogoutButton'])
    time.sleep(SHORT_DELAY)
    pyautogui.click(1900,5)

def processScenarios(phase=1):
    logging.info(f"processScenarios(): Loading Phase {phase} Scenarios...")
    HAS_OPENED_SNAPSHOT = False
    
    with open(os.path.join(ROOTDIR, 'testYIndices.txt'), 'r') as ResultCoords:
        CoordData = ResultCoords.readlines()
        for line in CoordData[1:]:
            CoordObj = TestYCoordinate(line)
                  
    logging.info(f"processScenarios(): {TestYCoordinate.AnalyteCounter} analyte Y-indices loaded from file.")
    
    Scenarios = []
    with open(os.path.join(ROOTDIR, f'Script18_Digest_Phase{phase}.tsv'), 'r') as ScenarioDataFile:
        ScenarioLines = ScenarioDataFile.readlines()
        ScenarioLines = ScenarioLines[1:]
        for line in ScenarioLines:
            try:
                Scenarios.append( TestingScenario(line) )
            except Exception:
                continue
    logging.info(f"processScenarios(): {len(Scenarios)} scenarios loaded from file.")

    logFileExists = os.path.isfile(os.path.join(ROOTDIR, "AutoTestingSession.log"))
    logFile = open(os.path.join(ROOTDIR, "AutoTestingSession.log"), 'a')
    logging.info(f"Opening output file '{ROOTDIR}/AutoTestingSession.log'...")
    if not logFileExists:
        logFile.write("Timestamp\tScenario ID\tPatient ID\tScenario\tRequisition\tInstrument\tSample ID\tComments\n")
        logFile.flush()

    logging.info("processScenarios(): Start processing Scenarios...")
    ScenarioCounter = 0
    for _TestingScenario in Scenarios: #Replace with num counter, get next scenario to clear list only when needed
        changedDOB = False
        ScenarioCounter = ScenarioCounter + 1
        logging.info(f"processScenarios(): Processing Scenario #{ScenarioCounter}, {_TestingScenario.ScriptID}")

        if not _TestingScenario.TestCode in TestYCoordinate.Analytes.keys():
            logging.error(f"No data for Test Code {_TestingScenario.TestCode} in Analytes dict. Please check input file and retry.")
            continue
        
        useSearch("Requisition Entry")

        logging.info(f"processScenarios(): Creating Request from '{DEFAULT_SUBMITTER}'...")
        pyautogui.click(x=300, y=156, clicks=1, button="left")
        pyautogui.typewrite(DEFAULT_SUBMITTER)
        pydirectinput.press('enter')
        time.sleep(SHORT_DELAY)
        logging.info(f"processScenarios(): Attempting to locate patient [{DEFAULT_LASTNAME}, {_TestingScenario.PatientID}], DOB {_TestingScenario.PatientDTBirth.strftime('%d/%m/%y %H:%M')}")
        pydirectinput.press('tab')
        pydirectinput.press('tab')
        pyautogui.typewrite(DEFAULT_LASTNAME)
        pydirectinput.press('tab')
        pyautogui.typewrite(_TestingScenario.PatientID)
        pydirectinput.press('enter')
        time.sleep(MED_DELAY)

        PatientSelectWindow = focusWindowIfExists("Patient Select")
        if PatientSelectWindow:
            logging.info("Patient appears to have been found. Loading patient profile...")
            pydirectinput.press('enter')
            time.sleep(MED_DELAY)

        PatientSearchWindow = False
        if not PatientSelectWindow:
            PatientSearchWindow = focusWindowIfExists("Patient Search")
            if PatientSearchWindow:
                logging.info("No matching patient was found. Creating entry...")
                pydirectinput.press('enter')
                pyautogui.click(PatientLookup['NoPatientMatchConfirmBtn'])
                pyautogui.click(PatientLookup['BirthDateField'])
                pyautogui.typewrite(_TestingScenario.PatientDTBirth.strftime("%d/%m/%Y"))
                pyautogui.click(PatientLookup['SexField'])
                pyautogui.typewrite(_TestingScenario.PatientSex)
                pyautogui.click(PatientLookup['NewPatientBtn'])
                time.sleep(MED_DELAY)

        if not PatientSelectWindow and not PatientSearchWindow:
            raise Exception("Neither Patient Selection nor Patient Search could be located. Rewrite software and try again.")

        #TODO: Birthday adjustment, reason code, bla bla bla
        SetDOB = retrieveViaClipboard(RequisitionEntry['DOBFieldStart'], RequisitionEntry['DOBFieldEnd'])
        SetDOB = datetime.datetime.strptime(SetDOB, "%d/%m/%Y")
        if not (_TestingScenario.PatientDTBirth.date() == SetDOB.date()):
            pyautogui.click(*RequisitionEntry['DOBFieldStart'], clicks=2)
            pyautogui.typewrite(_TestingScenario.PatientDTBirth.strftime("%d/%m/%Y"))
            changedDOB = True

        #Now to create the sample and request...
        logging.info(f"Patient loaded. Adjusting 'Time Of Birth' to {_TestingScenario.PatientDTBirth.strftime('%H:%M')}...")
        pyautogui.click(RequisitionEntry['TOBField'])
        pyautogui.typewrite(_TestingScenario.PatientDTBirth.strftime("%H:%M"))        
        pyautogui.click(RequisitionEntry['ProcedureField'])
        logging.info(f"Submitting request for test [{_TestingScenario.TestCode}]")
        pyautogui.typewrite(_TestingScenario.TestCode)
        pydirectinput.press('enter')
        time.sleep(MED_DELAY)
        
        #OrderWindow = focusWindowIfExists(r"Order Search - \\Remote")
        #if OrderWindow:
        #    pydirectinput.press('enter') #THIS HACK ASSUMES A GOOD MATCH, AND ALWAYS selects the FIRST entry!
        
        pyautogui.click(RequisitionEntry['CreateSpecimenButton'])

        _TestingScenario.specimenID = retrieveViaClipboard(RequisitionEntry['SpecimenIDStart'], RequisitionEntry['SpecimenIDStop'])
        logging.info(f"Specimen ID [{_TestingScenario.specimenID}] has been generated for scenario {_TestingScenario.ScriptID}.")

        if _TestingScenario.Phase == 1:
            SampleCollectionDT = datetime.datetime.now()-datetime.timedelta(minutes=45)
        else:
            SampleCollectionDT = datetime.datetime.now()
        pyautogui.click(RequisitionEntry['CollectionDateField'])
        pyautogui.typewrite(SampleCollectionDT.strftime("%d/%m/%y"))
        pyautogui.click(RequisitionEntry['CollectionTimeField'])
        pyautogui.typewrite(SampleCollectionDT.strftime("%H:%M"))
        #pyautogui.click(RequisitionEntry['ExternalIDField'])
        #pyautogui.typewrite()
        del(SampleCollectionDT)
        
        logging.debug("Receiving specimen...")
        pyautogui.click(RequisitionEntry['ReceiveButton'])
        logging.debug("Accepting requisition...")
        pyautogui.click(RequisitionEntry['AcceptAndNewButton'])
        time.sleep(MED_DELAY)

        DemographicsPrompt = focusWindowIfExists("Requisition Entry") #TODO: Could there be other popups with the same title?
        if DemographicsPrompt:
            logging.info(f"Located demographics update prompt, agreeing. Changed DOB: {changedDOB}")
            pyautogui.click(x=979, y=572, clicks=1, button="left") #agree to update demographics
            if changedDOB == True:
                time.sleep(MED_DELAY)
                #Sometimes - not sure what triggers it - it demands a reason for changing hte patient's identity. 
                #change of DOB more than 4 years - triggers
                logging.info("Checking for Identity Change prompt...")
                IDChangeWindow = focusWindowIfExists("Reason for Identity Change")
                if IDChangeWindow:
                    pyautogui.typewrite('106')
                    pydirectinput.press('enter')
                    time.sleep(SHORT_DELAY)
            time.sleep(SHORT_DELAY)
        
        logging.info("Proceeding to Results Entry...")
        useSearch("Result Entry and Verification")
        time.sleep(SHORT_DELAY)
        
        #TODO: ASSUMING the report window always pops - can you check?
        pydirectinput.press('esc')
        pyautogui.typewrite(_TestingScenario.specimenID)
        pydirectinput.press('enter')
        time.sleep(MED_DELAY)

        if HAS_OPENED_SNAPSHOT == False:
            pyautogui.click(*Results['OpenSnapshotButton'])
            HAS_OPENED_SNAPSHOT = True
            time.sleep(MED_DELAY)
        
        time.sleep(SHORT_DELAY)
        _TestingScenario.Instrument = retrieveViaClipboard(Results['InstrumentStart'], Results['InstrumentEnd'])
        logging.info(f"Instrument ID has been retrieved as {_TestingScenario.Instrument}.")
        time.sleep(SHORTEST_DELAY)
        _TestingScenario.Requisition = retrieveViaClipboard(Results['RequisitionStart'], Results['RequisitionEnd'])
        logging.info(f"Requisition has been retrieved as {_TestingScenario.Requisition}.")

        bootlegHotkey('alt', 'j')
        time.sleep(SHORTEST_DELAY)
        
        LAST_RESULT_SCROLLDOWN = False
        for result in _TestingScenario.targetResults:
            if not result.Analyte in TestYCoordinate.Analytes[_TestingScenario.TestCode]:
                raise KeyError(f"No data for Analyte {result.Analyte} under {_TestingScenario.TestCode} in Analytes dict. Please check input file and retry.")
            _YCoordObj = TestYCoordinate.Analytes[_TestingScenario.TestCode][result.Analyte]

            if _YCoordObj.CombinedSetIndex != 0:
                logging.info("Result is part of a combined set, selecting appropriate result")
                pyautogui.click(x=90, y=(300+(30*(_YCoordObj.CombinedSetIndex-1))))
                time.sleep(SHORT_DELAY)
                logging.debug("Re-launching result entry...")
                bootlegHotkey('alt', 'j')
                time.sleep(SHORT_DELAY)


            if _YCoordObj.PostScroll == True and LAST_RESULT_SCROLLDOWN == False:
                if HAS_OPENED_SNAPSHOT == True:
                    logging.debug("Result entry requires scrolling to bottom, snapshot pane OPEN")
                    pyautogui.moveTo(*Results['TestScrollDownSnapshot'])
                else:
                    logging.debug("Result entry requires scrolling to bottom, snapshot pane CLOSED")
                    pyautogui.moveTo(*Results['TestScrollDownNoSnapshot'])
                pyautogui.mouseDown()
                time.sleep(MED_DELAY)
                pyautogui.mouseUp()

            if _YCoordObj.PostScroll == False and LAST_RESULT_SCROLLDOWN == True:
                if HAS_OPENED_SNAPSHOT == True:
                    pyautogui.moveTo(*Results['TestScrollUpSnapshot'])
                    logging.debug("Result entry requires scrolling up, snapshot pane OPEN")
                else:
                    pyautogui.moveTo(*Results['TestScrollUpNoSnapshot'])
                    logging.debug("Result entry requires scrolling up, snapshot pane CLOSED")
                pyautogui.mouseDown()
                time.sleep(MED_DELAY)
                pyautogui.mouseUp()

            time.sleep(SHORT_DELAY)
        
            YPosition = _YCoordObj.calcYPos()
            logging.info(f"Aiming to enter value {result.Value} for analyte {result.Analyte}, x={RESULT_FIELD_X_POS}, y={YPosition}")
            pyautogui.click(RESULT_FIELD_X_POS, YPosition)
            time.sleep(SHORTEST_DELAY)

            pyautogui.typewrite(str(result.Value))
            LAST_RESULT_SCROLLDOWN = _YCoordObj.PostScroll
        
        #save results
        bootlegHotkey('alt', 'j')
        time.sleep(SHORT_DELAY)

        #this MAY OR MAY NOT open a popup...
        pydirectinput.press('enter') #TODO: Alternative is ALT O, esc

        #saving results collapses pane, so -
        pyautogui.click(*Results['OpenSnapshotButton'])
        time.sleep(LONG_DELAY)

        #open tracking pane
        logging.debug(f"Attempting to press Open Snapshot at {Results['OpenSnapshotButton']}...")
        pyautogui.click(*Results['TrackingPaneButton'])
        time.sleep(SHORT_DELAY)
        logging.debug(f"Scrolling to end of Snapshot pane...")

        pyautogui.moveTo(*Results['SnapshotScrollDownButton'])
        pyautogui.mouseDown()
        time.sleep(SHORT_DELAY)
        pyautogui.mouseUp()
        time.sleep(SHORT_DELAY)
        
        captureScreenshot(_TestingScenario.ScriptID, area=(680, 100, 1230, 890))

        if _TestingScenario.Phase == 2:
            pyautogui.click(*Results['Verify'])
            time.sleep(SHORT_DELAY)
            # try:
            #     ClinicalHoldLocation = pyautogui.locateOnScreen('Images/ClinApprovalRequired.png', 
            #                                                     region=(824, 320, 269, 24))
            #     if ClinicalHoldLocation:
            #         logging.debug(f"CLINICAL HOLD LOCATED: {ClinicalHoldLocation}")
            #     else:
            #         logging.debug("ClinicalHoldLocation object created but not True.")
            #         print(ClinicalHoldLocation)
            # except pyautogui.ImageNotFoundException:
            #     logging.debug("Attempted to locate Clinical Hold; we ain't found shit.")

            pyautogui.click(*Results['ClearClinicalHoldButton'])
            time.sleep(SHORT_DELAY)
            ClearClinicalHoldWindow = focusWindowIfExists("Clear Hold")
            if ClearClinicalHoldWindow:
                logging.info("Overriding Clinical Hold")
                pydirectinput.press('tab')
                pyautogui.typewrite('5')
                pydirectinput.press('tab')
                time.sleep(SHORTEST_DELAY)
                pyautogui.typewrite('Automated override for Beaker testing (LKB)')
                pydirectinput.press('tab')
                pydirectinput.press('enter')
                time.sleep(SHORT_DELAY)
            pyautogui.click(*Results['AcknowledgeWarnings'])
            time.sleep(SHORT_DELAY)
            pyautogui.click(*Results['ConfirmFinal']) #This disappears the sample
            time.sleep(SHORT_DELAY)

        bootlegHotkey('alt', 'c') #clear Worklist
        time.sleep(SHORT_DELAY)

        pyperclip.copy(_TestingScenario.specimenID)
        
        outStr = (f"{timestamp()}\t{_TestingScenario.ScriptID}\t{DEFAULT_LASTNAME}, {_TestingScenario.PatientID}\t"
                    f"{_TestingScenario.Scenario}\t{_TestingScenario.Requisition}\t"
                    f"{_TestingScenario.Instrument}\t{_TestingScenario.specimenID}\t\n")
        logging.info(outStr.strip())
        logFile.write(outStr)        
        logFile.flush()

        time.sleep(SHORT_DELAY)

    logging.info(f"All phase {phase} scenarios processed. Shutting down...")

#pyautogui.displayMousePosition()

#TO RUN THIS SCRIPT SUCCESSFULLY:
#Save fresh copy of Script 18 as Script18_Raw.xlsx
#Create 'worklist' by running script18_Extract.R
#run the below:
openBeaker()
processScenarios(phase=1)
#closeBeaker()

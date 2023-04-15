#Libraries====
.libPaths("C:/Users/BECKEL201/Documents/R Libraries/") #Sets local libraries to avoid MASSIVE network lag...
#Why do R libraries on managed machines by default install into a faraway network share? Because frick you that's why
library(assertthat)
library(readxl)
library(data.table)
library(stringr)
library(lubridate)

#Constants====
WEEKSPERYEAR = 52.25 #Rougher but gives "expected" dates. More or less.
DAYSPERMONTH = 30.5 # Sorry Ceri
RENAL_LOCATIONS = c("RSW124", "RSWHU", "RSWKU", "CHRU", "MDHRDU", "LHDU1", "LHDU2", "LHDU3", "LHDU4", "LHDU5", "LHDU")
minHour <- 09

#Functions====

SplitDataTableWithMultiRows<-function(DataTable, TargetColumnIndex, Separator=","){
    TEMPCOL <- NULL
    if(!("data.table" %in% class(DataTable))){ DT<- data.table(DataTable) }
    else { DT <- copy(DataTable) }
    if(class(TargetColumnIndex)=="character") {
        if (!TargetColumnIndex %in% colnames(DT)) { stop(paste0("Error: '", TargetColumnIndex,
                                                                "' is not a valid column in the table submitted to SplitDataTableWithMultiRows")) }
        TargetColumnIndex <- which(colnames(DT) == TargetColumnIndex)
    }
    DTColOrder <- copy(colnames(DT))
    TargetCol <- DTColOrder[TargetColumnIndex]
    DTF<-data.frame(DT)
    SplitRowData<-strsplit(as.character(DTF[,TargetColumnIndex]), Separator, fixed=TRUE)
    SplitRowData<-lapply(SplitRowData, function(x){if(length(x)==0){ return("")} else { return(x)}})
    if(TargetColumnIndex != 1){
        RemainingCols <- DTColOrder
        RemainingCols <- RemainingCols[-TargetColumnIndex]
        setcolorder(DT, c(TargetCol, RemainingCols))
    }
    nreps<-sapply(SplitRowData, function(x){max(length(x), 1)})
    out<-data.table(TEMPCOL= unlist(SplitRowData), DT[ rep(1:nrow(DT), nreps) , -1, with=F] )
    setnames(out, "TEMPCOL", TargetCol)
    if(TargetColumnIndex != 1){ setcolorder(out, DTColOrder) }
    rm(DT, DTF, DTColOrder, TargetCol, SplitRowData, nreps)
    return(out)
}

ts<-function(){ return(format(Sys.time(), "[%Y-%m-%d %H:%M:%S] -- ")) }

createTestScript <- function(Analyte, Excelfile, Excelrange, TestCoderange = TestCodeRange, 
                             Excelsheet="Scenarios", HasRunMode="Both", MakeMRISafe=TRUE){
    message(sprintf("createTestScript() - Running for:\nAnalyte\t%s\nExcelfile\t%s (%s - %s)\nHasRunMode\t%s",
                    Analyte, Excelfile, Excelsheet, Excelrange, HasRunMode))
    assert_that(assertthat::are_equal(HasRunMode=="Remove"|HasRunMode=="Increment"|HasRunMode=="Both", TRUE), 
                msg = "createTestScript(): HasRunMode must be one of 'Remove' or 'Increment'.")
    
    TODAY = Sys.Date()
    hour(TODAY) <- minHour
    LogFileOffset <- 0
    # ==== SCRIPT ====
    if(HasRunMode=="Increment"|HasRunMode=="Both"){
        if (file.exists("./Output/AutoTestingSession.log")){
            LogFile = fread("./Output/AutoTestingSession.log")
            LogFile = LogFile[grepl(Analyte, `Scenario ID`),unique(`Scenario ID`)]
            LogFileOffset = max(as.numeric(str_extract(LogFile, "\\d{3}")))
            message(sprintf("Loading AutoTestingSession.log - %d entries feature the current Analyte code. Setting LogFileOffset to %d",
                            length(LogFile), LogFileOffset+1))
            rm(LogFile)
            assert_that(is.numeric(LogFileOffset), 
                        msg=sprintf("HasRunMode 'Increment' cannot calculate LogFileOffset for analyte %s; getting {%s}", 
                                    Analyte, LogFileOffset))
        }
    }
    
    #Read in data====
    ScenarioData = readxl::read_xlsx(Excelfile, sheet = Excelsheet, range = Excelrange)
    ScenarioData = as.data.table(ScenarioData)
    ScenarioData[,Tab:=NULL] #Not needed
    
    ScenarioData[,SampleTaken := TODAY]
    ScenarioData[,SampleReceived := TODAY + as.difftime(10, units = "mins")]
    
    #Demarcate which scenarios need special attention
    ScenarioData[, Scenario:= str_replace_all(Scenario, ", prev", "<;><Prev>")]
    ScenarioData[grepl("<Prev>", Scenario), HasPrev:= TRUE]
    ScenarioData[is.na(HasPrev), HasPrev:= FALSE]
    
    setnames(ScenarioData, "Number", "PatientNo")
    ScenarioData[, PatientNo := PatientNo + LogFileOffset]
    #Need to pre-process patient numbers; post Phase-split processing fails to properly enumerate patients
    ScenarioData[!is.na(PtOverride), PatientNo := PtOverride + LogFileOffset]
    
    ScenarioData[,PatientID:= sprintf("%s-%03d", Analyte, PatientNo)]
    ScenarioData[,ScenarioID := sprintf("%s_%03d", PatientID, seq_len(.N)), by=PatientID]
    
    
    ScenarioData[,tmpResults:= substr(Scenario, str_locate(Scenario, "old with ")[,2]+1, str_length(Scenario))]
    ScenarioData[,tmpResults:= str_replace_all(tmpResults, " and ", ";")]
    
    ScenarioData[,Location:=str_match(toupper(Scenario), "/([A-Z0-9]+)$")[,2]]
    ScenarioData[Location=="RENAL", Location := sample(RENAL_LOCATIONS, 1)] #Replaces RENAL with random RENAL location
    message(sprintf("%d scenarios without assigned location are set to 'GP'.", ScenarioData[is.na(Location), .N]))
    ScenarioData[is.na(Location),Location:= "GP"] # all unknown / unassigned locations become GP patients
    
    #Extract Sex ====
    ScenarioData[,TargetSex:=str_match(Scenario, "^(F|M|X|U|I)( |,)")[,1]]
    ScenarioData[is.na(TargetSex), TargetSex:=ifelse(.I %% 2 == 0, "F", "M"), PatientNo]
    ScenarioData[,TargetSex:=str_trim(TargetSex, "both")]
    message(sprintf("%d scenarios without assigned sex remain.", ScenarioData[is.na(TargetSex), .N]))
    
    #Extract Age and calculate datetime ====
    #Age is defined at the start of the Scenario and always followed by " old", thus we extract:
    ScenarioData[,AgeStr:= substr(Scenario, 1, str_locate(Scenario, "old")[,2] - str_length(" old"))]
    ScenarioData[,tmpAge:=str_match(AgeStr, "(\\d{1,3}( )*(d|day|Days|hours old|H|D|Y|y|hour|m|month(s)*|M|W))")[,1]]
    
    message(sprintf("%d scenarios without assigned ages remain; assuming 'year'.", 
                    ScenarioData[is.na(tmpAge) & !HasPrev, .N]))
    
    ScenarioData[,tmpAgeNum:=as.numeric(str_extract(tmpAge, "\\d+"))]
    
    ScenarioData[grepl("H$|hour$|hours old$", tmpAge), TimeOfBirth:= TODAY - as.difftime(tmpAgeNum, units = "hours")]
    ScenarioData[is.na(TimeOfBirth) & grepl("d$|D$|day$|days$", tmpAge),       
                 TimeOfBirth:= TODAY - as.difftime(tmpAgeNum, units = "days")]
    ScenarioData[is.na(TimeOfBirth) & grepl("W$", tmpAge),                  
                 TimeOfBirth:= TODAY - as.difftime(tmpAgeNum, units = "weeks")]
    ScenarioData[is.na(TimeOfBirth) & grepl("M|m|month(s)*$", tmpAge),                  
                 TimeOfBirth:= TODAY - as.difftime(tmpAgeNum*DAYSPERMONTH, units = "days")]
    ScenarioData[is.na(TimeOfBirth) & grepl("Y$|y$|years$", tmpAge),        
                 TimeOfBirth:= TODAY - as.difftime(tmpAgeNum*WEEKSPERYEAR, units = "weeks")]
    ScenarioData[, TimeOfBirthStr := format(TimeOfBirth, "%Y-%m-%d")] #And format as string for the program
    
    #Extract PrevTime from scenarios with Prev
    ScenarioData[, PrevTime   := str_extract(Scenario, "\\(\\d{1,3} *days\\)")]
    ScenarioData[, PrevTime   := as.numeric(str_extract(PrevTime, "\\d{1,3}"))]
    ScenarioData[!is.na(PrevTime), PrevTimeDT := TODAY - as.difftime(PrevTime, units = "days")]
    
    #Ensure time does not go greater than _or equal to_ the n days of PrevTime (safety margin)
    ScenarioData[!is.na(PrevTime), PrevTimeDT := PrevTimeDT - as.difftime(1, units = "hours")] 
    
    # Adjust sets run multiple times in same patient to stay clear of Minimum Retest Intervals ====
    testCodes = readxl::read_xlsx("./221013_MinRetestTests v1.xlsx", sheet="TestCoords v2", range=TestCodeRange)
    testCodes = as.data.table(testCodes)
    ScenarioData[,tmpAnalyte := str_extract(tmpResults, "^[0-9\\w]+")]
    ScenarioData = merge(ScenarioData, testCodes[,.(TargetAnalyte, MinimumRetestInterval_Days)], 
                         by.x="tmpAnalyte", by.y="TargetAnalyte", all.x=T)
    setorder(ScenarioData, PatientID, ScenarioID)
    
    if (MakeMRISafe==TRUE){
        MultiProfileRuns = ScenarioData[, .N, .(PatientID, tmpAnalyte)][N>1]
        if (MultiProfileRuns[,.N] > 0) {
            for (i in 1:MultiProfileRuns[,.N]) {
                tmpAnalyte2 = MultiProfileRuns[i, tmpAnalyte]
                tmpPatient = MultiProfileRuns[i, PatientID]
                cases = ScenarioData[PatientID == tmpPatient & tmpAnalyte == tmpAnalyte2]
                safeMRI = cases[1, MinimumRetestInterval_Days] + 2
                assertthat::noNA(safeMRI)
                safeMRI_DT = as.difftime(safeMRI, units="days")
                minDT = cases[, max(SampleTaken)]
    
                for (j in cases[,.N]:1){ #HACK: go in reverse order to preserve causality.
                    tempScenario = cases[j, ScenarioID]
                    
                    ScenarioData[ScenarioID == tempScenario, SampleTaken := minDT]
                    if (!is.na(cases[j, PrevTime])){
                        minDT = min(minDT - safeMRI_DT, minDT - as.difftime(safeMRI + cases[j, PrevTime], units = "days"))
                    } else {
                        minDT = minDT - safeMRI_DT
                    }
                    
                    message(sprintf("Adjusting DateTimes for [%s], analyte [%s] by [%d] days, from [%s] to [%s].", 
                                    tempScenario, tmpAnalyte2, safeMRI, cases[j, SampleTaken], minDT))
                    
                }
            }
            rm(i,j, minDT, cases, safeMRI, safeMRI_DT, tmpAnalyte2, tmpPatient)
        }
    }
    ScenarioData[,MinimumRetestInterval_Days := NULL]
    
    #Split delta check scenarios into two rows====
    ScenarioData = SplitDataTableWithMultiRows(ScenarioData, "tmpResults", Separator="<;>") 
    #Split off prevs into their own line
    
    #Assign phased on a) Do you have a Prev? b) are you the Prev?
    ScenarioData[HasPrev==TRUE & grepl("\\<Prev\\>", tmpResults), Phase:= 1]
    ScenarioData[HasPrev==TRUE & is.na(Phase), Phase:= 2]
    ScenarioData[is.na(Phase), Phase:= 1]
    
    ScenarioData[HasPrev==TRUE & Phase == 1, SampleTaken := PrevTimeDT]
    ScenarioData[HasPrev==TRUE & Phase == 1, SampleReceived := PrevTimeDT + as.difftime(5, units = "mins")]
    
    #Ensure all phase 2 scenarios are *at least* 5 minutes after the phase 1 scenario unless already indicated in days
    ScenarioData[is.na(PrevTime) & Phase==2, SampleTaken := SampleTaken + as.difftime(5, units = "mins")]
    ScenarioData[is.na(PrevTime) & Phase==2, SampleReceived := SampleReceived + as.difftime(5, units = "mins")]
    
    ScenarioData[,HasPrev := NULL]
    
    #Cleanup====
    ScenarioData[, tmpResults := str_replace(tmpResults, pattern = "\\<Prev\\> ", replacement = "")]
    ScenarioData[, tmpResults := str_replace(tmpResults, pattern = " \\(\\d{1,3} *days\\)", replacement = "")]
    ScenarioData[, tmpResults := str_replace(tmpResults, "/(GP|OP|IP|gp|op|ip)$", "")]
    
    ScenarioData[, tmpResults2:= str_replace_all(tmpResults, pattern = "(\\w+) ((<|>)*\\d+\\.*\\d*|Yes|No)", replacement = "\\1=\\2")]
    ScenarioData[!grepl(pattern = "=", x = tmpResults2), 
                 tmpResults2:= str_replace_all(tmpResults2, pattern = "(\\w+) ((<|>)*\\w+)", replacement = "\\1=\\2")]
    
    ScenarioData[, tmpResults := tmpResults2]
    
    assertthat::are_equal(ScenarioData[!grepl("=", tmpResults), .N], 0)
    
    ScenarioData[, `:=`(AgeStr=NULL, tmpAgeNum=NULL, tmpAge=NULL, TimeOfBirth=NULL, 
                        PrevTime=NULL, PrevTimeDT=NULL, tmpResults2=NULL)]
    
    ScenarioData = SplitDataTableWithMultiRows(ScenarioData, "tmpResults", Separator=";")
    ScenarioData[,tmpResults := str_trim(tmpResults)]
    ScenarioData[,Analyte:=str_extract(tmpResults, "\\w+(?==)")]
    ScenarioData[,Result:=str_extract(tmpResults, "((<|>)*\\d+\\.*\\d*|Yes|No)$")]
    ScenarioData[,tmpResults := NULL]
    
    setcolorder(ScenarioData, c("PatientID", "Phase", "Analyte", "Result", "TargetSex", 
                                "Location", "SampleTaken", "SampleReceived", "TimeOfBirthStr", "Scenario", "ScenarioID"))
    
    #Load test codes and merge into table====
    ScenarioData = merge(ScenarioData, testCodes, by.x="Analyte", by.y="TargetAnalyte", all.x=T)
    ScenarioData[,CliniSysName:=NULL]
    rm(testCodes)
    
    #setorder(ScenarioData, PatientID, Phase, DisciplineIdx, ProfileIdx, AnalyteIdx) 
    #Ensures tests that need to be booked first _are_ booked first?
    
    #Check for any that do not yet have coordinates====
    NNoYCoords = ScenarioData[is.na(AnalyteIdx), .N]
    if (NNoYCoords > 0) {
        AnalytesWithoutY = ScenarioData[is.na(AnalyteIdx), unique(Analyte)]
        tmpMsg = sprintf("Cannot find Analyte Index for the following analytes / sets:\n\t%s",
                         paste0(AnalytesWithoutY, collapse="\n\t"))
        stop(tmpMsg)
    }
    
    #TODO: Process merges / offsets between different sets====
    #stop("Function not yet implemented")
    
    #Merge multi-set requests back into a single item, to be processed by Python====
    CompactScenario = copy(ScenarioData)
    if (!("PtOverride" %in% colnames(CompactScenario))){
        CompactScenario[,PtOverride := NA]
    }
    CompactScenario[,PyStr := sprintf("%s|%s|%s|%s|%s|%s|%s", 
                                      Profile, Analyte, Result, DisciplineIdx, ProfileIdx, AnalyteIdx, TotalOffset)]
    CompactScenario[,`:=`(Scenario=NULL, Profile=NULL, Analyte=NULL, 
                          Result=NULL, DisciplineIdx=NULL, ProfileIdx=NULL, AnalyteIdx=NULL)]
    CompactScenario = dcast.data.table(CompactScenario, 
                                       PatientID+ScenarioID+Phase+TargetSex+Location+SampleTaken+SampleReceived+
                                           TimeOfBirthStr+PtOverride+PatientNo+PtFlags+ClinDetails+ClinNotes~.,
                                       value.var = "PyStr", fun.aggregate = paste0, collapse=";")
    setnames(CompactScenario, ".", "ScenarioStr")
    
    #If PtOverride, process now that per-scenario test Code overrides are processed====
    #CompactScenario[!is.na(PtOverride), PatientNo := PtOverride]
    #CompactScenario[,PatientID:= sprintf("%s-%03d", ANALYTE, PatientNo + OFFSET)]
    CompactScenario[,ScenarioIndex := .I]
    #setorder(CompactScenario, PatientID,   ScenarioIndex)
    
    #TODO:Recalculate Sample Taken/Received times...
    CompactScenario[,SampleTaken := SampleTaken + as.difftime(10 * (seq_len(.N)-1), units="mins"), by=PatientID]
    CompactScenario[,SampleReceived := SampleTaken + as.difftime(5, units="mins"), by=PatientID]
    #CompactScenario[,ScenarioID := sprintf("%s_%03d", PatientID, seq_len(.N)), by=PatientID]
    CompactScenario[,`:=`(PtOverride=NULL, PatientNo=NULL, ScenarioIndex=NULL)]
    setcolorder(CompactScenario, c("PatientID", "Phase", "TargetSex", "Location", 
                                   "SampleTaken", "SampleReceived", "TimeOfBirthStr", "ScenarioStr",
                                   "ClinDetails", "PtFlags", "ScenarioID"))
    
    CompactScenario[, SampleTaken := format(SampleTaken, "%Y-%m-%d %H:%M")]
    CompactScenario[, SampleReceived := format(SampleReceived, "%Y-%m-%d %H:%M")]
    setorder(CompactScenario, SampleTaken, SampleReceived)
    
    #Remove those I've already done====
    if (HasRunMode=="Remove"|HasRunMode=="Both"){
        if (file.exists("./Output/AutoTestingSession.log")){
            AlreadyDone = fread("./Output/AutoTestingSession.log")
            AlreadyDone = AlreadyDone[,unique(`Scenario ID`)]
            DoneIntersect = intersect(CompactScenario[,ScenarioID], AlreadyDone)
            message(sprintf("Loading AutoTestingSession.log - %d entries, of which %d are still in current script.",
                            length(AlreadyDone), length(DoneIntersect)))
            CompactScenario = CompactScenario[!(ScenarioID %in% AlreadyDone)]
            rm(AlreadyDone)
        }
    }
    
    #CompactScenario = CompactScenario[grepl(pattern = "FBC", x = ScenarioStr)]
    message(sprintf("%s %d Scenarios remain in final file. Of these, %d are Phase 1 and %d are Phase 2 Scenarios.", 
                    ts(), CompactScenario[,.N], CompactScenario[Phase==1,.N], CompactScenario[Phase==2,.N]))
    
    #Write final table to files====
    fwrite(x = CompactScenario,  file=sprintf("./Output/ScriptDigest_%s.tsv", Analyte), sep = "\t")
}

#Variables====
TestCodeRange = "A1:M181"

#createTestScript(Analyte="VB12", Excelfile="./VB12.xlsm", Excelrange = "A3:G16", MakeMRISafe = FALSE)
#createTestScript(Analyte="CMRI-UHNM", Excelfile="./ComplexMRI-UHNM-r5.xlsm", Excelrange = "A3:G39", 
#MakeMRISafe=TRUE)

#createTestScript(Analyte="B12M", Excelfile="./B12M.xlsm", Excelrange = "A3:G17", 
#                 HasRunMode = "Both", MakeMRISafe=TRUE)
createTestScript(Analyte="CMRI-MCHT", Excelfile="./ComplexMRI-MCHT-r5.xlsm", 
                 Excelrange = "A3:G28", HasRunMode = "Both", MakeMRISafe=TRUE)


#createTestScript(Analyte="CREA", Excelfile="./Creatinine_Ceri.xlsm", Excelrange = "A3:G40")

#createTestScript(Analyte="XAN", Excelfile="./Xanthochromia r2.xlsm", Excelrange = "A3:G58", 
#                 MakeMRISafe = TRUE)
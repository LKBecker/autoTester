#Libraries====
.libPaths("C:/Users/BECKEL201/Documents/R Libraries/") #Sets local libraries to avoid MASSIVE network lag...
#Why do R libraries on managed machines by default install into a faraway network share? Because frick you that's why
library(readxl)
library(data.table)
library(stringr)
library(lubridate)

#Functions====
ffwrite <- function(x, FileName=NULL, Folder="./", ...) {
    Timestamp = format(Sys.time(), "%y%m%d-%H%M%S_")
    if (is.null(FileName)) { FileName = "QuickExport" }
    if ( !(substr(Folder, nchar(Folder), nchar(Folder)) == .Platform$file.sep) ) { Folder = paste0(Folder, .Platform$file.sep) }
    if (FileName!="clipboard") { FileName = paste0(Folder, Timestamp, FileName, ".tsv") }
    message(sprintf("Exporting data.table to '%s'.", FileName))
    write.table(x, file = FileName, quote=F, row.names = F, col.names = (length(colnames(x)) > 0),  sep="\t", ...)
}

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

#Constants====
WEEKSPERYEAR = 52.25 #Rougher but gives "expected" dates. More or less.
DAYSPERMONTH = 30.5 # Sorry Ceri
TODAY = Sys.Date()
hour(TODAY) <- 09
RENAL_LOCATIONS = c("RSW124", "RSWHU", "RSWKU", "CHRU", "MDHRDU", "LHDU1", "LHDU2", "LHDU3", "LHDU4", "LHDU5", "LHDU")

#Variables====
OFFSET = 0
ANALYTE = "VB12"
EXCELFILE = "./VB12.xlsm"
EXCELRANGE = "A3:G30"
TestCodeRange = "A1:M180"

# ==== SCRIPT ====
#Read in data====
ScenarioData = readxl::read_xlsx(EXCELFILE, sheet = "Scenarios", range = EXCELRANGE)
ScenarioData = as.data.table(ScenarioData)
ScenarioData[,Tab:=NULL] #Not needed

ScenarioData[,SampleTaken := TODAY]
ScenarioData[,SampleReceived := TODAY + as.difftime(10, units = "mins")]

#Demarcate which scenarios need special attention
ScenarioData[, Scenario:= str_replace_all(Scenario, ", prev", "<;><Prev>")]
ScenarioData[grepl("<Prev>", Scenario), HasPrev:= TRUE]
ScenarioData[is.na(HasPrev), HasPrev:= FALSE]

setnames(ScenarioData, "Number", "PatientNo")
#Need to pre-process patient numbers; post Phase-split processing fails to properly enumerate patients
ScenarioData[!is.na(PtOverride), PatientNo := PtOverride]

ScenarioData[,PatientID:= sprintf("%s-%03d", ANALYTE, PatientNo + OFFSET)]
ScenarioData[,ScenarioID := sprintf("%s_%03d", PatientID, seq_len(.N)), by=PatientID]


ScenarioData[,tmpResults:= substr(Scenario, str_locate(Scenario, "old with ")[,2]+1, str_length(Scenario))]
ScenarioData[,tmpResults:= str_replace_all(tmpResults, " and ", ";")]

ScenarioData[,Location:=str_match(toupper(Scenario), "/(IP|OP|GP|RENAL)$")[,2]]
ScenarioData[Location=="RENAL", Location := sample(RENAL_LOCATIONS, 1)] #Replaces RENAL with random RENAL location
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

#Split delta check scenarios into two rows====
ScenarioData = SplitDataTableWithMultiRows(ScenarioData, "tmpResults", Separator="<;>") 
#Split off prevs into their own line

#Assign phased on a) Do you have a Prev? b) are you the Prev?
ScenarioData[HasPrev==TRUE & grepl("\\<Prev\\>", tmpResults), Phase:= 1]
ScenarioData[HasPrev==TRUE & is.na(Phase), Phase:= 2]
ScenarioData[is.na(Phase), Phase:= 1]
ScenarioData[HasPrev==TRUE & Phase == 1, SampleTaken := PrevTimeDT]
ScenarioData[HasPrev==TRUE & Phase == 1, SampleReceived := PrevTimeDT + as.difftime(10, units = "mins")]
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

#Creatinine limits per age - currently NOT needed
if (F){
    #Fixing creatinine limits====
    CreaLimitScenarios = ScenarioData[grepl("Creatinine < LLN", `Result to Be Entered`)]
    CreaLimitScenarios[grepl("D$", TargetAge),         TargetAgeUnit:="days"]
    CreaLimitScenarios[grepl("Days$", TargetAge),      TargetAgeUnit:="days"]
    CreaLimitScenarios[grepl("H$", TargetAge),         TargetAgeUnit:="hours"]
    CreaLimitScenarios[grepl("hours old$", TargetAge), TargetAgeUnit:="hours"]
    CreaLimitScenarios[grepl("hour", TargetAge),       TargetAgeUnit:="hours"]
    CreaLimitScenarios[grepl("Y$|y$|years$", TargetAge),      TargetAgeUnit:="years"]
    CreaLimitScenarios[TargetAgeUnit=="hours", tmpAgeNum := tmpAgeNum / 24]
    CreaLimitScenarios[TargetAgeUnit=="hours", TargetAgeUnit := "days"]
    CreaLimitData = fread("./creaLimits.txt")
    
    for (counter in seq(1, CreaLimitScenarios[,.N])) {
        matchingLimits = CreaLimitData[Sex==CreaLimitScenarios[counter,TargetSex] & `Age unit` == CreaLimitScenarios[counter,TargetAgeUnit]]
        matchingLimits = matchingLimits[`Lower age limit` <= CreaLimitScenarios[counter,tmpAgeNum] & 
                                            `Upper age limit` >= CreaLimitScenarios[counter,tmpAgeNum]]
        setorder(matchingLimits, "Lower age limit", "Upper age limit")
        
        if (!assertthat::are_equal(matchingLimits[,.N], 1)){
            warning(paste("Could not isolate a single set of limits for line", counter, 
                          "of 'Creatinine < LLN' scenarios; using first set (sorted by min and max age)."))
            matchingLimits = matchingLimits[1]
        }
        if (!assertthat::are_equal(CreaLimitScenarios[counter,`Result to Be Entered`], "Creatinine < LLN")){
            stop(paste("Result is not just 'Creatinine < LLN', cannot handle line ", counter))
        }
        CreaLimitScenarios[counter,`Result to Be Entered`:= paste("Creatinine = ", as.numeric(matchingLimits[,`Lower limit`]) -1)]
    }
    rm(counter, matchingLimits)
    CreaLimitScenarios[,TargetAgeUnit:=NULL]
    
    ScenarioData = ScenarioData[!grepl("Creatinine < LLN", `Result to Be Entered`)] #Remove entries, and 
    ScenarioData = rbind(ScenarioData, CreaLimitScenarios) #Add altered entries
    rm(CreaLimitData, CreaLimitScenarios)
}

#Load test codes and merge into table====
testCodes = readxl::read_xlsx("./221013_MinRetestTests v1.xlsx", sheet="TestCoords v2", range=TestCodeRange)
testCodes = as.data.table(testCodes)
ScenarioData = merge(ScenarioData, testCodes, by.x="Analyte", by.y="TargetAnalyte", all.x=T)
ScenarioData[,CliniSysName:=NULL]
rm(testCodes)

setorder(ScenarioData, PatientID, Phase, DisciplineIdx, ProfileIdx, AnalyteIdx) 
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
setorder(CompactScenario, PatientID, ScenarioIndex)

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

#Remove those I've already done====
if (file.exists("./Output/AutoTestingSession.log")){
    AlreadyDone = fread("./Output/AutoTestingSession.log")
    AlreadyDone = AlreadyDone[,unique(`Scenario ID`)]
    DoneIntersect = intersect(CompactScenario[,ScenarioID], AlreadyDone)
    message(sprintf("Loading AutoTestingSession.log - %d entries, of which %d are still in current script.",
                    length(AlreadyDone), length(DoneIntersect)))
    CompactScenario = CompactScenario[!(ScenarioID %in% AlreadyDone)]
    rm(AlreadyDone)
}

#CompactScenario = CompactScenario[grepl(pattern = "FBC", x = ScenarioStr)]
message(sprintf("%s %d Scenarios remain in final file. Of these, %d are Phase 1 and %d are Phase 2 Scenarios.", 
                ts(), CompactScenario[,.N], CompactScenario[Phase==1,.N], CompactScenario[Phase==2,.N]))

#Write final table to files====
fwrite(CompactScenario, "./Output/ScriptDigest.tsv", sep = "\t", row.names = FALSE, col.names = T)

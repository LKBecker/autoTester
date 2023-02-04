#
# file to generate sets of normal results, to probe the minimum retest intervals
# by generating a series of samples [a,b,c,d] where the distance (in days) of specimen collection times
#   a->b = MRI + 1, 
#   b->c = MRI
#   c->d = MRI - 1
# the first should pass, the second fail; 
# if the second fails, the third passes (since "c" no longer counts and it's b->d)
#
# Required: a .txt file detailling test sets and their coordinates

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

#Constants====
WEEKSPERYEAR = 52.25 #Rougher but gives "expected" dates. More or less.
DAYSPERMONTH = 30.5 # Sorry Ceri
MAX_REQUESTS_PER_SAMPLE = 20
TODAY = Sys.Date()

#Variables====
TestCodeRange = "A1:M170"
Attempt_No = 7

#Script====
#TODO: read log, set to max of scenario id after extracting number
if (file.exists("./AutoTestingSession.log")==TRUE){
    Logfile = fread("./AutoTestingSession.log", sep="\t", header = T)
    Logfile = Logfile[grepl("^MRI-", `Scenario ID`), unique(`Scenario ID`)]
    Logfile = max(as.numeric(str_extract(Logfile, "\\d{3}")))
    
    if (Logfile >= Attempt_No){
        warning(sprintf("Logfile documents use of patient %02d, but current Attempt_No is %02d.
        To ensure clean Minimum Retest Interval testing, please increase Attempt_No,
        to create a new patient record.", Logfile, Attempt_No))
        #do nothing for now
    }
}

#Read in data====
testCodes = readxl::read_xlsx("./221013_MinRetestTests v1.xlsx", sheet="TestCoords v2", range=TestCodeRange)
testCodes = as.data.table(testCodes)
if (testCodes[is.na(MinimumRetestInterval_Days), .N] > 0){
    warning(sprintf("WARNING: %d test codes / sets do not have their Minimum Retest Interval set up, and won't run!",
                    testCodes[is.na(MinimumRetestInterval_Days), .N]))
}
MRITestData = testCodes[MinimumRetestInterval_Days!= 0, 
                        .(Profile, TestCode, DisciplineIdx, ProfileIdx, AnalyteIdx, TotalOffset,
                          MinimumRetestInterval_Days, NormalResult)]
rm(testCodes)

SpecialBehaviours = c("A1C","FERM","FERX","B12M", "VB12","FOLM","FOL",
                      "DHEA","EPO", "PTH","PTH","Prolactin","VITD","ACE","B2M",
                      "CU","HAPT","P3NP","ZN","TSH","AFP","C125","C153","C199","CEA","PSA")
SpecialBehaviours = MRITestData[TestCode %in% SpecialBehaviours, unique(Profile)]

MRI_Basics = MRITestData[!(Profile %in% SpecialBehaviours)]
MRI_Advanced = MRITestData[Profile %in% SpecialBehaviours]
rm(MRITestData, SpecialBehaviours)

#Duplicate if UCRE appears alone - does not require testing, is not named in MRIs.
MRI_Basics = MRI_Basics[Profile != "UCRE"]

#testCodes_Retesting[, sort(unique(Profile, MinimumRetestInterval_Days))]
#Create Scenarios for Basic Tests ====
setorder(MRI_Basics, Profile, TestCode)

#Generate Age as 21 years old ====
MRI_Basics[, TimeOfBirth:= TODAY - as.difftime(21*WEEKSPERYEAR, units = "weeks")]
MRI_Basics[, TimeOfBirth := format(TimeOfBirth, "%Y-%m-%d")] #And format as string for the program

#Assign location and create "patient" name
MRI_Basics[,PatientID:= sprintf("MRI-%03d", Attempt_No)]
MRI_Basics[,Location:= "GP"]

#Calculate when our simulated samples are taken, SampleDate1====
MRI_Basics[MinimumRetestInterval_Days<9999, SampleDate1 := 
               (TODAY - as.difftime((MinimumRetestInterval_Days * 4) + 1, units = "days"))]
#We allocate four MRIs as we do MRI - 1, MRI and MRI + 1 = 3x MRI, plus one extra for space

#For never-redo tests, we can just have a sample taken 28 days ago
MRI_Basics[is.na(SampleDate1), SampleDate1 := (TODAY - as.difftime(28, units = "days"))]
MRI_Basics[MinimumRetestInterval_Days>=9999, SampleDate2 := TODAY] #28 days

#For tests we try multiple times, each interval must then be followed by a new "phase 1" sample that is OUTSIDE the MRI
#This is due to  samples within MRI being auto-rejected, 
#and thus don't count for calculating the subsequent sample's MRI 
#i.e. for each test we need to first submit an 'initial' sample, which is also outside the MRI wrt/ all previous samples

MRI_Basics[MinimumRetestInterval_Days<9999, 
           SampleDate2 := SampleDate1 + as.difftime(MinimumRetestInterval_Days + 1, units = "days")]
MRI_Basics[MinimumRetestInterval_Days<9999, 
           SampleDate3 := SampleDate2 + as.difftime(MinimumRetestInterval_Days, units="days")  ]
MRI_Basics[MinimumRetestInterval_Days<9999, 
           SampleDate4 := SampleDate3 + as.difftime(MinimumRetestInterval_Days - 1, units = "days")]

#This gives us three successive timespans: 
#   SampleDate1 to SampleDate2 = MRI + 1 (should PASS, and thus Date 2 now counts)
#   SampleDate2 to SampleDate3 = MRI     (should work? and thus date 3 now counts)
#   SampleDate3 to SampleDate4 = MRI - 1 (should FAIL! unless previous has failed and disqualified Date 3)


tmpCols = colnames(MRI_Basics)[!grepl("(SampleDate\\d{1})", colnames(MRI_Basics))]
MRI_Basics_M = melt.data.table(MRI_Basics, id.vars = tmpCols, value.name = "SampleTaken")
rm(tmpCols)

MRI_Basics_M = MRI_Basics_M[!is.na(SampleTaken)]

MRI_Basics_M[, Phase := as.numeric(str_extract(variable, pattern = "\\d{1}$"))] #TODO replace for %%1 == 0``

MRI_Basics_M[TestCode=="FITS", NormalResult := format(SampleTaken - as.difftime(1, units = "days"), "%d/%m/%Y")]
MRI_Basics_M[TestCode=="FITR", NormalResult := format(SampleTaken, "%d/%m/%Y")]
MRI_Basics_M[TestCode=="FITD", NormalResult := format(SampleTaken, "%d/%m/%Y")]

MRI_Basics_M[, SampleReceived := format(SampleTaken, "%Y-%m-%d 12:10")]
MRI_Basics_M[, SampleTaken := format(SampleTaken, "%Y-%m-%d 12:00")]

MRI_Basics_M[,PyStr := sprintf("%s|%s|%s|%s|%s|%s|%s", 
                   Profile, TestCode, NormalResult, DisciplineIdx, ProfileIdx, AnalyteIdx, TotalOffset)]

MRI_Basics_M[,`:=`(variable=NULL, MinimumRetestInterval_Days=NULL)]



#Version 1: Compact, minimum number of samples====
if (F){
    MRI_Basics_Compact <- dcast.data.table(data=MRI_Basics_M, 
                                      formula = PatientID+TimeOfBirth+Location+Phase+SampleTaken+SampleReceived~.,
                                      value.var = "PyStr", fun.aggregate = paste0, collapse = ";")
    setnames(MRI_Basics_Compact, ".", "ScenarioStr")
    
    MRI_Basics_Compact[, TargetSex:=ifelse(Attempt_No %% 2 == 0, "F", "M")]
    MRI_Basics_Compact[,ScenarioID:= paste(PatientID, .I, sep="_")]
    setcolorder(MRI_Basics_Compact, c("PatientID", "Phase", "TargetSex","Location", "SampleTaken", "SampleReceived",
                                      "TimeOfBirth", "ScenarioStr", "ScenarioID"))
    
    fwrite(MRI_Basics_Compact, "./ScriptDigest-MRITest_Minimal.tsv", sep = "\t", row.names = FALSE, col.names = T)
}
#Version 2: One Set per sample====
MRI_Basics_PerProfile <- unique(MRI_Basics_M[, .(PatientID, TimeOfBirth, Location, SampleTaken, 
                      SampleReceived, ScenarioStr=paste0(PyStr, collapse=";")), .(Profile, Phase)])
setorder(MRI_Basics_PerProfile, Profile, Phase)

MRI_Basics_PerProfile[, TargetSex:=ifelse(Attempt_No %% 2 == 0, "F", "M")]
ScenarioIDs = MRI_Basics_PerProfile[order(Profile, Phase), .(unique(Profile))][, .(Profile=V1, ScenarioID = .I)] 
MRI_Basics_PerProfile = merge(MRI_Basics_PerProfile, ScenarioIDs, by="Profile")
MRI_Basics_PerProfile[, ScenarioID := paste(PatientID, ScenarioID, sep="_")]
MRI_Basics_PerProfile[,Profile:=NULL]
MRI_Basics_PerProfile[,ClinDetails:=""]
MRI_Basics_PerProfile[,PtFlags:=""]

setcolorder(MRI_Basics_PerProfile, c("PatientID", "Phase", "TargetSex","Location", "SampleTaken", "SampleReceived",
                                  "TimeOfBirth", "ScenarioStr", "ClinDetails", "PtFlags", "ScenarioID"))

fwrite(MRI_Basics_PerProfile, "./ScriptDigest-MRITest_PerProfile.tsv", sep = "\t", row.names = FALSE, col.names = T)
rm(MRI_Basics_M, MRI_Basics, ScenarioIDs)

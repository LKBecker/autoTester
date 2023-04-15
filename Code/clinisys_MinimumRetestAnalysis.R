#
# file to analyse the results of clinisys_MinimumRetestScenarios.R (run through the main_CliniSys.py script and WinPath)
#

#Libraries====
.libPaths("C:/Users/BECKEL201/Documents/R Libraries/") #Sets local libraries to avoid MASSIVE network lag...
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

Logfile    <- fread("./Output/AutoTestingSession.log", sep="\t", header = T)
MRI_Tests  <- Logfile[grepl("^MRI-016", `Scenario ID`)]

MRI_Tests  <- MRI_Tests[,.(ID=`Scenario ID`, Phase, CollDate=as.Date.character(`Sample Collected`, format="%d/%m/%Y"),
                           AuthQ=`Auth Queue`, Sample=`Sample ID`, Test=str_extract(`Target Result`, "(\\w+)\\]") )]
MRI_Tests[,Test := str_extract(Test, "\\w+")]
MRI_Tests[, PassFail := str_extract(AuthQ, "PASS|FAIL")]
MRI_Tests[, DupQueue := str_extract(AuthQ, "(DUP.+) ")]
MRI_Tests[,PrevScenario := shift(ID, n=1, type="lag")]
MRI_Tests[,PrevCollection := shift(CollDate, n=1, type="lag")]

MRI_Tests[PrevScenario != ID, PrevScenario := NA]
MRI_Tests <- MRI_Tests[!is.na(PrevScenario)]
MRI_Tests[, PrevScenario := NULL]

#MRI_Tests[,NDaysBetween := as.numeric(as.difftime(CollDate - PrevCollection, units = "days"))]
MRI_Tests[,NDaysBetween := days(CollDate - PrevCollection)]
MRI_Tests[,`:=`(Phase=NULL, CollDate=NULL, PrevCollection=NULL)]
setcolorder(MRI_Tests, c("ID", "Sample", "Test", "NDaysBetween", "PassFail", "DupQueue"))


write.table(MRI_Tests, "clipboard", sep="\t", row.names = F, col.names = T)
rm(Logfile)
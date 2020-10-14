### Title:    TiU Exam Merging Utility
### Author:   Kyle M. Lang
### Created:  2020-10-13
### Modified: 2020-10-14

rm(list = ls(all = TRUE))

source("subroutines.R")

library(stringr)
library(xlsx)
library(svDialogs)

if(FALSE) {
onlineFile <- "../../data/2020-10-09T1915_Grades-35B101-B-6.csv"
campusFile <- "../../data/6333 B_rep_cijferlijst.xlsx"
lookupFile <- "../../data/lookup_table.csv"
nQuestions <- 40
nOptions   <- 4
passNorm   <- 0.55
}

###--Get Inputs from User----------------------------------------------------###

## Define legal file types for online results file:
csvFilters <- matrix(c("Comma-Seperated Values", "*.csv;*.CSV"), 1, 2)

## Prompt the user to select the file path to the CSV file containing the online
## test results:
onlineFile <- dlgOpen(title = "Please select the file that contains the online test results.",
                      filters = csvFilters)$res

if(length(onlineFile) == 0)
    wrappedError("I cannot proceed without knowing where to find your online test results.")

## Define legal file types for on-campus results file:
xlsxFilters <- matrix(c("Office Open XML Spreadsheet", "*.xlsx;*.XLSX",
                     "Excel 2007 - 2019", "*.xlsx;*.XLSX"),
                   2, 2, byrow = TRUE)

## Prompt the user to select the file path to the XLSX file containing the
## on-campus test results:
campusFile <- dlgOpen(title = "Please select the file that contains the on-campus test results.",
                      filters = xlsxFilters)$res

if(length(campusFile) == 0)
    wrappedError("I cannot proceed without knowing where to find your on-campus test results.")

                                        #customScheme <-
                                        #    dlgList(choices = c("Yes", "No"),
                                        #            preselect = "No",
                                        #            title = "Did the instructor request a custom scoring scheme?")$res

customScheme <-
    dlgMessage(message = "Did the instructor request a custom scoring scheme?",
               type    = "yesno")$res

                                        #if(length(customScheme) == 0) {
                                        #    wrappedWarning("You have not told me if the instructor wants to use their own scoring scheme, so I will apply the University's default scoring rule.")
                                        #    customScheme <- "No"
                                        #}

if(customScheme == "yes") {
    ## Prompt the user to select the file path to the CSV file containing the
    ## lookup table describing the custom scoring scheme:
    tableFile <- dlgOpen(title = "Please select the file that contains the lookup table defining the custom scoring scheme.",
                         filters = csvFilters)$res

    tmp        <- prepScoringScheme(tableFile)
    scheme     <- tmp$scheme
    scoreTable <- tmp$table
    
    if(length(tableFile) == 0)
        wrappedError("I cannot proceed without a user-supplied scoring table.")
} else {
    nQuestions <- as.numeric(
        dlgInput("How many questions does this exam contain?")$res
    )
    nOptions   <- as.numeric(
        dlgInput("How many response options are available for each question?")$res
    )
    passNorm   <- as.numeric(
        dlgInput("What norm would you like to use to define a passing grade?",
                 default = 0.55)$res
    )
    check <- length(nQuestions) == 0 | length(nOptions) == 0 | length(passNorm) == 0
    if(check)
        wrappedError("I need to know how many questions this exam contains, the number of response options for each question, and the passing norm before I can proceed.") 
}

dlgMessage("Finally, I need you to tell me where you would like to save the results.")

outFile <- dlgSave(title = "Where would you like to save the results?")$res


###--Process Online Gradebook Data-------------------------------------------###

## Read in Online gradebook and column names:
tmp        <- autoReadCsv(onlineFile)
onlineData <- tmp$data

onlineNames <- as.character(
    read.table(paste0(onlineFile), nrows = 1, sep = tmp$sep)
)

## Drop metadata rows:
onlineData <- onlineData[-c(1 : 2), ]

## Find the exam column in the Online gradebook:
examCol <- findExam(data = onlineData, names = onlineNames)

## Extract metadata from online exam name:
oExamName <- onlineNames[examCol]
tmp       <- str_locate_all(oExamName, c("/", "\\("))

oExamDate <- tryCatch(substr(oExamName, 1, tmp[[1]][1, 1] - 1),
                      error = function(e) "Not Recovered")

courseCode <- tryCatch(substr(oExamName, tmp[[1]][1, 1] + 1, tmp[[1]][2, 2] - 1),
                       error = function(e) "Not Recovered")

tmp <- try(
    substr(oExamName, tmp[[1]][2, 2] + 1, tmp[[2]][1, 1] - 1),
    silent = TRUE
)
if(class(tmp) != "try-error") oExamName <- str_trim(tmp)

## Parse student names:
oNames <- sapply(onlineData$Student, parseName, USE.NAMES = FALSE)

## Extract the relevent columns:
online <- data.frame(onlineData[ , c("SIS.User.ID",
                                     colnames(onlineData)[examCol])
                                ],
                     surname   = oNames["name2", ],
                     firstName = oNames["name1", ],
                     version   = "",
                     source    = "Online"
                     )
colnames(online)[1 : 2] <- c("snr", "score")

## Remove any students without SNRs:
online <- online[!is.na(online$snr), ]


###--Process On-Campus Grade Data--------------------------------------------###

## Read in on-campus grades:
campusData <- read.xlsx(campusFile, sheetIndex = 1)

## Extract first and second columns:
c1 <- campusData[[1]]
c2 <- campusData[[2]]

## Find the rows containing individual grades:
gradeRows <- grep("\\d{6,7}", c1)

## Extract the individual grade information:
campusGrades           <- as.data.frame(campusData[gradeRows, ])
colnames(campusGrades) <- as.character(campusData[gradeRows[1] - 1, ])

## Check for students with broken SNRs:
extraCol <- campusGrades[[3]]
snrFlag  <- !is.na(extraCol)

## Replace any broken SNRs with the corrections provided by the SA:
if(any(snrFlag))
    for(i in which(snrFlag)) {
        x   <- extraCol[i]
        tmp <- unlist(str_locate_all(x, "\\d{6,7}\\s"))
        
        campusGrades[i, "S Nummer"]  <- substr(x, tmp[1], tmp[2] - 1)
        campusGrades[i, "Naam"]      <- substr(x, tmp[2] + 1, nchar(x))
    }

## Parse student names:
cNames <- sapply(campusGrades$Naam, parseName, USE.NAMES = FALSE)

## Extract metadata from the on-campus file:
cn        <- colnames(campusData)
cn        <- gsub("^X", "", cn)
faculty   <- tolower(cn[grep("Opleiding", cn) + 1])
batchId   <- as.numeric(cn[grep("Batch\\.ID", cn) + 1])
cExamName <- c2[grep("Toetsnaam", c1)]
cExamDate <- as.Date(as.numeric(cn[grep("Toetsdatum", cn) + 1]),
                     origin = "1899-12-30")

## Extract the relevent columns:
campus <- data.frame(
    campusGrades[ , c("S Nummer", "Score", "Versie")],
    surname   = cNames["name2", ],
    firstName = cNames["name1", ],
    source    = "Campus")
colnames(campus)[1 : 3] <- c("snr", "score", "version")

## Extract SA-computed result for testing purposes:
campus0 <- campusGrades[ , c("S Nummer", "Resultaat")]
colnames(campus0) <- c("snr", "result0")


###--Merge and Process Exam Grades-------------------------------------------###

## Check for duplicate students:
overlap <- intersect(campus$snr, online$snr)

if(length(overlap) > 0) {
    tmp <- as.matrix(
        online[online$snr %in% overlap, c("snr", "surname", "firstName")]
    )
    colnames(tmp) <- c("SNR", "Surname", "Initials/First Name")

    overlapFile <- paste(dirname(outFile), "duplicate_students.txt", sep = "/")
    write.table(tmp, file = overlapFile, sep = "\t", row.names = FALSE)
    
    msg <- paste0("It looks like ",
                  length(overlap),
                  " students are represented in both input files. I have saved their information in the file: ",
                  overlapFile,
                  ". Please correct this issue before trying to rerun this job.")    
    wrappedError(msg)
}

## Stack the relevent columns from the online and on-campus files:
pooled <- rbind(campus, online)

## Convert relevant columns to numeric:
pooled$snr      <- as.numeric(pooled$snr)
pooled$score    <- as.numeric(pooled$score)
campus0$result0 <- as.numeric(campus0$result0)

## Score the exams:
if(customScheme == "yes") {
    ## Make sure all observed scores are represented in the lookup table:
    extra <- setdiff(sort(unique(pooled$score)), names(scheme))
    if(length(extra) > 0) {
        msg <- paste0("This set of observed scores: {",
                      paste(extra, collapse = ", "),
                      "} is not represented in the supplied scoring table, so I cannot score this exam.")
        wrappedError(msg)
    }

    ## Score the exam:
    pooled$result <- sapply(pooled$score,
                            function(x, scheme) scheme[as.character(x)],
                            scheme = scheme)
} else {
    ## Score the exam:
    pooled$result <- scoreExam(score      = pooled$score,
                               nQuestions = nQuestions,
                               nOptions   = nOptions,
                               minGrade   = ifelse(faculty == "tisem", 0, 1),
                               pass       = passNorm)
    
    ## Generate a lookup table for the report:
    tmp        <- 1 : nQuestions
    scoreTable <- data.frame(Score = tmp,
                             Grade = scoreExam(score      = tmp,
                                               nQuestions = nQuestions,
                                               nOptions   = nOptions,
                                               minGrade   = ifelse(faculty == "tisem", 0, 1),
                                               pass       = passNorm)
                             )
}

## Check the results:
tmp   <- merge(pooled, campus0, by = "snr")
check <- all(with(tmp, result - result0) == 0)

if(!check)
    wrappedWarning("Something may have gone wrong. The exam results I've calculated do not match the results provided by the Student Administration.")
    

###--Create XLSX Output File-------------------------------------------------###

## Create an empty workbook and a new sheet therein:
wb1 <- createWorkbook()
s1  <- createSheet(wb = wb1, sheetName = "Results")

## Define a style to format headings in bold font:
BoldStyle <- CellStyle(wb1) + Font(wb1, isBold = TRUE)

## Populate column names for metadata block:
blockData <- c("Exam Name", "Exam Date", "Batch ID", "Course ID")
cb <- CellBlock(s1, 1, 1, 1, 4)
CB.setRowData(cb, blockData, 1, rowStyle = BoldStyle)

## Populate contents of metadata block:
blockData <- matrix(c(cExamName, oExamName,
                      as.character(cExamDate), as.character(oExamDate),
                      batchId, "",
                      courseCode, ""),
                    nrow = 2)
cb <- CellBlock(s1, 2, 1, 2, 4)
CB.setMatrixData(cb, blockData, 1, 1)

## Populate row names for summary measures block:
blockData <- c("Passed (%)", "Average Grade", "Average Score", "Students (N)")
cb <- CellBlock(s1, 6, 1, 4, 1)
CB.setColData(cb, blockData, 1, colStyle = BoldStyle)

## Populate contents of summary measures block:
blockData <- matrix(
    with(pooled,
         c(round(100 * mean(result > 5.5), 1),
           round(mean(result), 1),
           round(mean(score), 1),
           length(score)
           )
         )
)
cb <- CellBlock(s1, 6, 2, 4, 1)
CB.setColData(cb, blockData, 1)

## Populate contents of warning message:
blockData <- "Please note: Canvas cannot output special characters. Therefore, student names might be displayed improperly."
cb <- CellBlock(s1, 12, 1, 1, 1)
CB.setRowData(cb, blockData, 1)

## Populate column names for student results block:
blockData <- c("Surname",
               "Initials/First Name",
               "SNR",
               "Grade",
               "Score",
               "Source",
               "Version")
cb <- CellBlock(s1, 15, 1, 1, 7)
CB.setRowData(cb, blockData, 1, rowStyle = BoldStyle)

## Populate contents of student results block:
addDataFrame(pooled[c("surname",
                      "firstName",
                      "snr",
                      "result",
                      "score",
                      "source",
                      "version")
                    ],
             sheet = s1,
             col.names = FALSE,
             row.names = FALSE,
             startRow = 16,
             startCol = 1)

## Add another sheet to hold the scoring table:
s2 <- createSheet(wb = wb1, sheetName = "Scoring Table")
addDataFrame(scoreTable,
             sheet         = s2,
             row.names     = FALSE,
             colnamesStyle = BoldStyle)
    
## Save the final workbook to disk:
saveWorkbook(wb1, outFile)

### Title:    TiU Exam Merging Utility
### Author:   Kyle M. Lang
### Created:  2020-10-13
### Modified: 2020-10-14

rm(list = ls(all = TRUE))

verbose <- TRUE

source("subroutines.R")

library(stringr)
library(xlsx)
library(svDialogs)


###--Get Inputs from User----------------------------------------------------###

if(verbose)
    dlgMessage("First, I need to ask you for some input data. I'm hungry; please feed me!")

## Define legal file types for online results file:
oFilters <- matrix(c("Comma-Seperated Values", "*.csv;*.CSV"), 1, 2)

## Prompt the user to select the file path to the CSV file containing the online
## test results:
onlineFile <- dlgOpen(title = "Please select the file that contains the online test results.",
                      filters = oFilters)$res

## Define legal file types for on-campus results file:
cFilters <- matrix(c("Office Open XML Spreadsheet", "*.xlsx;*.XLSX",
                     "Excel 2007 - 2019", "*.xlsx;*.XLSX"),
                   2, 2, byrow = TRUE)

## Prompt the user to select the file path to the XLSX file containing the
## on-campus test results:
campusFile <- dlgOpen(title = "Please select the file that contains the on-campus test results.",
                      filters = cFilters)$res

if(verbose)
    dlgMessage("Now, I need to get some more information about this exam.")

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

if(verbose)
    dlgMessage("Finally, I need you to tell me where you would like to save the results.")

outFile <- dlgSave(title = "Where would you like to save the results?")$res


###--Process Online Gradebook Data-------------------------------------------###

## Read in Online gradebook and column names:
onlineData  <- read.csv2(onlineFile)
onlineNames <- as.character(
    read.table(paste0(onlineFile), nrows = 1, sep = ";")
)

## Drop metadata rows:
onlineData <- onlineData[-c(1 : 2), ]

## Find the exam column in the Online gradebook:
examCol <- grep("\\d{4}\\.\\d{2}\\.\\d{2}.*Remotely\\.Proctored",
                colnames(onlineData)
                )

## Extract metadata from online exam name:
oExamName <- onlineNames[examCol]

tmp <- str_locate_all(oExamName, c("/", "\\("))

oExamDate  <- as.Date(substr(oExamName, 1, tmp[[1]][1, 1] - 1))

courseCode <- substr(oExamName, tmp[[1]][1, 1] + 1, tmp[[1]][2, 2] - 1)

oExamName <- substr(oExamName, tmp[[1]][2, 2] + 1, tmp[[2]][1, 1] - 1)
oExamName <- str_trim(oExamName)

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
    
### DO SOMETHING ###############################################################
    
}

## Stack the relevent columns from the online and on-campus files:
pooled <- rbind(campus, online)

## Convert relevant columns to numeric:
pooled$snr      <- as.numeric(pooled$snr)
pooled$score    <- as.numeric(pooled$score)
campus0$result0 <- as.numeric(campus0$result0)

## Score the exams:
pooled$result <- scoreExam(score      = pooled$score,
                           nQuestions = nQuestions,
                           nOptions   = nOptions,
                           minGrade   = ifelse(faculty == "tisem", 0, 1),
                           pass       = passNorm)

## Check the results:
tmp   <- merge(pooled, campus0, by = "snr")
check <- all(with(tmp, result - result0) == 0)

if(!check) {

### DO SOMETHING ###############################################################

}


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

## Save the final workbook to disk:
saveWorkbook(wb1, outFile)

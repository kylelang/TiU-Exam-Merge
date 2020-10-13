### Title:    TiU Exam Merging Utility
### Author:   Kyle M. Lang
### Created:  2020-10-13
### Modified: 2020-10-13

rm(list = ls(all = TRUE))

nQuestions = 40
nOptions   = 4
passNorm   = 0.55
dataDir <- "../../data/"

source("subroutines.R")
library(readxl)
library(dplyr)
library(stringr)
library(xlsx)

### Process Online Gradebook Data ###

## Read in Online gradebook and column names:
onlineData  <-
    read.csv2(paste0(dataDir, "2020-10-09T1915_Grades-35B101-B-6.csv"))
onlineNames <- as.character(
    read.table(paste0(dataDir, "2020-10-09T1915_Grades-35B101-B-6.csv"),
               nrows = 1,
               sep = ";")
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
tmp   <- onlineData$Student
tmp   <- strsplit(tmp, ", ")
name1 <- sapply(tmp, "[", x = 2)
name2 <- sapply(tmp, "[", x = 1)

## Extract only the relevent columns:
online <- data.frame(onlineData[ , c("SIS.User.ID",
                                     colnames(onlineData)[examCol])
                                ],
                     surname   = name2,
                     firstName = name1,
                     version   = "O",
                     source    = "Online"
                     )
colnames(online)[1 : 2] <- c("snr", "score")

## Remove any students without SNRs:
online <- online[!is.na(online$snr), ]

### Process On-Campus Grade Data ###

## Read in on-campus grades:
campusData <- read_excel(paste0(dataDir, "6332 A_rep_cijferlijst.xlsx"))

## Extract first and second columns:
c1 <- pull(campusData, 1)
c2 <- pull(campusData, 2)

## Find the rows containing individual grades:
gradeRows <- grep("\\d{6,7}", c1)

## Extract the individual grade information:
campusGrades           <- as.data.frame(campusData[gradeRows, ])
colnames(campusGrades) <- as.character(campusData[gradeRows[1] - 1, ])

## Parse student names:
tmp   <- campusGrades$Naam[1]
name1 <- name2 <- c()
for(x in campusGrades$Naam) {
    tmp <- unlist(str_locate_all(x, "\\s\\w\\."))
    name2 <- c(name2, substr(x, 1, tmp[1] - 1))
    name1 <- c(name1, substr(x, tmp[1] + 1, nchar(x)))
}

## Extract metadata from on-campus file:
cn        <- colnames(campusData)
faculty   <- tolower(cn[grep("Opleiding", cn) + 1])
batchId   <- as.numeric(cn[grep("Batch ID", cn) + 1])
cExamName <- c2[grep("Toetsnaam", c1)]
cExamDate <- as.Date(as.numeric(cn[grep("Toetsdatum", cn) + 1]),
                     origin = "1899-12-30")

## Extract only the relevent columns:
campus <- data.frame(
    campusGrades[ , c("S Nummer", "Score", "Versie")],
    surname   = name2,
    firstName = name1,
    source    = "Campus")
colnames(campus)[1 : 3] <- c("snr", "score", "version")

## Extract SA-computed result for testing purposes:
campus0 <- campusGrades[ , c("S Nummer", "Resultaat")]
colnames(campus0) <- c("snr", "result0")


### Merge and Process Exam Grades ###

## Check for duplicate students:
overlap <- intersect(campus$snr, online$snr)

if(length(overlap) > 0) {
    
### DO SOMETHING
    
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

### DO SOMETHING

}


### Create Output File ###

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
saveWorkbook(wb1, paste0(dataDir, "testOut4.xlsx"))

### Title:    Combine and Process TiU Hybrid Exam Results
### Author:   Kyle M. Lang
### Created:  2020-10-13
### Modified: 2020-10-15


###--Process Online Gradebook Data-------------------------------------------###

## Read in Online gradebook and column names:
tmp        <- autoReadCsv(onlineFile, stringsAsFactors = FALSE)
onlineData <- tmp$data

onlineNames <- as.character(
    read.table(file             = paste0(onlineFile),
               nrows            = 1,
               sep              = tmp$sep,
               stringsAsFactors = FALSE)
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

## Remove any students without SNRs or scores:
online <- online[with(online, !is.na(snr) & !is.na(score)), ]


###--Process On-Campus Grade Data--------------------------------------------###

## Read in on-campus grades:
campusData <- read.xlsx(campusFile, sheetIndex = 1, stringsAsFactors = FALSE)

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

## Remove any students without SNRs or scores:
campus <- campus[with(campus, !is.na(snr) & !is.na(score)), ]

## Extract SA-computed result for testing purposes:
campus0 <- campusGrades[ , c("S Nummer", "Resultaat")]
colnames(campus0) <- c("snr", "result0")


###--Combine and Process Exam Grades-----------------------------------------###

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
    ## Define the minimum grade to use:
    minGrade <- ifelse(faculty == "tisem" & tisemMinGrade == 0, 0, 1)
    
    ## Score the exam:
    pooled$result <- scoreExam(score      = pooled$score,
                               nQuestions = nQuestions,
                               nOptions   = nOptions,
                               minGrade   = minGrade,
                               pass       = passNorm)

    ## Generate a lookup table for the report:
    tmp        <- 1 : nQuestions
    scoreTable <- data.frame(Score = tmp,
                             Grade = scoreExam(score      = tmp,
                                               nQuestions = nQuestions,
                                               nOptions   = nOptions,
                                               minGrade   = minGrade,
                                               pass       = passNorm)
                             )
}

## Calculate the minimum passing score:
cutoff <- with(scoreTable, Score[sum(Grade < 5.5) + 1])

## Check the results:
tmp   <- merge(pooled, campus0, by = "snr")
check <- all(with(tmp, result - result0) == 0)

if(!check)
    wrappedWarning("Something may have gone wrong. The exam results I've calculated do not match the results provided by the Student Administration.",
                   immediate. = TRUE)

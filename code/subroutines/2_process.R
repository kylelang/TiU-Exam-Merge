### Title:    Combine and Process TiU Hybrid Exam Results
### Author:   Kyle M. Lang
### Created:  2020-10-13
### Modified: 2020-12-09


###--Process Online Gradebook Data-------------------------------------------###

## Read in Online gradebook and column names:
tmp         <- autoReadCsv(onlineFile, stringsAsFactors = FALSE)
onlineData  <- tmp$data
onlineNames <- as.character(
    read.table(file             = paste0(onlineFile),
               nrows            = 1,
               sep              = tmp$sep,
               stringsAsFactors = FALSE)
)

## Check if the online exam was administered through Canvas or TestVision:
canvas <- any(grepl("Points Possible|(read only)", onlineData[2, ]))

if(canvas) {
    ## Drop metadata rows:
    onlineData <- onlineData[-c(1 : 2), ]
    
    ## Find the exam column(s) in the Online gradebook:
    examCol <- findExam(data = onlineData, names = onlineNames)
    
    ## Process the gradebook data:
    tmp <- lapply(X     = examCol,
                  FUN   = processCanvas,
                  data  = onlineData,
                  names = onlineNames,
                  codes = missingScoreCodes)
    
    onlineData <- do.call(rbind, lapply(tmp, "[[", x = "data"))
    onlineMeta <- lapply(tmp, "[", x = -1)
} else {
    tmp        <- processTestVision(onlineData)
    onlineData <- tmp$data
    onlineMeta <- list(tmp[-1])
}

###--Process On-Campus Grade Data--------------------------------------------###

if(campus) {
    tmp <- lapply(campusFile, processCampus)
    
    campusData  <- do.call(rbind, lapply(tmp, "[[", x = "data"))
    campusData0 <- do.call(rbind, lapply(tmp, "[[", x = "data0"))
    campusMeta  <- lapply(tmp, "[", x = -c(1, 2))
    faculty     <- campusMeta[[1]]$faculty
} else {
    campusData <- campusMeta <- NULL
}

###--Combine and Process Exam Grades-----------------------------------------###

## Stack the relevent columns from the online and on-campus files:
if(canvas) {
    pooled <- rbind(campusData, onlineData)
    
    ## Check for duplicate students:
    flag <- duplicated(pooled$snr)
    
    if(any(flag)) {
        tmp           <- as.matrix(pooled[flag, c("snr", "surname", "firstName")])
        colnames(tmp) <- c("SNR", "Surname", "Initials/First Name")
        
        ## Save the data on duplicate students:
        dupFile <- paste(dirname(outFile), "duplicate_students.txt", sep = "/")
        write.table(tmp, file = dupFile, sep = "\t", row.names = FALSE)
        
        msg <- paste0("It looks like ",
                      sum(flag),
                      " students are represented in multiple input files. I have saved their information in the file: ",
                      dupFile,
                      ". Please correct this issue before trying to rerun this job.")
        wrappedError(msg)
    }
} else {
    pooled <- rbind(campusData[ , -1], onlineData[ , -1])
    snr    <- c(campusData$snr, rep(NA, nrow(onlineData)))
    anr    <- c(rep(NA, nrow(campusData)), onlineData$anr)
    pooled <- data.frame(snr, anr, pooled)
}

## Merge the metadata lists:
meta <- c(campusMeta, onlineMeta)

## Convert relevant columns to numeric:
pooled$snr   <- as.numeric(pooled$snr)
pooled$score <- as.numeric(pooled$score)

if(campus)
    campusData0$result0 <- as.numeric(campusData0$result0)

## Score the exams:
if(scoreScheme == scoreOpts["wo3"]) {
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
    scoringIndex  <- which(scoreOpts %in% scoreScheme)
    
    ## Define the minimum grade to use:
    minGrade <- switch(scoringIndex,
                       ifelse(faculty == "tisem" & tisemMinGrade == 0, 0, 1),
                       1,
                       0)
    
    ## Score the exam:
    pooled$result <- scoreExam(score      = pooled$score,
                               what       = scoringIndex,
                               nQuestions = nQuestions,
                               nOptions   = nOptions,
                               minGrade   = minGrade,
                               pass       = passNorm)
    
    ## Generate a lookup table for the report:
    tmp        <- 0 : nQuestions
    scoreTable <- data.frame(Score = tmp,
                             Grade = scoreExam(score      = tmp,
                                               what       = scoringIndex,
                                               nQuestions = nQuestions,
                                               nOptions   = nOptions,
                                               minGrade   = minGrade,
                                               pass       = passNorm)
                             )
}

## Calculate the minimum passing score:
cutoff <- with(scoreTable, Score[sum(Grade < 5.5) + 1])

if(campus) {
    ## Check the results:
    tmp   <- merge(pooled, campusData0, by = "snr")
    check <- all(with(tmp, result - result0) == 0)
    
    if(!check)
        wrappedWarning("Something may have gone wrong. The exam results I've calculated do not match the results provided by the Student Administration.",
                       immediate. = TRUE)
}

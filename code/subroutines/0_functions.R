### Title:    Subroutines for TiU Exam Merging Utility
### Author:   Kyle M. Lang
### Created:  2020-10-13
### Modified: 2020-10-22

## Apply the exam committee's scoring rule for exams:
scoreExam <- function(score, nQuestions, nOptions, minGrade, pass = 0.55) {
    ## Compute score expected by guessing:
    guess <- nQuestions / nOptions

    ## Compute raw score with guessing correction:
    knowledge <- (score - guess) / (nQuestions - guess)
    knowledge[score < guess] <- 0 # Negative values are not allowed

    ## Convert raw score to a [minGrade, 10] scale:
    result <- rep(0, length(score))
    for(i in 1 : length(score)) {
        if(is.na(knowledge[i]))
            result[i] <- NA
        else if(knowledge[i] < pass)
            result[i] <- minGrade + knowledge[i] * (5.5 - minGrade) / pass
        else
            result[i] <- 5.5 + ((knowledge[i] - pass) / (1 - pass)) * (10 - 5.5)
    }
    round(result, 1)
}

###--------------------------------------------------------------------------###

## Prepare a scoring scheme based on a user-supplied lookup table:
prepScoringScheme <- function(file) {
    ## Read in the user-supplied lookup table:
    table           <- autoReadCsv(file, header = FALSE)$data
    colnames(table) <- c("Score", "Grade")

    ## Convert the lookup table into a named vector:
    scheme        <- table[[2]]
    names(scheme) <- table[[1]]

    list(scheme = scheme, table = table)
}

###--------------------------------------------------------------------------###

## Parse student names into surname and initials/first name:
parseName <- function(x) {
    ## Case 1.1: Surname, First Name
    ## Case 1.2: Surname, Initials
    tmp <- unlist(str_locate_all(x, ","))
    if(length(tmp) > 0) {
        name2 <- substr(x, 1, tmp[1] - 1)
        name1 <- str_trim(substr(x, tmp[1] + 1, nchar(x)))

        return(c(name1 = name1, name2 = name2))
    }

    ## Case 2.1: Surname Initials
    ## Case 2.2: Initials Surname
    tmp <- unlist(str_locate_all(x, "[A-Z]{1}\\."))
    if(length(tmp) > 0) {
        if(tmp[1] > 1) {
            name2 <- substr(x, 1, tmp[1] - 1)
            name1 <- substr(x, tmp[1], nchar(x))
        }
        else {
            name1 <- substr(x, 1, max(tmp))
            name2 <- str_trim(substr(x, max(tmp) + 1, nchar(x)))
        }
        return(c(name1 = name1, name2 = name2))
    }

    ## Case 3: First Name Surname
    tmp <- unlist(str_locate_all(x, "\\s"))
    if(length(tmp) > 0) {
        name1 <- substr(x, 1, tmp[1] - 1)
        name2 <- str_trim(substr(x, tmp[1], nchar(x)))

        return(c(name1 = name1, name2 = name2))
    }

    ## Return raw name when no case above is satisfied
    c(name1 = x, name2 = x)
}

###--------------------------------------------------------------------------###

## Wrap error messages:
wrappedError <- function(msg, width = 79)
    stop(paste(strwrap(msg, width), collapse = "\n"), call. = FALSE)

###--------------------------------------------------------------------------###

## Wrap warning messages:
wrappedWarning <- function(msg, width = 79, ...)
    warning(paste(strwrap(msg, width), collapse = "\n"), call. = FALSE, ...)

###--------------------------------------------------------------------------###

## Find the index of the column containing the online exam results:
findExam <- function(data,
                     names,
                     pattern = "\\d{4}\\.\\d{2}\\.\\d{2}.*Remotely\\.Proctored|OPT\\.OUT")
{
    ## Try to find the exam column in the Online gradebook:
    examCol <- grep(pattern, colnames(data))

    ## We found one candidate column:
    if(length(examCol) == 1) {
        ## Confirm the exam column with the user:
        msg <- paste0("I think I've found your online exam results in this column:\n\n",
                      names[examCol],
                      "\n\nAm I correct?")
        success <- dlgMessage(message = msg, type = "yesno")$res
    }

    ## We found multiple candiate columns:
    if(length(examCol) > 1) {
        ## Confirm the exam column with the user:
        dlgMessage("I've found multiple columns that seem like they may contain your online exam results. So, you'll need to select the appropriate column(s).")

        opt0      <- "None of these columns contains my online exam results."
        selection <- dlgList(choices  = c(names[examCol], opt0),
                             multiple = TRUE,
                             title    = "Select online exam results")$res

        fail <- length(selection) == 0 | opt0 %in% selection
        if(fail)
            success <- "no"
        else {
            success <- "yes"
            examCol <- which(names %in% selection)
        }
    }

    ## Ask the user to select the exam column, if we can't find it automatically:
    if(length(examCol) == 0 || length(success) == 0 || success == "no") {
        dlgMessage("I haven't been able to automatically detect your online exam results. So, I need you to identify the appropriate column(s).")

        selection <- dlgList(choices  = names,
                             multiple = TRUE,
                             title    = "Select online exam results.")$res

        if(length(selection) == 0)
            wrappedError("I cannot proceed without knowing which column contains your online exam results.")

        examCol <- which(names %in% selection)
    }

    ## Return the column index for the online exam results:
    examCol
}

###--------------------------------------------------------------------------###

## Automatically detect the field delimiter in a CSV file and read its contents.
autoReadCsv <- function(file, ...) {
    semiColonSize <-
        length(scan(file, what = "character", sep = ";", quiet = TRUE))
    commaSize     <-
        length(scan(file, what = "character", sep = ",", quiet = TRUE))

    sep <- ifelse(semiColonSize > commaSize, ";", ",")
    out <- read.csv(file, sep = sep, ...)

    list(data = out, sep = sep)
}

###--------------------------------------------------------------------------###

## Process one on-campus results file(s):
processCampus <- function(filePath) {
    ## Read in one set of on-campus grades:
    data <- read.xlsx(filePath, sheetIndex = 1, stringsAsFactors = FALSE)

    ## Extract first and second columns:
    c1 <- data[[1]]
    c2 <- data[[2]]

    ## Find the rows containing individual grades:
    gradeRows <- grep("\\d{6,7}", c1)

    ## Extract the individual grade information:
    grades           <- as.data.frame(data[gradeRows, ])
    colnames(grades) <- as.character(data[gradeRows[1] - 1, ])

    ## Check for students with broken SNRs:
    extraCol <- grades[[3]]
    snrFlag  <- !is.na(extraCol)

    ## Replace any broken SNRs with the corrections provided by the SA:
    if(any(snrFlag))
        for(i in which(snrFlag)) {
            x   <- extraCol[i]
            tmp <- unlist(str_locate_all(x, "\\d{6,7}\\s"))

            grades[i, "S Nummer"]  <- substr(x, tmp[1], tmp[2] - 1)
            grades[i, "Naam"]      <- substr(x, tmp[2] + 1, nchar(x))
        }

    ## Parse student names:
    stuNames <- sapply(grades$Naam, parseName, USE.NAMES = FALSE)

    ## Extract metadata from the on-campus file:
    cn       <- colnames(data)
    cn       <- gsub("^X", "", cn)
    faculty  <- tolower(cn[grep("Opleiding", cn) + 1])
    batchId  <- as.numeric(cn[grep("Batch\\.ID", cn) + 1])
    examName <- c2[grep("Toetsnaam", c1)]
    examDate <- as.character(
        as.Date(as.numeric(cn[grep("Toetsdatum", cn) + 1]),
                origin = "1899-12-30")
    )

    ## Extract the relevent columns:
    outData <- data.frame(
        grades[ , c("S Nummer", "Score", "Versie")],
        surname   = stuNames["name2", ],
        firstName = stuNames["name1", ],
        source    = "Campus")
    colnames(outData)[1 : 3] <- c("snr", "score", "version")

    ## Remove any students without SNRs or scores:
    outData <- outData[with(outData, !is.na(snr) & !is.na(score)), ]

    ## Extract SA-computed result for testing purposes:
    outData0           <- grades[ , c("S Nummer", "Resultaat")]
    colnames(outData0) <- c("snr", "result0")

    list(data    = outData,
         data0   = outData0,
         faculty = faculty,
         id      = batchId,
         name    = examName,
         date    = examDate)
}

###--------------------------------------------------------------------------###

                                        #index <- examCol
                                        #data  <- onlineData
                                        #names <- onlineNames

## Process the online results file:
processOnline <- function(index, data, names) {
    ## Extract metadata from online exam name:
    examName <- names[index]
    tmp      <- str_locate_all(examName, c("/", "\\(Remotely Proctored|OPT-OUT"))

    examDate <- tryCatch(substr(examName, 1, tmp[[1]][1, 1] - 1),
                          error = function(e) "Not Recovered")

    courseCode <- tryCatch(substr(examName, tmp[[1]][1, 1] + 1, tmp[[1]][2, 2] - 1),
                           error = function(e) "Not Recovered")

    if(grepl("proctored", examName, ignore.case = TRUE))
        version <- "Proctored"
    else if(grepl("opt-out", examName, ignore.case = TRUE))
        version <- "Not Proctored"
    else
        version <- ""

    tmp <- try(
        substr(examName, tmp[[1]][2, 2] + 1, tmp[[2]][1, 1] - 1),
        silent = TRUE
    )
    if(class(tmp) != "try-error") examName <- str_trim(tmp)


    ## Parse student names:
    stuNames <- sapply(data$Student, parseName, USE.NAMES = FALSE)

    ## Extract the relevent columns:
    outData <- data.frame(data[ , c("SIS.User.ID",
                                         colnames(data)[index])
                                    ],
                         surname          = stuNames["name2", ],
                         firstName        = stuNames["name1", ],
                         version          = version,
                         source           = "Online",
                         stringsAsFactors = FALSE)
    colnames(outData)[1 : 2] <- c("snr", "score")

    outData$snr
    outData$score
    
    ## Remove any students without SNRs or scores:
    drops   <- with(outData, empty(snr) | empty(score))
    outData <- outData[!drops, ]

    list(data = outData, name = examName, date = examDate, id = courseCode)
}

###--------------------------------------------------------------------------###

## Find different flavors of empty cell:
empty <- function(x, key = NULL)
    is.na(x) | length(x) == 0 | x == "" | x == "N/A" | x %in% key 

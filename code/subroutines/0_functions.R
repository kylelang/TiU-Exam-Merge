### Title:    Subroutines for TiU Exam Merging Utility
### Author:   Kyle M. Lang
### Created:  2020-10-13
### Modified: 2020-10-15

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
                     pattern = "\\d{4}\\.\\d{2}\\.\\d{2}.*Remotely\\.Proctored")
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
        dlgMessage("I've found multiple columns that seem like they may contain your online exam results. So, you'll need to select the appropriate column.")

        opt0      <- "None of these columns contains my online exam results."
        selection <- dlgList(choices = c(names[examCol], opt0),
                             title = "Select online exam results")$res

        if(selection == opt0)
            success <- "no"
        else {
            success <- "yes"
            examCol <- which(names %in% selection)
        }
    }

    ## Ask the user to select the exam column, if we can't find it automatically:
    if(length(examCol) == 0 || length(success) == 0 || success == "no") {
        dlgMessage("I haven't been able to automatically detect your online exam results. So, I need you to identify the appropriate column.")

        selection <- dlgList(choices = names,
                             title = "Select online exam results.")$res

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

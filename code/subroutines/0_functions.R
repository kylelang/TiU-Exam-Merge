### Title:    Subroutines for TiU Exam Merging Utility
### Author:   Kyle M. Lang
### Created:  2020-10-13
### Modified: 2020-12-15


## Score the exam according to one of the three functional scoring rules:
scoreExam <- function(score, what, nQuestions, nOptions, minGrade, pass = 0.55) 
    switch(what,
           ## Post-2020 standard guessing correction:
           scoreRule1(score      = score,
                      nQuestions = nQuestions,
                      nOptions   = nOptions,
                      minGrade   = minGrade,
                      pass       = pass),
           ## First option from work order form:
           scoreRule2(score      = score,
                      nQuestions = nQuestions,
                      nOptions   = nOptions,
                      minGrade   = 1),
           ## Second option from work order form:
           scoreRule1(score      = score,
                      nQuestions = nQuestions,
                      nOptions   = nOptions,
                      minGrade   = 0)
           )

###--------------------------------------------------------------------------###

## Function implementing the exam committee's post-2020 scoring rule:
scoreRule1 <- function(score, nQuestions, nOptions, minGrade, pass = 0.55) {
    ## Compute score expected by guessing:
    guess <- nQuestions / nOptions

    ## Compute raw score with guessing correction:
    knowledge <- (score - guess) / (nQuestions - guess)
    knowledge[score < guess] <- 0 # Negative values are not allowed

    ## Convert raw score to a [minGrade, 10] scale:
    grade <- rep(0, length(score))
    for(i in 1 : length(score)) {
        if(is.na(knowledge[i]))
            grade[i] <- NA
        else if(knowledge[i] < pass)
            grade[i] <- minGrade + knowledge[i] * (5.5 - minGrade) / pass
        else
            grade[i] <- 5.5 + ((knowledge[i] - pass) / (1 - pass)) * 4.5
    }
    round(grade, 1)
}

###--------------------------------------------------------------------------###

## Function implementing the first scoring option from the work order form:
scoreRule2 <- function(score, nQuestions, nOptions, minGrade = 1) {
    guess <- nQuestions / nOptions
    pass  <- guess + ((nQuestions - guess) / 2)
    
    grade <- ((5.5 - minGrade) / pass) * score + minGrade
    tmp   <- (4.5 / (nQuestions - pass)) * (score - pass) + 5.5

    grade[score >= pass] <- tmp[score >= pass]
    round(grade, 1)
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
        snr       = suppressWarnings(as.numeric(grades[ , "S Nummer"])),
        score     = suppressWarnings(as.numeric(grades$Score)),
        version   = grades$Versie,
        surname   = stuNames["name2", ],
        firstName = stuNames["name1", ],
        source    = "Campus"
    )
    
    ## Remove any students without SNRs or scores:
    outData <- outData[with(outData, !is.na(snr) & !is.na(score)), ]
    
    ## Extract SA-computed result for testing purposes:
    outData0 <- data.frame(
        snr     = suppressWarnings(as.numeric(grades[ , "S Nummer"])),
        result0 = suppressWarnings(as.numeric(grades$Resultaat))
    )
    
    list(data    = outData,
         data0   = outData0,
         faculty = faculty,
         id      = batchId,
         name    = examName,
         date    = examDate)
}

###--------------------------------------------------------------------------###

## Process the Canvas-based results file:
processCanvas <- function(index, data, names, ...) {
    ## Extract metadata from online exam name:
    examName <- names[index]
    tmp      <- str_locate_all(examName,
                               c("/", "\\(Remotely Proctored|OPT-OUT")
                               )
    
    examDate <- tryCatch(substr(examName, 1, tmp[[1]][1, 1] - 1),
                         error = function(e) "Not Recovered")
    
    courseCode <- tryCatch(substr(examName,
                                  tmp[[1]][1, 1] + 1,
                                  tmp[[1]][2, 2] - 1),
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
    outData <- data.frame(
        snr              = suppressWarnings(as.numeric(data$SIS.User.ID)),
        score            = suppressWarnings(as.numeric(data[ , index])),
        surname          = stuNames["name2", ],
        firstName        = stuNames["name1", ],
        version          = version,
        source           = "Canvas",
        stringsAsFactors = FALSE
    )
    
    ## Remove any students without SNRs or scores:
    drops   <- with(outData, empty(snr, ...) | empty(score, ...))
    outData <- outData[!drops, ]
    
    list(data = outData, name = examName, date = examDate, id = courseCode)
}

###--------------------------------------------------------------------------###

## Process the TestVision-based results file:
processTestVision <- function(data) {
    ## Extract metadata from online results file:
    examName   <- data[1, "ToetsNaam"]
    examDate   <- as.character(
        as.Date(data[1, "AfnameBegindatum"], format = "%d-%m-%Y")
    )
    courseCode <- data[1, "MapNaamToets"]
    version    <- ifelse(data$Proctoring == "Ja", "Proctored", "Not Proctored")
    
    ## Parse student names:
    stuNames <- sapply(data$KandidaatWeergavenaam, parseName, USE.NAMES = FALSE)

    ## Process ANRs:
    anr <- as.numeric(gsub("u", "", data$KandidaatAanmeldnaam))
    
    ## Extract the relevent columns:
    outData <- data.frame(anr              = anr,
                          score            = data$ToetsScore,
                          surname          = stuNames["name2", ],
                          firstName        = stuNames["name1", ],
                          version          = version,
                          source           = "TestVision",
                          stringsAsFactors = FALSE)
    
    list(data = outData, name = examName, date = examDate, id = courseCode)
}

###--------------------------------------------------------------------------###

## Find different flavors of empty cell:
empty <- function(x, codes = NULL)
    is.na(x) | length(x) == 0 | x == "" | x %in% codes

###--------------------------------------------------------------------------###

## Compute Cohen's H effect size measure:
cohenH <- function(x) as.numeric(2 * abs(diff(asin(sqrt(x)))))

###--------------------------------------------------------------------------###

## Compute Cohen's D effect size measure:
cohenD <- function(x, group) {
    n  <- table(group)
    m  <- tapply(x, group, mean)
    s2 <- tapply(x, group, var)
    
    as.numeric(
        abs(diff(m)) /
        sqrt(((n[1] - 1) * s2[1] + (n[2] - 1) * s2[2]) / (sum(n) - 2))
    )
}

###--------------------------------------------------------------------------###

## Compare online and on-campus exam scores to check for irregularities:
compareScores <- function(data) {
    ## Define the grouping variable:
    groups <- tolower(data$source)
    groups <- factor(gsub("canvas|testvision", "online", groups),
                     levels = c("online", "campus")
                     )

    ## Store groups sizes:
    n <- table(groups)
    
    ## How many students passed each exam?
    tmp  <- factor(data$result >= 5.5, levels = c("FALSE", "TRUE"))
    tab6 <- table(groups, pass = tmp)
    or6  <- (tab6["online", "TRUE"] / tab6["online", "FALSE"]) /
        (tab6["campus", "TRUE"] / tab6["campus", "FALSE"])
    
    ## How many students passed with a score of 8 or higher?
    tmp  <- factor(data$result >= 7.75, levels = c("FALSE", "TRUE"))
    tab8 <- table(groups, pass = tmp)
    or8  <- (tab8["online", "TRUE"] / tab8["online", "FALSE"]) /
        (tab8["campus", "TRUE"] / tab8["campus", "FALSE"])
    
    ## Compare mean passing rates:
    tOut <- t.test(data$result ~ groups, var.equal = FALSE)
    
    ## Compare proportions of passing students:
    if(sum(tab6[ , "TRUE"]) > 0) {
        if(all(tab6 >= 5))
            test6 <- prop.test(tab6, correct = FALSE)
        else
            test6 <- fisher.test(tab6)
    } else {
        test6 <- NA
    }

    ## Compare proportions of Cum Laude students:
    if(sum(tab8[ , "TRUE"]) > 0) {
        if(all(tab8 >= 5))
            test8 <- prop.test(tab8, correct = FALSE)
        else
            test8 <- fisher.test(tab8)
    } else {
        test8 <- NA
    }
    
    list(n      = n,
         mean   = tapply(data$result, groups, mean),
         count6 = tab6[ , "TRUE"],
         count8 = tab8[ , "TRUE"],
         or6    = or6,
         or8    = or8,
         tOut   = tOut,
         test6  = test6,
         test8  = test8,
         d      = cohenD(data$result, groups),
         h6     = cohenH(tab6[ , "TRUE"] / n),
         h8     = cohenH(tab8[ , "TRUE"] / n)
         )
}

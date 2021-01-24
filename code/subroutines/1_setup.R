### Title:    Define the Parameters of a TiU Exam Combination Job
### Author:   Kyle M. Lang
### Created:  2020-10-13
### Modified: 2021-01-24


## Define legal file types for online results file:
csvFilters <- matrix(c("Comma-Seperated Values", "*.csv;*.CSV"), 1, 2)

## Prompt the user to select the file path to the CSV file containing the online
## test results:
onlineFile <- dlgOpen(default = dir0,
                      title   = "Please select the file that contains the online test results.",
                      filters = csvFilters)$res
                      
## Derive the starting directory for future file-selection dialogs:
dir0 <- paste0(dirname(dirname(onlineFile)), "/*")

if(length(onlineFile) == 0)
    wrappedError("I cannot proceed without knowing where to find your online test results.")

## Define legal file types for on-campus results file:
xlsxFilters <- matrix(c("Office Open XML Spreadsheet", "*.xlsx;*.XLSX",
                     "Excel 2007 - 2019", "*.xlsx;*.XLSX"),
                   2, 2, byrow = TRUE)

## Ask the user if they need to process any on-campus exams:
campus <- dlgMessage(message = "Do you have any on-campus results to process?",
                     type    = "yesno")$res == "yes"

## Prompt the user to select the file path(s) to the XLSX file(s) containing the
## on-campus test results:
if(campus) {
    campusFile <-
        dlgOpen(default  = dir0,
                title    = "Please select the file(s) that contain(s) the on-campus test results.",
                multiple = TRUE,
                filters  = xlsxFilters)$res

    if(length(campusFile) == 0)
        wrappedError("I cannot proceed without knowing where to find your on-campus test results.")
}

## Prompt the user to select a scoring scheme:
scoreOpts <- list(new = "Post-2020 standard guessing correction",
                  wo1 = "Work order option 1",
                  wo2 = "Work order option 2",
                  wo3 = "Work order option 3 (Custom scoring scheme)")

scoreScheme <- dlgList(choices = scoreOpts,
                       title   = "Which scoring rule would you like to use?")$res

if(scoreScheme == scoreOpts["wo3"]) {
    ## Prompt the user to select the file path to the CSV file containing the
    ## lookup table describing the custom scoring scheme:
    tableFile <- dlgOpen(default = dir0,
                         title   = "Please select the file that contains the lookup table defining the custom scoring scheme.",
                         filters = csvFilters)$res

    ## Check the file type:
    check <- grepl("\\.csv$", tableFile, ignore.case = TRUE)
    if(!check)
        wrappedError("The lookup table must be a CSV file.")
    
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
    check <- length(nQuestions) == 0 | length(nOptions) == 0
    if(check)
        wrappedError("I need to know how many questions this exam contains and the number of response options for each question before I can proceed.")
}

if(scoreScheme == scoreOpts["new"]) {
    ## Prompt the user to define the passing norm to use when scoring the exam:
    passNorm   <- as.numeric(
        dlgInput("What norm would you like to use to define a passing grade?",
                 default = 0.55)$res
    )

    if(!campus) {
        ## Prompt the user to define the faculty to which the exam belongs:
        dlgMessage("Without any on-campus results, I can't automatically detect the faculty to which this exam belongs. So, you'll need to specify the appropriate faculty.")
        faculty <- tolower(
            dlgList(choices = list("TSB",
                                   "TSHD",
                                   "TLS",
                                   "TiSEM",
                                   "TST",
                                   "I don't know the faculty"),
                    title   = "Select the faculty")$res
        )
    }
} else {
    passNorm <- NULL
}

## Ask the user to define the ouput file(s):
dlgMessage("Finally, I need you to tell me where you would like to save the results.")

windows <- Sys.info()["sysname"] == "Windows"
outFile <- getOutputFile(windows, maxNameLength, dir0)

## Define an output file for the irregularity checks:
if(campus) {
    tmp        <- strsplit(outFile, ".xlsx", fixed = TRUE)[[1]]
    checksFile <- paste(tmp, "irregularity_checks.xlsx", sep = "-")

    ## Check for file length issues on Windows:
    fileLen <- nchar(checksFile)
    if(windows & fileLen > maxNameLength) {
        newFile <- dlgMessage(
            paste(
                "The filepath that I've automatically generated for your irregularity checks",
                paste0("(", checksFile, ")"),
                "is",
                fileLen,
                "characters long. Windows may not be able to handle filepaths of this length. Would you like to select a different location in which to save the results of your irregularity checks?"
            ),
            "yesno"
        )$res

        ## Ask the user to define a new file, if necessary:
        if(newFile == "yes")
            checksFile <- getOutputFile(windows, maxNameLength, dir0)
    }
}

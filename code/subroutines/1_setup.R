### Title:    Define the Parameters of a TiU Exam Combination Job
### Author:   Kyle M. Lang
### Created:  2020-10-13
### Modified: 2020-10-20


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

## Ask the user how many on-campus test files they want to read in:
campusCount <- as.numeric(
    dlgInput("How many sets (i.e., unique files) of on-campus results do you need to process?")$res
)

## Prompt the user to select the file path(s) to the XLSX file(s) containing the
## on-campus test results:
if(campusCount == 1) {
    msg <- "Please select the file that contains the on-campus test results."
} else {
    msg <- "Please select the file that contains the first set of on-campus test results."
}

campusFile <- list()
for(i in 1 : campusCount) {
    if(i == 2)
        msg <- gsub("first", "next", msg)
    
    campusFile[[i]] <- dlgOpen(title = msg, filters = xlsxFilters)$res
}

if(length(campusFile) == 0)
    wrappedError("I cannot proceed without knowing where to find your on-campus test results.")

customScheme <-
    dlgMessage(message = "Did the instructor request a custom scoring scheme?",
               type    = "yesno")$res

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

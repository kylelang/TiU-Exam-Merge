### Title:    Process TiU Hybrid Exams
### Author:   Kyle M. Lang
### Created:  2020-10-14
### Modified: 2020-10-22

rm(list = ls(all = TRUE)) # Clear the workspace

## By deault, missing scores are assumed to be represented by empty cells.
## Use this argument to define a character vector containing all of the special
## character strings used to indicate missing exam scores in the Canvas
## gradebook download:
missingScoreCodes <- c("N/A") 

# Set this argument to 1 if TiSEM adopts a 10-point grade scale:
tisemMinGrade <- 0 

source("subroutines/0_functions.R")

library(stringr)
library(xlsx)
library(svDialogs)

finished <- FALSE
while(!finished) {
    ## Interactively parameterize the job:
    source("subroutines/1_setup.R")

    ## Process and combine the exam results:
    source("subroutines/2_process.R")

    ## Generate an XLSX report:
    source("subroutines/3_report.R")
    
    finished <-
        dlgMessage("Do you want to process another exam?", "yesno")$res == "no"
}

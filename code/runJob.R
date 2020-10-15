### Title:    Process TiU Hybrid Exams
### Author:   Kyle M. Lang
### Created:  2020-10-14
### Modified: 2020-10-15

rm(list = ls(all = TRUE)) # Clear the workspace

tisemMinGrade <- 0 # Set this to 1 if TiSEM adopts a 10-point grade scale

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
    
    answer   <- dlgMessage("Do you want to process another exam?", "yesno")$res
    finished <- answer == "no"
}

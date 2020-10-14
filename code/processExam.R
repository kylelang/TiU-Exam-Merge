### Title:    Process TiU Hybrid Exams
### Author:   Kyle M. Lang
### Created:  2020-10-14
### Modified: 2020-10-14
### Note:     This script is basically a wrapper for 'combine.R' that allows
###           multiple exams to be processed sequentially.

rm(list = ls(all = TRUE)) # Clear the workspace

source("subroutines.R")

library(stringr)
library(xlsx)
library(svDialogs)

finished <- FALSE

while(!finished) {
    ## Interactively process one exam:
    source("combine.R")
    
    answer   <- dlgMessage("Do you want to process another exam?", "yesno")$res
    finished <- answer == "no"
}

### Title:    Create an XLSX Output File
### Author:   Kyle M. Lang
### Created:  2020-10-15
### Modified: 2021-01-24


## Create an empty workbook and a new sheet therein:
wb1 <- createWorkbook()
s1  <- createSheet(wb = wb1, sheetName = "Results")

## Define some useful formats:
BoldLeftStyle <- CellStyle(wb1) +
    Font(wb1, isBold = TRUE) +
    Alignment(horizontal = "ALIGN_LEFT")
BoldRightStyle <- CellStyle(wb1) +
    Font(wb1, isBold = TRUE) +
    Alignment(horizontal = "ALIGN_RIGHT")
LeftStyle  <- CellStyle(wb1) + Alignment(horizontal = "ALIGN_LEFT")
RightStyle <- CellStyle(wb1) + Alignment(horizontal = "ALIGN_RIGHT") 

###--------------------------------------------------------------------------###

### Meta Data ###

## Populate column names for metadata block:
blockData <- c("Exam Date", "Identifier", "Exam Name")
cb        <- CellBlock(s1, 1, 1, 1, 3)
CB.setRowData(cb, blockData, 1, rowStyle = BoldLeftStyle)

## Populate contents of metadata block:
blockData <- matrix(c(sapply(meta, "[[", x = "date"),
                      sapply(meta, "[[", x = "id"),
                      sapply(meta, "[[", x = "name")
                      ),
                    ncol = 3)

metaRows  <- nrow(blockData) + 1
cb        <- CellBlock(s1, 2, 1, metaRows, 3)
CB.setMatrixData(cb, blockData, 1, 1, cellStyle = LeftStyle)


### Summary Measures ###

## Populate row names for summary measures block:
blockData <- c("Cutoff Score",
               "Passed (%)",
               "Average Grade",
               "Average Score",
               "Students (N)")
cb        <- CellBlock(s1, metaRows + 3, 1, 5, 1)
CB.setColData(cb, blockData, 1, colStyle = BoldLeftStyle)

## Populate contents of summary measures block:
blockData <- matrix(
    c(cutoff,
      with(pooled,
           c(roundUp(100 * mean(result > 5.5), 1),
             roundUp(mean(result), 1),
             roundUp(mean(score), 1),
             length(score)
             )
           )
      )
)
cb        <- CellBlock(s1, metaRows + 3, 2, 5, 1)
CB.setColData(cb, blockData, 1, colStyle = LeftStyle)


### Warning Message ###

## Populate heading of warning message:
blockData <- "Please note:" 
cb        <- CellBlock(s1, metaRows + 10, 1, 1, 1)
CB.setRowData(cb, blockData, 1, rowStyle = BoldLeftStyle)

## Populate contents of warning message:
blockData <- "Neither Canvas nor TestVision can output special characters. Therefore, some student names might be displayed improperly."
cb        <- CellBlock(s1, metaRows + 10, 2, 1, 1)
CB.setRowData(cb, blockData, 1, rowStyle = LeftStyle)


### Student Results ###

## Populate column names for student results block:
blockData <- c("Surname", "Initials/First Name")

if(canvas | campus)
    blockData <- c(blockData, "SNR")

if(!canvas)
    blockData <- c(blockData, "ANR")

## How many columns will the first block fill?
firstCols <- length(blockData)

cb        <- CellBlock(s1, metaRows + 13, 1, 1, firstCols)
CB.setRowData(cb, blockData, 1, rowStyle = BoldLeftStyle)

blockData <- c("Grade", "Score")
cb        <- CellBlock(s1, metaRows + 13, firstCols + 1, 1, 2)
CB.setRowData(cb, blockData, 1, rowStyle = BoldRightStyle)

blockData <- c("Source", "Version")
cb        <- CellBlock(s1, metaRows + 13, firstCols + 3, 1, 2)
CB.setRowData(cb, blockData, 1, rowStyle = BoldLeftStyle)

## Populate contents of student results block:
addDataFrame(pooled[c("surname", "firstName")],
             sheet     = s1,
             col.names = FALSE,
             row.names = FALSE,
             startRow  = metaRows + 14,
             startCol  = 1,
             colStyle  = list("3" = LeftStyle)
             )

if(canvas | campus)
    addDataFrame(pooled["snr"],
                 sheet     = s1,
                 col.names = FALSE,
                 row.names = FALSE,
                 startRow  = metaRows + 14,
                 startCol  = 3,
                 colStyle  = list("3" = LeftStyle)
                 )

if(!canvas)
    addDataFrame(pooled["anr"],
                 sheet     = s1,
                 col.names = FALSE,
                 row.names = FALSE,
                 startRow  = metaRows + 14,
                 startCol  = ifelse(campus, 4, 3),
                 colStyle  = list("3" = LeftStyle)
                 )

addDataFrame(pooled[c("result", "score")],
             sheet     = s1,
             col.names = FALSE,
             row.names = FALSE,
             startRow  = metaRows + 14,
             startCol  = firstCols + 1)

addDataFrame(pooled[c("source", "version")],
             sheet     = s1,
             col.names = FALSE,
             row.names = FALSE,
             startRow  = metaRows + 14,
             startCol  = firstCols + 3)

## Set column widths:
setColumnWidth(sheet = s1, colIndex = 1, colWidth = 14)
setColumnWidth(sheet = s1, colIndex = 2, colWidth = 18)
setColumnWidth(sheet = s1, colIndex = 3, colWidth = 12)
setColumnWidth(sheet = s1, colIndex = 4 : 8, colWidth = 10)

###--------------------------------------------------------------------------###

## Add another sheet to hold the scoring table:
s2 <- createSheet(wb = wb1, sheetName = "Scoring Table")

## Add column headings:
cb <- CellBlock(s2, 1, 1, 1, 2)
CB.setRowData(cb, c("Score", "Grade"), 1, rowStyle = BoldRightStyle)

## Calculate the cutpoint for the pass/fail blocks:
cut <- which(scoreTable$Score == cutoff)

## Adding failing half of the score table:
cb <- CellBlock(s2, 2, 1, cut, 2)
CB.setMatrixData(cb,
                 as.matrix(scoreTable[1 : (cut - 1), ]),
                 startRow    = 1,
                 startColumn = 1,
                 cellStyle   = RightStyle + Fill(rgb(255, 200, 200, max = 255))
                 )

## Adding passing half of the score table:
cb <- CellBlock(s2, cut + 1, 1, nrow(scoreTable) + 1, 2)
CB.setMatrixData(cb,
                 as.matrix(scoreTable[cut : nrow(scoreTable), ]),
                 startRow    = 1,
                 startColumn = 1,
                 cellStyle   = RightStyle + Fill(rgb(100, 200, 255, max = 255))
                 )

###--------------------------------------------------------------------------###

## Save the final workbook to disk:
saveWorkbook(wb1, outFile)

###--------------------------------------------------------------------------###

if(campus && checkIrreg) {
    ## Create a new workbook and sheet to contain the results of the score
    ## comparisons:
    wb1 <- createWorkbook()
    s1  <- createSheet(wb = wb1, sheetName = "Irregularity Checks")
    
    ## Define some useful formats:
    BoldLeftStyle <- CellStyle(wb1) +
        Font(wb1, isBold = TRUE) +
        Alignment(horizontal = "ALIGN_LEFT")
    BoldCenterStyle <- CellStyle(wb1) +
        Font(wb1, isBold = TRUE) +
        Alignment(horizontal = "ALIGN_CENTER")
    
    ## Add column super-headings:
    blockData <- c(
        rep("", 4),
        "Sample Size",
        "",
        "Mean Grade",
        "",
        "Grades >= 6",
        "",
        "Grades >= 8",
        "",
        "Difference in Mean Grades",
        rep("", 3),
        "Difference in Proportion of Grades >= 6",
        rep("", 3),
        "Difference in Proportion of Grades >= 8",
        rep("", 3)
    )
    cb <- CellBlock(s1, 1, 1, 1, length(blockData))
    CB.setRowData(cb, blockData, rowIndex = 1, rowStyle = BoldCenterStyle)

    ## Merge cells in super-headings:
    addMergedRegion(s1, 1, 1, 5, 6)
    addMergedRegion(s1, 1, 1, 7, 8)
    addMergedRegion(s1, 1, 1, 9, 10)
    addMergedRegion(s1, 1, 1, 11, 12)
    addMergedRegion(s1, 1, 1, 13, 16)
    addMergedRegion(s1, 1, 1, 17, 20)
    addMergedRegion(s1, 1, 1, 21, 24)
  
    ## Add column sub-headings:
    blockData <-
        c("Faculty",
          "Exam Name",
          "Course Code",
          "Exam Date",
          rep(c("Campus", "Online"), 4),
          "T-Statistic",
          "DF",
          "P-Value",
          "Cohen's d",
          "Online:Campus Odds-Ratio",
          "Chi-Squared Statistic",
          "P-Value",
          "Cohen's h",
          "Online:Campus Odds-Ratio",
          "Chi-Squared Statistic",
          "P-Value",
          "Cohen's h"
          )
    cb <- CellBlock(s1, 2, 1, 1, length(blockData))
    CB.setRowData(cb, blockData, rowIndex = 1, rowStyle = BoldLeftStyle)

    ## Resize columns that won't be adjusted later to fit their headings:
    autoSizeColumn(s1, 1 : 2)
    
    ## Add comparison data:
    tmp <- meta[[length(meta)]]
    tmp <- c(toupper(meta[[1]]$faculty),
             tmp$name,
             tmp$id,
             tmp$date,
             with(scoreComps,
                  c(n[c("campus", "online")],
                    roundUp(mean[c("campus", "online")], 2),
                    count6[c("campus", "online")],
                    count8[c("campus", "online")],
                    with(tOut, c(roundUp(statistic, 2),
                                 roundUp(parameter, 2),
                                 roundUp(p.value, 3)
                                 )
                         ),
                    roundUp(d, 3)
                    )
                  )
             )
    
    check <- scoreComps$tOut$p.value < 0.001
    if(check) tmp[15] <- "<0.001"
    
    ## Data involving counts of passing students:
    check <- class(scoreComps$test6) == "htest"
    if(check) {
        tmp2 <- with(scoreComps,
                     c(roundUp(or6, 2),
                       with(test6,
                            c(ifelse(grepl("Fisher's", method),
                                     "Not Applicable",
                                     roundUp(statistic, 2)
                                     ),
                              roundUp(p.value, 3)
                              )
                            ),
                       roundUp(h6, 3)
                       )
                     )
    } else {
        tmp2 <- rep("Not Applicable", 4)
    }
    
    check <- check && scoreComps$test6$p.value < 0.001
    if(check) tmp2[3] <- "<0.001"

    if(tmp2[1] == "Inf") tmp2[1] <- "Undefined"

    tmp <- c(tmp, tmp2)
    
    ## Data involving counts of students score 8 or higher:
    check <- class(scoreComps$test8) == "htest"
    if(check) {
        tmp2 <- with(scoreComps,
                     c(roundUp(or8, 2),
                       with(test8,
                            c(ifelse(grepl("Fisher's", method),
                                     "Not Applicable",
                                     roundUp(statistic, 2)
                                     ),
                              roundUp(p.value, 3)
                              )
                            ),
                       roundUp(h8, 3)
                       )
                     )
    } else {
        tmp2 <- rep("Not Applicable", 4)
    }
    
    check <- check && scoreComps$test8$p.value < 0.001
    if(check) tmp2[3] <- "<0.001"
    
    if(tmp2[1] == "Inf") tmp2[1] <- "Undefined"
    
    blockData <- c(tmp, tmp2)
    cb        <- CellBlock(s1, 3, 1, 1, length(blockData))
    CB.setRowData(cb, blockData, rowIndex = 1)

    ## Resize the columns to fit the new data (except for the exam name):
    autoSizeColumn(s1, 3 : 24)
    
    ## Save the final workbook to disk:
    saveWorkbook(wb1, checksFile)
    
}# CLOSE if(campus)

### Title:    Create an XLSX Output File
### Author:   Kyle M. Lang
### Created:  2020-10-15
### Modified: 2020-10-24


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
           c(round(100 * mean(result > 5.5), 1),
             round(mean(result), 1),
             round(mean(score), 1),
             length(score)
             )
           )
      )
)
cb        <- CellBlock(s1, metaRows + 3, 2, 5, 1)
CB.setColData(cb, blockData, 1, colStyle = LeftStyle)

## Populate heading of warning message:
blockData <- "Please note:" 
cb        <- CellBlock(s1, metaRows + 10, 1, 1, 1)
CB.setRowData(cb, blockData, 1, rowStyle = BoldLeftStyle)

## Populate contents of warning message:
blockData <- "Canvas cannot output special characters. Therefore, some student names might be displayed improperly."
cb        <- CellBlock(s1, metaRows + 10, 2, 1, 1)
CB.setRowData(cb, blockData, 1, rowStyle = LeftStyle)

## Populate column names for student results block:
blockData <- c("Surname", "Initials/First Name", "SNR")
cb        <- CellBlock(s1, metaRows + 13, 1, 1, 3)
CB.setRowData(cb, blockData, 1, rowStyle = BoldLeftStyle)

blockData <- c("Grade", "Score")
cb        <- CellBlock(s1, metaRows + 13, 4, 1, 2)
CB.setRowData(cb, blockData, 1, rowStyle = BoldRightStyle)

blockData <- c("Source", "Version")
cb        <- CellBlock(s1, metaRows + 13, 6, 1, 2)
CB.setRowData(cb, blockData, 1, rowStyle = BoldLeftStyle)

## Populate contents of student results block:
addDataFrame(pooled[c("surname", "firstName", "snr")],
             sheet     = s1,
             col.names = FALSE,
             row.names = FALSE,
             startRow  = metaRows + 14,
             startCol  = 1,
             colStyle  = list("3" = LeftStyle)
             )

addDataFrame(pooled[c("result", "score")],
             sheet     = s1,
             col.names = FALSE,
             row.names = FALSE,
             startRow  = metaRows + 14,
             startCol  = 4)

addDataFrame(pooled[c("source", "version")],
             sheet     = s1,
             col.names = FALSE,
             row.names = FALSE,
             startRow  = metaRows + 14,
             startCol  = 6)

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

## Set column widths:
setColumnWidth(sheet = s1, colIndex = 1, colWidth = 14)
setColumnWidth(sheet = s1, colIndex = 2, colWidth = 18)
setColumnWidth(sheet = s1, colIndex = 3, colWidth = 12)
setColumnWidth(sheet = s1, colIndex = 4 : 8, colWidth = 10)

## Save the final workbook to disk:
saveWorkbook(wb1, outFile)

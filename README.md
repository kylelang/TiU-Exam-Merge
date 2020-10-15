# TiU-Exam-Merge

This repository provides a utility to combine results from the online and
on-campus version of hybrid exams at Tilburg University. The utility will output
a report for instructors as an XLSX workbook.

## Usage

To use this utility, you simply need to source the `code/runJob.R` script
from within an interactive R session and respond to the interactive dialog
boxes.

## Scoring Notes

The program will ask if the instructor has provided a custom scoring
scheme. After responding in the affirmative, you must supply a lookup table
defining the custom scoring scheme. This lookup table must be a CSV file with
two columns and no column names. The first column should contain all possible
exam scores, and the second column should contain the corresponding grades. As
example lookup table describing a simple linear scoring for an exam with 50
questions is available in `data/lookup_table.csv`.

When the instructor does not supply a custom scoring scheme (and, consequently,
the above question is answered in the negative), the exam will be scored with
Tilburg University's default scoring scheme. This scheme applies a guessing
correction so that all scores equal to or lower than the score expected from
guessing are mapped onto the minimum grade. A reference implementation of the
default scoring scheme (provided by the TiU Examination Committee) is available
in `reference/grading.R`.

When applying the default scoring scheme, you will be prompted to provide three
pieces of information:

1. The number of questions on the exam
1. The number of response options (i.e., alternatives) for each question
1. The passing norm to use in defining the minimum passing score

The passing norm must be a value between 0 and 1 that corresponds to the minimum
guessing-corrected proportional score that a student must achieve to earn a
grade of 5.5. By default, this value is set to 0.55. If you are not sure of what
value to choose, accept the default.

The final report will contain a sheet showing the scoring table corresponding to
whatever scoring scheme was applied to the exam. Cells in this table that
correspond to failing grades are filled with light red, and cells corresponding
to passing grades are filled with light blue.

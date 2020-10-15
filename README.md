# TiU-Exam-Merge

This repository provides a utility to combine results from the online and
on-campus version of hybrid exams at Tilburg University. The utility will output
a report for instructors as an XLSX workbook.

## Usage

To use this utility, you simply need to source the `code/runJob.R` script
from within an interactive R session and respond to the interactive dialog
boxes.

## Input Notes

The results of the online examination should be provided as part of a standard
Canvas gradebook download. This file must be in CSV format. The utility should
work with CSV files that use commas `","` or semi-colons `";"` as the field
delimiter.

The results of the on-campus examination should be provided in the standard
scoring report provided by the Student Administration. This file must be in XLSX
format.

The name of the column containing the online exam results should have the
following structure and format:

- `exam_date/course_code/exam_name(Remotely Proctored)extra_stuff`
- `YYYY-mm-dd/123456-{M,B}-{1-6}/whatever_name_you_like(Remotely Proctored)does_not_matter`

When these formatting requirements are not met, the utility should still work,
but some metadata may not be recoverable (i.e., date of the online exam, course
code, name of the online exam).

## Scoring Notes

The program will ask if the instructor has provided a custom scoring scheme. If
you respond in the affirmative, you will be asked to supply a lookup table
defining the custom scoring scheme. This lookup table must:

1. Be a CSV file
1. Contain exactly two columns
1. Not have column names

The first column of the lookup table should contain all possible exam scores,
and the second column should contain the corresponding grades. An example lookup
table describing a simple linear scoring scheme for an exam with 50 questions is
available in `data/lookup_table.csv`.

When the instructor does not supply a custom scoring scheme (and, consequently,
the aforementioned question is answered in the negative), the exam will be
scored with Tilburg University's default scoring scheme. This scheme applies a
guessing correction so that all scores equal to or lower than the score expected
from guessing are mapped onto the minimum grade. A reference implementation of
the default scoring scheme (provided by the TiU Examination Committee) is
available in `reference/grading.R`.

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

## Known Issues

The program expects the XLSX file containing the on-campus exam results to be
created on a Windows machine. If this file is created on a Mac, the origin used
to define the date of the on-campus exam will probably differ from the value
expected by the program. Consequently, the date reported for the on-campus exam
(in the final XLSX report) will be incorrect.

The program treats scores of 0 as valid results. Students who did not complete
an exam are expected to have empty cells in the input files (these students are
excluded from the output report). If a student has a score of zero because they
did not take the exam, they will still be assigned the minimum grade in the
final report.

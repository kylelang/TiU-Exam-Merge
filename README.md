# TiU-Exam-Merge

This repository provides a utility to combine results from the online and
on-campus version of hybrid exams at Tilburg University. The utility will output
a report for instructors as an XLSX workbook.

## Usage

There are multiple ways to use this program:

1. On Windows machines, you can execute the `code/runJob.bat` batch file.
1. On Linux (and Mac?) machines, you can execute the `code/runJob.sh` script.
1. On all platforms, you can source the `code/runJob.R` script from within an
interactive R session.

Regardless of how the program is executed, the utility operates by
parameterizing the job through a series of interactive dialog boxes. So, usage
amounts to executing the script via your chosen method and responding to the
dialog boxes when prompted.

## Input Notes

The results of the online examination should be provided as part of a standard
Canvas gradebook download. This file must be in CSV format. The utility should
work with CSV files that use commas `","` or semi-colons `";"` as the field
delimiter.

The results of the on-campus examinations should be provided in the standard
scoring report provided by the Student Administration. These files must be in
XLSX format.

The full set of on-campus results may be comprised of multiple files. In this
case, you will be prompted to locate each of the pertinent files when setting up
the job. Multiple sets of online results are also supported, but the program
assumes that these results will be represented as multiple columns in a single
input file. So, you may only read in one file of online results.

The name of the column containing the online exam results should have a specific
structure and format.

For proctored exam results:

- `exam_date/course_code/exam_name(Remotely Proctored)extra_stuff`
- `YYYY-mm-dd/123456-{M,B}-{1-6}/whatever_name_you_like(Remotely Proctored)does_not_matter`

For unproctored exam results:

- `exam_date/course_code/exam_name OPT-OUTextra_stuff`
- `YYYY-mm-dd/123456-{M,B}-{1-6}/whatever_name_you_like OPT-OUTdoes_not_matter`

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

When running the utility by sourcing the `code/runJob.sh` script interactively
through RStudio, the list selection dialog is semi-broken. RStudio (version <=
1.3.1093, at least) does not support a GUI list-selector dialog, so the script
will default to the text-based list selector (due to the behavior of
**svDialogs**::dlgList() version <= 1.0.0). The text-based list selector works
fine for selecting single elements but not for selecting multiple elements. When
selecting multiple options, the RStudio text-based list selector seems to force
selection of the default option (which cannot be disabled or chosen
intelligently by the program). This issue does not seem to occur when executing
the script via any means other than interactively sourcing `code/runJob.sh` from
within RStudio and may be resolved by future updates to either **svDialogs** or
RStudio.

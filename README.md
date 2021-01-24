# TiU-Exam-Merge

This repository provides a utility to combine results from the online and
on-campus version of hybrid exams at Tilburg University. The utility will output
a report for instructors as an XLSX workbook.

## Usage

There are multiple ways to use this program:

1. On Windows machines, you can execute the `code/runJob.bat` batch file (e.g.,
   by double-clicking the file from within the `code` directory).
1. On Linux (and Mac?) machines, you can execute the `code/runJob.sh` script
   (e.g., by navigating to the `code` directory and executing `./runJob.sh` from
   the shell).
1. On all platforms, you can source the `code/runJob.R` script from within an
   interactive R session.

Regardless of how the program is executed, the utility operates by
parameterizing the job through a series of interactive dialog boxes. So, usage
amounts to executing the script via your chosen method and responding to the
dialog boxes when prompted.

## Input Notes

The results of the online examination should be provided as either a standard
Canvas gradebook download or a TestVision results file. This file must be in CSV
format. The utility should work with CSV files that use commas `","` or
semi-colons `";"` as the field delimiter.

The results of the on-campus examinations should be provided in the standard
scoring report provided by the Student Administration. These files must be in
XLSX format.

The full set of on-campus results may comprise multiple files. In this case, you
will be prompted to locate each of the pertinent files when setting up the
job. Multiple sets of online results are supported if the online exam was
administered via Canvas. In this case, the program assumes that these results
will be represented as multiple columns in a single input file. So, you may only
read in one file of online results.

The program will also run with only online exam results. After indicating the
location of their online exam results, the user will be asked if they have any
on-campus results to process.

For online exams administered through Canvas, the name of the column containing
the online exam results should have a specific structure and format.

For proctored exam results:

- `exam_date/course_code/exam_name(Remotely Proctored)extra_stuff`
- `YYYY-mm-dd/123456-{M,B}-{1-6}/whatever_name_you_like(Remotely Proctored)does_not_matter`

For unproctored exam results:

- `exam_date/course_code/exam_name OPT-OUTextra_stuff`
- `YYYY-mm-dd/123456-{M,B}-{1-6}/whatever_name_you_like OPT-OUTdoes_not_matter`

When these formatting requirements are not met, the utility should still work,
but some metadata may not be recoverable (i.e., date of the online exam, course
code, name of the online exam).

The TestVision results file is automatically generated with known formatting, so
no special formatting considerations are necessary when processing TestVision
results.

## Output Notes

When running the script on a Windows machine, the lengths of the filepaths for
the report and irregularity checks are compared to the `maxNameLength` variable
defined in `runJob.R`. By default, `maxNameLength` is set to 260 characters,
which is the default maximum file name length on the last several versions of
Windows. On Windows 10, it is possible to enable long file name support and
extend the maximum file length to 32,767 characters. 

If any file name is found to be too long (i.e., to contain more than
`maxNameLength` characters), the user is asked if they would like to select a
new output location. If the user declines this offer, the extant file path is
retained.

If your machine has a maximum file name length higher than 260 characters, you
can either adjust the value of `maxNameLength` or decline the offer to select a
new file name when one of your output files is found to be too long.

## Scoring Notes

You will be prompted to select between four different scoring options:

1. The new standard guessing correction formula that will be adopted
   university-wide in 2021
1. The first scoring rule listed on the paper work order form
1. The second scoring rule listed on the paper work order form
1. A custom scoring scheme provided by the instructor as part of the third
   scoring option listed on the paper work order form

When applying any of the first three scoring schemes, you will be prompted to
indicate the number of questions on the exam and the number of response options
(i.e., alternatives) for each question. If you select the first scoring scheme,
you will also be prompted to indicate the passing norm to use in defining the
minimum passing score. 

The passing norm must be a value between 0 and 1 that corresponds to the minimum
guessing-corrected proportional score that a student must achieve to earn a
grade of 5.5. By default, this value is set to 0.55. If you are not sure of what
value to choose, accept the default.

A reference implementation of the new standard guessing correction formula
(provided by the TiU Examination Committee) is available in
`reference/grading.R`, and the formula is documented in
`reference/gradingFormula.pdf`.

If you opt to apply a custom scoring scheme, you will be asked to supply a
lookup table defining the custom scoring scheme. This lookup table must:

1. Be a CSV file
1. Contain exactly two columns
1. Not have column names

The first column of the lookup table should contain all possible exam scores,
and the second column should contain the corresponding grades. An example lookup
table describing a simple linear scoring scheme for an exam with 50 questions is
available in `data/lookup_table.csv`.

Regardless of how the exam is scored, the final report will contain a sheet
showing the scoring table corresponding to whatever scoring scheme was applied
to the exam. Cells in this table that correspond to failing grades are filled
with light red, and cells corresponding to passing grades are filled with light
blue.

The minimum score is determined by either the scoring rule used to score the
exam or the faculty to which the exam belongs. All scoring rules other than the
post-2020 standard guessing correction formula imply a fixed minimum grade. 

1. First scoring rule listed on the paper work order form: Minimum grade = 1
1. Second scoring rule listed on the paper work order form: Minimum grade = 0
1. Custom scoring scheme provided by the instructor: Minimum grade defined by
   the scheme
   
When using the post-2020 standard guessing correction formula, TiSEM exams will
be scored with a minimum grade of 0, and all other faculties will get a minimum
grade of 1. The faculty is determined by parsing the on-campus results file, so
the user will need to specify the faculty when processing only online exam
results. If the user cannot specify the faculty, the minimum score is set to 1.

## Irregularity Checks

When processing both online and on-campus results, the program will compare the
two grade distributions to check for irregularities. The results of these checks
will be saved as a separate XLSX workbook in the same directory the user chooses
for the final report. The name of this file is automatically generated by
appending the string *"-irregularity_checks"* onto the output file name
specified by the user.

If either the online or on-campus results contain fewer than 5 observations, the
irregularity checks are not conducted.

The distributions are compared via three tests:

### Mean Grades

The mean on-campus grade is compared to the mean online grade using an
independent samples t-test *without* assuming equal variances. The following
statistics are reported for this test:

1. The estimated t-statistic
1. The df of the t-test
1. The p-value of the t-test
1. The standardized mean difference (i.e, Cohen's d)

### Proportion of Grades >= 6 (i.e., Passing Students)

The proportions of the passing grades are compared using a chi-squared test for
independence when all cell counts in the *Exam Version* X *Passing* table
are at least 5. When any cell-count is less than 5, the comparison is made via
Fisher's exact test for independence. The following statistics are reported for
this test:

1. The ratio of the odds of passing the online exam to the odds of passing the
   on-campus exam.
1. The estimated chi-squared statistic (unless using Fisher's exact test)
1. The p-value for the test
1. The standardized difference in proportions (i.e., Cohen's h)

### Proportion of Grades >= 8 (i.e., Cum Laude Students)

The proportions of cum laude students are compared and reported in the same way
as the proportions of passing students.

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

When running the utility by sourcing the `code/runJob.R` script interactively
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

RStudio also does not allow multiple file selection, so only one set of
on-campus results can be read in when sourcing the `code/runJob.R` script
through RStudio.

The program will automatically append the ".xlsx" file extension to the output
file when the user does not include the extension. When applying this
functionality, however, duplicate filenames may not be detected, so the user may
not be warned about overwriting existing files. For example, if the user
specifies the output file name as: "myOuput", and a file exists with the name
"myOutput.xlsx", the existing file will be overwritten without a warning.

### Formula for grading MC exams

### NOTE: - This script represents the reference implementation of the scoring
###         rule implemented in the scoreExam() function.
###       - This script was provided by Marcel van Assen on 2020-10-13
###       - This script is not used for any actual calculations.


# Parameters characterizing test
K <- 40   # number of questions
a <- 4    # number of answering categories per question

zero <- K/a

# Parameters determined by the examiner / faculty
min.grade <- 1  # For TisEM: min.grade = 1
                # For all other faculties: min.grade = 0
pass <- 0.55    # Default knowledge = .55, users can select another value

# example INPUT SCORES of students
X <- 5:40
N <- length(X)
grade <- frac <- 0*1:N

# Grading based on example
# Two stages
# First stage: Determine fraction, before rounding stuff
knowledge <- (X-zero)/(K-zero)
for (i in 1:N) {
  knowledge[i] <- max(knowledge[i],0)
  if (knowledge[i] < pass) {frac[i] <- min.grade + knowledge[i]*(5.5-min.grade)/pass
  } else {
    frac[i] <- 5.5 + ((knowledge[i]-pass)/(1-pass)) * (10-5.5)
  }
}

# Stage 2: the rounding
grade <- round(2*frac)/2
for (i in 1:N) {
  if (grade[i] == 5.5 & frac[i] < 5.5) {grade[i] <- 5}
  if (grade[i] == 5.5 & frac[i] > 5.4999999) {grade[i] <- 6}      
}


cbind(X,knowledge,frac,grade)

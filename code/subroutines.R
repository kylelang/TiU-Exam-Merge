### Title:    Subroutines for TiU Exam Merging Utility
### Author:   Kyle M. Lang
### Created:  2020-10-13
### Modified: 2020-10-13

scoreExam <- function(score, nQuestions, nOptions, minGrade, pass = 0.55) {
    ## Compute score expected by guessing:
    guess <- nQuestions / nOptions
    
    ## Compute raw score with guessing correction:
    knowledge <- (score - guess) / (nQuestions - guess)
    knowledge[score < guess] <- 0 # Negative values are not allowed
    
    ## Convert raw score to a [minGrade, 10] scale:
    result <- rep(0, length(score))
    for(i in 1 : length(score)) {
        if(is.na(knowledge[i]))
            result[i] <- NA
        else if(knowledge[i] < pass)
            result[i] <- minGrade + knowledge[i] * (5.5 - minGrade) / pass
        else 
            result[i] <- 5.5 + ((knowledge[i] - pass) / (1 - pass)) * (10 - 5.5)
    }
    round(result, 1)
}

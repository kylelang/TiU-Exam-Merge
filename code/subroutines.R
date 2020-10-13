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

###--------------------------------------------------------------------------###

parseName <- function(x) {
    ## Case 1.1: Surname, First Name
    ## Case 1.2: Surname, Initials
    tmp <- unlist(str_locate_all(x, ","))
    if(length(tmp) > 0) {
        name2 <- substr(x, 1, tmp[1] - 1)
        name1 <- str_trim(substr(x, tmp[1] + 1, nchar(x)))

        return(c(name1 = name1, name2 = name2))
    }

    ## Case 2.1: Surname Initials
    ## Case 2.2: Initials Surname 
    tmp <- unlist(str_locate_all(x, "[A-Z]{1}\\."))
    if(length(tmp) > 0) {
        if(tmp[1] > 1) {
            name2 <- substr(x, 1, tmp[1] - 1)
            name1 <- substr(x, tmp[1], nchar(x))
        }
        else {
            name1 <- substr(x, 1, max(tmp))
            name2 <- str_trim(substr(x, max(tmp) + 1, nchar(x)))
        }
        return(c(name1 = name1, name2 = name2))
    }
    
    ## Case 3: First Name Surname
    tmp <- unlist(str_locate_all(x, "\\s"))
    if(length(tmp) > 0) {
        name1 <- substr(x, 1, tmp[1] - 1)
        name2 <- str_trim(substr(x, tmp[1], nchar(x)))
        
        return(c(name1 = name1, name2 = name2))
    }
    
    ## Return raw name when no case above is satisfied
    c(name1 = x, name2 = x)
}

###--------------------------------------------------------------------------###

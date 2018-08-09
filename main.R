# Author: Arnold Lestin
# Purpose: Extracts data from a Qualtrics survey evaluation CSV export file and outputs summary reports in "./reports"


# # DOWNLOAD SURVEY RESPONSE DATA USING QUALTRICS API #
# #install.packages("reticulate")
# library(reticulate)
# qualtrics.api <- source_python("src/qualtrics_api.py")
# # survey data will be saved to the same directory as this script
# download_survey("API_TOKEN", "SURVEY_ID", "DATACENTER_ID")

# opens a window to select the input file
winDialog(type = c("ok"),
          "Select the report file you exported from Qualtrics after the evaluation.")
evaluations.filename <- file.choose()

winDialog(
  type = c("ok"),
  "Select the contacts file you imported into Qualtrics before the evaluation."
)
student.contacts.filename <- file.choose()

#install.packages("scales")
library("scales")
evals <- read.csv(evaluations.filename)
student.contacts <- read.csv(student.contacts.filename)

# sets the directory to output reports files to
setwd("./reports")

# determines course sizes by counting student contacts per course
all.codes <- c()
for (i in 1:nrow(student.contacts)) {
  s.code <- as.character(student.contacts[i, "SUBJECT.CODE"])
  c.num <- as.character(student.contacts[i, "COURSE.."])
  s.num <- as.character(student.contacts[i, "SECTION.."])
  
  code <- paste(s.code, c.num, s.num, sep = ".")
  all.codes <- c(all.codes, code)
}

unique.codes <- unique(all.codes)
course.sizes <- list(all.codes)

for (code in 1:length(unique.codes)) {
  cur.code <- unique.codes[code]
  course.sizes[[cur.code]] <- length(which(all.codes == cur.code))
}


## CUSTOMIZE IF PROF COLS ARE GREATER THAN 16 ##

# creates vector of character representations of all professor names used in the file
contacts <- c(
  # [-c(1:2)] removes the headers from each column
  # as.character converts from levels (integer) representations to names (character)
  as.character(evals$PROF1[-c(1:2)]),
  as.character(evals$PROF2[-c(1:2)]),
  as.character(evals$PROF3[-c(1:2)]),
  as.character(evals$PROF4[-c(1:2)]),
  as.character(evals$PROF5[-c(1:2)]),
  as.character(evals$PROF6[-c(1:2)]),
  as.character(evals$PROF7[-c(1:2)]),
  as.character(evals$PROF8[-c(1:2)]),
  as.character(evals$PROF9[-c(1:2)]),
  as.character(evals$PROF10[-c(1:2)]),
  as.character(evals$PROF11[-c(1:2)]),
  as.character(evals$PROF12[-c(1:2)]),
  as.character(evals$PROF13[-c(1:2)]),
  as.character(evals$PROF14[-c(1:2)]),
  as.character(evals$PROF15[-c(1:2)]),
  as.character(evals$PROF16[-c(1:2)])
)

# removes empty contacts (not all classes have all 16 professor slots filled)
contacts <- contacts[-(which(contacts == ""))]
# removes duplicates
contacts <- unique(contacts)

# Gets all reviews for each professor in contacts and adds them to reviewl
reviewl <- list()
comment.files <- list()


##  COLUMN NUMBER CONFIGURATION ##
default.num.prof.cols <- 0
default.num.ta.cols <- 0
num.prof.cols <- default.num.prof.cols
num.ta.cols <- default.num.ta.cols

if (num.prof.cols > 0) {
  for (cur.eval in 3:nrow(evals)) {
    # for each P column, e.g, P1
    for (pctr in 1:num.prof.cols) {
      # converts the P number to a character for string pasting
      pctr.char <- as.character(pctr)
      # combines the character and number to have a valid column index, e.g., "P1"
      pcol <-  paste("PROF", pctr.char, sep = "")
      
      # adds one to the pctr because column Q2 corresponds to column P1's review
      qctr.char <- as.character(pctr + 1)
      qcol <- paste("Q", qctr.char, sep = "")
      
      # the question answer/review is the value at qcol
      review <- evals[cur.eval, qcol]
      # level (integer) to character
      review <- as.character(review)
      
      prof.name <- as.character(evals[cur.eval, pcol])
      # saves the current eval's course information into a variable
      course.title <- as.character(evals[cur.eval, "TITLE"])
      subject.code <- as.character(evals[cur.eval, "SUBJECT.CODE"])
      course.number <- as.character(evals[cur.eval, "COURSE.."])
      sequence.number <-
        as.character(evals[cur.eval, "SECTION.."])
      
      course.code <-
        paste(subject.code, course.number, sequence.number, sep = ".")
      
      if (prof.name != "") {
        # ignores empty reviews (some respondents skip questions)
        if (review != "") {
          reviewl[[prof.name]]$courses[[course.title]][[course.code]]$ratings <-
            c(reviewl[[prof.name]]$courses[[course.title]][[course.code]]$ratings, review)
          
        }
      }
    }
    
    # checks if ta reports was enabled by updating the variable num.ta.cols
    if (num.ta.cols > 0) {
      for (ta.ctr in 1:num.ta.cols) {
        ta.ctr.char <- as.character(ta.ctr)
        ta.review.col.char <- as.character(ta.ctr + 17)
        
        ta.col <- paste("TA", ta.ctr.char, sep = "")
        ta.review.col <- paste("Q", ta.review.col.char, sep = "")
        
        ta.name <- as.character(evals[cur.eval, ta.col])
        ta.review <- as.character(evals[cur.eval, ta.review.col])
        
        if (ta.name != "") {
          if (ta.review != "") {
            reviewl[[ta.name]]$courses[[course.title]][[course.code]]$ratings <-
              c(reviewl[[ta.name]]$courses[[course.title]][[course.code]]$ratings,  ta.review)
          }
        }
      }
    }
    
    eval.comment <- as.character(evals[cur.eval, "Q22"])
    
    alpha.course.title <- gsub("[[:punct:]]", ".", course.title)
    comment.file.name <-
      paste(paste(course.code, alpha.course.title, sep = "-"),
            ".txt",
            sep = "")
    
    if (eval.comment != "") {
      comment.block <- paste("Comment: ", eval.comment, "\n\n")
      comment.files[[comment.file.name]] <-
        c(comment.files[[comment.file.name]], comment.block)
    }
  }
} else {
  # error shows if script cannot be run because of invalid num.prof.cols value
  winDialog(
    type = c("ok"),
    "Variable num.prof.cols <= 0. Update script with the number of PROF[X] columns in the contacts file you uploaded to Qualtrics."
  )
  # quit(save = "ask")
}

# outputs comment text files for each course
for (cur.file in 1:length(names(comment.files))) {
  cur.file.name <- names(comment.files[cur.file])
  write(comment.files[[cur.file.name]], file = cur.file.name)
}

# combines reports for professors with similar names
similar <- c()
max.name.diff <- 4

for (i in 1:length(contacts)) {
  for (j in 1:length(contacts)) {
    fname <- contacts[i]
    sname <- contacts[j]
    
    fname <- gsub(" ", "", fname, fixed = TRUE)
    
    sname <- gsub(" ", "", sname, fixed = TRUE)
    
    
    if ((adist(fname, sname) < max.name.diff) &
        (contacts[i] != contacts[j])) {
      if ((length(which(similar == contacts[i])) == 0) &
          (length(which(similar == contacts[j])) == 0)) {
        similar <- rbind(similar, c(contacts[i], ">", contacts[j]))
        
        
      }
    }
  }
}

if (length(similar) != 0) {
  for (i in 1:nrow(similar)) {
    fname <- similar[i, 1]
    
    sname <- similar[i, 3]
    
    reviewl[[fname]]$courses <-
      c(reviewl[[fname]]$courses, reviewl[[sname]]$courses)
    reviewl[[fname]]$ratings <-
      c(reviewl[[fname]]$ratings, reviewl[[sname]]$ratings)
    reviewl[[sname]] <- NULL
  }
  
  # outputs text log so the user can check no names were merged unnecessarily
  write.table(
    similar,
    "merged-names.txt",
    col.names = FALSE,
    row.names = FALSE,
    quote = FALSE
  )
}

# outputs a file for each professor titled [Professor's name].csv in the format of (Course Code, Course Title, Reponse Rate, Num Evals, Course Size, Average, Frequencies)
summary.report <- list()
for (prof in 1:length(reviewl)) {
  num.courses <- length(reviewl[[prof]]$courses)
  prof.name <- names(reviewl[prof])
  
  prof.report <- c()
  summary.ratings.prod <- c()
  summary.num.ratings <- c()
  summary.course.names <- names(reviewl[[prof]]$courses)
  
  for (cur.course in 1:num.courses) {
    num.sections <- length(reviewl[[prof]]$courses[[cur.course]])
    
    for (cur.section in 1:num.sections) {
      reviews <-
        reviewl[[prof]]$courses[[cur.course]][[cur.section]]$ratings
      cur.course.code <-
        names(reviewl[[prof]]$courses[[cur.course]][cur.section])
      
      # counts frequencies of each possible review response
      if (any(reviews == "Excellent")) {
        ecount <- length(which(reviews == "Excellent"))
        
        vgcount <- length(which(reviews == "Very Good"))
        
        gcount <- length(which(reviews == "Good"))
        
        fcount <- length(which(reviews == "Fair"))
        
        pcount <- length(which(reviews == "Poor"))
      } else {
        ecount <- length(which(reviews == "5"))
        
        vgcount <- length(which(reviews == "4"))
        
        gcount <- length(which(reviews == "3"))
        
        fcount <- length(which(reviews == "2"))
        
        pcount <- length(which(reviews == "1"))
      }
      
      # vector of all rating counts
      freqs <- c(ecount, vgcount, gcount, fcount, pcount)
      
      reviewl[[prof]]$courses[[cur.course]][[cur.section]][["freqs"]] <-
        freqs
      
      # sums frequencies of all reviews in the section, i.e., total number of ratings the professor received
      num.ratings <- sum(freqs)
      
      # calculates a weighted average
      ratings.prod <-
        (5 * ecount + 4 * vgcount + 3 * gcount + 2 * fcount + 1 * pcount)
      
      reviewl[[prof]]$courses[[cur.course]][[cur.section]][["average"]] <-
        ratings.prod / num.ratings
      
      summary.ratings.prod <- c(summary.ratings.prod, ratings.prod)
      summary.num.ratings <- c(summary.num.ratings, num.ratings)
      
      cur.course.title <- names(reviewl[[prof]]$courses[cur.course])
      
      cur.course.size <- course.sizes[[cur.course.code]]
      
      # use the scales package to represent the response rate as a percent
      reviewl[[prof]]$courses[[cur.course]][[cur.section]][["response.rate"]] <-
        percent(num.ratings / cur.course.size)
      
      line <-
        c(
          cur.course.code,
          cur.course.title,
          reviewl[[prof]]$courses[[cur.course]][[cur.section]][["response.rate"]],
          num.ratings,
          cur.course.size,
          reviewl[[prof]]$courses[[cur.course]][[cur.section]][["average"]],
          reviewl[[prof]]$courses[[cur.course]][[cur.section]][["freqs"]]
        )
      prof.report <- rbind(prof.report, line)
    }
  }
  
  colnames(prof.report) <-
    c(
      "Course Code",
      "Course Title",
      "Response Rate",
      "Number of Evals",
      "Course Size",
      "Average Eval",
      "Excellent",
      "Very Good",
      "Good",
      "Fair",
      "Poor"
    )
  rownames(prof.report) <- NULL
  
  
  prof.name.formatted <- gsub("[[:punct:]]", ".", prof.name)
  
  prof.report.name <- paste(prof.name.formatted, ".csv", sep = "")
  
  write.table(prof.report,
              prof.report.name,
              sep = ",",
              row.names = FALSE)
  
  summary.ratings.prod <- sum(summary.ratings.prod)
  summary.num.ratings <- sum(summary.num.ratings)
  
  summary.average <- summary.ratings.prod / summary.num.ratings
  summary.line <-
    c(prof.name, summary.average, summary.course.names)
  summary.line.length <- length(summary.line)
  summary.report.ncol <- ncol(summary.report)
  
  # if the summary report is not empty
  if (is.null(summary.report.ncol) == FALSE) {
    if (summary.line.length < summary.ncol)
    {
      # pad the line with NA values to make equal width rows
      summary.line <-
        c(summary.line,
          rep(NA, summary.report.ncol - summary.line.length))
    }
    print(summary.line)
  }
  
  summary.report <- c(summary.report, list(summary.line))
}

# adds NA to each "row" of summary.report so a square table can be created
max.ncol <-  max(sapply(summary.report, length))
summary.report <- do.call(rbind, lapply(summary.report, function(z)
  c(z, rep(
    NA, max.ncol - length(z)
  ))))

col1.name <- "Name"
col2.name <- "Average Evaluation"
col3.name <- paste("C", (1:(max.ncol - 2)), sep = "")

colnames(summary.report) <- c(col1.name, col2.name, col3.name)

# @ symbol used to ensure the report is listed first in the directory
write.table(
  summary.report,
  "@Report.csv",
  sep = ",",
  row.names = FALSE,
  na = ""
)

# reminds user that the num.ta.cols variable was not configured
if (num.ta.cols <= 0) {
  winDialog(
    type = c("ok"),
    "Variable num.ta.cols was <= 0, so no TA reports were generated. Update script with the correct number if this is an error."
  )
}

winDialog(type = c("ok"),
          "Your reports have been generated in the reports folder.")

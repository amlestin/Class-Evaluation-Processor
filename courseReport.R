# Author: Arnold Lestin
# Purpose: Extracts data from a Qualtrics survey evaluation CSV export file and outputs summary reports in "./reports"

# # DOWNLOAD SURVEY RESPONSE DATA USING QUALTRICS API #
# #install.packages("reticulate")
# library(reticulate)
# qualtrics.api <- source_python("src/qualtrics_api.py")
# # survey data will be saved to the same directory as this script
# download_survey("API_TOKEN", "SURVEY_ID", "DATACENTER_ID")

# install required packages if they are not installed
if (!require("data.table", character.only = T, quietly = T)) {
  install.packages("data.table")
}
if (!require("openxlsx", character.only = T, quietly = T)) {
  install.packages("openxlsx")
}
if (!require("scales", character.only = T, quietly = T)) {
  install.packages("scales")
}

# load required packages
library("openxlsx")
library("scales")
library("data.table")

#TODO: Edit how directories work later
# directory_name <- dirname(rstudioapi::getSourceEditorContext()$path)
# setwd(directory_name)
# setwd("/..")
# print(getwd())

setwd("/Users/Scott Upman/Desktop/")

# Creates a reports file on to the desktop
if ("reports" %in% list.files() == FALSE)
  dir.create("reports")
setwd("reports")

#
# load helper functions from src directory
source("/Users/Scott Upman/Desktop/Github Repositories/Class-Evaluation-Processor/src/helpers.R")
#source("/src/helpers.R")
#source("src/helpers.R")
#source("C:/Users/Arnold/Documents/Class-Evaluation-Processor/src/helpers.R")

# opens an explorer window to select the input files
winDialog(type = c("ok"), "Select the report file you exported from Qualtrics after the evaluation.")
evaluations.filename <- file.choose()
winDialog(type = c("ok"), "Select the contacts file you imported into Qualtrics before the evaluation.")
student.contacts.filename <- file.choose()

#  read in the contact and distribution files
evals <- read.csv(evaluations.filename, stringsAsFactors = FALSE)
student.contacts <- read.csv(student.contacts.filename)

##  COLUMN NAME CONFIGURATION ##
#PROF
c.PROF <- "PROF"
#TA
c.TA <- "TA"
#UID
c.UID <- "UID."
#CRN
c.CRN <- "CRN"
#SUBJECT
c.SUB <- "SUBJECT.CODE"
#COURSE
c.CRS <- "COURSE.."
#SECTION
c.SEC <- "SECTION.."
#TITLE
c.TITLE <- "TITLE"
#COMMENTS
c.COM <- "W22"
# SEMESTER
c.SEM <- "TERM.DESCRIPTION"
#Course questions
c.Q1 <-
  "Description.of.course.objectives.and.assignments"
c.Q2 <-
  "Communication.of.ideas.and.information"
c.Q3 <-
  "Expression.of.expectation.for.performance.in.this.class"
c.Q4 <-
  "Availability.to.assist.students.in.or.out.of.class"
c.Q5 <-
  "Respect.and.concern.for.students"
c.Q6 <-
  "Stimulation.of.interest.in.the.course"
c.Q7 <-
  "Facilitation.of.learning"
c.Q8 <-
  "Overall.assessment.of.course"

# Column name validation in evals
valid.cns <- c(c.UID, c.CRN, c.SUB, c.CRS, c.SEC, c.TITLE, c.COM, c.SEM, c.Q1, c.Q2, c.Q3, c.Q4, c.Q5, c.Q6, c.Q7, c.Q8)
err <- all(valid.cns %in% names(evals))
if (!err) {
  cat("ERROR: There is a missing column in the input!\n")
  cat("Missing columns: \n")
  print(valid.cns[-which(valid.cns %in% names(evals))])
  while(1){}
}

##  COLUMN NUMBER CONFIGURATION ##
# calculate  the number of columns with PROF in their name
# calculate  the number of columns with TA in their name
num.prof.cols <- length(grep(c.PROF, names(evals)))
num.ta.cols <- length(grep(c.PROF, names(evals)))


# check for duplicate professors in the student contacts
# only checks this method if contacts file has CRNs
error.log <- c()
if (c.CRN %in% names(student.contacts) == TRUE) {
  library(data.table)
  DT <- data.table(student.contacts, stringsAsFactors = FALSE)
  #UID. may change semester
  DT <- unique(unique(DT, by = c.CRN), by = c.UID)
  for (course.ctr in 1:nrow(DT)) {
    profs.indices <- which(grepl(c.PROF, names(DT)))
    row <- DT[course.ctr, ]
    
    profs <- c()
    dups <- c()
    for (i in profs.indices) {
      p = as.character(row[[i]])
      
      if (p %in% profs) {
        dups <- c(dups, p)
      }
      profs <- c(profs, p)
    }
    
    if (length(dups) > 0) {
      dups <- dups[dups != ""]
      dups <- dups[dups != "NA"]
      dups <- dups[!is.na(dups)]
      dups.flag <- length(dups) > 0
      
      if (dups.flag == TRUE) {
        error.message <-
          paste(dups,
                as.character(DT[course.ctr, ..c.CRN]))
        error.log <- c(error.log, error.message)
      }
    }
  }
}

# determines course sizes by counting student contacts per course
all.codes <- c()
for (i in 1:nrow(student.contacts)) 
{
  # Subject Code
  subject_code <- as.character(student.contacts[i, c.SUB])
  
  # Course number code
  course_num <- as.character(student.contacts[i, c.CRS])
  
  # Section number code
  section_num <- as.character(student.contacts[i, c.SEC])
  
  code <- paste(subject_code, course_num, section_num, sep = ".")
  all.codes <- c(all.codes, code)
}

unique.codes <- unique(all.codes)
course.sizes <- list(all.codes)

for (code in 1:length(unique.codes)) {
  
  # Contains course code for each element in unique.codes vector
  cur.code <- unique.codes[code]
  
  # Inserts lengths of the courses into the course.sizes list
  course.sizes[[cur.code]] <- length(which(all.codes == cur.code)) ## number of students that are in the course
}
# Converts the all.codes vector to a list
course.codes.to.crn <- list(all.codes)


# creates vector of character representations of all professor names used in the file
prof.cols <- grep(c.PROF, colnames(evals))
contacts <-
  mapply(function (row.index)
    mapply(function(col.index)
      as.character(evals[row.index, col.index]), prof.cols),
    seq(3, nrow(evals)))
# removes empty contacts (not all classes have all 16 professor slots filled)
contacts <- contacts[-(which(contacts == ""))]
# removes duplicates
contacts <- unique(contacts)

# Gets all reviews for each professor in contacts and adds them to reviewl
reviewl <- list()
reviews.by.course.code <- list()
if (num.prof.cols > 0) {
  for (cur.eval in 3:nrow(evals)) {
    # for each P column, e.g, P1
    previously.seen.profs <- c()
    for (pctr in 1:num.prof.cols) {
      # converts the P number to a character for string pasting
      pctr.char <- as.character(pctr)
      # combines the character and number to have a valid column index, e.g., "P1"
      pcol <-  paste(c.PROF, pctr.char, sep = "")
      
      # adds one to the pctr because column Q2 corresponds to column P1's review
      qctr.char <- as.character(pctr + 1)
      qcol <- paste("Q", qctr.char, sep = "")
      
      # the question answer/review is the value at qcol
      review <- evals[cur.eval, qcol]
      # level (integer) to character
      review <- as.character(review)
      
      prof.name <- as.character(evals[cur.eval, pcol])
      
      
      if (!is.name.valid(prof.name))
        next()
      
      # check for duplicate professors in a course
      if (prof.name %in% previously.seen.profs) {
        error.log <-
          c(error.log,
            paste(prof.name, as.character(evals[cur.eval, c.CRN])))
        next()
      }
      previously.seen.profs <- c(previously.seen.profs, prof.name)
      
      # saves the current eval's course information into a variable
      course.title <- as.character(evals[cur.eval, c.TITLE])
      subject.code <- as.character(evals[cur.eval, c.SUB])
      course.number <- as.character(evals[cur.eval, c.CRS])
      sequence.number <-
        as.character(evals[cur.eval, c.SEC])
      
      course.code <-
        paste(subject.code, course.number, sequence.number, sep = ".")
      
      if (length(prof.name) > 0) {
        # ignores empty reviews (some respondents skip questions)
        if (length(review) > 0) {
          reviews.by.course.code[[course.title]][[course.code]][[prof.name]]$ratings <-
            c(reviews.by.course.code[[course.title]][[course.code]][[prof.name]]$ratings, review)
          reviewl[[prof.name]]$courses[[course.title]][[course.code]]$ratings <-
            c(reviewl[[prof.name]]$courses[[course.title]][[course.code]]$ratings, review)
          
        }
      }
    }
    
    # checks if ta reports was enabled by updating the variable num.ta.cols
    if (num.ta.cols > 0) {
      for (ta.ctr in 1:num.ta.cols) {
        ta.ctr.char <- as.character(ta.ctr)
        ###############################
        #                             #
        #          MAGIC NUMBER       #
        #                             #
        ###############################
        ta.review.col.char <- as.character(ta.ctr + 17)
        
        ta.col <- paste(c.TA, ta.ctr.char, sep = "")
        ta.review.col <- paste("Q", ta.review.col.char, sep = "")
        
        ta.name <- as.character(evals[cur.eval, ta.col])
        ta.review <- as.character(evals[cur.eval, ta.review.col])
        
        if (!is.name.valid(ta.name))
          next()
        
        if (length(ta.name) > 0) {
          if (length(ta.review) > 0) {
            reviewl[[ta.name]]$courses[[course.title]][[course.code]]$ratings <-
              c(reviewl[[ta.name]]$courses[[course.title]][[course.code]]$ratings,  ta.review)
          }
        }
      }
    }
  }
} else {
  # error shows if script cannot be run because of invalid num.prof.cols value
  # winDialog(
  #   type = c("ok"),
  #   "Variable num.prof.cols <= 0. Update script with the number of PROF[X] columns in the contacts file you uploaded to Qualtrics."
  # )
  # quit(save = "ask")
}

# outputs a file for each professor titled [Professor's name].csv in the format of (Course Code, Course Title, Reponse Rate, Num Evals, Course Size, Average, Frequencies)
summary.report <- list()
for (prof in 1:length(reviewl)) {
  num.courses <- length(reviewl[[prof]]$courses)
  prof.name <- names(reviewl[prof])
  
  # prof.report <- c()
  prof.report <- data.frame()
  line <- data.frame()
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
      
      # vector of all rating counts
      freqs <- evals.to.freqs(reviews)
      
      reviewl[[prof]]$courses[[cur.course]][[cur.section]][["freqs"]] <-
        freqs
      
      # sums frequencies of all reviews in the section, i.e., total number of ratings the professor received
      num.ratings <- sum(freqs)
      
      # calculates a weighted average
      ratings.prod <-
        (5 * freqs[1] + 4 * freqs[2] + 3 * freqs[3] + 2 * freqs[4] + 1 * freqs[5])
      
      reviewl[[prof]]$courses[[cur.course]][[cur.section]][["average"]] <-
        ratings.prod / num.ratings
      
      summary.ratings.prod <- c(summary.ratings.prod, ratings.prod)
      summary.num.ratings <- c(summary.num.ratings, num.ratings)
      
      cur.course.title <- names(reviewl[[prof]]$courses[cur.course])
      
      cur.course.size <- course.sizes[[cur.course.code]]
      
      # use the scales package to represent the response rate as a percent
      #      if (cur.course.size < 1 || is.null(cur.course.size)) {
      if (TRUE) {
        cur.course.size <- 1234567
        s <- paste(
          "Invalid course size! Setting to 1234567.",
          cur.course.code,
          cur.course.title,
          " - Size:",
          cur.course.size
        )
        #        print(s)
        #        winDialog(type = c("ok"), s)
      }
      
      reviewl[[prof]]$courses[[cur.course]][[cur.section]][["response.rate"]] <-
        percent(num.ratings / cur.course.size)
      
      line <-
        cbind(
          cur.course.code,
          cur.course.title,
          reviewl[[prof]]$courses[[cur.course]][[cur.section]][["response.rate"]],
          num.ratings,
          cur.course.size,
          reviewl[[prof]]$courses[[cur.course]][[cur.section]][["average"]],
          matrix(as.character(reviewl[[prof]]$courses[[cur.course]][[cur.section]][["freqs"]]), nrow = 1)
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
      "Poor",
      "Fair",
      "Good",
      "Very Good",
      "Excellent"
    )
  rownames(prof.report) <- NULL
  
  
  prof.name.formatted <- gsub("[[:punct:]]", ".", prof.name)
  prof.report.name <- paste(prof.name.formatted, ".csv", sep = "")
  
  summary.ratings.prod <- sum(summary.ratings.prod)
  summary.num.ratings <- sum(summary.num.ratings)
  summary.average <- summary.ratings.prod / summary.num.ratings
  
  total.students.taught <- sum(as.numeric(prof.report[, 5]))
  num.courses.taught <- length(summary.course.names)
  
  summary.line <-
    c(
      prof.name,
      round(summary.average, digits = 2),
      summary.num.ratings,
      total.students.taught,
      num.courses.taught,
      summary.course.names
    )
  
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
col4.name <- "Total Number of Evals"
col5.name <- "Total Number of Students"
col6.name <- "Total Number of Courses Taught"
col3.name <- paste("C", (1:(max.ncol - 5)), sep = "")

colnames(summary.report) <-
  c(col1.name,
    col2.name,
    col4.name,
    col5.name,
    col6.name,
    col3.name)

if (length(error.log) > 0) {
  # outputs error log for duplicated professors by course
  error.log <- unique(error.log)
  write.table(
    error.log,
    "error-log.txt",
    col.names = FALSE,
    row.names = FALSE,
    quote = FALSE
  )
  # winDialog(type = c("ok"),
  #           "Please check error-log.txt to repair the input data!")
  
}

# Calls the semester.summary function and passes the review list as an argument 
semester.summary <- create.semester.summary(reviewl)
sci <-
  which(lengths(lapply(1:length(semester.summary), function(x)
    unique(find.sections.by.course(x, semester.summary)))) > 1)

sc <- semester.summary[sci]

# Calls the export.semester.summary passing sc as an argument
export.semester.summary(sc)

# Semester.summary is a list
semester.summary <- semester.summary[-sci]

# Calls the function to output the final semester.summary
export.semester.summary(semester.summary)

if (num.ta.cols <= 0) {
  # winDialog(
  #   type = c("ok"),
  #   "Variable num.ta.cols was <= 0, so no TA reports were generated. Update script with the correct number if this is an error."
  # )
}

# winDialog(type = c("ok"),
#           "Your reports have been generated in the reports folder.")

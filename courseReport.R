# Author: Arnold Lestin
# Purpose: Extracts data from a Qualtrics survey evaluation CSV export file and outputs summary reports in "./reports"

# # DOWNLOAD SURVEY RESPONSE DATA USING QUALTRICS API #
# #install.packages("reticulate")
# library(reticulate)
# qualtrics.api <- source_python("src/qualtrics_api.py")
# # survey data will be saved to the same directory as this script
# download_survey("API_TOKEN", "SURVEY_ID", "DATACENTER_ID")

if (!require("openxlsx", character.only = T, quietly = T)) {
  install.packages("openxlsx")
}
library(openxlsx)

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
evals <- read.csv(evaluations.filename, stringsAsFactors = FALSE)
student.contacts <- read.csv(student.contacts.filename)

##  COLUMN NUMBER CONFIGURATION ##
num.prof.cols <- length(grep("PROF", names(evals)))
num.ta.cols <- length(grep("TA", names(evals)))
num.ta.cols <- 0


# sets the directory to output reports files to
setwd("~/Class-Evaluation-Processor/reports")

# check for duplicate professors in the student contacts
# only checks this method if contacts file has CRNs
error.log <- c()
if ("CRN" %in% names(student.contacts) == TRUE) {
  library(data.table)
  DT <- data.table(student.contacts, stringsAsFactors = FALSE)
  DT <- unique(unique(DT, by = "CRN"), by = "UID.")
  for (course.ctr in 1:nrow(DT)) {
    profs.indices <- which(grepl("PROF", names(DT)))
    row <- DT[course.ctr,]
    
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
                as.character(DT[course.ctr, "CRN"]))
        error.log <- c(error.log, error.message)
      }
    }
  }
}

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

course.codes.to.crn <- list(all.codes)

create.semester.summary <- function(reviewl) {
  unique.titles <- as.character(unique(student.contacts[, "TITLE"]))
  unique.titles <- sort(unique.titles)
  
  # 0) sort everything so section 1 come before 2
  # 1) List of titles of all courses
  # 2) For each title, find relevant unique course codes
  # 3) For each course  code, print prof stats
  
  prof.cols <- grep("PROF", names(evals))
  profs.by.course <-
    mapply(
      function(course)
        course <- course[which(course != "")],
      mapply(
        
        function(course)
          as.character(unlist(course)),
        mapply(function(course.title)
          #          evals[which(evals$TITLE == course.title), ][1, ][, prof.cols]
          unique(evals[which(evals$TITLE == course.title),][, prof.cols])
          , unique.titles, SIMPLIFY = F),
        SIMPLIFY = F
      ),
      SIMPLIFY = F
    )
  
  semester.summary <- c()
  for (i in seq(length(profs.by.course))) {
    course.name <- names(profs.by.course)[[i]]
    
    course.summary <- list()
    for (j in seq(length(profs.by.course[[i]]))) {
      prof.name <- profs.by.course[[i]][[j]]
      prof.evals <-
        reviewl[[which(names(reviewl) == prof.name)]][["courses"]][[course.name]]
      prof.evals <-
        prof.evals[sort(names(prof.evals), index.return = T)$ix]
      course.summary[[prof.name]] <- prof.evals
    }
    
    semester.summary[[course.name]] <- course.summary
  }
  return(semester.summary)
}

# creates vector of character representations of all professor names used in the file
prof.cols <- grep("PROF", colnames(evals))
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
comment.files <- list()

if (num.prof.cols > 0) {
  for (cur.eval in 3:nrow(evals)) {
    # for each P column, e.g, P1
    previously.seen.profs <- c()
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
      if (length(prof.name) == 0 ||
          is.na(prof.name) || prof.name == "")
        next()
      
      # check for duplicate professors in a course
      if (prof.name %in% previously.seen.profs) {
        error.log <-
          c(error.log,
            paste(prof.name, as.character(evals[cur.eval, "CRN"])))
        next()
      }
      previously.seen.profs <- c(previously.seen.profs, prof.name)
      
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
        ###############################
        #                             #
        #          MAGIC NUMBER       #
        #                             #
        ###############################
        ta.review.col.char <- as.character(ta.ctr + 17)
        
        ta.col <- paste("TA", ta.ctr.char, sep = "")
        ta.review.col <- paste("Q", ta.review.col.char, sep = "")
        
        ta.name <- as.character(evals[cur.eval, ta.col])
        ta.review <- as.character(evals[cur.eval, ta.review.col])
        
        if (length(ta.name) < 1)
          next(0)
        
        if (ta.name != "") {
          if (ta.review != "") {
            reviewl[[ta.name]]$courses[[course.title]][[course.code]]$ratings <-
              c(reviewl[[ta.name]]$courses[[course.title]][[course.code]]$ratings,  ta.review)
          }
        }
      }
    }
    
    eval.comment <- as.character(evals[cur.eval, "W22"])
    
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
      if (any(reviews == "Excellent") |
          any(reviews == "Very Good") |
          any(reviews == "Good") |
          any(reviews == "Fair") | any(reviews == "Poor")) {
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
      #      summary.course.sizes <- c(summary.course.sizes, cur.course.size)
      #      reviewl[[prof]]$total.students.taught <-  reviewl[[prof]]$total.students.taught + cur.course.size
      
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
  
  # write.table(prof.report,
  #             prof.report.name,
  #             sep = ",",
  #             row.names = FALSE)
  
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
  
  # summary.line <-
  #   c(prof.name, summary.average, summary.course.names)
  
  
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

# @ symbol used to ensure the report is listed first in the directory
# write.table(
#   summary.report,
#   paste("@Report", evals[3, "TERM.DESCRIPTION"], ".csv"),
#   sep = ",",
#   row.names = FALSE,
#   na = ""
# )

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
  winDialog(type = c("ok"),
            "Please check error-log.txt to repair the input data!")
  
}


# reminds user that the num.ta.cols variable was not configured
if (num.ta.cols <= 0) {
  winDialog(
    type = c("ok"),
    "Variable num.ta.cols was <= 0, so no TA reports were generated. Update script with the correct number if this is an error."
  )
}

winDialog(type = c("ok"),
          "Your reports have been generated in the reports folder.")

semester.summary <- create.semester.summary(reviewl)
export.semester.summary <- function(semester.summary) {
  report.col.names  <-
    c("Professor", "Average", "Responses")
  semester.report <- c()
  for (course.index in seq(length(semester.summary))) {
    course.summary <- semester.summary[[course.index]]
    course.name <- names(semester.summary)[[course.index]]
    
    sheet.number <- 1
    wb <- createWorkbook("Admin")
    addWorksheet(wb, sheet.number) # add modified report to a worksheet
    
    course.report <- c()
    for (prof.index in seq(length(course.summary))) {
      freqs <- c(0, 0, 0, 0, 0)
      for (sec.index in seq(length(semester.summary[[course.index]][[prof.index]]))) {
        freqs <-
          freqs + semester.summary[[course.index]][[prof.index]][[sec.index]][["freqs"]]
      }
      
      
      # sums frequencies of all reviews in the section, i.e., total number of ratings the professor received
      num.ratings <- sum(freqs)
      
      # calculates a weighted average
      ratings.prod <-
        (5 * freqs[1] + 4 * freqs[2] + 3 * freqs[3] + 2 * freqs[4] + 1 * freqs[5])
      
      average <- ratings.prod / num.ratings
      average <- round(average, digits = 2)
      row <-
        c(
          names(semester.summary[[course.index]])[[prof.index]],
          average,
          paste(num.ratings, "/", course.sizes[[course.code]], sep =
                  "")
#          Reduce(paste, names(semester.summary[[course.index]][[prof.index]]), course.name)
        )
      
      course.report <- rbind(course.report, row)
    }
    
    course.code.array <- unlist(strsplit(course.code, "\\."))
    subject.code <- course.code.array[1]
    course <- course.code.array[2]
    section <- course.code.array[3]
    course.respondents <-
      length(
        which(
          evals$SUBJECT.CODE == subject.code &
            evals$COURSE.. == course & evals$SECTION.. == section
        )
      )
    
    q1 <-
      evals[which(
        evals$SUBJECT.CODE == subject.code &
          evals$COURSE.. == course &
          evals$SECTION.. == section
      ), "Description.of.course.objectives.and.assignments"]
    q2 <-
      evals[which(
        evals$SUBJECT.CODE == subject.code &
          evals$COURSE.. == course &
          evals$SECTION.. == section
      ), "Communication.of.ideas.and.information"]
    q3 <-
      evals[which(
        evals$SUBJECT.CODE == subject.code &
          evals$COURSE.. == course &
          evals$SECTION.. == section
      ), "Expression.of.expectation.for.performance.in.this.class"]
    q4 <-
      evals[which(
        evals$SUBJECT.CODE == subject.code &
          evals$COURSE.. == course &
          evals$SECTION.. == section
      ), "Availability.to.assist.students.in.or.out.of.class"]
    q5 <-
      evals[which(
        evals$SUBJECT.CODE == subject.code &
          evals$COURSE.. == course &
          evals$SECTION.. == section
      ), "Respect.and.concern.for.students"]
    q6 <-
      evals[which(
        evals$SUBJECT.CODE == subject.code &
          evals$COURSE.. == course &
          evals$SECTION.. == section
      ), "Stimulation.of.interest.in.the.course"]
    q7 <-
      evals[which(
        evals$SUBJECT.CODE == subject.code &
          evals$COURSE.. == course &
          evals$SECTION.. == section
      ), "Facilitation.of.learning"]
    q8 <-
      evals[which(
        evals$SUBJECT.CODE == subject.code &
          evals$COURSE.. == course &
          evals$SECTION.. == section
      ), "Overall.assessment.of.course"]
    
    
    evals.to.average <- function(reviews) {
      # counts frequencies of each possible review response
      if (any(reviews == "Excellent") |
          any(reviews == "Very Good") |
          any(reviews == "Good") |
          any(reviews == "Fair") | any(reviews == "Poor")) {
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
      
      # sums frequencies of all reviews in the section, i.e., total number of ratings the professor received
      num.ratings <- sum(freqs)
      
      # calculates a weighted average
      ratings.prod <-
        (5 * ecount + 4 * vgcount + 3 * gcount + 2 * fcount + 1 * pcount)
      
      average <- ratings.prod / num.ratings
      
      return (average)
    }
    
    q1.average <- round(evals.to.average(q1), digits = 2)
    q2.average <- round(evals.to.average(q2), digits = 2)
    q3.average <- round(evals.to.average(q3), digits = 2)
    q4.average <- round(evals.to.average(q4), digits = 2)
    q5.average <- round(evals.to.average(q5), digits = 2)
    q6.average <- round(evals.to.average(q6), digits = 2)
    q7.average <- round(evals.to.average(q7), digits = 2)
    q8.average <- round(evals.to.average(q8), digits = 2)
    
    q.cols <-
      rbind(
        q1.average,
        q2.average,
        q3.average,
        q4.average,
        q5.average,
        q6.average,
        q7.average,
        q8.average
      )
    
    table <- rbind(
      "Description of course objectives and assignments" ,
      "Communication of ideas and information" ,
      "Expression of expectation for performance in this class" ,
      "Availability to assist students in or out of class" ,
      "Respect and concern for students" ,
      "Stimulation of interest in the course" ,
      "Facilitation of learning" ,
      "Overall assessment of course"
    )
    
    table.col.names <- c("Field", "Average", "", "")
    table <- cbind(table, q.cols, "", "")
    table <- rbind(table.col.names, table)
    colnames(table) <- table.col.names
    
    course.report <-
      rbind(
        c(course.name, "", "", ""),
        c("Summer 2019 Evals", "", "", ""),
        "",
        table,
        "",
        report.col.names,
        course.report
      )
    
    firstStyle <-
      createStyle(fontSize = 18,
                  fontColour = "#333333",
                  halign = "center")
    secondStyle <-
      createStyle(fontSize = 14,
                  fontColour = "#333333",
                  halign = "center")
    tableStyle <-
      createStyle(
        fontSize = 12,
        fontColour = "#000000",
        textDecoration = c("bold", "underline"),
        halign = "left"
      )
    
    addStyle(
      wb,
      sheet.number,
      firstStyle,
      rows = 1,
      cols = seq(1, nrow(course.report)),
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      secondStyle,
      rows = 2,
      cols = seq(1, nrow(course.report)),
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      tableStyle,
      rows = 4,
      cols = seq(1, nrow(course.report)),
      stack = TRUE
    )
    
    file.name <- make.names(course.name)
    file.name <- gsub("\\.", " ", file.name)
    file.name <-
      gsub("(?<=[\\s])\\s*|^\\s+|\\s+$", "", file.name, perl = TRUE)
    
#    course.report <- data.frame(course.report)
    writeData(wb,
              sheet = sheet.number,
              course.report,
              colNames = FALSE) # add the new worksheet to the workbook
    # resizes column widths to fit contents
    setColWidths(wb, sheet.number, cols = 1:4, widths = "auto")
    # makes sure sheet fits on one printable page
    pageSetup(wb,
              sheet.number,
              fitToWidth = TRUE,
              fitToHeight = FALSE)
    
    saveWorkbook(wb, paste(file.name, ".xlsx", sep = ""), overwrite = TRUE) # writes a workbook containing all reports inputted
    
    semester.report <- rbind(semester.report, "", course.report)
  }
  
  colnames(semester.report)  <- report.col.names
  
  write.csv(semester.report,
            "semester-report.csv",
            row.names = FALSE,
            na = "")
}

export.semester.summary(semester.summary)

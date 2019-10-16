# Author: Arnold Lestin
# Purpose: Extracts data from a Qualtrics survey evaluation CSV export file and outputs summary reports in "./reports"

# # DOWNLOAD SURVEY RESPONSE DATA USING QUALTRICS API #
# #install.packages("reticulate")
# library(reticulate)
# qualtrics.api <- source_python("src/qualtrics_api.py")
# # survey data will be saved to the same directory as this script
# download_survey("API_TOKEN", "SURVEY_ID", "DATACENTER_ID")
get_worksheet_entries <- function(wb, sheet) {
  #  https://github.com/awalker89/openxlsx/pull/382/commits/7e9cfd09934ac61491b71e91785c415a22b3dc31
  # get worksheet data
  dat <- wb$worksheets[[sheet]]$sheet_data
  # get vector of entries
  val <- dat$v
  # get boolean vector of text entries
  typ <- (dat$t == 1) & !is.na(dat$t)
  # get text entry strings
  str <- unlist(wb$sharedStrings[as.integer(val)[typ] + 1])
  # remove xml tags
  str <- gsub("<.*?>", "", str)
  # write strings to vector of entries
  val[typ] <- str
  # return vector of entries
  val
}

auto_heights <-
  function(wb,
           sheet,
           selected,
           fontsize = NULL,
           factor = 1.0,
           base_height = 15,
           extra_height = 12) {
    #    https://github.com/awalker89/openxlsx/pull/382/commits/7e9cfd09934ac61491b71e91785c415a22b3dc31
    # get base font size
    if (is.null(fontsize)) {
      fontsize <- as.integer(openxlsx::getBaseFont(wb)$size$val)
    }
    # set factor to adjust font width (empiricially found scale factor 4 here)
    factor <- 4 * factor / fontsize
    # get worksheet data
    dat <- wb$worksheets[[sheet]]$sheet_data
    # get columns widths
    colWidths <- wb$colWidths[[sheet]]
    # select fixed (non-auto) and visible (non-hidden) columns only
    specified <-
      (colWidths != "auto") & (attr(colWidths, "hidden") == "0")
    # return default row heights if no column widths are fixed
    if (length(specified) == 0) {
      message("No column widths specified, returning default row heights.")
      cols <- integer(0)
      heights <- rep(base_height, length(selected))
      return(list(cols, heights))
    }
    # get fixed column indices
    cols <- as.integer(names(specified)[specified])
    # get fixed column widths
    widths <- as.numeric(colWidths[specified])
    # get all worksheet entries
    val <- get_worksheet_entries(wb, sheet)
    # compute optimal height per selected row
    heights <- sapply(selected, function(row) {
      # select entries in given row and columns of fixed widths
      index <- (dat$rows == row) & (dat$cols %in% cols)
      # remove line break characters
      chr <- gsub("\\r|\\n", "", val[index])
      # measure width of entry (in pixels)
      wdt <-
        strwidth(chr, unit = "in") * 20 / 1.43 # 20 px = 1.43 in
      # compute optimal height
      if (length(wdt) == 0) {
        base_height
      } else {
        base_height + extra_height * as.integer(max(wdt / widths * factor))
      }
    })
    # return list of indices of columns with fixed widths and optimal row heights
    list(cols, heights)
  }



if (!require("openxlsx", character.only = T, quietly = T)) {
  install.packages("openxlsx")
}
if (!require("scales", character.only = T, quietly = T)) {
  install.packages("scales")
}
library(openxlsx)
library("scales")

# opens a window to select the input file
winDialog(type = c("ok"),
          "Select the report file you exported from Qualtrics after the evaluation.")
evaluations.filename <- file.choose()
winDialog(
  type = c("ok"),
  "Select the contacts file you imported into Qualtrics before the evaluation."
)
student.contacts.filename <- file.choose()

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
# for (cur.file in 1:length(names(comment.files))) {
#   cur.file.name <- names(comment.files[cur.file])
#   write(comment.files[[cur.file.name]], file = cur.file.name)
# }

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

semester.summary <- create.semester.summary(reviewl)
export.semester.summary <- function(semester.summary) {
  report.col.names  <-
    c("Professor", "Average", "Responses")
  for (course.index in seq(length(semester.summary))) {
    course.summary <- semester.summary[[course.index]]
    course.name <- names(semester.summary)[[course.index]]
    course.codes <- c()
    
    sheet.number <- 1
    wb <- createWorkbook("Admin")
    addWorksheet(wb, sheet.number) # add modified report to a worksheet
    
    course.report <- c()
    course.respondents <- c()
    for (prof.index in seq(length(course.summary))) {
      freqs <- c(0, 0, 0, 0, 0)
      for (sec.index in seq(length(semester.summary[[course.index]][[prof.index]]))) {
        freqs <-
          freqs + semester.summary[[course.index]][[prof.index]][[sec.index]][["freqs"]]
      }
      
      prof.course.codes <-
        names(semester.summary[[course.index]][[prof.index]])
      course.codes <- c(course.codes, prof.course.codes)
      # sums frequencies of all reviews in the section, i.e., total number of ratings the professor received
      num.ratings <- sum(freqs)
      
      # calculates a weighted average
      ratings.prod <-
        (5 * freqs[1] + 4 * freqs[2] + 3 * freqs[3] + 2 * freqs[4] + 1 * freqs[5])
      
      average <- ratings.prod / num.ratings
      average <- round(average, digits = 2)
      average <- sprintf("%0.2f", average)
      
      course.size.from.sections <- function(course.codes) {
        section.sizes <- c()
        for (code in course.codes) {
          section.size <-  course.sizes[[code]]
          section.sizes <- c(section.sizes, section.size)
        }
        course.size <- sum(section.sizes)
      }
      
      course.size <- course.size.from.sections(prof.course.codes)
      
      row <-
        c(
          names(semester.summary[[course.index]])[[prof.index]],
          average,
          paste(num.ratings, "/",
                course.size, sep =
                  "")
        )
      
      course.report <- rbind(course.report, row)
      course.respondents <- c(course.respondents, num.ratings)
    }
    
    course.codes <- sort(unique(course.codes))
    
    q1 <- c()
    q2 <- c()
    q3 <- c()
    q4 <- c()
    q5 <- c()
    q6 <- c()
    q7 <- c()
    q8 <- c()
    
    course.comments <- c()
    for (i in seq(length(course.codes))) {
      course.code <- course.codes[i]
      
      course.code.array <- unlist(strsplit(course.code, "\\."))
      subject.code <- course.code.array[1]
      course <- course.code.array[2]
      section <- course.code.array[3]
      
      q1 <- c(q1,
              evals[which(
                evals$SUBJECT.CODE == subject.code &
                  evals$COURSE.. == course &
                  evals$SECTION.. == section
              ), "Description.of.course.objectives.and.assignments"])
      q2 <- c(q2,
              evals[which(
                evals$SUBJECT.CODE == subject.code &
                  evals$COURSE.. == course &
                  evals$SECTION.. == section
              ), "Communication.of.ideas.and.information"])
      q3 <- c(q3,
              evals[which(
                evals$SUBJECT.CODE == subject.code &
                  evals$COURSE.. == course &
                  evals$SECTION.. == section
              ), "Expression.of.expectation.for.performance.in.this.class"])
      q4 <- c(q4,
              evals[which(
                evals$SUBJECT.CODE == subject.code &
                  evals$COURSE.. == course &
                  evals$SECTION.. == section
              ), "Availability.to.assist.students.in.or.out.of.class"])
      q5 <- c(q5,
              evals[which(
                evals$SUBJECT.CODE == subject.code &
                  evals$COURSE.. == course &
                  evals$SECTION.. == section
              ), "Respect.and.concern.for.students"])
      q6 <- c(q6,
              evals[which(
                evals$SUBJECT.CODE == subject.code &
                  evals$COURSE.. == course &
                  evals$SECTION.. == section
              ), "Stimulation.of.interest.in.the.course"])
      q7 <- c(q7,
              evals[which(
                evals$SUBJECT.CODE == subject.code &
                  evals$COURSE.. == course &
                  evals$SECTION.. == section
              ), "Facilitation.of.learning"])
      q8 <- c(q8,
              evals[which(
                evals$SUBJECT.CODE == subject.code &
                  evals$COURSE.. == course &
                  evals$SECTION.. == section
              ), "Overall.assessment.of.course"])
      
      
      current.comments <- evals[which(
        evals$SUBJECT.CODE == subject.code &
          evals$COURSE.. == course
        & evals$SECTION.. == section
      ), "W22"]
      
      if (all(nchar(current.comments) == 0))
        next()
      
      
      current.comments <-
        current.comments[which(nchar(current.comments) != 0)]
      current.comments <-
        paste(">>> ", current.comments, "\n", sep = "")
      
      Encoding(current.comments) <- "UTF-8"
      
      course.comments <-
        c(course.comments, current.comments)
    }
    
    #    course.comments <- c("", course.comments)
    course.respondents <-
      course.respondents[course.respondents != ""]
    
    q1 <- q1[q1 != ""]
    q2 <- q2[q2 != ""]
    q3 <- q3[q3 != ""]
    q4 <- q4[q4 != ""]
    q5 <- q5[q5 != ""]
    q6 <- q6[q6 != ""]
    q7 <- q7[q7 != ""]
    q8 <- q8[q8 != ""]
    
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
      average <- round(average, digits = 2)
      average <- sprintf("%0.2f", average)
      
      return (average)
    }
    
    q1.average <- evals.to.average(q1)
    q2.average <- evals.to.average(q2)
    q3.average <- evals.to.average(q3)
    q4.average <- evals.to.average(q4)
    q5.average <- evals.to.average(q5)
    q6.average <- evals.to.average(q6)
    q7.average <- evals.to.average(q7)
    q8.average <- evals.to.average(q8)
    
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
    
    course.respondents <- max(
      length(q1),
      length(q2),
      length(q3),
      length(q4),
      length(q5),
      length(q6),
      length(q7),
      length(q8),
      course.respondents
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
    
    #    table.col.names <- c("Field", "Average", "")
    course.size <- course.size.from.sections(course.codes)
    response.rate <- percent(course.respondents / course.size, 2)
    response.rate.heading <-
      paste(course.respondents, "/", course.size, sep = "")
    response.rate.heading <-
      paste(response.rate.heading,
            " = ",
            response.rate,
            " response rate",
            sep = "")
    table.col.names <- c("Field", "Average", "Responses")
    
    
    q.responses <-
      rbind(
        length(q1),
        length(q2),
        length(q3),
        length(q4),
        length(q5),
        length(q6),
        length(q7),
        length(q8)
      )
    
    table <- cbind(table, q.cols, q.responses)
    table <- rbind(table.col.names, table)
    colnames(table) <- table.col.names
    
    course.codes.heading <-
      unlist(sapply(sapply(course.codes, function(x)
        strsplit(x, "\\.")), function (y)
          paste(paste(y[1], y[2], sep = ""), sprintf("%03d", as.numeric(y[3])), sep = "."), simplify = F))
    course.codes.heading <- Reduce(paste, course.codes.heading)
    
    course.comments <-
      cbind(course.comments, rep("", length(course.comments)))
    course.comments <-
      cbind(course.comments, rep("", length(course.comments)))
    num.profs <- nrow(course.report)
    course.report <-
      rbind(
        c(course.name, "", ""),
        c("Summer 2019 Evals", "", ""),
        c(course.codes.heading, "", ""),
        c(response.rate.heading, "", ""),
        "",
        table,
        "",
        report.col.names,
        course.report,
        "",
        c("Comments", "", ""),
        course.comments
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
        border = "bottom",
        borderColour = "#000000",
        borderStyle = "thick",
        textDecoration = "bold",
        halign = "left"
      )
    columnStyle <-
      createStyle(
        fontSize = 12,
        fontColour = "#000000",
        border = "left",
        borderColour = "#333333",
        borderStyle = "thin",
        halign = "left"
      )
    rowStyle <-
      createStyle(
        fontSize = 12,
        fontColour = "#000000",
        border = "bottom",
        borderColour = "#333333",
        borderStyle = "thin",
        halign = "left"
      )
    
    textStyle <-
      createStyle(fontSize = 12,
                  fontColour = "#000000",
                  wrapText = TRUE)
    
    # page header style
    addStyle(
      wb,
      sheet.number,
      firstStyle,
      rows = 1,
      cols = 1:4,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      secondStyle,
      rows = 2,
      cols = 1:4,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      secondStyle,
      rows = 3,
      cols = 1:4,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      secondStyle,
      rows = 4,
      cols = 1:4,
      stack = TRUE
    )
    
    # table headings
    addStyle(
      wb,
      sheet.number,
      tableStyle,
      rows = 6,
      cols = 1:3,
      stack = TRUE
    )
    
    addStyle(
      wb,
      sheet.number,
      tableStyle,
      rows = 16,
      cols = 1:3,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      tableStyle,
      rows = (16 + num.profs + 2),
      cols = 1:3,
      stack = TRUE
    )
    
    # bottom borders for each field
    addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 7:14,
      cols = 1,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 7:14,
      cols = 2,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 7:14,
      cols = 3,
      stack = TRUE
    )
    
    # bottom borders for each professor
    addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 17:(16 + num.profs),
      cols = 1,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 17:(16 + num.profs),
      cols = 2,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 17:(16 + num.profs),
      cols = 3,
      stack = TRUE
    )
    
    addStyle(
      wb,
      sheet.number,
      columnStyle,
      rows = 16:(16 + num.profs),
      cols = 2,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      columnStyle,
      rows = 16:(16 + num.profs),
      cols = 3,
      stack = TRUE
    )
    
    
    addStyle(
      wb,
      sheet.number,
      columnStyle,
      rows = 6:14,
      cols = 2,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      columnStyle,
      rows = 6:14,
      cols = 3,
      stack = TRUE
    )
    
    addStyle(
      wb,
      sheet.number,
      textStyle,
      rows = seq(16 + num.profs + 3 , nrow(course.report)),
      cols = 1,
      stack = TRUE,
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
    #setColWidths(wb, sheet.number, cols = 1:4, widths = "auto")
    setColWidths(wb, sheet.number, cols = 1, widths = 70)
    #    setRowHeights(wb, sheet.number, rows = seq(16 + num.profs + 3, nrow(course.report)), heights = "auto")
    #    setRowHeights(wb, sheet.number, rows = seq(16 + num.profs + 3, nrow(course.report)), heights = 50)
    # makes sure sheet fits on one printable page
    pageSetup(wb,
              sheet.number,
              fitToWidth = TRUE,
              fitToHeight = FALSE)
    
    mergeCells(wb, 1, cols = 1:4, rows = 1)
    mergeCells(wb, 1, cols = 1:4, rows = 2)
    mergeCells(wb, 1, cols = 1:4, rows = 3)
    mergeCells(wb, 1, cols = 1:4, rows = 4)
    
    auto_heights(
      wb,
      sheet.number,
      seq(16 + num.profs + 3, nrow(course.report)),
      base_height = 150,
      extra_height = 20
    )
    
    saveWorkbook(wb, paste(file.name, ".xlsx", sep = ""), overwrite = TRUE) # writes a workbook containing all reports inputted
  }
}

export.semester.summary(semester.summary)

# reminds user that the num.ta.cols variable was not configured
if (num.ta.cols <= 0) {
  winDialog(
    type = c("ok"),
    "Variable num.ta.cols was <= 0, so no TA reports were generated. Update script with the correct number if this is an error."
  )
}

winDialog(type = c("ok"),
          "Your reports have been generated in the reports folder.")

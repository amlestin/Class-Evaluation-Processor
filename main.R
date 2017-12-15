# Author: Arnold Lestin
# Purpose: Extracts data from a Qualtrics survey evaluation CSV export file and outputs a Report.csv file with average reviews per professor

# opens a window to select the input file
winDialog(type = c("ok"),
          "Select the report file you exported from Qualtrics after the evaluation.")
evaluations.filename <- file.choose()

winDialog(
  type = c("ok"),
  "Select the contacts file you imported into Qualtrics before the evaluation."
)
student.contacts.filename <- file.choose()

evaluations <- read.csv(evaluations.filename)
student.contacts <- read.csv(student.contacts.filename)

all.codes <- c()
for (i in 1:nrow(student.contacts)) {
  s.code <- as.character(student.contacts[i, "Subject.Code"])
  c.num <- as.character(student.contacts[i, "Course.Number"])
  s.num <- as.character(student.contacts[i, "Sequence.Number"])
  
  code <- paste(s.code, c.num, s.num, sep = ".")
  all.codes <- c(all.codes, code)
}

unique.codes <- unique(all.codes)
course.sizes <- list(all.codes)

for (code in 1:length(unique.codes)) {
  cur.code <- unique.codes[code]
  course.sizes[[cur.code]] <- length(which(all.codes == cur.code))
}

# vector of character representations of all professor names used in the file
contacts <- c(
  # [-c(1:2)] removes the headers from each column
  # as.character converts from levels (integer) representations to names (character)
  as.character(d$P1[-c(1:2)]),
  as.character(d$P2[-c(1:2)]),
  as.character(d$P3[-c(1:2)]),
  as.character(d$P4[-c(1:2)]),
  as.character(d$P5[-c(1:2)]),
  as.character(d$P6[-c(1:2)]),
  as.character(d$P7[-c(1:2)]),
  as.character(d$P8[-c(1:2)]),
  as.character(d$P9[-c(1:2)]),
  as.character(d$P10[-c(1:2)]),
  as.character(d$P11[-c(1:2)]),
  as.character(d$P12[-c(1:2)]),
  as.character(d$P13[-c(1:2)]),
  as.character(d$P14[-c(1:2)]),
  as.character(d$P15[-c(1:2)]),
  as.character(d$P16[-c(1:2)])
)

# remove empty contacts (not all classes have all 16 professor slots filled)
contacts <- contacts[-(which(contacts == ""))]
# remove duplicates
contacts <- unique(contacts)

# Gets all reviews for each professor in contacts and adds them to reviewl
reviewl <- list()
for (cur.eval in 3:nrow(d)) {
  # for each P column, e.g, P1
  for (pctr in 1:16) {
    # convert the P number to a character
    pctr.char <- as.character(pctr)
    # combine the character and number to have a valid column index, e.g., "P1"
    pcol <-  paste("P", pctr.char, sep = "")
    
    # add one to the pctr because Q3 corresponds to P1's review
    qctr.char <- as.character(pctr + 1)
    qcol <- paste("Q", qctr.char, sep = "")
    
    # the question answer/review is the value at qcol
    review <- d[cur.eval, qcol]
    # level (integer) to character
    review <- as.character(review)
    
    prof.name <- as.character(d[cur.eval, pcol])
    # save the current eval's course title into a variable
    course.title <- as.character(d[cur.eval, "Course.Title"])
    subject.code <- as.character(d[cur.eval, "Subject.Code"])
    course.number <- as.character(d[cur.eval, "Course.Number"])
    sequence.number <- as.character(d[cur.eval, "Sequence.Number"])
    
    course.code <-
      paste(subject.code, course.number, sequence.number, sep = ".")
    
    if (prof.name != "") {
      # some respondents skip questions
      if (review != "") {
        # reviewl[[prof]]$courses[[course.title]]$ratings <- c(reviewl[[prof]]$courses[[course.title]]$ratings, review)
        reviewl[[prof.name]]$courses[[course.title]][[course.code]]$ratings <-
          c(reviewl[[prof.name]]$courses[[course.title]][[course.code]]$rating, review)
        
      }
    }
  }
  
  eval.comment <- as.character(d[cur.eval, "Q22"])
  
  comment.file.name <-
    paste(paste(course.code, sep = " "), ".txt", sep = "")
  
  if (eval.comment != "") {
    write("Comment: ", file = comment.file.name, append = TRUE)
    write(eval.comment, file = comment.file.name, append = TRUE)
    write("\n\n", file = comment.file.name, append = TRUE)
  }
}


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
        similar <- rbind(similar, c(contacts[i], contacts[j]))
        
        
      }
    }
  }
}

for (i in 1:nrow(similar)) {
  fname <- similar[i, 1]
  
  sname <- similar[i, 2]
  
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

# Outputs a report in the format [Professor's name].csv in Course Code, Course Title, Reponse Rate, Num Evals, Course Size, Average, Frequencies format
# for each professor in reviewl
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
      
      # number of "Excellent" reviews in row
      ecount <- length(which(reviews == "Excellent"))
      
      vgcount <- length(which(reviews == "Very Good"))
      
      gcount <- length(which(reviews == "Good"))
      
      fcount <- length(which(reviews == "Fair"))
      
      pcount <- length(which(reviews == "Poor"))
      
      # vector of all rating counts
      freqs <- c(ecount, vgcount, gcount, fcount, pcount)
      #names(freqs) <- c("Excellent", "Very Good", "Good", "Fair", "Poor")
      
      reviewl[[prof]]$courses[[cur.course]][[cur.section]][["freqs"]] <-
        freqs
      
      # sum of all ratings, i.e., number of ratings the professor received
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
      
      
      # use the scales package to represent response rate as a percent
      library(scales)
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
  prof.name.formatted <- gsub(" ", ".", prof.name)
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
  
  if (is.null(summary.report.ncol) == FALSE) {
    if (summary.line.length < summary.ncol)
    {
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

winDialog(type = c("ok"),
          "Your reports have been generated in the program directory.")
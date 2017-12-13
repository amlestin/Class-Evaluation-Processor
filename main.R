# Author: Arnold Lestin
# Purpose: Extracts data from a Qualtrics survey evaluation CSV export file and outputs a Report.csv file with average reviews per professor

# opens a window to select the input file
winDialog(type = c("ok"),
          "Select the report file you exported from Qualtrics after the evaluation.")
EXPORTED_FILENAME <- file.choose()

d <- read.csv(EXPORTED_FILENAME)

winDialog(
  type = c("ok"),
  "Select the contacts file you imported into Qualtrics before the evaluation."
)
CONTACTS_FILENAME <- file.choose()

student_contacts <- read.csv(CONTACTS_FILENAME)

all_codes <- c()
for (i in 1:nrow(student_contacts)){
  s_code <- as.character(student_contacts[i, "Subject.Code"])
  c_num <- as.character(student_contacts[i, "Course.Number"])
  s_num <- as.character(student_contacts[i, "Sequence.Number"])
  
  code <- paste(s_code, c_num, s_num, sep=".") 
  all_codes <- c(all_codes, code)
}

unique_codes <- unique(all_codes)
course_sizes <- list(all_codes)

for (code in 1:length(unique_codes)){
  cur_code <- unique_codes[code]
  course_sizes[[cur_code]] <- length(which(all_codes==cur_code))  
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
for (cur_eval in 3:nrow(d)) {
  # for each P column, e.g, P1
  for (pctr in 1:16) {
    # convert the P number to a character
    pctr_char <- as.character(pctr)
    # combine the character and number to have a valid column index, e.g., "P1"
    pcol <-  paste("P", pctr_char, sep = "")
    
    # add one to the pctr because Q3 corresponds to P1's review
    qctr_char <- as.character(pctr + 1)
    qcol <- paste("Q", qctr_char, sep = "")
    
    # the question answer/review is the value at qcol
    review <- d[cur_eval, qcol]
    # level (integer) to character
    review <- as.character(review)
    
    prof_name <- as.character(d[cur_eval, pcol])
    # save the current eval's course title into a variable
    course_title <- as.character(d[cur_eval, "Course.Title"])
    subject_code <- as.character(d[cur_eval, "Subject.Code"])
    course_number <- as.character(d[cur_eval, "Course.Number"])
    sequence_number <- as.character(d[cur_eval, "Sequence.Number"])
    
    course_code <-
      paste(subject_code, course_number, sequence_number, sep = ".")
    
    if (prof_name != "") {
      # some respondents skip questions
      if (review != "") {
        # reviewl[[prof]]$courses[[course_title]]$ratings <- c(reviewl[[prof]]$courses[[course_title]]$ratings, review)
        reviewl[[prof_name]]$courses[[course_title]][[course_code]]$ratings <-
          c(reviewl[[prof_name]]$courses[[course_title]][[course_code]]$rating, review)
        
      }
    }
  }
}


similar <- c()
MAX_NAME_DIFF <- 4

for (i in 1:length(contacts)) {
  for (j in 1:length(contacts)) {
    fname <- contacts[i]
    sname <- contacts[j]
    
    fname <- gsub(" ", "", fname, fixed = TRUE)
    
    sname <- gsub(" ", "", sname, fixed = TRUE)
    
    
    if ((adist(fname, sname) < MAX_NAME_DIFF) &
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
write.table(similar,
            "merged-names.txt",
            col.names = FALSE,
            row.names = FALSE,
            quote=FALSE)

# Outputs a report in the format [Professor's name].csv in Course Code, Course Title, Reponse Rate, Num Evals, Course Size, Average, Frequencies format
# for each professor in reviewl

for (prof in 1:length(reviewl)) {
  num_courses <- length(reviewl[[prof]]$courses)
  prof_name <- names(reviewl[prof])
  
  prof_report <- c()
  for (cur_course in 1:num_courses) {
    num_sections <- length(reviewl[[prof]]$courses[[cur_course]])
    
    for (cur_section in 1:num_sections) {
      reviews <-
        reviewl[[prof]]$courses[[cur_course]][[cur_section]]$ratings
      cur_course_code <-
        names(reviewl[[prof]]$courses[[cur_course]][cur_section])
      
      # number of "Excellent" reviews in row
      ecount <- length(which(reviews == "Excellent"))
      
      vgcount <- length(which(reviews == "Very Good"))
      
      gcount <- length(which(reviews == "Good"))
      
      fcount <- length(which(reviews == "Fair"))
      
      pcount <- length(which(reviews == "Poor"))
      
      # vector of all rating counts
      freqs <- c(ecount, vgcount, gcount, fcount, pcount)
      #names(freqs) <- c("Excellent", "Very Good", "Good", "Fair", "Poor")
      
      reviewl[[prof]]$courses[[cur_course]][[cur_section]][["freqs"]] <-
        freqs
      
      # sum of all ratings, i.e., number of ratings the professor received
      num_responses <- sum(freqs)
      
      # calculates a weighted average
      average <-
        (5 * ecount + 4 * vgcount + 3 * gcount + 2 * fcount + 1 * pcount) / num_responses
      reviewl[[prof]]$courses[[cur_course]][[cur_section]][["average"]] <-
        average
      
      cur_course_title <- names(reviewl[[prof]]$courses[cur_course])
      
      cur_course_size <- course_sizes[[cur_course_code]]
      
      
      # use the scales package to represent response rate as a percent
      library(scales)
      reviewl[[prof]]$courses[[cur_course]][[cur_section]][["response_rate"]] <-
        percent(num_responses / cur_course_size)
      
      line <-
        c(
          cur_course_code,
          cur_course_title,
          reviewl[[prof]]$courses[[cur_course]][[cur_section]][["response_rate"]],
          num_responses,
          cur_course_size,
          reviewl[[prof]]$courses[[cur_course]][[cur_section]][["average"]],
          reviewl[[prof]]$courses[[cur_course]][[cur_section]][["freqs"]]
        )
      prof_report <- rbind(prof_report, line)
    }
  }
    
  colnames(prof_report) <-
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
  rownames(prof_report) <- NULL
  prof_name <- gsub(" ", "_", prof_name)
  PROF_REPORT_NAME <- paste(prof_name, ".csv", sep = "")
  write.table(prof_report,
              PROF_REPORT_NAME,
              sep = ",",
              row.names= FALSE)
}

winDialog(
  type = c("ok"),
  "Your reports have been generated in the program directory."
)

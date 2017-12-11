# Author: Arnold Lestin
# Purpose: Extracts data from a Qualtrics survey evaluation CSV export file and outputs a Report.csv file with average reviews per professor

# opens a window to select the input file
INPUT_FILENAME <- file.choose()

# read the csv file in ("FILENAME.csv")
d <- read.csv(INPUT_FILENAME)

# vector of character representation of all professor names used in the file
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

# remove empty contacts (not all classes have 16 all professors)
contacts <- contacts[-(which(contacts == ""))]
# remove duplicates
contacts <- unique(contacts)



# Gets all reviews for each professor in contacts and adds them to reviewl, e.g., reviewl[[prof]] gives
# [1] "Good" "Fair"
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
    
    # retrieve the professor's past reviews in reviewl and add this one to them
    prof <- as.character(d[cur_eval, pcol])
    # save the current eval's course title into a variable
    course_title <- as.character(d[cur_eval, "Course.Title"])
    
    if (prof != "") {
      # some respondents skip questions
      if (review != "") {
        # reviewl[[prof]]$courses[[course_title]]$ratings <- c(reviewl[[prof]]$courses[[course_title]]$ratings, review)
        reviewl[[prof]]$courses[[course_title]]$ratings <- c(reviewl[[prof]]$courses[[course_title]]$rating, review);
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
  
  
  #reviewl[[fname]] <- c(reviewl[[fname]], reviewl[[sname]])
  #reviewl[[fname]]$courses <- c(reviewl[[fname]]$courses, reviewl[[sname]]$courses)
  
  reviewl[[fname]]$courses <-
    c(reviewl[[fname]]$courses, reviewl[[sname]]$courses)
  reviewl[[fname]]$ratings <-
    c(reviewl[[fname]]$ratings, reviewl[[sname]]$ratings)
  reviewl[[sname]] <- NULL
}

profnames <- names(reviewl)
# makes reviewl easier to iterate over
reviewl <- lapply(reviewl, unlist)
# find the most number of reviews any professor received
max <- max(sapply(reviewl, length))

# pads the reviewl list with NA values until all professors have max reviews

reviewl <-
  do.call(rbind, lapply(reviewl, function(z)
    c(z, rep(NA, max - length(
      z
    )))))

#reviewl[[5]]["ratings"]




# Finds each professor's numeric average review and the rating counts used to calculate the average
Averages <- c()
Counts <- c()

# for each professor in reviewl
for (i in 1:nrow(reviewl)) {
  # row is all of their reviews
  row <- reviewl[i, ]
  
  # number of "Excellent" reviews in row
  ecount <- length(which(row == "Excellent"))
  
  vgcount <- length(which(row == "Very Good"))
  
  gcount <- length(which(row == "Good"))
  
  fcount <- length(which(row == "Fair"))
  
  pcount <- length(which(row == "Poor"))
  
  # vector of all rating counts
  all <- c(ecount, vgcount, gcount, fcount, pcount)
  
  # adds the rating counts as a row of matrix Counts
  Counts <- rbind(Counts, all)
  
  # sum of all ratings, i.e., number of ratings the professor received
  sum <- sum(all)
  
  # calculates a weighted average
  average <-
    (5 * ecount + 4 * vgcount + 3 * gcount + 2 * fcount + 1 * pcount) / sum
  
  Averages <- c(Averages, average)
}

out <- cbind(profnames, Averages, Counts, reviewl)

# column labels for the Counts variables
ratings <- c("Excellent", "Very Good", "Good", "Fair", "Poor")


colnames(out)[1] = "Professor"
colnames(out)[2] = "Average"
colnames(out)[3:7] = ratings

# there is an arbitrary number of other columns representing the number of reviews each professor got
# ncol(out) is eqaul to the max number of reviews a professor received
# delete 7 to make the lengths equal
colnames(out)[8:ncol(out)] = 1:(ncol(out) - 7)


write.table(similar,
            "merged-names.txt",
            col.names = FALSE,
            row.names = FALSE)

# output a csv file
# do not include the table's row.names column
INPUT_FILENAME <- sub('\\..*$', '', basename(INPUT_FILENAME))

OUTPUT_FILENAME <-
  paste("Report-", INPUT_FILENAME, ".csv", sep = "")
write.table(out,
            OUTPUT_FILENAME,
            sep = ",",
            row.names = FALSE)

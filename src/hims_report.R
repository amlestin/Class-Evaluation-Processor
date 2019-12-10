# HIM Simplified Reports

library(openxlsx)
wb <- createWorkbook()
addWorksheet(wb, 1)

dist <- read.csv(file.choose(), stringsAsFactors = FALSE)

# find the columns corresponding to questions
response.cols <- grep("X[0-9]", colnames(dist))
# remove comments column
response.cols <- response.cols[-length(response.cols)]

# collect responses
responses <- dist[3:nrow(dist), response.cols]

# question types in order of rating from least desirable to mosts desirable responses
a <- c("Never", "Sometimes", "About half the time", "Most of the time", "Always")
b <- c("Definitely not", "Probably not", "Might or might not", "Probably yes", "Definitely yes")
c <- c("Terrible", "Poor", "Average", "Good", "Excellent")
d <- c("Very difficult", "Difficult", "Neither easy nor difficult", "Easy", "Very easy")
e <- c("Not helpful", "Not very helpful", "No opinion", "Sometimes helpful", "Definitely helpful")
question.types <- list(a, b, c, d, e)

q.a <- c()
for (col.num in 1:length(response.cols)){
  question.col <- as.vector(dist[, response.cols[col.num]])
  question.col <- question.col[-c(1:2)]

  #ratings <- vapply(question.col, function(a) grep(a, question.types[[grep(a, question.types)]]), c(1))
  #ratings <- vapply(question.col, function(a) { print(grep(a, question.types))}, c(1))
  #names(ratings) <- NULL
  getmode <- function(v) {
    #https://www.tutorialspoint.com/r/r_mean_median_mode.htm
    uniqv <- unique(v)
    uniqv[which.max(tabulate(match(v, uniqv)))]
  }
  
  q.type <- c()
  for (a in question.col) {
    q.type <- c(q.type, grep(a, question.types))
  }
  q.type <- getmode(q.type)
  ratings <- vapply(question.col, function(a) grep(a, question.types[[q.type]]), c(1))
  question.average <- round(mean(ratings), 2)
  
  q.a <- rbind(q.a, c(dist[1, response.cols[col.num]], question.average))
  
 # dist[nrow(dist)+1, response.cols[col.num]] <- question.average
}


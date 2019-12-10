# HIM Simplified Reports

dist <- read.csv(file.choose(), stringsAsFactors = FALSE)


getmode <- function(v) {
  #https://www.tutorialspoint.com/r/r_mean_median_mode.htm
  uniqv <- unique(v)
  uniqv[which.max(tabulate(match(v, uniqv)))]
}

library(openxlsx)
wb <- createWorkbook()
addWorksheet(wb, 1)

# find the columns corresponding to questions
response.cols <- grep("X[0-9]", colnames(dist))
# remove comments column
response.cols <- response.cols[-length(response.cols)]

# collect responses
responses <- dist[3:nrow(dist), response.cols]

# question types in order of rating from least desirable to mosts desirable responses
a <-
  c("Never",
    "Sometimes",
    "About half the time",
    "Most of the time",
    "Always")
b <-
  c("Definitely not",
    "Probably not",
    "Might or might not",
    "Probably yes",
    "Definitely yes")
c <- c("Terrible", "Poor", "Average", "Good", "Excellent")
d <-
  c("Very difficult",
    "Difficult",
    "Neither easy nor difficult",
    "Easy",
    "Very easy")
e <-
  c(
    "Not helpful",
    "Not very helpful",
    "No opinion",
    "Sometimes helpful",
    "Definitely helpful"
  )
question.types <- list(a, b, c, d, e)

q.a <- c()
for (col.num in 1:length(response.cols)) {
  question.col <- as.vector(dist[, response.cols[col.num]])
  question.col <- question.col[-c(1:2)]
  
  #ratings <- vapply(question.col, function(a) grep(a, question.types[[grep(a, question.types)]]), c(1))
  #ratings <- vapply(question.col, function(a) { print(grep(a, question.types))}, c(1))
  #names(ratings) <- NULL
  
  
  q.type <- c()
  for (a in question.col) {
    q.type <- c(q.type, grep(a, question.types))
  }
  q.type <- getmode(q.type)
  ratings <-
    vapply(question.col, function(a)
      grep(a, question.types[[q.type]]), c(1))
  question.average <- round(mean(ratings), 2)
  
  q.a <-
    rbind(q.a, c(dist[1, response.cols[col.num]], question.average))
  
  # dist[nrow(dist)+1, response.cols[col.num]] <- question.average
}

names(q.a) <- NULL
q.a[1, ] <- c("Question", "Average")
q.a <- rbind(c("HIMS Course Evaluations",""), q.a)


setColWidths(wb, 1, cols = 1:2, widths = "auto")
headerStyle <-
  createStyle(fontSize = 18,
              fontColour = "#333333",
              halign = "center")
tableStyle <-
  createStyle(
    fontSize = 12,
    fontColour = "#000000",
    border = "TopBottomLeftRight",
    borderColour = "#000000",
    borderStyle = "thin",
    #    textDecoration = "bold",
    halign = "left"
  )


addStyle(wb, 1, headerStyle, cols = 1, rows = 1)
mergeCells(wb, 1, cols = 1:3, rows = 1)

for (i in 2:nrow(q.a)) {
  addStyle(wb, 1, tableStyle, cols = 1, rows = i)
  addStyle(wb, 1, tableStyle, cols = 2, rows = i)
}

writeData(wb, 1, q.a, colNames = FALSE)
saveWorkbook(wb, "HIMS-report.xlsx", overwrite = TRUE)


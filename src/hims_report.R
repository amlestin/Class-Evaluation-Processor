# HIM Simplified Reports

dist <- read.csv(file.choose(), stringsAsFactors = FALSE)

library(openxlsx)
wb <- createWorkbook()
addWorksheet(wb, 1)

# find the columns corresponding to questions
response.cols <- grep("X[0-9]", colnames(dist))
# remove comments column
response.cols <- response.cols[-length(response.cols)]

q.a <- c()
for (col.num in 1:length(response.cols)) {
  question.col <- as.vector(dist[, response.cols[col.num]])
  question.col <- as.numeric(question.col[-c(1:2)])
  
  
  question.average <- round(mean(question.col), 2)
  
  q.a <-
    rbind(q.a, c(dist[1, response.cols[col.num]], question.average))
}

names(q.a) <- NULL
q.a[1, ] <- c("Question", "Average")
q.a <- rbind(c("HIMS Course Evaluations", ""), q.a)


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
    halign = "left"
  )


addStyle(wb, 1, headerStyle, cols = 1, rows = 1)
mergeCells(wb, 1, cols = 1:2, rows = 1)

for (i in 2:nrow(q.a)) {
  addStyle(wb, 1, tableStyle, cols = 1, rows = i)
  addStyle(wb, 1, tableStyle, cols = 2, rows = i)
}

writeData(wb, 1, q.a, colNames = FALSE)
saveWorkbook(wb, "HIMS-report.xlsx", overwrite = TRUE)

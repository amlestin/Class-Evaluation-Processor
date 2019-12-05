# HIM Simplified Reports

dist <- read.csv(file.choose())

response.cols <- grep("X[0-9]", colnames(dist))
responses <- dist[3:nrow(dist), response.cols]

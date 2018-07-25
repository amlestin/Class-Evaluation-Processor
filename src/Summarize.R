f.path <- file.choose()
first <- read.csv(f.path)
first.names <- first$Name
  
s.path <- file.choose()
second <- read.csv(s.path)
second.names <- second$Name

View(first)
View(second)


max.name.diff <- 8
for (left in first.names) {
  other.name.index = which(adist(left, second.names) < max.name.diff)
  
  other.name = as.character(second.names[other.name.index])
  
  other.exists <- length(other.name > 0)
  
  if(other.exists) {
    print(paste("Found a match: ", noquote(left), " = ", other.name))
  } else {
    print(paste("                No match for: ", left))
  }
  
}

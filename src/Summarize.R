f.path <- file.choose()
first <- read.csv(f.path)
first.names <- first$Name
  
s.path <- file.choose()
second <- read.csv(s.path)
second.names <- second$Name

View(first)
View(second)

for (left in first.names) {
  max.name.diff <- 4
  
  other.name.index = which(left == second.names)
  other.name.index = which(left == second.names)
  
  other.name = as.character(second.names[other.name.index])
  
  other.exists <- adist(left, other.name) < max.name.diff
  
  if(other.exists) {
    print(paste("Found a match for", noquote(left)))
  }
  
}

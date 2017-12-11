INPUT_FILENAME <- file.choose()

c <- read.csv(INPUT_FILENAME)

all_titles <- as.character(c$Course.Title)
unique_titles <- unique(all_titles)
counts <- list(unique_titles)

for (title in 1:length(unique_titles)){
  cur_title <- unique_titles[title]
  counts[[cur_title]] <- length(which(all_titles==cur_title))  
}

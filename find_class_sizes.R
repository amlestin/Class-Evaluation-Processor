INPUT_FILENAME <- file.choose()

c <- read.csv(INPUT_FILENAME)

all_codes <- c()
for (i in 1:nrow(c)){
 s_code <- as.character(c[i, "Subject.Code"])
 c_num <- as.character(c[i, "Course.Number"])
 s_num <- as.character(c[i, "Sequence.Number"])
 
 code <- paste(s_code, c_num, s_num, sep=".") 
 all_codes <- c(all_codes, code)
}

unique_codes <- unique(all_codes)
course_sizes <- list(all_codes)

for (code in 1:length(unique_codes)){
  cur_code <- unique_codes[code]
  course_sizes[[cur_code]] <- length(which(all_codes==cur_code))  
}

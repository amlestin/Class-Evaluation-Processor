winDialog(
  type = c("ok"),
  "Select the file with the course rosters/students for all your courses to be evaluated. (CSV format)"
)

contacts.file <- file.choose()

winDialog(type = c("ok"),
          "Select the file with all instructors for each course. (CSV format)")

profs.by.crn.file <- file.choose()

profs.by.crn <-
  read.csv(profs.by.crn.file, colClasses = "character")
contacts <- read.csv(contacts.file, colClasses = "character")


contacts["PROF1"] <- NA
contacts["PROF2"] <- NA
contacts["PROF3"] <- NA
contacts["PROF4"] <- NA
contacts["PROF5"] <- NA
contacts["PROF6"] <- NA
contacts["PROF7"] <- NA
contacts["PROF8"] <- NA
contacts["PROF9"] <- NA
contacts["PROF10"] <- NA
contacts["PROF11"] <- NA
contacts["PROF12"] <- NA
contacts["PROF13"] <- NA
contacts["PROF14"] <- NA
contacts["PROF15"] <- NA
contacts["PROF16"] <- NA

contacts["TA1"] <- NA
contacts["TA2"] <- NA
contacts["TA3"] <- NA
contacts["TA4"] <- NA
contacts["TA5"] <- NA
contacts["TA6"] <- NA
contacts["TA7"] <- NA
contacts["TA8"] <- NA
contacts["TA9"] <- NA
contacts["TA10"] <- NA
contacts["TA11"] <- NA
contacts["TA12"] <- NA
contacts["TA13"] <- NA
contacts["TA14"] <- NA

convert.lastname.comma.firstname.to.firstname.space.lastname <-
  function(pair.of.strings) {
    if (length(pair.of.strings) == 0) {
      return("")
    }
    return(paste(pair.of.strings[2], pair.of.strings[1]))
  }

error.log <- c()
profs <- list()
for (i in 1:nrow(contacts)) {
  current.crn <- as.character(contacts[i, "CRN"])
  course.index <- which(profs.by.crn["CRN"] == current.crn)
  
  # skip blank lines
  if (is.na(contacts[i, "TITLE"]))
    next
  
  # skip courses that do not have records in profs.by.crn
  # logs error to file
  if (identical(course.index, integer(0))) {
    error <-
      paste("NOT INCLUDING ",
            as.character(contacts[i, "TITLE"]),
            " CRN: ",
            current.crn)
    
    print(error)
    error.log <- c(error.log, error)
    original.prof.cols <- grep("PROF", colnames(contacts[i, ]))
    
    course.director <- as.character(contacts[i, "INSTRUCTOR"])
    
    split.name <- unlist(strsplit(course.director, ","))
    first.name <- trimws(split.name[2])
    last.name <- trimws(split.name[1])
    
    spaced.full.name <- paste(first.name, last.name)
    
    contacts[i, original.prof.cols] <-
      c(course.director, rep(NA, length(original.prof.cols) - 1))
    contacts[i,] <- NA
    next
  }
  
  course.info <- profs.by.crn[course.index, ]
  
  #  prof.cols <- grep("INSTRUCTOR[0-9]", colnames(course.info))
  prof.cols <- grep("PROF[0-9]", colnames(course.info))
  ta.cols <- grep("TA[0-9]", colnames(course.info))
  
  course.profs <- course.info[prof.cols]
  course.tas <- as.character(course.info[ta.cols])
  original.ta.cols <- grep("TA[0-9]", colnames(contacts[i, ]))
  
  # course.profs[which(is.na(course.profs))] <- ""
  course.tas[which(is.na(course.tas))] <- ""
  
  contacts[i, original.ta.cols] <- course.tas
  
  if (length(course.tas) != 0) {
    # TODO not needed because of name format F19
    # course.tas <-
    #   unlist(
    #     lapply(
    #       strsplit(course.tas, ", "),
    #       convert.lastname.comma.firstname.to.firstname.space.lastname
    #     )
    #   )
    contacts[i, original.ta.cols] <- course.tas
  }
  
  # TODO: may need to AND parameters to prevent
  valid.profs <-
    as.character(course.profs[unique(c(which(course.profs != ""), which(!is.na(course.profs))))])
  
  # TODO: not needed because name format F19
  # valid.profs <-
  #   unlist(
  #     lapply(
  #       strsplit(valid.profs, ", "),
  #       convert.lastname.comma.firstname.to.firstname.space.lastname
  #     )
  #   )
  #
  # TODO: should not occur because of check above
  if (length(which(valid.profs == "")) > 0) {
    valid.profs <- valid.profs[-c(which(valid.profs == ""))]
  }
  
  # TODO: add to error log
  if (length(valid.profs) == 0) {
    #   profs[[current.crn]] <- "NO VALID PROFS"
  } else {
    #  profs[[current.crn]] <- c(current.crn, valid.profs)
    original.prof.cols <- grep("PROF", colnames(contacts[i, ]))
    
    contacts[i, original.prof.cols] <-
      c(valid.profs, rep(NA, length(original.prof.cols) - length(valid.profs)))
  }
}


colnames(contacts)[which(colnames(contacts) == "EMAIL.ADDRESS")] = "EMAIL"
colnames(contacts)[which(colnames(contacts) == "FIRST.NAME")] = "FirstName"
colnames(contacts)[which(colnames(contacts) == "LAST.NAME")] = "LastName"

contacts <-
  contacts[!apply(is.na(contacts) | contacts == "", 1, all), ]

working.dir <- trimws(dirname(contacts.file))
setwd(working.dir)

write.table(
  contacts,
  "Instructor Added Contacts.csv",
  sep = ",",
  row.names = FALSE,
  na = ""
)

write.table(
  unique(error.log),
  "error log.txt",
  col.names = FALSE,
  row.names = FALSE,
  quote = FALSE
)

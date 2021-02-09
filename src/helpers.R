
# 
find.sections.by.course <-
  function(course.index, semester.summary) {
    sections <- c()
    for (prof.section in 1:length(semester.summary[[course.index]])) {
      sections <-
        c(sections, names(semester.summary[[course.index]][[prof.section]]))
    }
    course.sections <- unique(sections)
    
    return(course.sections)
  }

evals.to.freqs <- function(reviews) {
  # counts frequencies of each possible review response
  if (any(reviews == "Excellent") |
      any(reviews == "Very Good") |
      any(reviews == "Good") |
      any(reviews == "Fair") | any(reviews == "Poor")) {
    ecount <- length(which(reviews == "Excellent"))
    vgcount <- length(which(reviews == "Very Good"))
    gcount <- length(which(reviews == "Good"))
    fcount <- length(which(reviews == "Fair"))
    pcount <- length(which(reviews == "Poor"))
  } else {
    ecount <- length(which(reviews == "5"))
    vgcount <- length(which(reviews == "4"))
    gcount <- length(which(reviews == "3"))
    fcount <- length(which(reviews == "2"))
    pcount <- length(which(reviews == "1"))
  }
  
  return(c(ecount, vgcount, gcount, fcount, pcount))
}

evals.to.average <- function(reviews) {
  freqs <- evals.to.freqs(reviews)
  # sums frequencies of all reviews in the section, i.e., total number of ratings the professor received
  num.ratings <- sum(freqs)
  
  # calculates a weighted average
  ratings.prod <-
    (5 * freqs[1] + 4 * freqs[2] + 3 * freqs[3] + 2 * freqs[4] + 1 * freqs[5])
  
  average <- ratings.prod / num.ratings
  average <- round(average, digits = 2)
  average <- sprintf("%0.2f", average)
  
  return (average)
}

course.size.from.sections <- function(course.codes) {
  section.sizes <- c()
  for (code in course.codes) {
    section.size <-  course.sizes[[code]]
    section.sizes <- c(section.sizes, section.size)
  }
  course.size <- sum(section.sizes)
  
  return(course.size)
}

# The prompt should be in the main file
split.course.summary <- function(course.index, semester.summary) {
  course.title <- names(semester.summary)[[course.index]]
  course.sections <-
    find.sections.by.course(course.index, semester.summary)
  
  original.num.sections <- length(course.sections)
  if (original.num.sections == 1) {
    print("Error: There is only 1 section in this course.")
    return(NULL)
  }
  
  print.sections(course.title, course.sections)
  
  print("How many reports would you like to separate this course into?")
  prompt.message <- "Enter # reports to split into: "
  num.reports <- readline(prompt.message)
  num.reports <- as.integer(num.reports)
  #TODO handle invalid num.reports
  
  if (num.reports > original.num.sections) {
    print("Error: More reports than sections. Restart script and try again.")
    return(NULL)
  }
  
  if (num.reports <= 1) {
    print("Error: Invalid number of reports. Restart script and try again.")
    return(NULL)
  }
  
  # print the course section options
  
  reports <- list()
  for (r.number in 1:num.reports) {
    print.sections(course.title, course.sections)
    
    print("Separate input by commas. Ex: [1,2,3]")
    prompt.message <-
      paste(
        "Enter the section numbers (#?) you want included in report #",
        r.number,
        "/",
        num.reports,
        ": ",
        sep = ""
      )
    
    section.numbers <- readline(prompt.message)
    
    #TODO: validate section numbers and reset prompt if invalid
    
    section.numbers <-
      as.integer(unlist(strsplit(section.numbers, ",")))
    # if (length(section.numbers) < r.number)
    
    
    
    
    if (all(section.numbers %in% 1:length(course.sections))) {
      chosen.sections <- course.sections[section.numbers]
      for (section in chosen.sections) {
        course.sections <-
          course.sections[-which(course.sections == section)]
        
        if (length(reports) >= r.number) {
          # TODO: do not append create sublist
          reports[[r.number]] <-
            c(reports[[r.number]], section)
        }
        else {
          reports[[r.number]] <- section
          # names(reports)[r.number] <- r.number
        }
        
        
        
        if (length(course.sections) == 0 &
            length(reports) < num.reports) {
          print("ERROR: No more sections to allocate!")
          return(NULL)
        }
        
      }
    }
    else {
      print("ERROR: invalid section index. Please restart script.")
      return(NULL)
    }
  }
  report.name <- names(semester.summary)[course.index]
  names(reports) <- report.name
  return(reports)
}

# TODO: Function has subscript error with css, also tried to change brackets with the dups
create.export.ss <- function(reports.by.codes) {

  course.title <- names(reports.by.codes)[1]
  css <- list()
  # TODO: remove dependency on reviews.by.course.code
  reviews <- reviews.by.course.code[course.title]
  reviews.by.section <-
    unlist(reviews)
  
  for (report in reports.by.codes) {
    ss <- list()
    for (section in report) {
      section.ind <-
        which(grepl(section, names(reviews.by.section)) == TRUE)
      section.reviews <- reviews.by.section[section.ind]
      # print(evals.to.average(section.reviews))
      
      # vector of index of columns with PROF in their column name
      prof.cols.fmt <- paste("^",c.PROF, "[0-9]|", c.PROF, "[0-9][0-9]", sep = '')
      ta.cols.fmt <- paste("^", c.TA, "[0-9]|", c.TA, "[0-9][0-9]", sep = '')
        
      prof.cols <- grep(prof.cols.fmt, names(evals))
      # vector of index of columns with TA in their column name
      ta.cols <- grep(ta.cols.fmt, names(evals))
      
      pcols <- c(prof.cols, ta.cols)
      
      # TEST
      # associate the evaluated professors with each course title
      
      # separate section into three parts: code, course, section
      v <- strsplit(section, ".")
      v <- unlist(strsplit(section, split = ".", fixed = TRUE))
      
      sjc <- v[1]
      crs <- v[2]
      sc <- v[3]
      
      eval.ind <-
        which(evals$SUBJECT.CODE == sjc &
                evals$COURSE.. == crs & evals$SECTION.. == sc)
      
      profs.by.course <-
        mapply(function(eval.ind) {
          # extract unique professors from evals
          evals[eval.ind, pcols]
        }
        , eval.ind, SIMPLIFY = T)
      
      profs.by.course <- unique(unlist(profs.by.course))
      # remove empty course titles
      profs.by.course <-
        profs.by.course[which(profs.by.course != "")]
      
      # boolean flag that a course had no evaluations
      unevaluated.course.found <-
        length(which(lengths(profs.by.course) == 0)) > 0
      
      # if a course was not evaluated
      if (unevaluated.course.found)
        profs.by.course <-
        profs.by.course[-which(lengths(profs.by.course) == 0)]
      
      # only one course at a time so i is always 1
      i <- 1
      # get the ith course name from profs.by.course
      course.summary <- list()
      for (j in 1:length(profs.by.course)) {
        # get the jth prof for the ith course
        prof.name <- profs.by.course[[j]]
        
        # get evals for the professor's name in that course
        eval.index <- which(names(reviewl) == prof.name)
        prof.sections <-
          names(reviewl[[eval.index]]$courses[[course.title]])
        
        #this professor was not in the given section
        if ((section %in% prof.sections) == FALSE) {
          next()
        }
        
        if (length(eval.index) == 0) {
          cat(paste(prof.name, "has no evaluations in", course.title))
          next()
        }
        
        prof.evals <-
          reviewl[[eval.index]][["courses"]][[course.title]][[section]]
        
        # save the courses evals to course summary for a prof in this section
        if (is.null(course.summary[[prof.name]][[section]]))
          course.summary[[prof.name]][[section]] <- prof.evals
      }
      
      # append the section summary to the semester summary
      t <- course.title
      ss[[t]] <- c(ss[[t]], course.summary)
    }
    # names(ss) <- paste(names(ss), paste(sort(report), collapse = " "))
    # print(ss)
    
    # TODO: check
    dups <- which(table(names(ss[[1]])) > 1)
    dups.ss.ind <- list()
    
    
    if (length(dups) > 0) {
      for (prof in names(dups)) {
        course <- ss[[1]]
        dups.ss.ind <- which(names(course) == prof)
        
        # print(dups.ss.ind)
        
        e <- course[dups.ss.ind]
        
        ratings <- c()
        freqs <- c()
        
        combined.sections <- c()
        for (section in e) {
          data <- section[[1]]
          section.name <- names(section)
          combined.sections <- c(combined.sections, section.name)
          ratings <- c(ratings, data$ratings)
        }
        combined.average <- evals.to.average(ratings)
        combined.freqs <- evals.to.freqs(ratings)
        combined.size <-
          course.size.from.sections(combined.sections)
        combined.sections <-
          paste(combined.sections, collapse = " ")
        
        # Error: Empty List when being iterated through
        css[[course.title]][[prof]][[combined.sections]][["ratings"]] <-
          ratings
        css[[course.title]][[prof]][[combined.sections]][["freqs"]] <-
          combined.freqs
        css[[course.title]][[prof]][[combined.sections]][["average"]] <-
          as.numeric(combined.average)
        css[[course.title]][[prof]][[combined.sections]][["response.rate"]] <-
          percent(length(ratings) / combined.size)
        
 
        # ss[dups.ss.ind] <- NULL
        # ss[[(length(ss) + 1)]] <- css
        
        # ss <- css
        print(prof)
        print(combined.average)
      }
    }
    
    # This leads into a call passing css as an empty parameter.
    export.semester.summary(css, s = TRUE)
  }
}

create.semester.summary <- function(reviewl) {
  # character vector of the title of every course evaluated
  unique.titles <- as.character(unique(student.contacts[, c.TITLE]))
  unique.titles <- sort(unique.titles)
  
  # vector of index of columns with PROF in their column name
  prof.cols <- grep(c.PROF, names(evals))
  
  # associate the evaluated professors with each course title
  profs.by.course <-
    mapply(
      function(course)
        # remove empty course titles
        course <- course[which(course != "")],
      mapply(
        function(course)
          # unlist and convert to character
          as.character(unlist(course)),
        mapply(function(course.title)
          # extract unique professors from evals with the given course title
          unique(evals[which(evals$TITLE == course.title),][, prof.cols])
          , unique.titles, SIMPLIFY = F),
        SIMPLIFY = F
      ),
      SIMPLIFY = F
    )
  
  # boolean flag that a course had no evaluations
  unevaluated.course.found <-
    length(which(lengths(profs.by.course) == 0)) > 0
  
  # if a course was not evaluated
  if (unevaluated.course.found)
    profs.by.course <-
    profs.by.course[-which(lengths(profs.by.course) == 0)]
  
  semester.summary <- c()
  for (i in seq(length(profs.by.course))) {
    # get the ith course name from profs.by.course
    course.name <- names(profs.by.course)[[i]]
    
    course.summary <- list()
    for (j in seq(length(profs.by.course[[i]]))) {
      # get the jth prof for the ith course
      prof.name <- profs.by.course[[i]][[j]]
      
      # get evals for the professor's name in that course
      eval.index <- which(names(reviewl) == prof.name)
      
      if (length(eval.index) == 0) {
        cat(paste(prof.name, "has no evaluations in", course.name))
        next()
      }
      prof.evals <-
        reviewl[[eval.index]][["courses"]][[course.name]]
      
      # extract prof evals in the order of the alphabetically sorted professor names
      prof.evals <-
        prof.evals[sort(names(prof.evals), index.return = T)$ix]
      
      # save the courses evals to course summary
      course.summary[[prof.name]] <- prof.evals
    }
    
    # append the course summary to the semester summary
    semester.summary[[course.name]] <- course.summary
  }
  return(semester.summary)
}

export.semester.summary <- function(semester.summary, s = FALSE) {
  report.col.names  <-
    c("Professor", "Average", "Responses", "Poor", "Fair", "Good", "Very Good", "Excellent")
  for (course.index in seq(length(semester.summary))) {
    course.summary <- semester.summary[[course.index]]
    course.name <- names(semester.summary)[[course.index]]
    course.codes <- c()
    
    unique.sections <-
      unique(find.sections.by.course(course.index, semester.summary))
    
    
    
    uscs.formatted <- paste(unique.sections, collapse = ' ')
    if ((s == FALSE) && length(unique.sections) > 1) {
      cat(
        paste(
          "Course can be split: ",
          course.name,
          " (",
          course.index,
          ") ",
          "- ",
          uscs.formatted,
          "\n"
          ,
          sep = ''
        )
      )
      
      j <-
        readline(
          "Press enter to continue without splitting. Press [S] to split course into multiple reports: "
        )
      
      if (j == "S") {
        # Create split report
        
        reports.by.codes <- split.course.summary(course.index, semester.summary)
        
        ### This may be unnecessary
        # reports.by.codes <- NULL
         
        # while (is.null(reports.by.codes)) {
        #  reports.by.codes <-
        #    split.course.summary(course.index, semester.summary)
        }
        
        # This call gives subscript error
        # Passes reports.by.codes and then within create.export.ss function, export.semester.summary gets called again and since css is length 0...
        # The call breaks
        # It's as if this is a recursive function which I wouldn't think would be optimal.
        create.export.ss(reports.by.codes)
        
        # split the relevant sections manually and remove them from semester.summary
        # semester.summary <- semester.summary[-course.index]
        next()
      }
      
    }
    
    sheet.number <- 1
    wb <- createWorkbook("Admin")
    addWorksheet(wb, sheet.number) # add modified report to a worksheet
    
    course.report <- c()
    course.respondents <- c()
    for (prof.index in seq(length(course.summary))) {
      freqs <- c(0, 0, 0, 0, 0)
      for (sec.index in seq(length(semester.summary[[course.index]][[prof.index]]))) {
        freqs <-
          freqs + semester.summary[[course.index]][[prof.index]][[sec.index]][["freqs"]]
      }
      
      
      prof.course.codes <-
        names(semester.summary[[course.index]][[prof.index]])
      
      if (s == TRUE) {
        prof.course.codes <-
          unlist(strsplit(names(semester.summary[[course.index]][[prof.index]]), split = " "))
      }
      
      cat(
        "Processing: ",
        course.name,
        "-",
        prof.course.codes,
        "-",
        names(semester.summary[[course.index]])[prof.index],
        "\n"
      )
      
      course.codes <- c(course.codes, prof.course.codes)
      # sums frequencies of all reviews in the section, i.e., total number of ratings the professor received
      num.ratings <- sum(freqs)
      
      # calculates a weighted average
      ratings.prod <-
        (5 * freqs[1] + 4 * freqs[2] + 3 * freqs[3] + 2 * freqs[4] + 1 * freqs[5])
      
      average <- ratings.prod / num.ratings
      average <- round(average, digits = 2)
      average <- sprintf("%0.2f", average)
      
      
      course.size <- course.size.from.sections(prof.course.codes)
      
      row <-
        c(
          names(semester.summary[[course.index]])[[prof.index]],
          average,
          paste(num.ratings, "/",
                course.size, sep =
                  ""),
          freqs[5],
          freqs[4],
          freqs[3],
          freqs[2],
          freqs[1]
        )
      
      course.report <- rbind(course.report, row)
      course.respondents <- c(course.respondents, num.ratings)
    }
    
    course.codes <- sort(unique(course.codes))
    
    q1 <- c()
    q2 <- c()
    q3 <- c()
    q4 <- c()
    q5 <- c()
    q6 <- c()
    q7 <- c()
    q8 <- c()
    
    course.comments <- c()
    for (i in seq(length(course.codes))) {
      course.code <- course.codes[i]
      
      course.code.array <- unlist(strsplit(course.code, "\\."))
      subject.code <- course.code.array[1]
      course <- course.code.array[2]
      section <- course.code.array[3]
      
      q1 <- c(q1,
              evals[which(
                evals$SUBJECT.CODE == subject.code &
                  evals$COURSE.. == course &
                  evals$SECTION.. == section
              ), c.Q1])
      q2 <- c(q2,
              evals[which(
                evals$SUBJECT.CODE == subject.code &
                  evals$COURSE.. == course &
                  evals$SECTION.. == section
              ), c.Q2])
      q3 <- c(q3,
              evals[which(
                evals$SUBJECT.CODE == subject.code &
                  evals$COURSE.. == course &
                  evals$SECTION.. == section
              ), c.Q3])
      q4 <- c(q4,
              evals[which(
                evals$SUBJECT.CODE == subject.code &
                  evals$COURSE.. == course &
                  evals$SECTION.. == section
              ), c.Q4])
      q5 <- c(q5,
              evals[which(
                evals$SUBJECT.CODE == subject.code &
                  evals$COURSE.. == course &
                  evals$SECTION.. == section
              ), c.Q5])
      q6 <- c(q6,
              evals[which(
                evals$SUBJECT.CODE == subject.code &
                  evals$COURSE.. == course &
                  evals$SECTION.. == section
              ), c.Q6])
      q7 <- c(q7,
              evals[which(
                evals$SUBJECT.CODE == subject.code &
                  evals$COURSE.. == course &
                  evals$SECTION.. == section
              ), c.Q7])
      q8 <- c(q8,
              evals[which(
                evals$SUBJECT.CODE == subject.code &
                  evals$COURSE.. == course &
                  evals$SECTION.. == section
              ), c.Q8])
      
      current.comments <- evals[which(
        evals$SUBJECT.CODE == subject.code &
          evals$COURSE.. == course
        & evals$SECTION.. == section
      ), c.COM]
      
      if (all(nchar(current.comments) == 0))
        next()
      
      current.comments <-
        current.comments[which(nchar(current.comments) != 0)]
      current.comments <-
        paste(">>> ", current.comments, "\n", sep = "")
      
      Encoding(current.comments) <- "UTF-8"
      
      course.comments <-
        c(course.comments, current.comments)
    }
    
    course.respondents <-
      course.respondents[course.respondents != ""]
    
    q1 <- q1[q1 != ""]
    q2 <- q2[q2 != ""]
    q3 <- q3[q3 != ""]
    q4 <- q4[q4 != ""]
    q5 <- q5[q5 != ""]
    q6 <- q6[q6 != ""]
    q7 <- q7[q7 != ""]
    q8 <- q8[q8 != ""]
    
    
    
    q1.average <- evals.to.average(q1)
    q1.freqs <- evals.to.freqs(q1)
    q2.average <- evals.to.average(q2)
    q2.freqs <- evals.to.freqs(q2)
    q3.average <- evals.to.average(q3)
    q3.freqs <- evals.to.freqs(q3)
    q4.average <- evals.to.average(q4)
    q4.freqs <- evals.to.freqs(q4)
    q5.average <- evals.to.average(q5)
    q5.freqs <- evals.to.freqs(q5)
    q6.average <- evals.to.average(q6)
    q6.freqs <- evals.to.freqs(q6)
    q7.average <- evals.to.average(q7)
    q7.freqs <- evals.to.freqs(q7)
    q8.average <- evals.to.average(q8)
    q8.freqs <- evals.to.freqs(q8)
    
    q1.freqs <- rev(q1.freqs)
    q2.freqs <- rev(q2.freqs)
    q3.freqs <- rev(q3.freqs)
    q4.freqs <- rev(q4.freqs)
    q5.freqs <- rev(q5.freqs)
    q6.freqs <- rev(q6.freqs)
    q7.freqs <- rev(q7.freqs)
    q8.freqs <- rev(q8.freqs)
    
    q.average.cols <-
      rbind(
        q1.average,
        q2.average,
        q3.average,
        q4.average,
        q5.average,
        q6.average,
        q7.average,
        q8.average
      )
    
     q.freqs.cols <-
      rbind(
        q1.freqs,
        q2.freqs,
        q3.freqs,
        q4.freqs,
        q5.freqs,
        q6.freqs,
        q7.freqs,
        q8.freqs
      )
    
    
    course.respondents <- max(
      length(q1),
      length(q2),
      length(q3),
      length(q4),
      length(q5),
      length(q6),
      length(q7),
      length(q8),
      course.respondents
    )
    
    table <- rbind(
      "Description of course objectives and assignments" ,
      "Communication of ideas and information" ,
      "Expression of expectation for performance in this class" ,
      "Availability to assist students in or out of class" ,
      "Respect and concern for students" ,
      "Stimulation of interest in the course" ,
      "Facilitation of learning" ,
      "Overall assessment of course"
    )
    
    #    table.col.names <- c("Field", "Average", "")
    course.size <- course.size.from.sections(course.codes)
    response.rate <- percent(course.respondents / course.size, 2)
    response.rate.heading <-
      paste(course.respondents, "/", course.size, sep = "")
    response.rate.heading <-
      paste(response.rate.heading,
            " = ",
            response.rate,
            " response rate",
            sep = "")
    table.col.names <- c("Field", "Average", "Responses", "Poor", "Fair", "Good", "Very Good", "Excellent")
    
    q.responses <-
      rbind(
        length(q1),
        length(q2),
        length(q3),
        length(q4),
        length(q5),
        length(q6),
        length(q7),
        length(q8)
      )
    
    table <- cbind(table, q.average.cols, q.responses, q.freqs.cols)
    table <- rbind(table.col.names, table)
    colnames(table) <- table.col.names
    
    course.codes.heading <-
      unlist(sapply(sapply(course.codes, function(x)
        strsplit(x, "\\.")), function (y)
          paste(paste(y[1], y[2], sep = ""), sprintf("%03d", as.numeric(y[3])), sep = "."), simplify = F))
    course.codes.heading <- Reduce(paste, course.codes.heading)
    
    if (length(course.comments) != 0) {
      course.comments <-
        cbind(course.comments, "", "", "", "", "", "", "")
    } else {
      course.comments <- c("", "", "", "", "", "", "", "")
    }
    
    semester <- student.contacts[1, c.SEM]
    num.profs <- nrow(course.report)
    
    course.report <-
      rbind(
        c(course.name, "", "", "", "", "", "", ""),
        c(paste(semester, "Course Evaluations"), "", "", "", "", "", "", ""),
        c(course.codes.heading, "", "", "", "", "", "", ""),
        c(response.rate.heading, "", "", "", "", "", "", ""),
        "",
        table,
        "",
        report.col.names,
        course.report,
        "",
        c("Comments", "", "", "", "", "", "", ""),
        course.comments
      )
    
    firstStyle <-
      createStyle(fontSize = 18,
                  fontColour = "#333333",
                  halign = "center")
    secondStyle <-
      createStyle(fontSize = 14,
                  fontColour = "#333333",
                  halign = "center")
    tableStyle <-
      createStyle(
        fontSize = 12,
        fontColour = "#000000",
        border = "bottom",
        borderColour = "#000000",
        borderStyle = "thick",
        textDecoration = "bold",
        halign = "left"
      )
    columnStyle <-
      createStyle(
        fontSize = 12,
        fontColour = "#000000",
        border = "left",
        borderColour = "#333333",
        borderStyle = "thin",
        halign = "left"
      )
    rowStyle <-
      createStyle(
        fontSize = 12,
        fontColour = "#000000",
        border = "bottom",
        borderColour = "#333333",
        borderStyle = "thin",
        halign = "left"
      )
    textStyle <-
      createStyle(fontSize = 12,
                  fontColour = "#000000",
                  wrapText = TRUE)
    # page header style
    addStyle(
      wb,
      sheet.number,
      firstStyle,
      rows = 1,
      cols = 1:4,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      secondStyle,
      rows = 2,
      cols = 1:4,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      secondStyle,
      rows = 3,
      cols = 1:4,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      secondStyle,
      rows = 4,
      cols = 1:4,
      stack = TRUE
    )
    # table headings
    addStyle(
      wb,
      sheet.number,
      tableStyle,
      rows = 6,
      cols = 1:8,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      tableStyle,
      rows = 16,
      cols = 1:8,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      tableStyle,
      rows = (16 + num.profs + 2),
      cols = 1:8,
      stack = TRUE
    )
    # bottom borders for each field
    addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 7:14,
      cols = 1,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 7:14,
      cols = 2,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 7:14,
      cols = 3,
      stack = TRUE
    )
     addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 7:14,
      cols = 4,
      stack = TRUE
    )   
         addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 7:14,
      cols = 5,
      stack = TRUE
    )
             addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 7:14,
      cols = 6,
      stack = TRUE
    )
                 addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 7:14,
      cols = 7,
      stack = TRUE
    )
                     addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 7:14,
      cols = 8,
      stack = TRUE
    )
    # bottom borders for each professor
    addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 17:(16 + num.profs),
      cols = 1,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 17:(16 + num.profs),
      cols = 2,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 17:(16 + num.profs),
      cols = 3,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 17:(16 + num.profs),
      cols = 4,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 17:(16 + num.profs),
      cols = 5,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 17:(16 + num.profs),
      cols = 6,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 17:(16 + num.profs),
      cols = 7,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      rowStyle,
      rows = 17:(16 + num.profs),
      cols = 8,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      columnStyle,
      rows = 16:(16 + num.profs),
      cols = 2,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      columnStyle,
      rows = 16:(16 + num.profs),
      cols = 1,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      columnStyle,
      rows = 6:14,
      cols = 2,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      columnStyle,
      rows = 6:14,
      cols = 3,
      stack = TRUE
    )
    addStyle(
      wb,
      sheet.number,
      textStyle,
      rows = seq(16 + num.profs + 3 , nrow(course.report)),
      cols = 1,
      stack = TRUE,
    )
    
    # file.name <- course.name
    file.name <- paste(course.name, course.codes.heading)
    file.name <- make.names(file.name)
    file.name <- gsub("\\.", " ", file.name)
    file.name <-
      gsub("(?<=[\\s])\\s*|^\\s+|\\s+$", "", file.name, perl = TRUE)
    file.name <- paste(file.name, semester)
    
    writeData(wb,
              sheet = sheet.number,
              course.report,
              colNames = FALSE) # add the new worksheet to the workbook
    
    # resizes column widths to fit contents
    setColWidths(wb, sheet.number, cols = 1, widths = 70)
    # makes sure sheet fits on one printable page
    pageSetup(wb,
              sheet.number,
              fitToWidth = TRUE,
              fitToHeight = FALSE)
    
    mergeCells(wb, 1, cols = 1:4, rows = 1)
    mergeCells(wb, 1, cols = 1:4, rows = 2)
    mergeCells(wb, 1, cols = 1:4, rows = 3)
    mergeCells(wb, 1, cols = 1:4, rows = 4)
    
    auto_heights(
      wb,
      sheet.number,
      seq(16 + num.profs + 3, nrow(course.report)),
      base_height = 150,
      extra_height = 20
    )
    
    # writes a workbook containing all reports inputted
    saveWorkbook(wb, paste(file.name, ".xlsx", sep = ""), overwrite = TRUE)
  }




is.name.valid <- function(name) {
  if (length(name) == 0 ||
      is.na(name) || name == "") {
    return(FALSE)
  }
  
  return(TRUE)
}

get_worksheet_entries <- function(wb, sheet) {
  #  https://github.com/awalker89/openxlsx/pull/382/commits/7e9cfd09934ac61491b71e91785c415a22b3dc31
  # get worksheet data
  dat <- wb$worksheets[[sheet]]$sheet_data
  # get vector of entries
  val <- dat$v
  # get boolean vector of text entries
  typ <- (dat$t == 1) & !is.na(dat$t)
  # get text entry strings
  str <- unlist(wb$sharedStrings[as.integer(val)[typ] + 1])
  # remove xml tags
  str <- gsub("<.*?>", "", str)
  # write strings to vector of entries
  val[typ] <- str
  # return vector of entries
  val
}

auto_heights <-
  function(wb,
           sheet,
           selected,
           fontsize = NULL,
           factor = 1.0,
           base_height = 15,
           extra_height = 12) {
    #    https://github.com/awalker89/openxlsx/pull/382/commits/7e9cfd09934ac61491b71e91785c415a22b3dc31
    # get base font size
    if (is.null(fontsize)) {
      fontsize <- as.integer(openxlsx::getBaseFont(wb)$size$val)
    }
    # set factor to adjust font width (empiricially found scale factor 4 here)
    factor <- 4 * factor / fontsize
    # get worksheet data
    dat <- wb$worksheets[[sheet]]$sheet_data
    # get columns widths
    colWidths <- wb$colWidths[[sheet]]
    # select fixed (non-auto) and visible (non-hidden) columns only
    specified <-
      (colWidths != "auto") & (attr(colWidths, "hidden") == "0")
    # return default row heights if no column widths are fixed
    if (length(specified) == 0) {
      message("No column widths specified, returning default row heights.")
      cols <- integer(0)
      heights <- rep(base_height, length(selected))
      return(list(cols, heights))
    }
    # get fixed column indices
    cols <- as.integer(names(specified)[specified])
    # get fixed column widths
    widths <- as.numeric(colWidths[specified])
    # get all worksheet entries
    val <- get_worksheet_entries(wb, sheet)
    # compute optimal height per selected row
    heights <- sapply(selected, function(row) {
      # select entries in given row and columns of fixed widths
      index <- (dat$rows == row) & (dat$cols %in% cols)
      # remove line break characters
      chr <- gsub("\\r|\\n", "", val[index])
      # measure width of entry (in pixels)
      wdt <-
        strwidth(chr, unit = "in") * 20 / 1.43 # 20 px = 1.43 in
      # compute optimal height
      if (length(wdt) == 0) {
        base_height
      } else {
        base_height + extra_height * as.integer(max(wdt / widths * factor))
      }
    })
    # return list of indices of columns with fixed widths and optimal row heights
    list(cols, heights)
  }

print.sections <- function(course.title, course.sections) {
  print(paste(course.title, "sections:"))
  j <- 1
  for (section.index in 1:length(course.sections)) {
    print(paste("#", j, ": ", course.sections[section.index], sep = ""))
    j <- j + 1
  }
}

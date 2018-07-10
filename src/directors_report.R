evals <- evals[-c(1:3), ]
course.number <-
  paste(evals$Subject.Code, evals$Course.Number, sep = ".")
course.name <- evals$Course.Title
course.instructor <- evals$P1
course.communication.score <-
  evals$Q1_Communication.of.ideas.and.information
course.instructor.availability <-
  evals$Q1_Availability.to.assist.students.in.or.out.of.class
course.instructor.rating <- evals$Q2
course.satisfaction <- evals$Q1_Overall.assessment.of.course

table <-
  cbind(
    course.number,
    as.character(course.name),
    as.character(course.instructor),
    as.character(course.communication.score),
    as.character(course.instructor.availability),
    as.character(course.instructor.rating),
    as.character(course.satisfaction)
  )

colnames(table) <- c("Course #", "Course Title", "Course Director", "Communication of Ideas and Information", "Instructor Availability", "Instructor Rating", "Course Satisfaction") 

write.table(table,
            "Directors_Report.csv",
            sep = ",",
            row.names = FALSE,
            na = "")

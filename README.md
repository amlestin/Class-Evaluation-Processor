# Class-Evaluation-Processor
Qualtrics Class Evaluation Processor

Qualtrics provides an easy way of producing and distributing surveys, but the data collected can be exported to be interpreted in other ways than provided in their ecosystem. This project is a tool created to simplify the process of summarizing course evaluation data collected by a Qualtrics survey distributed at the end of each semester.

Follow the following process to prepare a survey compatible with this tool:

1. Add extra variables (embedded data) to your survey contacts CSV with the column names:
  PROF1
  PROF2
  PROF3
  PROF4
  PROF5
  PROF6
  PROF7
  PROF8
  PROF9
  PROF10
  PROF11
  PROF12
  PROF13
  PROF14
  PROF15
  TA1
  TA2
  TA3
  TA4
  
    The number of PROF and TA columns is arbitrary, but the setup above works for most courses.

2. Create question templates in Qualtrics to customize each question based on the embedded data in the contact record for each survey taker. This enable one survey template to be used for arbitrary course titles, semesters, names, etc. as are embedded for the survey taker's contact record or embedded in the unique link they clicked.

3. Distribute the survey then download the data collected as a CSV file.

4. Place the survey results in the same folders as the main.R file.

5. If you do not have the package scales installed, uncomment the line


``` 
install.packages("scales")
```




6. Drag the main.R file into an R session or run all of main.R in an IDE.


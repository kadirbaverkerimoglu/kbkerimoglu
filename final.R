library(tidyverse)
library(dplyr)
library(openxlsx)
library(stringi)

## OVERVIEW PREPROCESS 

HOP <- "C:/Users/PIT80/Downloads/10141346_2020-2021_Ozet_Tablolar.xlsx"

overview<- read.xlsx(
  xlsxFile = HOP, sheet = 2, startRow = 5,  skipEmptyRows = TRUE, fillMergedCells = TRUE, detectDates = FALSE)
overview <- overview[ -c(1,2,5,10,16) ]
overview <- overview[ -c(4) ]
overview1 <- overview %>% drop_na()
overview <- overview1

names(overview) <- c('LevelOfEducation', 'School', 'NumberOfStudentsTotal', 'NumberOfStudentsMale', 'NumberOfStudentsFemale', 'NumberOfTeachersTotal', 
                      'NumberofTeacherPermanentMale', 'NumberofTeacherContractMale', 'NumberofTeacherPermanentFemale', 'NumberofTeacherContractFemale', 'Classroom')

overview <- overview %>% filter(!row_number() %in% c(1, 37: 42))
overview1 <- overview %>%  mutate(School = stri_replace_all_fixed(School, "(1)", ""))
overview1 <- overview1 %>%  mutate(School = stri_replace_all_fixed(School, " ", ""))
overview1 <- overview1 %>%  mutate(LevelOfEducation = stri_replace_all_fixed(LevelOfEducation, " ", ""))
overview1 <- overview %>%  mutate(NumberOfTeachersTotal = stri_replace_all_fixed(NumberOfTeachersTotal, "(2)", ""))
overview1 <- overview1 %>%  mutate(NumberOfTeachersTotal = stri_replace_all_fixed(NumberOfTeachersTotal, "(1)", ""))
overview1 <- overview1 %>%  mutate(NumberOfTeachersTotal = stri_replace_all_fixed(NumberOfTeachersTotal, " ", ""))
overview1 <- overview1 %>%  mutate(NumberOfTeachersTotal = stri_replace_all_fixed(NumberOfTeachersTotal, "-", "0"))


overview1 <- overview1 %>%  mutate(NumberofTeacherPermanentMale = stri_replace_all_fixed(NumberofTeacherPermanentMale, "(2)", ""))
overview1 <- overview1 %>%  mutate(NumberofTeacherPermanentMale = stri_replace_all_fixed(NumberofTeacherPermanentMale, "(1)", ""))
overview1 <- overview1 %>%  mutate(NumberofTeacherPermanentMale = stri_replace_all_fixed(NumberofTeacherPermanentMale, " ", ""))
overview1 <- overview1 %>%  mutate(NumberofTeacherPermanentMale = stri_replace_all_fixed(NumberofTeacherPermanentMale, "-", "0"))

overview1 <- overview1 %>%  mutate(NumberofTeacherContractMale = stri_replace_all_fixed(NumberofTeacherContractMale, "(2)", ""))
overview1 <- overview1 %>%  mutate(NumberofTeacherContractMale = stri_replace_all_fixed(NumberofTeacherContractMale, "(1)", ""))
overview1 <- overview1 %>%  mutate(NumberofTeacherContractMale = stri_replace_all_fixed(NumberofTeacherContractMale, " ", ""))
overview1 <- overview1 %>%  mutate(NumberofTeacherContractMale = stri_replace_all_fixed(NumberofTeacherContractMale, "-", "0"))

overview1 <- overview1 %>%  mutate(NumberofTeacherPermanentFemale = stri_replace_all_fixed(NumberofTeacherPermanentFemale, "(2)", ""))
overview1 <- overview1 %>%  mutate(NumberofTeacherPermanentFemale = stri_replace_all_fixed(NumberofTeacherPermanentFemale, "(1)", ""))
overview1 <- overview1 %>%  mutate(NumberofTeacherPermanentFemale = stri_replace_all_fixed(NumberofTeacherPermanentFemale, " ", ""))
overview1 <- overview1 %>%  mutate(NumberofTeacherPermanentFemale = stri_replace_all_fixed(NumberofTeacherPermanentFemale, "-", "0"))

overview1 <- overview1 %>%  mutate(NumberofTeacherContractFemale = stri_replace_all_fixed(NumberofTeacherContractFemale, "(2)", ""))
overview1 <- overview1 %>%  mutate(NumberofTeacherContractFemale = stri_replace_all_fixed(NumberofTeacherContractFemale, "(1)", ""))
overview1 <- overview1 %>%  mutate(NumberofTeacherContractFemale = stri_replace_all_fixed(NumberofTeacherContractFemale, " ", ""))
overview1 <- overview1 %>%  mutate(NumberofTeacherContractFemale = stri_replace_all_fixed(NumberofTeacherContractFemale, "-", "0"))

overview1 <- overview1 %>%  mutate(Classroom = stri_replace_all_fixed(Classroom, "(2)", ""))
overview1 <- overview1 %>%  mutate(Classroom = stri_replace_all_fixed(Classroom, "(1)", ""))
overview1 <- overview1 %>%  mutate(Classroom = stri_replace_all_fixed(Classroom, " ", ""))
overview1 <- overview1 %>%  mutate(Classroom = stri_replace_all_fixed(Classroom, "-", "0"))

overview1 <- overview1 %>%  mutate(School = stri_replace_all_fixed(School, "(1)", ""))
overview1 <- overview1 %>%  mutate(School = stri_replace_all_fixed(School, " ", ""))


overview1 <- overview1 %>%  
  mutate(LevelOfEducation = stri_replace_all_fixed(LevelOfEducation, "                                                                    ", "-"))
overview1 <- overview1 %>%  
  mutate(LevelOfEducation = stri_replace_all_fixed(LevelOfEducation, "                  ", ""))

overview1$School <- as.numeric(overview1$School)
overview1$NumberOfStudentsTotal <- as.numeric(overview1$NumberOfStudentsTotal)
overview1$NumberOfStudentsMale <- as.numeric(overview1$NumberOfStudentsMale)
overview1$NumberOfStudentsFemale <- as.numeric(overview1$NumberOfStudentsFemale)
overview1$NumberOfTeachersTotal <- as.numeric(overview1$NumberOfTeachersTotal)
overview1$NumberofTeacherPermanentMale <- as.numeric(overview1$NumberofTeacherPermanentMale)
overview1$NumberofTeacherPermanentFemale <- as.numeric(overview1$NumberofTeacherPermanentFemale)
overview1$NumberofTeacherContractFemale <- as.numeric(overview1$NumberofTeacherContractFemale)
overview1$NumberofTeacherContractMale <- as.numeric(overview1$NumberofTeacherContractMale)
overview1$Classroom <- as.numeric(overview1$Classroom)

overview <- overview1

## PREPRIMARY PREPROCESS

preprimary<- read.xlsx(
  xlsxFile = HOP, sheet = 3, startRow = 6,  skipEmptyRows = TRUE, fillMergedCells = TRUE, detectDates = FALSE)
preprimary <- preprimary[ -c(1:4,7,11,15,17) ]

names(preprimary) <- c('TypeOfSchool', 'NumberofSchools', 'NumberOfStudentsTotal', 'NumberOfStudentsMale', 'NumberOfStudentsFemale', 'NumberOfTeachersTotal', 
                                     'NumberofTeachersMale', 'NumberofTeachersFemale', 'NumberofClassroom')

preprimary <- preprimary %>% filter(!row_number() %in% c(1, 34: 36))
preprimary1 <- preprimary %>% drop_na()

preprimary1[preprimary1 == " -"] <- 0
preprimary <- preprimary1

preprimary$Level <- 'PREPRIMARY'

## Numeric Column Names

num_col_names <- c( 'NumberofSchools', 'NumberOfStudentsTotal', 'NumberOfStudentsMale', 'NumberOfStudentsFemale', 'NumberOfTeachersTotal', 
                   'NumberofTeachersMale', 'NumberofTeachersFemale', 'NumberofClassroom')


preprimary[num_col_names] <- sapply(preprimary[num_col_names],as.numeric)

## PRIMARY PREPROCESS

primary<- read.xlsx(
  xlsxFile = HOP, sheet = 4, startRow = 6,  skipEmptyRows = TRUE, fillMergedCells = TRUE, detectDates = FALSE)

primary <- primary[ -c(1,2,5,9,13,15) ]

primary[primary == " -"]<- 0
primary <- primary %>% filter(!row_number() %in% c(1))

names(primary) <- c('TypeOfSchool', 'NumberofSchools', 'NumberOfStudentsTotal', 'NumberOfStudentsMale', 'NumberOfStudentsFemale', 'NumberOfTeachersTotal', 
                       'NumberofTeachersMale', 'NumberofTeachersFemale', 'NumberofClassroom')

primary[num_col_names] <- sapply(primary[num_col_names],as.numeric)

primary$Level <- 'PRIMARY'

## LOWER SECONDARY SCHOOL

secondary_lower<- read.xlsx(
  xlsxFile = HOP, sheet = 5, startRow = 5,  skipEmptyRows = TRUE, fillMergedCells = TRUE, detectDates = FALSE)

secondary_lower <- secondary_lower[ -c(1,2,5,9,13) ]

secondary_lower[secondary_lower == " -"]<- 0
secondary_lower <- secondary_lower %>% filter(!row_number() %in% c(24))

names(secondary_lower) <- c('TypeOfSchool', 'NumberofSchools', 'NumberOfStudentsTotal', 'NumberOfStudentsMale', 'NumberOfStudentsFemale', 'NumberOfTeachersTotal', 
                    'NumberofTeachersMale', 'NumberofTeachersFemale', 'NumberofClassroom')
secondary_lower <- secondary_lower %>% filter(!row_number() %in% c(1))

secondary_lower[num_col_names] <- sapply(secondary_lower[num_col_names],as.numeric)

secondary_lower$Level <- 'LOWER SECONDARY'

## SECONDARY(VOCATIONAL) SCHOOL

secondary1<- read.xlsx(
  xlsxFile = HOP, sheet = 6, startRow = 5,  skipEmptyRows = TRUE, fillMergedCells = TRUE, detectDates = FALSE)

secondary1 <- secondary1[ -c(1,4,8,12) ]

secondary1 <- secondary1 %>% drop_na()

secondary1[secondary1 == " -"]<- 0

names(secondary1) <- c('TypeOfSchool', 'NumberofSchools', 'NumberOfStudentsTotal', 'NumberOfStudentsMale', 'NumberOfStudentsFemale', 'NumberOfTeachersTotal', 
                            'NumberofTeachersMale', 'NumberofTeachersFemale', 'NumberofClassroom')

secondary1 <- secondary1 %>% filter(!row_number() %in% c(1))

secondary1[num_col_names] <- sapply(secondary1[num_col_names],as.numeric)

secondary1$Level <- 'VOCATIONAL SECONDARY'

## SECONDARY2 SCHOOL

secondary2<- read.xlsx(
  xlsxFile = HOP, sheet = 7, startRow = 5,  skipEmptyRows = TRUE, fillMergedCells = TRUE, detectDates = FALSE)

secondary2 <- secondary2[ -c(1,2,5,9, 13, 15) ]

secondary2 <- secondary2 %>% drop_na()

secondary2 <- secondary2 %>% filter(!row_number() %in% c(56:59))

secondary2 <- secondary2 %>% filter(!row_number() %in% c(35:38))

secondary2[secondary2 == " -"]<- 0

names(secondary2) <- c('TypeOfSchool', 'NumberofSchools', 'NumberOfStudentsTotal', 'NumberOfStudentsMale', 'NumberOfStudentsFemale', 'NumberOfTeachersTotal', 
                       'NumberofTeachersMale', 'NumberofTeachersFemale', 'NumberofClassroom')


secondary2 <- secondary2 %>% filter(!row_number() %in% c(1))

secondary2[num_col_names] <- sapply(secondary2[num_col_names],as.numeric)

secondary2$Level <- 'SECONDARY'

## NON-FORMAL EDUCATION

non_formal<- read.xlsx(
  xlsxFile = HOP, sheet = 8, startRow = 5,  skipEmptyRows = TRUE, fillMergedCells = TRUE, detectDates = FALSE)

non_formal <- non_formal[ -c(2,4,8,12,14) ]

non_formal <- non_formal %>% drop_na()
non_formal <- non_formal %>% filter(!row_number() %in% c(17:21))

non_formal[non_formal == " -"]<- 0

names(non_formal) <- c('TypeOfSchool', 'NumberofSchools', 'NumberOfStudentsTotal', 'NumberOfStudentsMale', 'NumberOfStudentsFemale', 'NumberOfTeachersTotal', 
                       'NumberofTeachersMale', 'NumberofTeachersFemale', 'NumberofClassroom')
non_formal <- non_formal %>% filter(!row_number() %in% c(1))
non_formal <- non_formal %>% filter(!row_number() %in% c(1))

non_formal[num_col_names] <- sapply(non_formal[num_col_names],as.numeric)

non_formal$Level <- 'NON_FORMAL'

## FINAL DATAFRAME

final_dataframe <- bind_rows(preprimary, primary, secondary_lower, secondary1, secondary2, non_formal)

# Removing Unnecessary Ones

rm(overview1)
rm(preprimary1)
rm(non_formal)
rm(primary)
rm(secondary_lower)
rm(secondary1)
rm(secondary2)
rm(preprimary)


  
















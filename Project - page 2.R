library(utils)
library(stringr)
library(readxl)
library(tidyverse)
library(dplyr)
library(writexl)
library(pdftools)
library(openxlsx)

## add second sheet
Files <- list.files(Output_file)
File_path <- paste0(Output_file, Files)
for (n in 1:length(Files)) {
  text <- read_xlsx(File_path[n])
  Blank <- createWorkbook()
  addWorksheet(Blank, "Sheet1")
  addWorksheet(Blank, "Sheet2")
  writeData(Blank, sheet = "Sheet1", text)
  saveWorkbook(Blank, File_path[n], overwrite=TRUE)
}

#X is Data frame
#Y name of file
Export <- function(X, Y) {
  Data.df <- as.data.frame(X)
  write_xlsx(Data.df,paste0(Output_file, Y, ".xlsx"))
}

Source_file <- "I:/Actuarial/IfNSW/TMF/General Lines/Contributions/2021-22/Mock Run/MM/Input/BHI Downloads/BHI NSW Quarterly Reports/"
Output_file <- "I:/Actuarial/IfNSW/TMF/General Lines/Contributions/2021-22/Mock Run/MM/Input/R Output/Project/"

#X is file number
#Y is page of file
#Z is end of lines to be selected one or two characters
#XX is start of word to be selected

cleanp2 <- function(X, Y, Z) {
  File <- list.files (Source_file)[X]
  File_path <- paste0(Source_file, File)
  Quarter <- str_remove(File, "BHI_HQ_")
  Quarter <- str_remove(Quarter, ".pdf")
  Last_Quarter <- paste0((as.numeric(str_sub(Quarter, end=4))-1), str_sub(Quarter, start=-1))
  
  text <- pdftools::pdf_text(File_path[i])[Y]
  #cat(text)
  text <-strsplit(text," ")
  text<- lapply(text, function(z){ z[!is.na(z) & z != ""]})
  A<- unlist(text)
  
  
  #Into lines
  line <- " "
  Paper <- " "
  for (n in 1:length(A)) {
    if (str_detect(A[n], "\n") == TRUE) {
      if (str_sub(A[n], start=-1)=="n") {
        A[n] <- str_remove(A[n], "\n")
        line <- paste0(line, " ", A[n])
        Paper <- rbind(Paper, line)
        line <- " "
        n <- n+1
      } else {
        B<- unlist(str_split(A[n], "\n"))
        line <-paste0(line, " ", B[1])
        Paper <- rbind(Paper, line)
        line <- B[2]
        n <- n+1
      }
    } else {
      line <- paste0(line, " ", A[n])
      n <- n+1
    }
  }
  
  
  #Selected lines
  #change length of end characters to be detected
  Paper1 <- " "
  for (n in 1:length(Paper)) {
    if(str_sub(Paper[n], start=-2) %in% Z) {
      Paper1 <- rbind(Paper1, Paper[n])
      n <- n+1
    } else {
      n <- n+1
    }
  }
  
  
  Category <- " "
  Information <- " "
  Line <- c(1, 2, 3)
  for (n in 1:length(Paper1)) {
    word <- unlist(str_split(Paper1[n], " "))
    k <- length(word)
    if (str_detect(word[k], "points")==TRUE)  {
      Category <- " "
      for (l in 1:(k-5)) {
        Category <- paste0(Category, " ", word[l])
        l <- l +1
      }
      Line <- c(Category, word[k-4], word[k-3])
    } else if (str_detect(word[k], "mins")==TRUE) {
      Category <- " "
      for (l in 1:(k-6)) {
        Category <- paste0(Category, " ", word[l])
        l <- l +1
      }
      Line <- c(Category, word[k-5], word[k-3])
    } else if (str_detect(word[k], "mins")==TRUE) {
      Category <- " "
      for (l in 1:(k-6)) {
        Category <- paste0(Category, " ", word[l])
        l <- l +1
      }
      Line <- c(Category, word[k-5], word[k-3])
    } else if (str_detect(word[k], "days")==TRUE) {
      Category <- " "
      for (l in 1:(k-6)) {
        Category <- paste0(Category, " ", word[l])
        l <- l +1
      }
      Line <- c(Category, word[k-5], word[k-3])
    } else if (str_detect(word[k], "unchanged")==TRUE) {
      Category <- " "
      for (l in 1:(k-5)) {
        Category <- paste0(Category, " ", word[l])
        l <- l +1
      }
      Line <- c(Category, word[k-4], word[k-2])
    } else if (str_sub(word[k], start=-1) == "m") {
      Category <- " "
      for (l in 1:(k-3)) {
        Category <- paste0(Category, " ", word[l])
        l <- l +1
      }
      Line <- c(Category, word[k-2], word[k-1])
    } 
    else {
      Line <- c("", "", "")
    }
    n <- n+1
    Information <- rbind(Information, Line) 
  }
  
  colnames(Information) <- c("Variable", Quarter, Last_Quarter)
  
  Data.df <- as.data.frame(Information)
  text <- read_xlsx(paste0(Output_file, Quarter, ".xlsx"))
  Blank <- createWorkbook()
  addWorksheet(Blank, "Sheet1")
  addWorksheet(Blank, "Sheet2")
  writeData(Blank, sheet = "Sheet1", text)
  writeData(Blank, sheet="Sheet2", Data.df)
  saveWorkbook(Blank, paste0(Output_file, Quarter, ".xlsx"), overwrite=TRUE)
}

cleanp2(1, 11, Z)



X<-10
Y<-11
XX <- c(as.character(0:9), "*")
int <- c(as.character(0:9))
Z<- c("ts", "ns", "ed", "ys", "1m", "5m", "3m", "6m", "2m")

clean(1, 9, Z, int)



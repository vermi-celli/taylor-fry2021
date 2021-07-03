#install package if needed
library(utils)
library(readxl)
library(writexl)
library(stringr)
library(dplyr)
library(tidyverse)

Quarter <- 20204
Type <- "Elective"

Output_file <- paste0("I:Actuarial/IfNSW/TMF/General Lines/Contributions/2021-22/Mock Run/MM/Input/R Output/Project/LHD/", Type, "/")
Source_file <- paste0("I:/Actuarial/IfNSW/TMF/General Lines/Contributions/2021-22/Mock Run/MM/Input/BHI Downloads/BHI by LHD Quarterly Excel/", Type, "/")
list.files(Source_file)



for (k in 15:18) {
  File_path <- paste0(Source_file, list.files(Source_file[k]))
  File <- read_xlsx(File_path)
  Quarter <- str_sub(File_path, start=8, end=12)
}


Source_file <- paste0(Source_file, list.files(Source_file)[15:18])
File <- read_xlsx(Source_file)
File <- read_xlsx(Source_file, sheet=3)
read_xl
clean <- function(Keyword, Nsheet, Col, Data_Type) {
  Data <- read_xls(File, sheet = Nsheet)
  Label <- Data[1]
  Label<-as.vector(unlist(Label))
  Label[is.na(Label)] <- 0
  
  A<-1:19
  B<-1:19
  n<-1
  k<-1
  for (n in 1:length(Label)) {
    if ((str_detect(Label[n], Keyword) == TRUE)) {
      A[k] <- Label[n]
      B[k] <- Data[n,Col]
      k <- k+1
      n <- n+1
    } else if (Label[n]=="New South Wales") {
      A[k] <- Label[n]
      B[k] <- Data[n,Col]
      k <- k+1
      n <- n+1
    } else if (k==20) {
      break
    } else {
      n <- n+1
    }
  }
  Data <- cbind(as.character(A), str_remove(B, ","))
  colnames(Data) <- c("LHD", Data_Type)  
  return(Data)
}

Filechange <- function(Quarter, Type) {
  Source_file <- paste0("I:/Actuarial/IfNSW/TMF/General Lines/Contributions/2021-22/Mock Run/MM/Input/BHI Downloads/BHI by LHD Quarterly Excel/", Type, "/BHI_HQ_", Quarter, "_", Type, ".xls")
  return(Source_file)
}

File <- Filechange(20203, "Emergency")
Emergency <- function(A){
  Table <- clean("TOTAL", 2, 8, "ED Attendances")
  Table <- cbind(Table, clean("TOTAL", 2, 6, "Transferred within 30 minutes %")[,2])
  Table <- cbind(Table, clean("Total", 3, 2, "Start in timeframe: Triage 2 %")[,2])
  Table <- cbind(Table, clean("Total", 3, 4, "Start in timeframe: Triage 3 %")[,2])
  Table <- cbind(Table, clean("Total", 3, 6, "Start in timeframe: Triage 4 %")[,2])
  Table <- cbind(Table, clean("Total", 3, 8, "Start in timeframe: Triage 5 %")[,2])
  colnames(Table) <- c("LHD", "ED Presentations", "Transferred within 30 minutes %","Start in timeframe: Triage 2 %", "Start in timeframe: Triage 3 %", "Start in timeframe: Triage 4 %", "Start in timeframe: Triage 5 %" )
  return(Table)
}

Table1<-cbind(Table1, Table)


k <- 20124
File <- Filechange(k, "Elective")
#read_xlsx(File, sheet=2)
Stacked <- Elective(1)
Stacked <- cbind(Stacked, Elective(1))
Data.df <- as.data.frame(Stacked)
Type <- "Elective"
write_xlsx(Data.df, paste0(Output_file, "BHI_LHD_", Type, "_2012.xlsx"))



Admission <- function(A) {
  AA <- clean("Total", 1, 6, "Acute - sameday")
  B <- clean("Total", 1, 7, "Acute-Overnight")
  Total_Acute <- as.numeric(AA[,2])+as.numeric(B[,2])
  D <- clean("Total", 1, 8, "Total bed days")
  Table <- cbind(AA, B[,2], Total_Acute , D[,2])
  colnames(Table) <- c("LHD", "Acute - sameday", "Acute - overnight","Total acute episodes","Total bed days")
  return(Table)
}

Elective <- function(A) {
  AA <- clean("Total", 1, 2, "Elective performed")
  B <- clean("Total", 1, 4, "Elective Urgent performed")
  C <- clean("Total", 1, 5, "Elective semiurgent performed")
  D <- clean("Total", 2, 2, "Elective ontime")
  E <- clean("Total", 2, 4, "Elective Urgent ontime")
  G <- clean("Total", 2, 5, "Elective semiurgent ontime")
  H <- clean("Total", 3, 2, "Median Waiting time urgent")
  I <- clean("Total", 3, 5, "Median Waiting time semiurgent")
  Table <- cbind(AA, B[,2], C[,2] , D[,2], E[,2], G[,2], H[,2], I[,2])
  colnames(Table) <- c("LHD", "Elective performed", "Elective Urgent performed", "Elective Semiurgent performed","Elective ontime","Elective urgent ontime","Elective semiurgent ontime", "Median Waiting time urgent", "Median Waiting time semiurgent")
  return(Table)
}




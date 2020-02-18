# Load Excel data into R

list.of.packages <- c("xts","tidyverse",
    "timetk","readxl")
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages)
suppressMessages(library(xts)) 
suppressMessages(library(tidyverse))
suppressMessages(library(timetk))
suppressMessages(library(readxl))

DM_By_Year <-function(cell="C2:C14"){
    read_excel_allsheets <- function(filename, tibble = FALSE) {
    sheets <- readxl::excel_sheets(filename)
    x <- lapply(sheets, function(X) readxl::read_excel(filename, sheet = X,range = cell ))
    if(!tibble) x <- lapply(x, as.data.frame)
    names(x) <- sheets
    x
    }
    DM <- read_excel_allsheets("Details_Monthly.xls")
    DM <- as.data.frame(DM)
    DM <- DM[c(6:1)]
    DM2 <- data.frame(NULL)
    for (i in 1:6){
        DM2 <- c(DM2,DM[,i])
    }
    DM_data <<- as.vector(unlist(DM2))
}

DM_By_Year("B2:B14")
DM_data

data_A <- ts(DM_data, 
             start=c(2012,1), 
             end = c(2017,12),
             frequency=12)
# saveRDS(data_A,"Total_A_M.Rds")



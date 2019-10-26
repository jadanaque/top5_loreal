# Helpers for reading L'Oréal Data into R

library(tidyverse)
library(readxl)
library(lubridate)
library(magrittr)
library(data.table)

# Create function to read last dataset in specified directory (do not include file name)
# If a pattern is supplied, it will look for that specific pattern first.

read_loreal <- function(directory = "data", pattern = NULL){
    
    # Loading data
    files <- list.files(directory, pattern, full.names = T, recursive = T)
    
    filePath <- files %>%
        file.mtime %>%
        which.max %>%
        files[.]
    
    # Choose method for data reading
    if(grepl("\\.csv$", filePath)){
        
        myDF <- fread(filePath, sep = ",", colClasses = "character", check.names = TRUE) %>% tbl_df()
        
    } else if(grepl("\\.xlsx$", filePath)){
        
        data_types <- rep("text", ncol(read_excel(filePath)))
        
        myDF <- read_excel(filePath, sheet=1, col_types = data_types)
        
    } else {
        stop("I only accept CSV or Excel files")
    }
    
    
    # Data Cleaning
    names(myDF) <- make.names(names(myDF), unique = T)
    names(myDF) <- iconv(names(myDF), to='ASCII//TRANSLIT')
    names(myDF) <- gsub("[^[:alnum:]]", "", names(myDF))
    
    myDF
    
}

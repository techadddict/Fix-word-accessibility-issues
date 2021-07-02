library(RDCOMClient)
library(here)
library(tidyverse)

my_local_path <- paste0(getwd(),"/Factsheets") #CHANGE THIS TO YOUR LOCAL PATH BEFORE ACCESSIBILITY STEP!!!!



all_files <- list.files(here(my_local_path), pattern = ".docx")
#all_files<-all_files[2:16]


for (i in seq_len(length(all_files))) {
  
  
      path<-paste0(getwd(),paste0("/Factsheets/",all_files[i]))
      file_path <- normalizePath(path)
      wordApp  <- COMCreate("Word.Application")  # create a new instance of a registered COM server class
      
      file <- wordApp[["Documents"]]$Open(file_path) #opens your docx in wordApp
      
      
      table_num <-as.numeric(file[["Tables"]]$Count())
      
      for (t in 1: file[["Tables"]]$Count()) { # used the identified amount of charts begins assigning a decorative tag 
        skip_to_next <- FALSE
        tryCatch({
          tble<- file[["Tables"]]$Item(t)
          #page_start<-  tble$Range()$Start()
         # page_end<-  tble$Range()$End()
          tble[["Style"]] <- "List Table 6 Colorful"
            # #specify table style here
          tble[["Shading"]][["BackgroundPatternColorIndex"]] = 8  #specify backround colour in table
          tble_num <- paste0("Table " ,t)
          tble[["Title"]] <- tble_num  #specify title for table 
          tble[["Descr"]] <- paste0("Regional trade in goods and services ",tble_num) #provide table description
          #specify columns and resize width
          
          #wdAutoFitWindow" "wdAutoFitContent" 
          
          tble[["AutoFitBehavior"]]<-"wdAutoFitWindow"
          
        
          
          
          tble[["Selection"]][["ParagraphFormat"]][["Alignment"]] = "wdAlignParagraphLeft"

          }, error = function(e) { skip_to_next <<- TRUE})

          if(skip_to_next) { next } }
        
          
          
        
          #chart and images
      for (k in 1:file[["InlineShapes"]]$Count()) { # used the identified amount of charts begins assigning a decorative tag 
          
              my_chart_1 <- file[["InlineShapes"]]$Item(k)
              my_chart_1[["Decorative"]] <- TRUE}


  
  
    
file$Close(SaveChanges =TRUE)
    #wordApp$Quit()
}
   
 

  
  
  
  



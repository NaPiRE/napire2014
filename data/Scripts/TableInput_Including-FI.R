library("XLConnect")

#Remember to export data with both option "Replace numeric codes with their labels (Excel and CSV exports only)" checked and not checked
# Remember to have the working directory in setwd("/napire/Publications/EMSE (2015)/Scripts")

## FUNCTIONS ##
import.xls <- function(xlsfile){
  
  wb <- loadWorkbook(xlsfile, create = FALSE)
  data <- readWorksheet(wb, sheet = "Export 1.1")
  return(data)
}

codify.na <- function(data.subset){
  
  data.subset <- as.matrix(data.subset)
  
  data.subset[data.subset==-77 | data.subset==-66 | data.subset==-99 | data.subset=="Please make a selection"]<-""
  data.subset <- as.data.frame(data.subset)
  
  return(data.subset)
  
}


codify.quotes <- function(data.subset){
  
  data.subset <- as.matrix(data.subset)
  
  data.subset[data.subset==quoted]<-1
  data.subset[data.subset==not.quoted]<-0
  
  data.subset <- as.data.frame(data.subset)
  
  return(data.subset)
  
}

## VARIABLES ##

not.quoted <- "not quoted"
quoted <- "quoted"

problems.ranks <- c("v_342", "v_344", "v_346", "v_348", "v_350")
causes <- c("v_391","v_392", "v_393","v_394", "v_395")
effects <- c("v_508", "v_509", "v_510", "v_511", "v_512")
failures <- c("v_381","v_382" , "v_383", "v_384", "v_385")
mitigations <- c("v_513","v_514","v_515","v_516","v_517")

problems.colnames <- paste("Problem_", seq(1:5), sep="")
causes.colnames   <-  paste("Cause_", seq(1:5), sep="")
effects.colnames  <-  paste("Effect_", seq(1:5), sep="")
failures.colnames <- paste("Failure_", seq(1:5), sep="")
mitigations.colnames <- paste("Mitigation_", seq(1:5), sep="")

surveys.path <- "../../../Surveys/2014/"
raw.data.folder <- "/Rawdata Exports/Export_"
file.extension.excel <- ".xls"

## -- ##


## INPUT ##   The working directory should be napire/Publications/EMSE (2015)/Scripts

codebook <- read.csv2("../../IST (2015)/Data analysis/Antonio/operational-codebooks/operational-codebook.csv")

countries.names <- c("Austria", "Brazil", "Canada", "Estonia", "Finland", "Germany", "Ireland", "Norway" , "Sweden", "USA") 
countries.codes <- c("AT", "BR", "CA", "EE", "FI", "DE", "IE", "NO", "SE", "US") 


for(c in countries.names){ 

  code <- countries.codes[ which( countries.names %in% c) ]
  
  
  data.all <- import.xls( paste(surveys.path,c,raw.data.folder,code,file.extension.excel,sep="") )
  data <- data.all[ data.all$dispcode == "Beendet (31)" | data.all$dispcode == "Beendet nach Unterbrechung (32)" | 
                       data.all$dispcode == "Completed (31)"  |  data.all$dispcode == "Completed after break (32)", ]

      ## FIRST SHEET "Raw data" ##
      
    
      
      new.df <- paste(code,"-",data$lfdn,sep="") #subject.id
      new.df <- cbind( new.df, codify.na(data[ ,problems.ranks])) # top 5 problems 
      new.df <- cbind( new.df, codify.na(data[, causes])) #causes
      new.df <- cbind( new.df, codify.na(data[, effects])) #effects
      new.df <- cbind( new.df, codify.na(codify.quotes(data[, failures]))) #failures
      new.df <- cbind( new.df, codify.na(data[, mitigations])) #mitigations
      
      area.code <- rep("",nrow(new.df))
      country.code <- rep(code, nrow(new.df))
      
      new.df <- cbind( area.code, country.code, new.df )
      colnames(new.df) <- c("AreaCode", "CountryCode", "Subject", problems.colnames, causes.colnames, 
                            effects.colnames, failures.colnames, mitigations.colnames)
      
      res <- loadWorkbook(paste("../Data/DataForCoding_",code,".xlsx", sep=""), create = TRUE)
      createSheet(res, name="RAW_data")
      writeWorksheet(object = res, data = new.df, sheet = "RAW_data")
      
      
      
      ## SECOND SHEET "Coding" ##
      
      new.df$Subject <- as.factor(as.character(new.df$Subject))
      
      # BUILD THE STRUCTURE SECTION BY SECTION #  
      
      second.sheet <- data.frame(AreaCode=character( nrow(new.df) *5 ), CountryCode=character(nrow(new.df) *5), Subject=character(nrow(new.df) *5),Coded=character(nrow(new.df) *5), Validated=character(nrow(new.df) *5) )
      
      # section problems #
  
      if(code == "BR"){
        
        wb.br <- loadWorkbook("problems-br.xlsx", create = FALSE)
        ws.br <- readWorksheet(wb.br, sheet = "Sheet1")
        problems <-ws.br$Problems
      }else{ 
        problems <- codebook[ codebook$OPTION_ID == problems.ranks[1] , ]$VARNAME #take list of problems
        problems <- problems[-c(1)] #the first element of the vector is not a problem (see codebook) 
        problems <- as.character(problems)
      }
      
      for(i in seq(1 : length(problems) )) {
      
        second.sheet[, ncol(second.sheet) + 1 ] <- rep("",nrow(second.sheet))
        
      }
        
      colnames(second.sheet)[ ( ncol(second.sheet) - length(problems) + 1 ) : ncol(second.sheet)  ] <- problems
      
      
      second.sheet$Causes.Original.Statements <- rep("",nrow(second.sheet))
      second.sheet$Effects.Original.Statements <- rep("",nrow(second.sheet))
      second.sheet$Project.Failure.0.1 <- rep(0,nrow(second.sheet))
      second.sheet$Mitigation.Original.Statements <- rep("",nrow(second.sheet))
      
      
      # FILL THE STRUCTURE # 
      
      
      second.sheet$AreaCode <- as.character(second.sheet$AreaCode)
      second.sheet$CountryCode <- as.character(second.sheet$CountryCode)
      second.sheet$Subject <- as.character(second.sheet$Subject)
      
      counter <- 1
      
      #TODO minor: levels on subject id is not working
      for( j in seq(1:nrow(new.df))) {
       
        s <- as.character(new.df$Subject[j])
        
        print(s)  
       
        for(i in seq(1:5)) {
            
            #ACTHUNG : Area Code, Country: hard coded
            second.sheet[counter,]$AreaCode <- ""
            second.sheet[counter,]$CountryCode <- c(code) 
            second.sheet[counter,]$Subject <- s
            
            #section problems
            p <- new.df[ new.df$Subject == s , problems.colnames[i] ]
            second.sheet[counter, which( colnames(second.sheet) %in% p)] <- "YES"
          
            
            #section causes #
            c <- as.character(new.df[ new.df$Subject == s , causes.colnames[i] ]) 
            second.sheet[counter,]$Causes.Original.Statements <- c
            
            #section effects#
            e <- as.character(new.df[ new.df$Subject == s , effects.colnames[i] ] )
            second.sheet[counter,]$Effects.Original.Statements  <- e
            
            #section project failure#
            f <- as.character(new.df[ new.df$Subject == s , failures.colnames[i] ] ) 
            second.sheet[counter,]$Project.Failure.0.1  <- f
            
            #section mitigation#
            m <- as.character(new.df[ new.df$Subject == s , mitigations.colnames[i] ] )
            second.sheet[counter,]$Mitigation.Original.Statements  <- m
            
            counter <- counter + 1
        }
        
      }
      

      createSheet(res, name="coding")

      writeWorksheet(object = res, data = second.sheet, sheet = "coding")
            
      saveWorkbook(res)
      
      
}
  
## -- ##




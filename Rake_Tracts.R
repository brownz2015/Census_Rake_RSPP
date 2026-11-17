# Rake_Tracts.R
# Loops iterative proportional fitting (raking) over every census tract of interest,
# generating tract-specific weights for ACS PUMS person records. The weights are merged
# back onto the PUMS data for the PUMA they came from, producing one dataframe per PUMA
# (in PUMA_Master) with a scaled_weight_<tract GEOID> column for each tract. Summing those
# weights reproduces tract totals from ACS tables and supports custom cross-tabulations.
# See README.md for setup and the three spreadsheet templates in templates/.

library(openxlsx)
library(stats)
library(stringr)
library(tidycensus)
library(qualV)
#library(fuzzyjoin)
library(writexl)
library(utils)
library(readxl)
library(dplyr)
library(varhandle)
library(tibble)
library(anesrake)
library(survey)

#	Clears workspace of old variable names
rm(list = ls())

#	Lines 17-48 should be the only lines that you need to modify
#	Most importantly, update the filepath
filepath <- "templates/"  # relative path; see README

#	You might need to replace the API key with your own
census_api_key(Sys.getenv("CENSUS_API_KEY"))  # set your own key first: Sys.setenv(CENSUS_API_KEY = "...") or tidycensus::census_api_key("...", install = TRUE)
setwd(filepath)

#	You can customize the names of input and output files
binsheetname <- "Bin_Template_Final_FiveGroup"
pumvarname <- "PUMA_Vars_FiveGroup"
binvars <- "Conditionals_Final_FiveGroup"

#Helper Variables Used throughout
#	These variables just specify geographies
#	aCode is the PUMA number
# aCode <- c("02102","02101","02300","02400","02200","02000","00900")
# aCode <- c("02102","02101","02300","02400")
aCode <- c("02102")
geog <- "tract"
st <- "Michigan"
# cnty <- c("Kalamazoo","Cass","Berrien","Van Buren","Allegan","Barry","Calhoun","Branch","St. Joseph")
cnty <- c("Kalamazoo")
yr <- 2023
surv <- "acs1"
microdatasurv <- "acs5"


#For now this doesn't work, manually feed tracts by puma
# #	You can choose multiple Census tracts
# #	Just make sure they're in quotations in a list
# library(tigris)
# options(tigris_use_cache = TRUE)
# 
# # Step 1: Get PUMAs and tracts for the state
# mi_pumas <- pumas(state = "MI", year = 2021)
# mi_tracts <- tracts(state = "MI", year = 2021)
# 
# # Step 2: Filter PUMAs of interest
# selected_pumas <- mi_pumas %>% filter(PUMACE10 %in% aCode)
# 
# # Step 3: Spatial join to find which tracts intersect selected PUMAs
# library(sf)
# tracts_in_pumas <- st_join(mi_tracts, selected_pumas, left = FALSE)
# 
# #Then create the PUMA-Tract Container
# tracts_in_pumas$PUMACE10 <- as.character(tracts_in_pumas$PUMACE10)
# tract_list_by_puma <- split(tracts_in_pumas$GEOID, tracts_in_pumas$PUMACE10)

# tract_list_by_puma<-list("02102"=c("26077000100","26077000201","26077000202","26077000300","26077000500","26077000601","26077000900","26077001001","26077001002","26077001100","26077001200","26077001504","26077001508","26077001509","26077001510","26077001511","26077001601","26077001603","26077001604","26077001701","26077001702","26077001801","26077001802","26077001905","26077001906","26077001907","26077002002","26077002003","26077002004","26077002005","26077002102","26077002103","26077002104","26077003500","26077005501"), "02101"=c("26077001300","26077001501","26077001502","26077001503","26077001803","26077002201","26077002203","26077002204","26077002601","26077002701","26077002702","26077002801","26077002802","26077002901","26077002906","26077002907","26077002908","26077002909","26077002910","26077002911","26077003002","26077003005","26077003006","26077003007","26077003008","26077003302","26077003400","26077005502","26077006102","26077006103","26077006601","26077006701","26077006702"), "02300"=c("26027001000","26027001100","26027001200","26027001500","26027001600","26027001701","26027001702","26027001801","26027001802","26027001901","26027001902","26027002000","26027002100","26027002201","26027002202","26159010101","26159010102","26159010201","26159010202","26159010300","26159010400","26159010500","26159010600","26159010901","26159010902","26159011003","26159011004","26159011300","26159011400","26159011500","26159011601","26159011602","26159011801","26159011802","26159011900","26159012001","26159012002"),"02400"=c("26021000300","26021000400","26021000500","26021000600","26021000700","26021000801","26021000802","26021000900","26021001000","26021001100","26021001300","26021001400","26021001500","26021001600","26021001700","26021001800","26021001900","26021002000","26021002100","26021002200","26021002300","26021002400","26021002500","26021010100","26021010201","26021010202","26021010300","26021010400","26021010500","26021010600","26021011000","26021011101","26021011102","26021011201","26021011202","26021011301","26021011302","26021011400","26021011500","26021011600","26021020100","26021020200","26021020300","26021020400","26021020500","26021020600","26021020700","26021020900","26021021000","26021021100","26021021200","26021021300","26021021400"))

tract_list_by_puma<-list("02102"=c("26077000100","26077000201","26077000202","26077000300","26077000500","26077000601","26077000900","26077001001","26077001002","26077001100","26077001200","26077001504","26077001508","26077001509","26077001510","26077001511","26077001601","26077001603","26077001604","26077001701","26077001702","26077001801","26077001802","26077001905","26077001906","26077001907","26077002002","26077002003","26077002004","26077002005","26077002102","26077002103","26077002104","26077003500","26077005501"))


PUMA_Master <- list()  # a named list of dataframes, one per PUMA


#This is the start of the giant loop
#Cycles through PUMAs and then tracts within PUMAs, running a bazillion times
#Going to try to cache the PUMA data do it doesn't need to be downloaded 449 times
for (puma in names(tract_list_by_puma)) {
 
  #	Reads in user interface sheet
  bin <- read.xlsx(paste0(binsheetname,".xlsx"))
  pumavars <- read.xlsx(paste0(pumvarname,".xlsx"), colNames=FALSE)
  nVars <- ncol(bin)/3 
  uservars <- matrix(0)
  for(t in 1:nVars)
  {
    uservars[t] <- names(bin[(3 * t)-2])
  }
  
  #	Reads in user-inputted microdata variable names ***************************
  pumsvars <- pumavars$X1
  pumsvars <- c("PUMA","TYPEHUGQ",pumsvars)
  
  #	Obtains microdata from tidycensus package
  nmicrodata <- get_pums(
    variables = pumsvars,
    state = st,
    survey = microdatasurv,
    year = yr)
  
  #	Filters out data from other PUMAs	
  npuma <- nmicrodata[ which(nmicrodata$PUMA == puma), ]
  
  # AHHHHHH WE NEED TO DROP GROUP QUARTERS OBSERVATIONS BEFORE THE DATA IS READIED
  npuma <- npuma[ which(npuma$TYPEHUGQ=="1"), ]
  
  contbox <- list()
  cleaned <- data.frame(matrix(0,ncol = nVars,nrow = nrow(npuma)))
  target <- rep(NA,nVars)
  
  ztest <- read.xlsx(paste0(binvars,".xlsx"), sheet = 1)
  automation2 <- read.xlsx(paste0(binvars,".xlsx"), sheet = 2)
  automation3 <- read.xlsx(paste0(binvars,".xlsx"), sheet = 3)
  automation4 <- read.xlsx(paste0(binvars,".xlsx"), sheet = 4)
  automation5 <- read.xlsx(paste0(binvars,".xlsx"), sheet = 5)
  # automation6 <- read.xlsx(paste0(binvars,".xlsx"), sheet = 6)
  zvars <- ncol(ztest)/2
  zvars2 <- ncol(automation2)/5
  zvars3 <- ncol(automation3)/5
  zvars4 <- ncol(automation4)/8
  zvars5 <- ncol(automation5)/4
  # zvars6 <- ncol(automation6)/3
  sodumb <- as.data.frame(npuma)
  
  #From here will need to adjust npuma to sodumb at the end of this section
  
  #1 Variable (Race and Hispanic)
  for(x in 1:zvars)
  {
    newVar <- as.character(names(ztest[2*x]))
    sodumb[[newVar]] <- NA
    for(i in 1:length(ztest[, (2*x)-1][!is.na(ztest[, (2*x)-1])]))
    {
      checker <- as.character(ztest[i,(2*x)-1])
      checker2 <- as.character(ztest[i,(2*x)])
      newCol <- as.character(names(ztest[2*x]))
      oldCol <- as.character(names(ztest[(2*x)-1]))
      sodumb[newCol][sodumb[[oldCol]]==checker,] <- checker2
    }
  }
  
  #2 Variable (Income to Poverty Ratio)
  for(x in 1:zvars2)
  {
    newVar <- as.character(names(automation2[5*x]))
    sodumb[[newVar]] <- NA
    for(i in 1:length(automation2[, (5*x)-1][!is.na(automation2[, (5*x)-1])]))
    {
      condition1 <- automation2[i,(5*x)-4]
      condition2 <- automation2[i,(5*x)-3]
      condition3 <- automation2[i,(5*x)-2]
      condition4 <- automation2[i,(5*x)-1]
      
      
      
      newValue <- as.character(automation2[i,(5*x)])
      newCol <- as.character(names(automation2[5*x]))
      oldCol <- as.character(names(automation2[(5*x)-4]))
      oldCol2 <- as.character(names(automation2[(5*x)-2]))
      sodumb[newCol][sodumb[[oldCol]]>=condition1 & sodumb[[oldCol]]<condition2 & sodumb[[oldCol2]]>=condition3 & sodumb[[oldCol2]]<condition4,] <- newValue
    }
    #Temp Fix for Raking to total Pop
    sodumb[newCol][sodumb[oldCol]<18 & is.na(sodumb[newCol]),]<-"a37"
    sodumb[newCol][sodumb[oldCol]>65 & is.na(sodumb[newCol]),]<-"a38"
  }
  
  #3 Variable (Education)
  for(x in 1:zvars3)
  {
    newVar <- as.character(names(automation3[5*x]))
    sodumb[[newVar]] <- NA
    for(i in 1:length(automation3[, (5*x)-1][!is.na(automation3[, (5*x)-1])]))
    {
      condition1 <- automation3[i,(5*x)-4]
      condition2 <- automation3[i,(5*x)-3]
      condition3 <- as.character(automation3[i,(5*x)-2])
      condition4 <- as.character(automation3[i,(5*x)-1])
      
      
      
      newValue <- as.character(automation3[i,(5*x)])
      newCol <- as.character(names(automation3[5*x]))
      oldCol <- as.character(names(automation3[(5*x)-4]))
      oldCol2 <- as.character(names(automation3[(5*x)-2]))
      oldCol3 <- as.character(names(automation3[(5*x)-1]))
      sodumb[newCol][sodumb[[oldCol]]>=condition1 & sodumb[[oldCol]]<condition2 & sodumb[[oldCol2]]==condition3 & sodumb[[oldCol3]]==condition4,] <- newValue
    }
    #Temp Fix for Raking to total Pop
    sodumb[newCol][sodumb[oldCol]<18 & sodumb[oldCol2]==1,]<-"e33"
    sodumb[newCol][sodumb[oldCol]>65 & sodumb[oldCol2]==1,]<-"e34"
    sodumb[newCol][sodumb[oldCol]<18 & sodumb[oldCol2]==2,]<-"e35"
    sodumb[newCol][sodumb[oldCol]>65 & sodumb[oldCol2]==2,]<-"e36"
  }
  
  #4 Variable (Work)
  for(x in 1:zvars4)
  {
    newVar <- as.character(names(automation4[8*x]))
    sodumb[[newVar]] <- NA
    for(i in 1:length(automation4[, (8*x)-1][!is.na(automation4[, (8*x)-1])]))
    {
      condition1 <- automation4[i,(8*x)-7]
      condition2 <- automation4[i,(8*x)-6]
      condition3 <- automation4[i,(8*x)-5]
      condition4 <- automation4[i,(8*x)-4]
      condition5 <- automation4[i,(8*x)-3]
      condition6 <- automation4[i,(8*x)-2]
      condition7 <- automation4[i,(8*x)-1]
      
      
      newValue <- as.character(automation4[i,(8*x)])
      newCol <- as.character(names(automation4[8*x]))
      oldCol <- as.character(names(automation4[(8*x)-7]))
      oldCol2 <- as.character(names(automation4[(8*x)-5]))
      oldCol3 <- as.character(names(automation4[(8*x)-4]))
      oldCol4 <- as.character(names(automation4[(8*x)-2]))
      
      sodumb[newCol][sodumb[[oldCol]]>=condition1 & sodumb[[oldCol]]<condition2 & sodumb[[oldCol2]]==condition3 & sodumb[[oldCol3]]>=condition4 & sodumb[[oldCol3]]<condition5 & sodumb[[oldCol4]]>=condition6 & sodumb[[oldCol4]]<condition7,] <- newValue
    }
    #Temp Fix for Raking to total Pop
    sodumb[newCol][sodumb[oldCol]<16 & sodumb[oldCol2]==1,]<-"d10"
    sodumb[newCol][sodumb[oldCol]>65 & sodumb[oldCol2]==1,]<-"d11"
    sodumb[newCol][sodumb[oldCol]<16 & sodumb[oldCol2]==2,]<-"d12"
    sodumb[newCol][sodumb[oldCol]>65 & sodumb[oldCol2]==2,]<-"d13"
    
  }
  
  #5 Variable (AgeSex)
  for(x in 1:zvars5)
  {
    
    newVar <- as.character(names(automation5[4*x]))
    sodumb[[newVar]] <- NA
    for(i in 1:length(automation5[, (4*x)-1][!is.na(automation5[, (4*x)-1])]))
    {
      condition1 <- automation5[i,(4*x)-3]
      condition2 <- automation5[i,(4*x)-2]
      condition3 <- automation5[i,(4*x)-1]
      
      newValue <- as.character(automation5[i,(4*x)])
      newCol <- as.character(names(automation5[4*x]))
      oldCol <- as.character(names(automation5[(4*x)-3]))
      oldCol2 <- as.character(names(automation5[(4*x)-1]))
      sodumb[newCol][sodumb[[oldCol]]>=condition1 & sodumb[[oldCol]]<condition2 & sodumb[[oldCol2]]==condition3,] <- newValue
    
    }
  }  
  
  #6 Variable (WorkType)
  # for(x in 1:zvars6)
  # {
  #   
  #   newVar <- as.character(names(automation6[3*x]))
  #   sodumb[[newVar]] <- NA
  #   for(i in 1:length(automation6[, (3*x)-1][!is.na(automation6[, (3*x)-1])]))
  #   {
  #     condition1 <- automation6[i,(3*x)-2]
  #     condition2 <- automation6[i,(3*x)-1]
  #     
  #     newValue <- as.character(automation6[i,(3*x)])
  #     newCol <- as.character(names(automation6[3*x]))
  #     oldCol <- as.character(names(automation6[(3*x)-2]))
  #     oldCol2 <- as.character(names(automation6[(3*x)-1]))
  #     sodumb[newCol][sodumb[[oldCol]]==as.character(condition1) & sodumb[[oldCol2]]==as.character(condition2),] <- newValue
  #     
  #     #Temp Fix for Raking to total Pop
  #     sodumb[newCol][(sodumb[oldCol2]=="b" | sodumb[oldCol2]=="9"),]<-"g11"
  # 
  #   }
  # } 
  
  #Reset Npuma
  npuma<-sodumb
  
  #Initializes the PUMA data in finalized form for later merging in tract weights
  if (!puma %in% names(PUMA_Master)) {
    PUMA_Master[[puma]] <- npuma  # your input PUMS data
  }
  
  #With npuma set, we can now run the tract loop over and over again. Gonna be a bumpy ride!!
  
  for (tract in tract_list_by_puma[[puma]]) {
    
    age_table <- get_acs(
      geography = geog,
      state = st,
      county = cnty,
      table = "B01001",
      year = yr,
      cache_table = FALSE
    )
    age_table <- age_table[age_table$GEOID == tract, ]
    
    under_16_M<-c("B01001_003","B01001_004","B01001_005")
    under_18_M<-c("B01001_003","B01001_004","B01001_005","B01001_006")
    under_16_F<-c("B01001_027","B01001_028","B01001_029")
    under_18_F<-c("B01001_027","B01001_028","B01001_029","B01001_030")
    over_65_M<-c("B01001_020","B01001_021","B01001_022","B01001_023","B01001_024","B01001_025")
    over_65_F<-c("B01001_044","B01001_045","B01001_046","B01001_047","B01001_048","B01001_049")
    
    under_16_M_total <- sum(age_table$estimate[age_table$variable %in% under_16_M], na.rm = TRUE)
    under_16_F_total <- sum(age_table$estimate[age_table$variable %in% under_16_F], na.rm = TRUE)
    under_18_M_total <- sum(age_table$estimate[age_table$variable %in% under_18_M], na.rm = TRUE)
    under_18_F_total <- sum(age_table$estimate[age_table$variable %in% under_18_F], na.rm = TRUE)
    under_18_total <- sum(age_table$estimate[age_table$variable %in% under_18_M | age_table$variable %in% under_18_F], na.rm = TRUE)
    over_65_M_total<-sum(age_table$estimate[age_table$variable %in% over_65_M], na.rm = TRUE)
    over_65_F_total<-sum(age_table$estimate[age_table$variable %in% over_65_F], na.rm = TRUE)
    over_65_total<-sum(age_table$estimate[age_table$variable %in% over_65_M | age_table$variable %in% over_65_F], na.rm = TRUE)
    total_pop<-sum(age_table$estimate[age_table$variable=="B01001_001"], na.rm = TRUE)
    
    
    for(i in 1:length(uservars))
    {
      #	Breaks spreadsheet input data down into individual lists of variables
      f <- (i * 3) - 2
      h <- (i * 3)
      q <- bin[f:h]
      
      #	Pulls ACS data for variable with API and tidycensus
      allcodes <- unlist(strsplit(names(q[2]),",")[[1]])
      acs <- matrix(NA,nrow=1,ncol=5)
      acstotal <- matrix(NA,nrow=1,ncol=5)
      code <- for(p in 1:length(allcodes))
      {				
        code <- allcodes[p]
        pacs <- get_acs(geography = geog,
                        state = st,
                        county = cnty,
                        table = code,
                        year = yr,
                        cache_table = FALSE)
        
        pacs<-pacs[pacs$GEOID==tract,]
        
        #	Checks if Census table has letter at end
        #	If the Census table code ends in a letter 
        #	(eg Hispanic householder table B25003I)
        #	Removes letter to obtain universe total, not just sub-populations
        ifelse(check.numeric(substr(code,2,nchar(code)))==TRUE,
               totcode <- paste(code,"_001", sep = ""),
               totcode <- paste(substr(code,1,nchar(code)-1),"_001", sep = ""))

        #	Obtain tract population totals from API
        pacstotal <- get_acs(geography = geog,
                             state = st,
                             county = cnty,
                             variable = totcode,
                             year = yr,
                             cache_table = FALSE)

        pacstotal<-pacstotal[pacstotal$GEOID==tract,]

        #	Combines ACS data for multiple variables, if requested by user
        if(p == 1)
        {
          acs <- pacs
          acstotal <- pacstotal
        }
        if(p != 1)
        {
          acs <- rbind(acs,pacs)
          acstotal <- rbind(acstotal,pacstotal)
        }
      }
      # #	Creates a data frame that sums universe totals by tract
      # acstotal <- aggregate(estimate ~ GEOID, data = acstotal,sum)
        #New Attempt at this: Recreate old acstotal using total tract pop
        acstotal <- data.frame(GEOID = tract, none = total_pop)
      
      #	Renames the dataframe containing universe population totals
      # names(acstotal) <- c("GEOID","none")
      
      #	Creates complete list of variables in Census table
      var <- unique(acs$variable)
      var <- var[!is.na(var)]
      
      #	Pulls user-generated variable codes
      bincode <- unique(q[1])
      bincode <- bincode[!is.na(bincode)]		
      
      #	Filters out Census disregarded by the user 
      filt <- acs
      filt[which(unlist(filt$variable) %in% bincode,arr.ind=TRUE),]
      
      #	Reformats data for the format desired by anesrake
      #	(Tracts as columns and variables in rows)
      geo <- unique(acs$GEOID)
      refo <- data.frame(row.names = geo)
      for(v in 1:length(bincode))
      {
        id <- bincode[v]
        h <- filt[ which(filt$variable == id), ]
        refo <- cbind(refo,h$estimate)
      }
      
      #	Attaches user-generated bin names
      if(is.na(q[1,3]) == FALSE)
      {
        namecode <- q[1]
      }  else
      {
        namecode <- q[2]
      }
      namecode <- namecode[!is.na(namecode)]
      names(refo) <- namecode
      
      #	Collapses columns that belong to same bin (ie 'ot' in "race")
      bin_names <- names(refo)
      
      # Aggregate values by bin name manually
      collapsed <- vapply(unique(bin_names), function(bin) {
        cols <- which(bin_names == bin)
        rowSums(refo[, cols, drop = FALSE], na.rm = TRUE)
      }, numeric(1))
      
      # Turn into 1-row data frame with clean bin names
      refo <- as.data.frame(t(collapsed))
      rownames(refo) <- geo  # Assign tract GEOID as row name
      
      #Current Manual Adding of Additional <18 >65 Columns
      if(i==1){
        
        refo$a37<-under_18_total
        refo$a38<-over_65_total
        
      }
      if(i==4){
        
        refo$d10<-under_16_M_total
        refo$d11<-over_65_M_total
        refo$d12<-under_16_F_total
        refo$d13<-over_65_F_total
      }
      if(i==5){
        
        refo$e33<-under_18_M_total
        refo$e34<-over_65_M_total
        refo$e35<-under_18_F_total
        refo$e36<-over_65_F_total
      }
      # if(i==7){
      #   
      #   refo$g11<-acstotal$none - as.numeric(rowSums(refo))
      # }
      
      #	Compare Census population totals with filtered values
      #	(Necessary for anesrake to operate properly)
      #	Creates a none category to ameliorate discrepancies
      discrepancy <- (acstotal[2] - rowSums(refo))
      refo <- cbind(refo,discrepancy)
      
      # #	Aggregate bins comprised of multiple variable codes
      # #	Determines if variable is categorical or continuous	
      # #	Categorical variables will have an empty third column
      # #	and text in the second column 
      # if(is.na(q[1,3]) == TRUE)
      # {
      #   refof <- t(refo)
      #   refof <- by(refof, INDICES=row.names(refof), FUN=colSums)
      #   refof <- as.data.frame(do.call(cbind,refof)) 
      #   refof <- refof[,order(colnames(refof))]
      # }
      # if(is.na(q[1,3]) == FALSE)
      # {
      #   refof <- refo
      # }
      
      #	Eliminates all but the tract(s) designated for analysis by user	
      selected <- refo[rownames(refo) %in% tract,]
      selected <- selected[ ,names(selected)!="none"]
      target[i] <- list(colSums(selected)/sum(selected))
      
    }
    
    #	Remove empty categories from target (tract-level)
    target <- lapply(target,function(x){x[x!=0]})
    names(target) <- uservars
    #	Drops unnused categories (levels) from survey data
    nsurvey <- droplevels(npuma)
    tsurvey <- na.omit(nsurvey)
    uniquservars <- unique(uservars)
    
    #	Removes categories not present in tract-level targets from microdata
    for(b in 1:length(uniquservars))
    {
      #	Creates objects: excluded = variables absent from tract
      #	uniquservars = unduplicated user-generated variable names 
      #	(in case user wanted to variables merged like "hous")
      #	fullvars = all levels found in microdata
      excluded <- c()
      a <- uniquservars[b]
      
      fullvars <- na.exclude(unique(sort(tsurvey[[a]])))
      
      #	Checks if level appears in target dataset
      for(g in 1:length(fullvars))
      {
        smooth <- paste0('\\b', fullvars[g], '\\b')
        checker <- grepl(smooth,names(target[[a]]))
        if(any(checker)==FALSE)
        {
          excluded[g] <- fullvars[g]
        }
      }
      excluded <- na.omit(excluded)
      
      #	Removes extraneous levels from microdata
      if(length(excluded) > 0)
      {
        for(h in 1:length(excluded))
        {
          #Okay, I think I get it. The original is being commented out and I'm going to try to redo it
          smooth2 <- paste0('\\b', excluded[h], '\\b')
          tsurvey <- tsurvey[!grepl(smooth2,tsurvey[[a]]),]
          #tsurvey <- tsurvey[!grepl(fullvars[[excluded[h]]],tsurvey[[a]]),]
        }
      }
      
      #	Reformatting (eg alphabetizing and dropping unused factors)
      #	Necessary for anesrake to run properly
      original_levels <- levels(tsurvey[[a]])
      tsurvey <- droplevels(tsurvey)
      tsurvey[[a]] <- factor(as.character(tsurvey[[a]]))
      new_levels <- levels(tsurvey[[a]])
      dropped_levels <- setdiff(original_levels, new_levels)
      if (length(dropped_levels) > 0) {
        cat("Dropped levels for variable", a, ":", paste(dropped_levels, collapse = ", "), "\n")
      }
    }
    
    #	Removes NAs from target object	
    for(r in 1:length(uniquservars))
    {
      if(any(is.na(names(target[[r]]))) == TRUE)
      {
        target[[r]] <- target[[r]][-which(is.na(names(target[[r]])))]
      }
    }
    #Here's what you need afterwards
    
    tester <- as.data.frame(tsurvey)
    tester$uniqueid <- 1:length(tester$SERIALNO)
    
    
    #	Finally raking!!!!!!
    raking <- anesrake(inputter = target,
                       dataframe = tester,
                       caseid = tester$uniqueid,
                       cap = 10,
                       verbose = TRUE,
                       choosemethod = "total",
                       convcrit = 0.01,
                       type = "pctlim",
                       pctlim = 0.05)
    
    caseweights <- data.frame(uniqueid=raking$caseid, weights=raking$weightvec)
    output <- merge(caseweights, tester, by="uniqueid")
    
    #This generate the pop total age 18-65 to rescale weights
    # eighteenUp<-c("B01001_007", "B01001_008", "B01001_009", "B01001_010", "B01001_011", "B01001_012","B01001_013","B01001_014","B01001_015","B01001_016","B01001_017","B01001_018","B01001_019","B01001_031", "B01001_032", "B01001_033", "B01001_034", "B01001_035", "B01001_036","B01001_037","B01001_038","B01001_039","B01001_040","B01001_041","B01001_042","B01001_043")
    # pop18_65 <- get_acs(
    #   geography = "tract",
    #   state = st,
    #   county = cnty,
    #   variables = eighteenUp,
    #   year = yr
    # ) %>%
    #   filter(GEOID == tract) %>%
    #   group_by(GEOID) %>%
    #   summarise(tract_pop_18_65 = sum(estimate))
    
    
    
    raw_weights <- as.numeric(raking$weight)
    
    # Scale
    scaled_weights <- raw_weights / sum(raw_weights) * acstotal$none
    
    weight_df <- tester %>%
      select(SERIALNO, SPORDER) %>%
      mutate(!!paste0("scaled_weight_", tract) := scaled_weights)
    
    # Merge into PUMA_Master[[puma]]
    PUMA_Master[[puma]] <- left_join(
      PUMA_Master[[puma]],
      weight_df,
      by = c("SERIALNO", "SPORDER")
    )
     
  }
  
  
  
  
  
  
  
   
}


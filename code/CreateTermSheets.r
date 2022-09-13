# process lists of keywords and relationships to create spreadsheets for each site
# John Porter, 2022
rm(list=ls())
setwd("c:/users/John/box sync/meetings/LTER_ASM_2022/SemanticAnalysis")
library(tidyverse)
library(openxlsx)

# read in termHierarchies

hierDf1<-read_csv("termHierarchies.csv",col_names=c("keyword","hierarchyString"))
hierDf2<-read_csv("termHierarchiesEDI.csv",col_names=c("keyword","hierarchyString"))
hierDf<-rbind(hierDf1,hierDf2)
hierDf<-hierDf[!duplicated(hierDf),]
hierCodesList<-str_split(hierDf$hierarchyString,fixed("|"))
hierDf$hierCode<-as.numeric(lapply(hierCodesList,'[',2))

# Now read in top level terms and merge them onto the df
topDf<-read_csv("TopLevelTerms.csv",col_names=c("topTerm","hierCode"))
hierDf<-merge(hierDf,topDf)

# read in the data packages and keyword data
datasetsDf1<-read_csv("resultSetKeywords.csv")
datasetsDf2<-read_csv("resultSetKeywordsEDI.csv")
datasetsDf<-rbind(datasetsDf1,datasetsDf2)
datasetsDf<-datasetsDf[!duplicated(datasetsDf),]
datasetsDf<-merge(datasetsDf,hierDf,all.x=T)
datasetsDf$topTerm<-ifelse(is.na(datasetsDf$topTerm),"No_category",datasetsDf$topTerm)
datasetsWideDf<-pivot_wider(datasetsDf,id_cols=c("packageid","title"),
                            names_from="topTerm",values_from="keyword",values_fill=list(""))
# change lists of terms into strings
listToString<-function(inList){
x<-lapply(inList,paste,sep="; ")
y<-lapply(x,paste,collapse='; ',sep="; ")
return(y)
}
datasetsWideDf$No_category<-listToString(datasetsWideDf$No_category)
datasetsWideDf$measurements<-listToString(datasetsWideDf$measurements)
datasetsWideDf$processes<-listToString(datasetsWideDf$processes)
datasetsWideDf$substances<-listToString(datasetsWideDf$substances)
datasetsWideDf$organisms<-listToString(datasetsWideDf$organisms)
datasetsWideDf$methods<-listToString(datasetsWideDf$methods)
datasetsWideDf$disciplines<-listToString(datasetsWideDf$disciplines)
datasetsWideDf$substrates<-listToString(datasetsWideDf$substrates)
datasetsWideDf$ecosystems<-listToString(datasetsWideDf$ecosystems)
datasetsWideDf$events<-listToString(datasetsWideDf$events)
datasetsWideDf$"organizational units"<-listToString(datasetsWideDf$"organizational units")

datasetsWideDf<-datasetsWideDf[order(datasetsWideDf$packageid),c("packageid","title","measurements","substances","processes","disciplines","organisms","ecosystems","organizational units","substrates","methods","events","No_category")]

openxlsx::write.xlsx(datasetsWideDf,"All_Keyword_Table.xlsx",asTable=T,colNames=T,
           firstActiveRow=2,firstActiveCol=3,
           colWidths=list(c(20,55,25,25,25,25,25,25,25,25,25,25,50)))
### After saving the spreadsheet, use EXCEL to change it to a smaller font (10 pt) and turn on 
### wrapping for all the columns. That can't be done here. 

# Now generate a workbook with separate sheets for each site
# extract the site name from the scope of the packageID
datasetsWideDf$site<- sub("knb-lter-","",datasetsWideDf$packageid,ignore.case=T)
datasetsWideDf$site<- gsub("[.][0-9]+[.][0-9]+","",datasetsWideDf$site)

# add some new columns for adding terms
datasetsWideDf$new_measurements<-""
datasetsWideDf$new_processes<-""
datasetsWideDf$new_substances<-""
datasetsWideDf$new_organisms<-""
datasetsWideDf$new_methods<-""
datasetsWideDf$new_disciplines<-""
datasetsWideDf$new_substrates<-""
datasetsWideDf$new_ecosystems<-""
datasetsWideDf$new_events<-""
datasetsWideDf$new_organizational_units<-""

# now put them in order
datasetsWideDf<-datasetsWideDf[,c("packageid","title",
                                  "ecosystems","new_ecosystems",
                                  "processes","new_processes",
                                  "disciplines","new_disciplines",
                                  "substances","new_substances",
                                  "measurements","new_measurements",
                                  "methods","new_methods",
                                  "organisms","new_organisms",
                                  "substrates","new_substrates",
                                  "events","new_events",
                                  "organizational units","new_organizational_units",
                                  "No_category","site")]

datasetsWideDf$site<-toupper(datasetsWideDf$site)

library(xlsx)
rm(wb)
wb=xlsx::createWorkbook(type='xlsx')
#myCellStyle <- xlsx::CellStyle(wb,alignment=xlsx::Alignment(wrapText=T),font=xlsx::Font(wb,heightInPoints=10))
myCellStyle <- xlsx::CellStyle(wb)+xlsx::Alignment(wrapText=T)+xlsx::Font(wb,heightInPoints=10)
dfColIndex <- rep(list(myCellStyle), dim(datasetsWideDf)[2]) 

names(dfColIndex) <- seq(1, dim(datasetsWideDf)[2], by = 1)
for (i in levels(as.factor(datasetsWideDf$site))){
  mysheet<-xlsx::createSheet(wb,sheetName=toupper(i))
  
  xlsx::addDataFrame(list(datasetsWideDf[datasetsWideDf$site == i,]),sheet=mysheet,
                     col.names=T, row.names=F,colStyle=dfColIndex)
  
}
sheets<-xlsx::getSheets(wb)
for (i in 1:31){
xlsx::setColumnWidth(sheets[[i]], colIndex=1:ncol(datasetsWideDf), colWidth=25)
xlsx::setColumnWidth(sheets[[i]], colIndex=2, colWidth=55)
xlsx::setColumnWidth(sheets[[i]], colIndex=23, colWidth=60)
}
xlsx::saveWorkbook(wb,"Keyword_Table_sheets.xlsx")

# See if we can add some data validation, we'll need to go back to openxlsx for this one
library(openxlsx)
wb1<-openxlsx::loadWorkbook("Keyword_Table_sheets.xlsx")
# read the valid values for each category
validDf<-openxlsx::read.xlsx("KeywordsByTopCategory.xlsx",sheet=1)
# add valids as a sheet
openxlsx::addWorksheet(wb1,"valids")
openxlsx::writeData(wb1,"valids",validDf)
# add validation for each category
for (i in 1:31){
# activeSheet(wb1)<-i
  # ecosystems
 openxlsx::dataValidation(wb1,sheet=i,cols=4,rows=2:700,showInputMsg=T,
                          showErrorMsg=F,type="list",value="'valids'!$B$2:$B$37")
  # processes
 openxlsx::dataValidation(wb1,sheet=i,cols=6,rows=2:700,showInputMsg=T,
                           showErrorMsg=F,type="list",value="'valids'!$H$2:$H$137")
 #disciplines
 openxlsx::dataValidation(wb1,sheet=i,cols=8,rows=2:700,showInputMsg=T,
                           showErrorMsg=F,type="list",value="'valids'!$A$2:$A$54")
 # substances
 openxlsx::dataValidation(wb1,sheet=i,cols=10,rows=2:700,showInputMsg=T,
                          showErrorMsg=F,type="list",value="'valids'!$I$2:$I$112")
 #measurements
 openxlsx::dataValidation(wb1,sheet=i,cols=12,rows=2:700,showInputMsg=T,
                          showErrorMsg=F,type="list",value="'valids'!$D$2:$D$171")
 # methods
 openxlsx::dataValidation(wb1,sheet=i,cols=14,rows=2:700,showInputMsg=T,
                          showErrorMsg=F,type="list",value="'valids'!$E$2:$E$35")
 # organisms
 openxlsx::dataValidation(wb1,sheet=i,cols=16,rows=2:700,showInputMsg=T,
                          showErrorMsg=F,type="list",value="'valids'!$F$2:$F$98")
 # substrates
 openxlsx::dataValidation(wb1,sheet=i,cols=18,rows=2:700,showInputMsg=T,
                          showErrorMsg=F,type="list",value="'valids'!$J$2:$J$29")
 # events
 openxlsx::dataValidation(wb1,sheet=i,cols=20,rows=2:700,showInputMsg=T,
                          showErrorMsg=F,type="list",value="'valids'!$C$2:$C$19")
 #organizational units
 openxlsx::dataValidation(wb1,sheet=i,cols=22,rows=2:700,showInputMsg=T,
                          showErrorMsg=F,type="list",value="'valids'!$G$2:$G$16")
 
}

openxlsx::saveWorkbook(wb1,"Keyword_Table_sheets1.xlsx",overwrite=T)
# you still need to edit the spreadsheet to change all the NA's to blanks

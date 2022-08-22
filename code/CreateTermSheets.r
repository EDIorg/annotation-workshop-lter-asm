# process lists of keywords and relationships to create spreadsheets for each site
# John Porter, 2022
rm(list=ls())
setwd("c:/users/John/box sync/meetings/LTER_ASM_2022/SemanticAnalysis")
library(tidyverse)
library(openxlsx)

# read in termHierarchies

hierDf<-read_csv("termHierarchies.csv",col_names=c("keyword","hierarchyString"))
hierCodesList<-str_split(hierDf$hierarchyString,fixed("|"))
hierDf$hierCode<-as.numeric(lapply(hierCodesList,'[',2))

# Now read in top level terms and merge them onto the df
topDf<-read_csv("TopLevelTerms.csv",col_names=c("topTerm","hierCode"))
hierDf<-merge(hierDf,topDf)

# read in the data packages and keyword data
datasetsDf<-read_csv("resultSetKeywords.csv")

datasetsDf<-merge(datasetsDf,hierDf,all.x=T)
datasetsDf$topTerm<-ifelse(is.na(datasetsDf$topTerm),"None",datasetsDf$topTerm)
datasetsWideDf<-pivot_wider(datasetsDf,id_cols=c("packageid","title"),
                            names_from="topTerm",values_from="keyword",values_fill=list(""))
# change lists of terms into strings
listToString<-function(inList){
x<-lapply(inList,paste,sep="; ")
y<-lapply(x,paste,collapse='; ',sep="; ")
return(y)
}
datasetsWideDf$None<-listToString(datasetsWideDf$None)
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

datasetsWideDf<-datasetsWideDf[order(datasetsWideDf$packageid),c("packageid","title","measurements","substances","processes","disciplines","organisms","ecosystems","organizational units","substrates","methods","events","None")]

write.xlsx(datasetsWideDf,"All_Keyword_Table.xlsx",asTable=T,colNames=T,
           firstActiveRow=2,firstActiveCol=3,
           colWidths=list(c(20,55,25,25,25,25,25,25,25,25,25,25,50)))
### After saving the spreadsheet, use EXCEL to change it to a smaller font (10 pt) and turn on 
### wrapping for all the columns. That can't be done here. 

           
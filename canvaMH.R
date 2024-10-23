# Load libraries ----
library(tidyverse)
library(readxl)
library(lubridate)
library(dplyr)
library(gdata)
library(stringr)
library(ggplot2)
library(xts)
library(tableHTML)
library(kableExtra)
library(tableone)
library(gdata)
library(scales)
# Load Current Data ----
load("../Data/currentData.Rdata")
library("googledrive")
library("googlesheets4")
drive_auth(email = "hqtorontooperations@gmail.com")
gs4_auth(email = "hqtorontooperations@gmail.com")
#drive_find()

##################################
dirPercentageRounding <- function(newNumber){
  if (round(newNumber,digits=2)*100<10)
  {
    newNumber <- round(newNumber,digits=3)*100
  }else{
    newNumber <- round(newNumber,digits=2)*100
  }
  return (newNumber)
}

dirRounding <- function(newNumber){
  if (round(newNumber,digits=2)<10)
  {
    newNumber <- round(newNumber,digits=1)
  }else{
    newNumber <- round(newNumber,digits=0)
  }
  return (newNumber)
}
##################################

MHReturnVisits <- AllVisitData %>% filter(visitType == "Return Visit (MH)" | visitType == "First Visit (MH)")
slide2a<-c("visits",n_distinct(MHReturnVisits))
slide2b<-c("individuals",n_distinct(MHReturnVisits$EmrID))
MHUniqueVisitors <- MHReturnVisits %>% select(EmrID)
MHUniqueVisitors <-MHUniqueVisitors %>% distinct()


returnVisitsData <- MHReturnVisits %>% left_join(Demographics,join_by(EmrID)) # %>% group_by(EmrID) %>% slice_max(EntryDate)
meanAgeRV <- dirRounding(mean(returnVisitsData$Age, na.rm = T))

#testingThis <- AllVisitData %>% left_join(demogInfo) %>% select(EmrID,Age,ReasonForVisit)
rvNumber <- length(returnVisitsData$EmrID)
#HIV status calculations
countHIVposRV<-sum(returnVisitsData$HIVpos, na.rm=T)
percHIVposRV<-paste(dirPercentageRounding(countHIVposRV/rvNumber),"\\%",sep="")
#Gender calculations
countCisManRV<-sum(returnVisitsData$genderCisM, na.rm = T)
percCisManRV<-paste(dirPercentageRounding(countCisManRV/rvNumber),"\\%",sep="")
countNonBinaryRV<-sum(returnVisitsData$genderNB, na.rm = T)
percNonBinaryRV<-paste(dirPercentageRounding(countNonBinaryRV/rvNumber),"\\%",sep="")
countOtherGenderRV<-sum(returnVisitsData$genderOther, na.rm = T)
percOtherGenderRV<-paste(dirPercentageRounding(countOtherGenderRV/rvNumber),"\\%",sep="")
countTransWomanRV<-sum(returnVisitsData$genderTransW, na.rm = T)
percTransWomanRV<-paste(dirPercentageRounding(countTransWomanRV/rvNumber),"\\%",sep="")
countTransManRV<-sum(returnVisitsData$genderTransM, na.rm = T)
percTransManRV<-paste(dirPercentageRounding(countTransManRV/rvNumber),"\\%",sep="")

############################################
currentMonth <- format(Sys.time(), "%B %Y")
currentMonth
## Calculation----
### ----
#MHUniqueVisitors <- MHReturnVisits %>% mutate(newMultiEthnicity=if_else(is.na(newMultiEthnicity),"Prefer not to answer",newMultiEthnicity))
MHUniqueVisitors <- returnVisitsData[!duplicated(returnVisitsData$EmrID),]
filteredVisit2<-MHUniqueVisitors

n_distinct(MHUniqueVisitors)


distinctRecords<-filteredVisit2 %>%
  select(EmrID) %>%
  distinct()

recordNumber<-dim(filteredVisit2)[1]

distinctRecordsDecember<-filteredVisit2 %>%
  select(EmrID) %>%
  distinct()

recordNumberDecember<-dim(distinctRecordsDecember)[1]

#   b)  mean/median age
# Creation of age has been moved to ReadInitialData.r
# Demographics<-Demographics %>%
#   mutate(Age=floor(as.numeric(difftime(EntryDate,`Birth date`,units="weeks"))/52.25))
meanAge <- dirRounding(mean(Demographics$Age,na.rm = T))
medianAge <- median(Demographics$Age,na.rm = T)


#   c)  count and percentages of RFV
#       i) create a barchart

RF1VTable<-FirstVisitInfo%>%
  group_by(ReasonForVisit) %>%
  count()
RFVTable<-AllVisitData%>%
  group_by(ReasonForVisit) %>%
  count()
#   d) count and percentages of ethnicity (include Census 2021 for reference)
ethnicity<-c("White - North American (e.g. Canadian, American)",
             "White - European (e.g. English, Italian, Portuguese, Russian)",
             "Latin American (e.g. Dominican Republic, Chilean, Salvadorian)",
             "Black - Caribbean (e.g. Barbadian, Jamaican)",
             "Black - African (e.g. Ghanaian, Kenyan, Somali)",
             "Black - North American (e.g. Canadian, American)",
             "Indian - Caribbean (e.g. Guyanese with origins in India)",
             "Asian - East (e.g. Chinese, Japanese, Korean)",
             "Asian - South East (e.g. Malaysian, Filipino, Vietnamese)",
             "Asian - South (e.g. Indian, Pakistani, Sri Lankan)",
             "Middle Eastern (e.g. Egyptian, Iranian, Lebanese)",
             "First Nations",
             "M\u00e9tis",
             "Inuit",
             "Indigenous/Aboriginal not included elsewhere",
             "Multiple ethnicities selected",
             "Prefer not to answer",
             "Other(s) (Please specify)","Do not know")
ethnicityLabel<-c("White - North American",
                  "White - European",
                  "Latin American",
                  "Black - Caribbean",
                  "Black - African",
                  "Black - North American",
                  "Indian - Caribbean",
                  "Asian - East",
                  "Asian - South East",
                  "Asian - South",
                  "Middle Eastern",
                  "First Nations",
                  "M\u00e9tis",
                  "Inuit",
                  "Indigenous/Aboriginal",
                  "Multiple ethnicities selected",
                  "Prefer not to answer",
                  "Other","Do not know")
ASO<-c(NA,NA,"CSSP","BlackCAP","BlackCAP","BlackCAP",NA,"ACAS","ACAS","ASAAP","ASAAP","2-Spirits","2-Spirits","2-Spirits","2-Spirits",NA,NA,NA,NA)
`Ethnicity Groups`<-c("White","White","Latinx","Black","Black","Black","Indo-Caribbean","Asian","Asian","Asian","Middle Eastern",
                      "Indigenous","Indigenous","Indigenous","Indigenous","Mixed race","Prefer not to answer","Other","Prefer not to answer")
groupCode<-c(1,1,2,3,3,3,9,4,4,4,5,6,6,6,6,7,8,10,8)
groupOrderCode<-c(18,18,15,16,16,16,14,17,17,17,13,12,12,12,12,11,20,19,20)
ethnicityOrder<-1:19
#`Toronto Census 2016`=c("50.2%",NA,"2.8%","8.5%",NA,NA,NA,"12.7%","7.0%","12.3%","1.1%","0.5%","0.2%",NA,NA,NA,NA,"3.3%",NA)
`Toronto Census 2016`=c("50.2%",NA,"2.8%","8.5%",NA,NA,"2.0%","12.7%","7.0%","12.3%","1.1%","0.5%","0.2%",NA,NA,"1.5%",NA,"1.3%",NA)
#`Toronto Census`=c("50.2%","50.2%","2.8%","8.5%","8.5%","8.5%",NA,"32.0%","32.0%","32.0%","1.1%","0.7%","0.7%","0.7%","0.7%",NA,NA,"3.3%",NA)
`Toronto Census`=c("50.2%","50.2%","2.8%","8.5%","8.5%","8.5%","2.0%","32.0%","32.0%","32.0%","1.1%","0.7%","0.7%","0.7%","0.7%","1.5%",NA,"1.3%",NA)

EthnicityReferenceTable<-as_tibble(data.frame(ethnicity=ethnicity,
                                              ethnicityLabel=ethnicityLabel,
                                              ethnicityOrder=ethnicityOrder,
                                              groupOrderCode=groupOrderCode,
                                              ASO=ASO,
                                              groupCode=groupCode,
                                              `Toronto Census 2016`=`Toronto Census 2016`))
MiniEthnicityReferenceTable<-as_tibble(data.frame(ethnicity=ethnicity,
                                                  ethnicityLabel=ethnicityLabel,
                                                  ethnicityOrder=ethnicityOrder,
                                                  groupOrderCode=groupOrderCode,
                                                  `Ethnicity Groups`=`Ethnicity Groups`,
                                                  groupCode=groupCode,
                                                  `Toronto Census`=`Toronto Census`))


EthnicityTable <- filteredVisit2 %>%
  group_by(multiEthnicity) %>%
  count() %>%
  mutate(percent = round(100*n/recordNumber,digits=1),
         percenttext=paste0(percent," %")) %>%
  left_join(EthnicityReferenceTable,by=c("multiEthnicity"="ethnicity")) %>%
  ungroup()

MiniEthnicityTable <- filteredVisit2 %>%
  group_by(multiEthnicity) %>%
  count() %>%
  mutate(percent = dirPercentageRounding(n/recordNumber),
         percenttext=paste0(percent," %")) %>%
  left_join(MiniEthnicityReferenceTable,by=c("multiEthnicity"="ethnicity")) %>%
  ungroup()
#!!

EthnicitySummary<-filteredVisit2 %>%
  left_join((EthnicityReferenceTable %>% select(-c(ethnicityOrder,`Toronto.Census.2016`,ASO))),
            by=c("multiEthnicity"="ethnicity")) %>%
  group_by(groupCode) %>% count() %>%
  mutate(subtotalPercent = round(100*n/recordNumber,digits=1),
         subtotalPercenttext=paste0(subtotalPercent," %")) %>%
  select(-n) %>% ungroup()

MiniEthnicitySummary<-filteredVisit2 %>%
  left_join((MiniEthnicityReferenceTable %>% select(-c(ethnicityOrder,`Toronto.Census`))),
            by=c("multiEthnicity"="ethnicity")) %>%
  group_by(`Ethnicity.Groups`) %>% count() %>%
  mutate(subtotalPercent = dirPercentageRounding(n/recordNumber),
         subtotalPercenttext=paste0(subtotalPercent," %")) %>%
  select(-n) %>% ungroup()
#!!!
EthnicityTable<-EthnicityTable %>%
  left_join(EthnicitySummary) %>%
  arrange(ethnicityOrder)


####
MiniNewEthnicitySummary<-filteredVisit2 %>%
  left_join((MiniEthnicityReferenceTable %>% select(-c(ethnicityOrder,`Toronto.Census`))),
            by=c("newMultiEthnicity"="ethnicity")) %>%
  group_by(`Ethnicity.Groups`) %>% count() %>%
  mutate(subtotalPercentNew = dirPercentageRounding(n/recordNumber),
         subtotalPercenttextNew=paste0(subtotalPercentNew," %")) %>%
  select(-n) %>% ungroup()
####


MiniEthnicitySummary2<-MiniEthnicitySummary %>%
  left_join(MiniEthnicityReferenceTable%>%select(Ethnicity.Groups,Toronto.Census,groupOrderCode) %>%
              distinct()) %>%
  arrange(groupOrderCode)


MiniEthnicitySummary2 <- MiniEthnicitySummary2 %>% left_join(MiniNewEthnicitySummary) %>% mutate(censusNumbers=as.numeric(gsub('%','',Toronto.Census)),
                                                                                                 Difference=subtotalPercentNew-censusNumbers)

LastMonthEthnicitySummary<-filteredVisit2 %>%
  left_join((MiniEthnicityReferenceTable %>% select(-c(ethnicityOrder,`Toronto.Census`))),
            by=c("multiEthnicity"="ethnicity")) %>%
  group_by(`Ethnicity.Groups`) %>% count() %>%
  mutate(subtotalPercent = dirPercentageRounding(n/recordNumberDecember),
         subtotalPercenttext=paste0(subtotalPercent," %")) %>%
  select(-n) %>% ungroup()


## Output ----
slide6 <- MiniEthnicitySummary2 %>% select(Ethnicity.Groups,subtotalPercenttext,Toronto.Census)
slide6 <- slide6 %>% rename(`HQ Percentage` = subtotalPercenttext,
                            `Toronto Census`=Toronto.Census,
                            ` ` = Ethnicity.Groups)

slide6 %>% View()

####
# CAP staff assignments/stepped care assessment
####
AllApptsDataFixed <- AllApptsData %>%
  select(ApptDate=EntryDate,LastApptUpdate=`Booking Date`,EmrID,ProviderID=`Provider ID`,ProviderName=Provider,
  ApptDuration=Duration,ApptDetails=Details,UserInitials=Comments,ApptType=Late)
################################
View(MHReturnVisits %>% group_by(format(as.Date(EntryDate),"%Y-%m (%y %b)")) %>% count())
MHReturnVisits %>% group_by(MonthServed) %>% count() %>% View()

MHReferralDataDemog <- MHReferralData  %>% left_join(Demographics, join_by(EmrID))
MHReferralDataPos <- MHReferralDataDemog %>% filter(HIVpos==T) #%>% group_by(HIVpos) %>% count() %>% ungroup()
MHReferralDataNonCis <- MHReferralDataDemog %>% filter(genderCisM==F)
MHReferralDataNonCis <- MHReferralDataDemog %>% filter(genderCisW==F)

MHReferralDataPos %>% group_by(mdReferral) %>% count()
MHReferralDataNonCis %>% group_by(mdReferral) %>% count()
TDVisits <- AllVisitData %>% filter(EntryDate > "2022-10-31" & EntryDate < "2023-11-01")
TDVisits %>% group_by(EmrID) %>% slice_max(EntryDate)
TDReferral <- MHReferralDataDemog %>%  filter(Date > "2022-10-31" & Date < "2023-11-01")
TDReferral %>% group_by(EmrID) %>% slice_max(Date)
TDReferral$EmrID %>% n_distinct()
TDMHScreening <- MHScreeningData %>%  filter(EntryDate > "2022-10-31" & EntryDate < "2023-11-01")
TDMHScreening %>% group_by(clickedYes) %>% count()
TDMHScreening %>% group_by(SelfHarm_Suicide) %>% count()
LabResults %>% filter(EntryDate > "2023-10-11" & EntryDate < "2023-11-01" &TESTS=="LABHIV_Antigen" &results== " reactive")
TDMHScreening %>% group_by(clickedYes) %>% count()
TDNoMHSupport<-TDMHScreening %>% filter(clickedYes==T & Receiving_MH_Support!=T)

TDNoMHSupport <- TDNoMHSupport %>% left_join(TDReferral,join_by(EmrID))
View(TDNoMHSupport)
TDNoMHSupport <- TDNoMHSupport %>% group_by(`EmrID`) %>%  slice_max(`Status Date`) %>% ungroup()
TDNoMHSupport <- TDNoMHSupport %>% group_by(`Referral Status`)

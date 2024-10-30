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
  library(officer)
  library(magrittr) 

# Load Current Data ----
load("../../Data/currentData.Rdata")
#library("googledrive")
#library("googlesheets4")
#drive_auth(email = "hqtorontooperations@gmail.com")
#gs4_auth(email = "hqtorontooperations@gmail.com")
#drive_find()

#backupSheet <- gs4_create(name = paste("Canva", format(Sys.time(), "%Y-%m-%d")))
#currentSheet <- gs4_create("Canva Gilead")
#currentSheet <- drive_get("Canva 3")
#googleSheet<-currentSheet
#canva1 <- read_sheet(googleSheet,trim_ws = T)
#range_clear(googleSheet,sheet = "Sheet2")





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

## Slide 1 ----
currentMonth <- format(Sys.time(), "%B %Y")
currentMonth

# Define a function to add a slide and populate the content
add_custom_slide <- function() {
  
  # Open the PowerPoint presentation
  ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\Board-Report\\Template.pptx")
  
  # Add a slide to the PowerPoint presentation using layout placeholders
  ppt <- ppt %>%
    add_slide(layout = "Introduction Layout", master = "HQ Master Style Slide")
  
  # Format the date to today's date
  formatted_date <- format(Sys.Date(), "%B %Y")
  
  # Define text formatting for the date
  date_style <- fp_text(font.size = 51.6, font.family = "Garnett 1", bold = FALSE, color = "black")
  
  # Create a formatted text object using fpar() and ftext()
  formatted_text <- fpar(ftext(formatted_date, prop = date_style))
  
  # Add the formatted date to a placeholder for the date (assume there's a specific placeholder like 'Date Placeholder')
  ppt <- ph_with(ppt, value = formatted_text, 
                 location = ph_location_label(ph_label = "Date")) 

  # Save the presentation after adding the date to the slide
  print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\Board-Report\\CanvaTrial.pptx")
}

# Call the function to create the PowerPoint
add_custom_slide()



# Slide 2 ----

# layout = "Two text columns"
# Template for Slide 2

# Define a function to add a slide and populate the content
add_custom_slide <- function(ppt, title, text1, text2, date_label = Sys.Date()) {
  
  # Open the PowerPoint presentation
  ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\Board-Report\\CanvaTrial.pptx")
  
  ppt %>%
    add_slide(layout = "Two text columns", master = "HQ Master Style Slide") %>%
    ph_with(value = title, location = ph_location_label(ph_label = "BigTitle")) %>%
    ph_with(value = text1, location = ph_location_label(ph_label = "Text1")) %>%
    ph_with(value = text2, location = ph_location_label(ph_label = "Text2")) %>%
    ph_with(value = format(date_label, "%B %d, %Y"), location = ph_location_label(ph_label = "Date"))
  
  
  # Save the updated PowerPoint
  print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\Board-Report\\CanvaTrial.pptx")
  
  
}

# Example usage
ppt <- read_pptx() # Reading the ppt that I want to add texts
ppt <- add_custom_slide(ppt, 
                        title = "Sexual Health Key Messages", 
                        text1 = "Highest weekly average last week
178 sexual health visits / day
Highest number of daily service users on Monday July 8, 2024 at 220 users
New visits have stabilized.
Our testing is significantly less expensive than Public Health. ", 
                        text2 = "Successful  Code for Canada Hackathon
Idea 1:
Improve kiosk check-in process
Idea 2:
Automate mass sending of all negative results (80% of patients) ")


# Slide 3 ----
SHReturnVisits <- AllVisitData %>% filter((visitType == "Return Visit (SH)" | visitType == "First Visit (SH)"))
slide2a <- paste(n_distinct(SHReturnVisits), "visits")
slide2b <- paste(n_distinct(SHReturnVisits$EmrID), "individuals")

# Combine the two with a newline character between them
slide2_text <- paste(slide2a, slide2b, sep = "\n")

SHUniqueVisitors <- SHReturnVisits %>% select(EmrID)
SHUniqueVisitors <- SHUniqueVisitors %>% distinct()


#range_clear(googleSheet,sheet = "Sheet2")
df <- data.frame(visits=slide2a,individuals=slide2b)
#sheet_append(ss= googleSheet, data =df,sheet = "Sheet2")


# Function for the template Number of Visits
# Using Layout Number of Visits

# Define a function to add a slide and populate the content
add_Title_Text <- function(Text1, Text2 = Sys.Date()) {
  
  # Open the PowerPoint presentation
  ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\Board-Report\\CanvaTrial.pptx")  
  
  # Add slide and populate content
  ppt <- ppt %>%
    add_slide(layout = "Number of Visits", master = "HQ Master Style Slide") %>%
    ph_with(value = Text1, location = ph_location_label(ph_label = "Text1")) %>%
    ph_with(value = Text2, location = ph_location_label(ph_label = "Text2"))
    
    # Save the updated PowerPoint
    print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\Board-Report\\CanvaTrial.pptx") 
  
  # Return the PowerPoint object
  return(ppt)
}

# Example usage
ppt <- add_Title_Text(
  Text1 = paste("Cumulative data on sexual health visit from July 2022 to", format(Sys.Date(), "%B %Y")),
  Text2 = paste(slide2a, slide2b, sep = "\n")
)





# Slide 4 ----
SHReturnVisits %>% group_by(year(EntryDate)) %>% count()
totalMonthlyCount <- SHReturnVisits %>% group_by(format(as.Date(EntryDate), "%Y-%m (%y %b)")) %>% count()
totalMonthlyCount <- totalMonthlyCount %>% rename(" " = `format(as.Date(EntryDate), "%Y-%m (%y %b)")`)
totalMonthlyCount
monthlyCount <- SHReturnVisits %>% group_by(visitType,format(as.Date(EntryDate), "%Y-%m (%y %b)")) %>% count()
monthlyCount <- monthlyCount %>% rename(visitMonth = `format(as.Date(EntryDate), "%Y-%m (%y %b)")`)
monthlyCount <- monthlyCount %>% pivot_wider(names_from = visitType, values_from = n)
monthlyCount

#range_clear(googleSheet,sheet = "Sheet3")
#sheet_append(ss= googleSheet, data =totalMonthlyCount,sheet = "Sheet3")


# Create a plot using the totalMonthlyCount data
plot <- ggplot(totalMonthlyCount, aes(x = ` `, y = n)) +
  geom_col(fill = "purple") +
  labs(title = NULL, x = NULL, y = NULL) +  # Remove x and y labels
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

# Save the plot as an image (PNG)
plot_file <- "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\Board-Report\\thisPlot_transparent.png"
ggsave(filename = plot_file, plot = plot, width = 8, height = 5, bg = "transparent")  # Save with transparent background



# Define a function to add a slide and populate the content
add_Title_Graphic_Legend <- function(ppt, title, graphic_path, legend = Sys.Date()) {
  
  
  # Open the PowerPoint presentation
  ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\Board-Report\\CanvaTrial.pptx")  
  
  # Add a new slide and populate placeholders
  ppt <- ppt %>%
    add_slide(layout = "Title/Graphic/Legend", master = "HQ Master Style Slide") %>%
    
    # Add title text to the placeholder labeled "Title"
    ph_with(value = title, location = ph_location_label(ph_label = "Title")) %>%
    
    # Add Graphic to the placeholder labeled "Legend"
    ph_with(value = legend, location = ph_location_label(ph_label = "Legend")) %>%
    
    # Add the picture to the placeholder labeled "Picture"
    ph_with(value = external_img(graphic_path), location = ph_location_label(ph_label = "Graphicpicture"))
  
  
  # Save the updated PowerPoint
  print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\Board-Report\\CanvaTrial.pptx") 
  
  return(ppt)
}

# Call the function with the title, text, and picture
ppt <- add_Title_Graphic_Legend(
  ppt, 
  title = "First vs Return
 July 2023 - June 2024", 
  legend = "04,725 visits in 2022
27,025 visits in 2023
18,986 visits in 2024
",
  graphic_path = plot_file
)


# Slide 10 ----
currentMonth
Demographics$Age <- Demographics$Age %>% as.numeric()
meanAge <- (Demographics$Age %>% mean(na.rm = TRUE))
dirRounding(meanAge)
SHUniqueVisitors <- SHUniqueVisitors %>% left_join(Demographics)
SHUniqueCount <- n_distinct(SHReturnVisits$EmrID)
HivStatus <- SHUniqueVisitors %>% filter(HIVpos == TRUE) %>% group_by(HIVpos) %>% count()
hivPercentage <- paste(dirPercentageRounding(HivStatus$n / n_distinct(SHReturnVisits$EmrID)), "%", sep = "")
countCisManRV <- sum(SHUniqueVisitors$genderCisM, na.rm = TRUE)
cisManPercentage <- paste(dirPercentageRounding(countCisManRV / SHUniqueCount), "%", sep = "")
countTransManRV <- sum(SHUniqueVisitors$genderTransM, na.rm = TRUE)
transManPercentage <- paste(dirPercentageRounding(countTransManRV / SHUniqueCount), "%", sep = "")
countTransWomanRV <- sum(SHUniqueVisitors$genderTransW, na.rm = TRUE)
transWomanPercentage <- paste(dirPercentageRounding(countTransWomanRV / SHUniqueCount), "%", sep = "")
countNonBinaryRV <- sum(SHUniqueVisitors$genderNB, na.rm = TRUE)
nonBinaryPercentage <- paste(dirPercentageRounding(countNonBinaryRV / SHUniqueCount), "%", sep = "")
countOtherGenderRV <- sum(SHUniqueVisitors$genderOther, na.rm = TRUE)
otherGenderPercentage <- paste(dirPercentageRounding(countOtherGenderRV / SHUniqueCount), "%", sep = "")


# Combine results into a single text string for text1, including the average age
text1 <- paste("Average Age: ", dirRounding(meanAge))
text2 <- paste("HIV Positive: ", hivPercentage)
text3 <- paste("Cis Male: ", cisManPercentage)
text4 <- paste("Trans Male: ", transManPercentage)
text5 <- paste("Trans Woman: ", transWomanPercentage)
text6 <- paste("Non-Binary: ", nonBinaryPercentage)
text7 <- paste("Other Gender: ", otherGenderPercentage)

# Function for the layout Demographics

# Define a function to add a slide and populate the content
add_Title_Text <- function(ppt, title, text1, text2, text3, text4, text5, text6, text7) {
  
  # Create the dynamic title with a static start date and current date
  title <- paste("Demographics July 2022 to", format(Sys.Date(), "%B %Y"))  
  
  # Open the PowerPoint presentation
  ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\Board-Report\\CanvaTrial.pptx") 
  
  
  ppt %>%
    add_slide(layout = "Demographics", master = "HQ Master Style Slide") %>%
    ph_with(value = title, location = ph_location_label(ph_label = "title")) %>%
    ph_with(value = text1, location = ph_location_label(ph_label = "text1")) %>%
    ph_with(value = text2, location = ph_location_label(ph_label = "text2")) %>%
    ph_with(value = text3, location = ph_location_label(ph_label = "text3")) %>%
    ph_with(value = text4, location = ph_location_label(ph_label = "text4")) %>%
    ph_with(value = text5, location = ph_location_label(ph_label = "text5")) %>%
    ph_with(value = text6, location = ph_location_label(ph_label = "text6")) %>%
    ph_with(value = text7, location = ph_location_label(ph_label = "text7")) 
  
  # Save the updated PowerPoint
  print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\Board-Report\\CanvaTrial.pptx") 
}

ppt <- add_Title_Text(ppt, 
                      text1 = text1,
                      text2 = text2,
                      text3 = text3,
                      text4 = text4,
                      text5 = text5,
                      text6 = text6,
                      text7 = text7)


# Slide 5 ----
currentMonth


# Slide 6 ----
currentMonth
## Calculation----
### ----
DemographicsTemp <- Demographics

Demographics <- Demographics20240917
SHUniqueVisitors <- SHUniqueVisitors %>% mutate(newMultiEthnicity=if_else(is.na(newMultiEthnicity),"Prefer not to answer",newMultiEthnicity))
SHUniqueVisitors <- SHUniqueVisitors[!duplicated(SHUniqueVisitors$EmrID),]
filteredVisit2<-SHUniqueVisitors

n_distinct(SHUniqueVisitors)


distinctRecords<-EligiblePatients %>%
  select(EmrID) %>%
  distinct()

recordNumber<-dim(distinctRecords)[1]

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


EthnicityTable <- Demographics %>%
  group_by(multiEthnicity) %>%
  count() %>%
  mutate(percent = round(100*n/recordNumber,digits=1),
         percenttext=paste0(percent," %")) %>%
  left_join(EthnicityReferenceTable,by=c("multiEthnicity"="ethnicity")) %>%
  ungroup()

MiniEthnicityTable <- Demographics %>%
  group_by(multiEthnicity) %>%
  count() %>%
  mutate(percent = dirPercentageRounding(n/recordNumber),
         percenttext=paste0(percent," %")) %>%
  left_join(MiniEthnicityReferenceTable,by=c("multiEthnicity"="ethnicity")) %>%
  ungroup()
#!!

EthnicitySummary<-Demographics%>%
  left_join((EthnicityReferenceTable %>% select(-c(ethnicityOrder,`Toronto.Census.2016`,ASO))),
            by=c("multiEthnicity"="ethnicity")) %>%
  group_by(groupCode) %>% count() %>%
  mutate(subtotalPercent = round(100*n/recordNumber,digits=1),
         subtotalPercenttext=paste0(subtotalPercent," %")) %>%
  select(-n) %>% ungroup()

MiniEthnicitySummary<-Demographics%>%
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
MiniNewEthnicitySummary<-Demographics%>%
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

LastMonthEthnicitySummary<-filteredVisit2%>%
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


#range_clear(googleSheet,sheet = "Sheet6")
#sheet_append(ss= googleSheet, data =slide6,sheet = "Sheet6")



# Creating a bar plot for the ethnicity percentages
ethnicity_plot <- ggplot(MiniEthnicitySummary2, aes(x = `Ethnicity.Groups`)) +
  geom_bar(aes(y = subtotalPercentNew, fill = "HQ Toronto"), stat = "identity", position = "dodge", color = "black") +
  geom_bar(aes(y = censusNumbers, fill = "Toronto Census"), stat = "identity", position = "dodge", color = "black") +
  scale_fill_manual(values = c("HQ Toronto" = "orange", "Toronto Census" = "yellow")) +
  labs(title = "Ethnicity Distribution: HQ Toronto vs Toronto Census",
       x = "Ethnicity Groups", 
       y = "Percentage (%)", 
       fill = "Legend") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))


# Print the plot
print(ethnicity_plot)

# Save the plot as a PNG file
ggsave("ethnicity_distribution.png", plot = ethnicity_plot, width = 10, height = 6, dpi = 300)


# Function for the template Ethnicity
# Define a function to add a slide and populate the content
add_Title_Graphic <- function(ppt, title, image_path = "ethnicity_distribution.png") {
  
  # Open the PowerPoint presentation
  ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\Board-Report\\CanvaTrial.pptx")  
  
  # Add a new slide and populate placeholders
  ppt <- ppt %>%
    add_slide(layout = "Ethnicity", master = "HQ Master Style Slide") %>% 
    
    # Add title text to the placeholder labeled "Title"
    ph_with(value = title, location = ph_location_label(ph_label = "title")) %>% 
    
    # Add the saved image (ethnicity_distribution.png) to the placeholder labeled "Picture"
    ph_with(value = external_img(image_path), location = ph_location_label(ph_label = "Picture"))
  
  # Save the updated PowerPoint
  print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\Board-Report\\CanvaTrial.pptx")
}

# Call the function with the title and image path
ppt <- add_Title_Graphic(
  ppt, 
  title = "Ethnicity by individuals\nJuly 2022 to June 2024", 
  image_path = "ethnicity_distribution.png"
)



#Slide7

slide7 <- MiniEthnicitySummary2 %>% 
  select(Ethnicity.Groups, Difference) %>%
  rename(`Ethnicity Group` = Ethnicity.Groups) %>%
  arrange(desc(Difference))

# Create the bar plot using ggplot2 with custom colors
ethnicity_plot <- ggplot(slide7, aes(x = reorder(`Ethnicity Group`, Difference), y = Difference, fill = `Ethnicity Group`)) +
  geom_bar(stat = "identity") +
  coord_flip() +  # To flip the axes
  labs(title = "Difference Between Toronto 2020 Census and HQ",
       x = "Ethnicity Group", y = "Difference") +
  scale_fill_manual(values = rep(c("steelblue", "orange", "green"), length.out = nrow(slide7))) +  # Rotating colors
  theme_minimal()


# Print the plot
print(ethnicity_plot)


# Save the plot as an image
plot_path <- "ethnicity_plot.png"
ggsave(plot_path, plot = ethnicity_plot, width = 8, height = 6)


# Define a function to add a slide and populate the content
add_Title_Image <- function(ppt, title, image_path) {
  
  # Add a new slide with "Ethnicity2020" layout 
  ppt <- add_slide(ppt, layout = "Ethnicity2020", master = "HQ Master Style Slide") 
  
  # Add the plot image using the label "Imagem"
  ppt <- ph_with(ppt, value = external_img(image_path), location = ph_location_label(ph_label = "Imagem"))
  
  # Add the title
  ppt <- ph_with(ppt, value = title, location = ph_location_label(ph_label = "title"))
  
  # Return modified ppt object
  return(ppt)
}

# Load the existing PowerPoint presentation
ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\Board-Report\\CanvaTrial.pptx")

# Call the function with the title and image path
ppt <- add_Title_Image(
  ppt, 
  title = "Ethnicity (Difference between Toronto 2020 Census and HQ)", 
  image_path = "ethnicity_plot.png"
)

# Save the modified PowerPoint presentation
print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\Board-Report\\CanvaTrial.pptx")



# Slide 8 ----
## TODO: request calculation method ----



# Slide 9 ----
## Experience Slide

# Slide 10 ----
## TODO: Add batch email count ----
currentMonth
tableone::CreateTableOne(vars = c("overall","efficiency","welcoming","recommend"),data=surveyData)
surveyCount <- nrow(surveyData)
surveySMSCount<-smsData %>% filter(grepl("survey",Body,fixed = T)) %>% count()
paste(dirPercentageRounding((surveyCount/surveySMSCount)),"%", sep="")
surveySMSCount$n

((surveyData %>% filter(Date > "2023-09-30" & Date < "2023-11-01")) %>% filter(welcoming<3) %>% count()/surveyCount)*100

((surveyData %>% filter(Date > "2023-09-30" & Date < "2023-11-01")) %>% filter(welcoming==3) %>% count()/surveyCount)*100
((surveyData %>% filter(Date > "2023-09-30" & Date < "2023-11-01")) %>% filter(recommend ==3) %>% count()/surveyCount)*100

((surveyData %>% filter(Date > "2023-09-30" & Date < "2023-11-01")) %>% filter(recommend <3) %>% count()/surveyCount)*100
((surveyData %>% filter(Date > "2023-09-30" & Date < "2023-11-01")) %>% filter(efficiency <3) %>% count()/surveyCount)*100

((surveyData %>% filter(Date > "2023-09-30" & Date < "2023-11-01")) %>% filter(efficiency==3) %>% count()/surveyCount)*100
((surveyData %>% filter(Date > "2023-09-30" & Date < "2023-11-01")) %>% filter(overall==3) %>% count()/surveyCount)*100

((surveyData %>% filter(Date > "2023-09-30" & Date < "2023-11-01")) %>% filter(overall<3) %>% count()/surveyCount)*100

tableone::CreateTableOne(vars = c("overall","efficiency","welcoming","recommend"),data=((surveyData %>% filter(Date > "2023-09-30" & Date < "2023-11-01")) %>% filter(Date > "2023-09-30" & Date < "2023-11-01")))
surveyCount <- nrow((surveyData %>% filter(Date > "2023-09-30" & Date < "2023-11-01")))
surveySMSCount<-smsData %>% filter(EntryDate > "2023-09-30" & EntryDate < "2023-11-01" & grepl("survey",Body,fixed = T)) %>% count()
paste(dirPercentageRounding((surveyCount/surveySMSCount)),"%", sep="")
surveySMSCount$n

((surveyData %>% filter(Date > "2023-09-30" & Date < "2023-11-01")) %>% filter(welcoming<3) %>% count()/surveyCount)*100

((surveyData %>% filter(Date > "2023-09-30" & Date < "2023-11-01")) %>% filter(welcoming==3) %>% count()/surveyCount)*100
((surveyData %>% filter(Date > "2023-09-30" & Date < "2023-11-01")) %>% filter(recommend ==3) %>% count()/surveyCount)*100

((surveyData %>% filter(Date > "2023-09-30" & Date < "2023-11-01")) %>% filter(recommend <3) %>% count()/surveyCount)*100
((surveyData %>% filter(Date > "2023-09-30" & Date < "2023-11-01")) %>% filter(efficiency <3) %>% count()/surveyCount)*100

((surveyData %>% filter(Date > "2023-09-30" & Date < "2023-11-01")) %>% filter(efficiency==3) %>% count()/surveyCount)*100
((surveyData %>% filter(Date > "2023-09-30" & Date < "2023-11-01")) %>% filter(overall==3) %>% count()/surveyCount)*100

((surveyData %>% filter(Date > "2023-09-30" & Date < "2023-11-01")) %>% filter(overall<3) %>% count()/surveyCount)*100

tableone::CreateTableOne(vars = c("overall","efficiency","welcoming","recommend"),data=((surveyData ) ))
surveyCount <- nrow((surveyData ))
surveySMSCount<-smsData %>% filter(grepl("survey",Body,fixed = T)) %>% count()
paste(dirPercentageRounding((surveyCount/surveySMSCount)),"%", sep="")
surveySMSCount$n
((surveyData ) %>% filter(welcoming<3) %>% count()/surveyCount)*100
((surveyData ) %>% filter(welcoming==3) %>% count()/surveyCount)*100
((surveyData ) %>% filter(recommend ==3) %>% count()/surveyCount)*100
((surveyData ) %>% filter(recommend <3) %>% count()/surveyCount)*100
((surveyData ) %>% filter(efficiency <3) %>% count()/surveyCount)*100
((surveyData ) %>% filter(efficiency==3) %>% count()/surveyCount)*100
((surveyData ) %>% filter(overall==3) %>% count()/surveyCount)*100
((surveyData ) %>% filter(overall<3) %>% count()/surveyCount)*100


# Slide 11 ----
## Experience Slide

# Slide 12
                                          ## Wait times ----

# Slide 13 ----

Test_Results_Count <- nrow(LabResults) %>% as.character()

Treatments_Count <- nrow(SHAnPTxsData) %>% as.character()

paste1a <- Test_Results_Count
paste1b <- Treatments_Count

#Improving the Health of Our Community

# Function for the template Number of Test Results and Treatment

# Define a function to add a slide and populate the content
add_Title_Text <- function(Text1, Text2 = Sys.Date()) {
  
  # Open the PowerPoint presentation
  ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\Board-Report\\CanvaTrial.pptx")  
  
  # Add slide and populate content
  ppt <- ppt %>%
    add_slide(layout = "Test and Results", master = "HQ Master Style Slide") %>%
    ph_with(value = Text1, location = ph_location_label(ph_label = "Title")) %>%
    ph_with(value = Text2, location = ph_location_label(ph_label = "text1"))

# Save the updated PowerPoint
print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\Board-Report\\CanvaTrial.pptx") 

# Return the PowerPoint object
return(ppt)

}
# Example usage
ppt <- add_Title_Text(
  Text1 = "Improving the Health of Our Community",
  Text2 = paste(paste1a, "Test Results", paste1b, "Treatment:", sep = " ")
)
 
 # Slide 11 ----
 ## Experience Slide
 
 # Slide 12 ----
 ## Wait times ----
labMonthlyCount <- LabResults %>%
  filter(!(TESTS %in% c("LABCT_R", "LABCT_T", "LABCT_U"))) %>%
  group_by(Month = format(as.Date(EntryDate), "%Y-%m (%y %b)")) %>%
  count()

# Check the labMonthlyCount dataframe
print(labMonthlyCount)

# Assign the plot to ethnicity_plot using geom_bar
ethnicity_plot <- ggplot(labMonthlyCount, aes(x = Month, y = n)) +  # Set correct aesthetics
  geom_bar(stat = "identity", fill = "#FF8C00") +  # Bar plot with orange color
  labs(title = "Number of Tests", x = NULL, y = NULL) +  # Remove x and y labels
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))  # Rotate x-axis labels

# Print the plot
print(ethnicity_plot)

# Save the plot as a PNG file named "monthlycount"
ggsave("monthlycount.png", plot = ethnicity_plot, width = 8, height = 6)

# Add a new slide and populate placeholders (Create only ONE slide)
ppt <- ppt %>%
  add_slide(layout = "STBBI", master = "HQ Master Style Slide") %>% 
  
  # Define title and add it to the title placeholder
  ph_with(value = "Cumulative data on sexual health visit", location = ph_location_label(ph_label = "title")) %>% 
  
  # Add the saved image (plot) to the placeholder labeled "Picture"
  ph_with(value = external_img("monthlycount.png"), location = ph_location_label(ph_label = "Picture"))

# Save the updated PowerPoint with only one slide
print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\Board-Report\\CanvaTrial.pptx")



# Slide 15 ----
#LabResults <- LabResults %>% filter()
LabResults<-separate_wider_delim(LabResults,delim = "_",cols=TESTS, names=c("TESTS","Site"),too_few = "align_start")
LabResults<-LabResults
LabResults$TESTS <- gsub("LAB","",LabResults$TESTS)
LabResults
LabResults$results <- gsub(" nonreactive"," negative",LabResults$results)
LabResults$results <- gsub(" reactive"," positive",LabResults$results)
LabResults$results <- gsub(" negative","Negative",LabResults$results)
LabResults$results <- gsub(" positive","Positive",LabResults$results)
resultsCount <- LabResults %>% filter(results == "Positive" |
                                        results == "Negative") %>%
  group_by(TESTS,results) %>% count()
resultsCount <- resultsCount %>% pivot_wider(names_from = results, values_from = n)


slide15<-resultsCount


resultsCount %>% View()

GCPos2023 <- LabResults %>% filter(year(EntryDate) == 2023 &results == "Positive" & TESTS == "GC") %>% select(-Site)
CTPos2023 <- LabResults %>% filter(year(EntryDate) == 2023 &results == "Positive" & TESTS == "CT") %>% select(-Site)
SyphPos2023 <- LabResults %>% filter(year(EntryDate) == 2023 &results == "Positive" & TESTS == "SYPH") %>% select(-Site)
HIVPos2023 <- LabResults %>% filter(year(EntryDate) == 2023 &results == "Positive" & TESTS == "HIV") %>% select(-Site)


HIVPosTenMonths2023 <- LabResults %>% filter(EntryDate <= "2023-10-31" & EntryDate >= "2023-01-01"  &results == "Positive" & TESTS == "HIV") %>% select(-Site)
GCPosTenMonths2023 <- LabResults %>% filter(EntryDate <= "2023-10-31" & EntryDate >= "2023-01-01"  &results == "Positive" & TESTS == "GC") %>% select(-Site)
CTPosTenMonths2023 <- LabResults %>% filter(EntryDate <= "2023-10-31" & EntryDate >= "2023-01-01"  &results == "Positive" & TESTS == "CT") %>% select(-Site)




range_clear(googleSheet,sheet = "Sheet15")
sheet_append(ss= googleSheet, data =slide15,sheet = "Sheet15")

# Slide 16 ----
currentMonth
positivityRate <- resultsCount %>% mutate(Positivity = paste(dirPercentageRounding(Positive/(Negative+Positive)),"%",sep =""))
positivityRate <- positivityRate %>% select(TESTS, Positivity) %>% arrange(desc(Positivity))
slide16 <- data.frame(TESTS=c("Tests"),Positivity=c("Positivity rate")) %>% rbind(positivityRate)

slide16

range_clear(googleSheet,sheet = "Sheet16")
sheet_append(ss= googleSheet, data =slide16,sheet = "Sheet16")

# Slide 17 ----
# title page

#Slide 18 ----
lastMonth <- format(Sys.Date() %m+% months(-1),"%Y-%m")
lastMonth
LabResults %>% filter(format(EntryDate, "%Y-%m")==lastMonth & !(TESTS == "CT")) %>% group_by(TESTS) %>% count()
lastMonthLabCount <- LabResults %>% filter(format(EntryDate, "%Y-%m")==lastMonth & TESTS != "CT") %>% group_by(format(EntryDate, "%Y-%m")) %>% count()
lastMonthLabCount$n
labourCost <- round(21000/lastMonthLabCount$n,2)
paste0("$",labourCost,sep="")
1.85+labourCost
1.34+labourCost
9.5+labourCost
25.00+labourCost
labourCost
#Slide 19 ----
SHVisitData <- AllVisitData %>% filter(ReasonForVisitAbbrev != "Mental Health Visit")
cumulativeExpressTesting <- (SHVisitData %>% filter(ReasonForVisitAbbrev == "Express testing") %>%
                               group_by(ReasonForVisitAbbrev) %>% count() %>% arrange(desc(n)))$n
dollar(cumulativeExpressTesting*37.60)

savingsMonthly <- SHVisitData %>% filter(ReasonForVisitAbbrev == "Express testing") %>%
    group_by(format(EntryDate, "%Y-%m")) %>% count() %>% arrange(desc(n))

savingsMonthly <- savingsMonthly %>% mutate(n = n*37.60)
slide19<-savingsMonthly

range_clear(googleSheet,sheet = "Sheet19")
sheet_append(ss= googleSheet, data =slide19,sheet = "Sheet19")

# Slide 20 ----
SHVisitData <- AllVisitData
slide20<-SHVisitData %>% filter(ReasonForVisitAbbrev != "Mental Health Visit") %>%
  group_by(ReasonForVisitAbbrev) %>% count() %>% arrange(desc(n))
slide20
range_clear(googleSheet,sheet = "Sheet20")
sheet_append(ss= googleSheet, data =slide20,sheet = "Sheet20")

# Slide 21 ----
currentMonth
SHReturnVisits <- AllVisitData %>% filter(visitType == "Return Visit (SH)" | visitType == "First Visit (SH)")
SHVisitsMonthly <- SHReturnVisits %>% group_by(visitType,format(EntryDate, "%y-%b")) %>% count()
SHVisitsMonthly <- SHVisitsMonthly %>% pivot_wider(names_from = visitType, values_from = n)
SHVisitsMonthly <- SHVisitsMonthly %>% rename(EntryDate = `format(EntryDate, "%y-%b")`)
slide21<-SHVisitsMonthly
slide21

range_clear(googleSheet,sheet = "Sheet21")
sheet_append(ss= googleSheet, data =slide21,sheet = "Sheet21")


# Slide 22 ----
OHIP <- read_tsv(file = "DemogGCReport.txt")
OHIP <- OHIP %>% rename(EmrID = `Patient #`)
OHIP <- OHIP %>% select(EmrID,`Health Number`)
noOHIP <- OHIP %>% filter(startsWith(OHIP$`Health Number`,"ON 0001"))
nrow(noOHIP)
noOHIPLabResults <- LabResults %>% right_join(noOHIP)
noOHIPLabResults %>% filter(results == "Positive") %>%
  group_by(TESTS) %>% count() %>% arrange(desc(n))
noOHIPTx <- SHAnPTxsData
noOHIPTx <- noOHIPTx %>% right_join(noOHIP)
noOHIPTxSyph <- noOHIPTx %>% select(EmrID,EntryDate,contains("Syph"))
noOHIPTxSyph <- noOHIPTxSyph %>% filter_at(vars(-EmrID, -EntryDate), any_vars( . == T))
nrow(noOHIPTxSyph)

noOHIPTxGCCT <- noOHIPTx %>% select(EmrID,EntryDate,contains("CT"),contains("GC"),contains("Chlam"),contains("LGV"))
noOHIPTxGCCT <- noOHIPTxGCCT %>% filter_at(vars(-EmrID, -EntryDate), any_vars( . == T))
nrow(noOHIPTxGCCT)

noOHIPTxAny <- noOHIPTx
noOHIPTxAny <- noOHIPTxAny %>% filter_at(vars(-EmrID, -EntryDate, -`Health Number`), any_vars( . == T))
nrow(noOHIPTxAny)
n_distinct(noOHIPTxAny$EmrID)


#TODO: Add IFH number----

# slide 23 ----
tempHIVpos <- noOHIP %>% left_join((Demographics %>% select(EmrID,HIVpos)))
tempHIVpos <- tempHIVpos %>% filter(HIVpos==T)
nrow(tempHIVpos)

# Slide 24 ----
noOHIPAge <- noOHIP %>% left_join((Demographics %>% select(EmrID,Age)))
histogramAgeNoOHIP <- hist(noOHIPAge$Age, breaks = 4)
histogramAgeNoOHIP
slide24<-data.frame(Age=c("18-25","26-35","36-45","46-55","56-65",">66"),`Number of individuals`=histogramAgeNoOHIP$counts)
slide24

range_clear(googleSheet,sheet = "Sheet24")
sheet_append(ss= googleSheet, data =slide24,sheet = "Sheet24")


# Slide 25 ----
noOHIPCountry <- noOHIP %>% left_join((Demographics %>% select(EmrID,countrySpecify)))
slide25<-noOHIPCountry %>% group_by(countrySpecify) %>% count() %>% arrange(desc(n))
slide25

range_clear(googleSheet,sheet = "Sheet25")
sheet_append(ss= googleSheet, data =slide25,sheet = "Sheet25")

# Slide 26 ----
noOHIPYear <- noOHIP %>% left_join((Demographics %>% select(EmrID,dateOfArrival)))
slide26<-noOHIPYear %>% group_by(dateOfArrival) %>% count() %>% arrange(desc(n))
slide26

range_clear(googleSheet,sheet = "Sheet26")
sheet_append(ss= googleSheet, data =slide26,sheet = "Sheet26")


## -----------------
SHVisitDataUnique <- SHVisitData #%>% select(EmrID,EntryDate)
View(SHVisitDataUnique)
tempDemogMOH <- SHVisitDataUnique %>% inner_join(Demographics,by=join_by(EmrID))
tempHIVpos <- SHVisitDataUnique %>% inner_join(Demographics %>% filter(HIVpos==T),by=join_by(EmrID))
View(tempHIVpos)
tempHIVpos %>% distinct(EmrID,.keep_all = T) %>% group_by(gender) %>% count()

tempHIVneg <- SHVisitDataUnique %>% inner_join(Demographics %>% filter(HIVpos!=T),by=join_by(EmrID))
View(tempHIVneg %>% distinct(EmrID,.keep_all = T) %>% group_by(gender) %>% count() %>% arrange(desc(n)))
View(tempHIVpos %>% distinct(EmrID,.keep_all = T) %>% group_by(gender,Age) %>% count() )
AgeGenderPHA <- tempHIVpos %>% distinct(EmrID,.keep_all = T) %>% group_by(gender,Age) %>% count()
AgeGenderPHA %>% filter(Age < 18) %>% group_by(gender) %>% summarise(nn=sum(n))
AgeGenderPHA %>% filter(Age >= 18 & Age < 25) %>% group_by(gender) %>% summarise(nn=sum(n)) %>% View()
AgeGenderPHA %>% filter(Age >= 26 & Age < 35) %>% group_by(gender) %>% summarise(nn=sum(n)) %>% View()
AgeGenderPHA %>% filter(Age >= 36 & Age < 45) %>% group_by(gender) %>% summarise(nn=sum(n)) %>% View()
AgeGenderPHA %>% filter(Age >= 46 & Age < 55) %>% group_by(gender) %>% summarise(nn=sum(n)) %>% View()
AgeGenderPHA %>% filter(Age >= 56 & Age < 65) %>% group_by(gender) %>% summarise(nn=sum(n)) %>% View()
AgeGenderPHA %>% filter(Age >= 66 & Age < 75) %>% group_by(gender) %>% summarise(nn=sum(n)) %>% View()
AgeGenderPHA %>% filter(Age >= 75) %>% group_by(gender) %>% summarise(nn=sum(n)) %>% View()
AgeGenderPHAneg <- tempHIVneg %>% distinct(EmrID,.keep_all = T) %>% group_by(gender,Age) %>% count()
AgeGenderPHAneg %>% filter(Age < 18) %>% group_by(gender) %>% summarise(nn=sum(n))
AgeGenderPHAneg %>% filter(Age >= 18 & Age < 25) %>% group_by(gender) %>% summarise(nn=sum(n)) %>% View()
AgeGenderPHAneg %>% filter(Age >= 26 & Age < 35) %>% group_by(gender) %>% summarise(nn=sum(n)) %>% View()
AgeGenderPHAneg %>% filter(Age >= 36 & Age < 45) %>% group_by(gender) %>% summarise(nn=sum(n)) %>% View()
AgeGenderPHAneg %>% filter(Age >= 46 & Age < 55) %>% group_by(gender) %>% summarise(nn=sum(n)) %>% View()
AgeGenderPHAneg %>% filter(Age >= 56 & Age < 65) %>% group_by(gender) %>% summarise(nn=sum(n)) %>% View()
AgeGenderPHAneg %>% filter(Age >= 66 & Age < 75) %>% group_by(gender) %>% summarise(nn=sum(n)) %>% View()
AgeGenderPHAneg %>% filter(Age >= 75) %>% group_by(gender) %>% summarise(nn=sum(n)) %>% View()

AllApptsData %>% group_by(Late) %>% count() %>% View()
OceanPrEP <- read_tsv("OceanPrEP.txt")
OceanPrEP <- OceanPrEP %>% rename(EmrID = `Patient #`) %>% select(EmrID) %>% mutate(onPrEP = T)

##HIRI


HIRIpos <- HIRIDataSummary %>% filter(OUThScore >= 10)
HIRIpos$OUThScore %>% mean()


#Sheet 33 - Weekly Average
slide33 <- WeeklyAverages %>% filter(WeekServed > 48)



save.image("canva_2024-06-25.Rdata")

########################3
# Deselect Time column
attach("../Data/currentData.Rdata")
AllVisitData <- AllVisitData
detach()

AllVisitData <- AllVisitData %>%
  select(-Time) %>%
  distinct()
# Define quarters starting from Q2 2022 to Q4 2024
quarters <- seq(as.Date("2022-04-01"), as.Date("2024-12-31"), by = "quarter")
quarter_labels <- paste0("Q", quarter(quarters), " ", year(quarters))

# Initialize an empty data frame for the results
results <- data.frame(Quarter = quarter_labels,
                      TotalCumulativePatients = integer(length(quarter_labels)),
                      TotalActivePatients = integer(length(quarter_labels)),
                      TotalActiveVisits = integer(length(quarter_labels)),
                      NewDiagnosed = integer(length(quarter_labels)))

# Initialize start_date for the calculation
start_date <- as.Date("2022-04-01")

for (i in 1:length(quarters)) {
  end_date <- quarters[i] + months(3) - days(1)

  # Total cumulative patients since start_date
  results$TotalCumulativePatients[i] <- Demographics %>%
    filter(EntryDate <= end_date) %>%
    n_distinct()

  # Total active patients seen in the quarter, distinct by EmrID
  results$TotalActivePatients[i] <- AllVisitData %>%
    filter(ReasonForVisitAbbrev != "Mental Health Visit") %>%
    filter(EntryDate >= start_date & EntryDate <= end_date) %>%
    distinct(EmrID) %>%  # Select distinct EmrID
    nrow()  # Count the number of distinct EmrID


  # Total active patients seen in the quarter
  results$TotalActiveVisits[i] <- AllVisitData %>%
    filter(ReasonForVisitAbbrev != "Mental Health Visit")%>%
    filter(EntryDate >= start_date & EntryDate <= end_date) %>%
    n_distinct()

  # New diagnosed patients in the quarter
  # results$NewDiagnosed[i] <- NewHIVDx %>%
  #   filter(EntryDate >= start_date & EntryDate <= end_date) %>%
  #   n_distinct()

  # Update start_date for the next quarter
  start_date <- end_date + days(1)
}

# Print the results
View(results)


# Initialize an empty data frame for the results
demographics_results <- data.frame(Quarter = quarter_labels)

# Initialize empty lists to store gender and age counts
gender_counts <- list()
age_counts <- list()

# Define gender columns and age groups
gender_cols <- c("genderCisM", "genderCisW", "genderTransM", "genderTransW","genderNB", "genderInter", "genderPreferNA", "genderOther")
age_groups <- list(
  "<25" = c(0, 24),
  "25-29" = c(25, 29),
  "30-34" = c(30, 34),
  "35-39" = c(35, 39),
  "40-44" = c(40, 44),
  "45-49" = c(45, 49),
  "50-54" = c(50, 54),
  "55-59" = c(55, 59),
  "60-69" = c(60, 69),
  "70-79" = c(70, 79),
  "80+" = c(80, Inf)
)

# Calculate values for each quarter
start_date <- as.Date("2022-04-01")
for (i in 1:length(quarters)) {
  end_date <- quarters[i] + months(3) - days(1)

  # Filter active patients for the quarter
  active_patients <- AllVisitData %>%
    filter(ReasonForVisitAbbrev != "Mental Health Visit")%>%
    filter(EntryDate >= start_date & EntryDate <= end_date) %>%
    distinct(EmrID)

  # Join with demographics to get gender and age information
  active_demographics <- active_patients %>%
    inner_join(Demographics, by = "EmrID")

  # Calculate gender counts
  gender_count <- active_demographics %>%
    summarise(across(all_of(gender_cols), ~ sum(.x, na.rm = TRUE)))
  gender_counts[[i]] <- gender_count

  # Calculate age counts using the Age column directly
  active_demographics <- active_demographics %>%
    mutate(Age = if_else(!is.na(Age), Age, Age))

  # Calculate age counts using the Age column directly
  age_count <- map_dfr(age_groups, ~ active_demographics %>%
                         filter(Age >= .x[1] & Age <= .x[2]) %>%
                         summarise(Count = n()))

  age_counts[[i]] <- age_count

  # Update start_date for the next quarter
  start_date <- end_date + days(1)
}

# Combine gender counts and age counts into the results data frame
for (j in 1:length(quarters)) {
  demographics_results[j, gender_cols] <- gender_counts[[j]]
  demographics_results[j, names(age_groups)] <- age_counts[[j]]$Count
}

# Print the demographics results
View(demographics_results)

save.image("canva_2024-06-25_counts.Rdata")



# Save the results to a CSV file (if needed)
write_csv(demographics_results, "HIV_Demographics_Quarterly_Data.csv")


###############################33333
# Define years for the analysis
years <- c(2022, 2023, 2024)
year_labels <- as.character(years)

# Initialize an empty data frame for the results
demographics_yearly_results <- data.frame(Year = year_labels)

# Initialize empty lists to store gender and age counts
gender_counts_yearly <- list()
age_counts_yearly <- list()

# Calculate values for each year
for (i in 1:length(years)) {
  start_date <- as.Date(paste0(years[i], "-01-01"))
  end_date <- as.Date(paste0(years[i], "-12-31"))

  # Filter demographics for new patients in the year
  new_demographics <- Demographics %>%
    filter(EntryDate >= start_date & EntryDate <= end_date) %>%
    ungroup()


  # Calculate gender counts
  gender_count <- new_demographics %>%
    summarise(across(all_of(gender_cols), ~ sum(.x, na.rm = TRUE)))
  gender_counts_yearly[[i]] <- gender_count

  # Calculate age counts using the Age column directly
  age_count <- map_dfr(age_groups, ~ new_demographics %>%
                         filter(Age >= .x[1] & Age <= .x[2]) %>%
                         summarise(Count = n()))
  age_counts_yearly[[i]] <- age_count
}

# Combine gender counts and age counts into the results data frame
for (j in 1:length(years)) {
  demographics_yearly_results[j, gender_cols] <- gender_counts_yearly[[j]]
  demographics_yearly_results[j, names(age_groups)] <- age_counts_yearly[[j]]$Count
}

View(demographics_yearly_results)
############################################

# Initialize an empty data frame for the results
substance_use_results <- data.frame(Quarter = quarter_labels)

# Define a mapping for substance use columns
substance_mapping <- list(
  "Tobacco" = "tobaccoUse",
  "Alcohol_heavy" = "alcoholUse",
  "Current_recreational_drug_use_all" = c("cannabisUse", "methUse", "otherStimUse", "cocaineUse", "opioidUse", "benzoUse", "ketaUse", "ghbUse", "popperUse", "injectionUse", "hallucUse"),
  "Injectables" = "injectionUse",
  "Crystal_meth" = "methUse",
  "Cannabis" = "cannabisUse",
  "Opioids" = "opioidUse",
  "Other_GHB_GBL_etc" = c("ghbUse", "otherStimUse")
)

# Calculate values for each quarter
start_date <- as.Date("2022-04-01")
for (i in 1:length(quarters)) {
  end_date <- quarters[i] + months(3) - days(1)

  # Filter active patients for the quarter
  active_patients <- AllVisitData %>%
    filter(ReasonForVisitAbbrev != "Mental Health Visit")%>%
    filter(EntryDate >= start_date & EntryDate <= end_date) %>%
    distinct(EmrID)

  # Join with demographics to get substance use information
  active_substance_use <- active_patients %>%
    inner_join(Demographics, by = "EmrID")

  # Initialize a data frame to store substance use counts for the quarter
  quarter_counts <- data.frame(matrix(0, ncol = length(substance_mapping), nrow = 1))
  colnames(quarter_counts) <- names(substance_mapping)

  # Calculate substance use counts
  for (substance in names(substance_mapping)) {
    cols <- substance_mapping[[substance]]
    if (is.character(cols)) {
      cols <- list(cols)
    }
    # Ensure cols is a character vector
    if (is.list(cols)) {
      cols <- unlist(cols)
    }
    # Calculate substance use counts ensuring we count patients, not instances
    if (substance == "Alcohol_heavy") {
      quarter_counts[[substance]] <- active_substance_use %>%
        filter(alcoholUse == 2) %>%  # Only count heavy use for alcohol
        distinct(EmrID) %>%
        n_distinct()  # Count distinct patients
    } else if (substance == "Current_recreational_drug_use_all") {
      quarter_counts[[substance]] <- active_substance_use %>%
        select(EmrID, all_of(cols)) %>%
        group_by(EmrID) %>%
        summarise(any_use = any(c_across(all_of(cols)) %in% c(1, 2))) %>%
        filter(any_use) %>%
        distinct(EmrID) %>%
        n_distinct()  # Count distinct patients
    } else {
      quarter_counts[[substance]] <- active_substance_use %>%
        select(EmrID, all_of(cols)) %>%
        group_by(EmrID) %>%
        summarise(any_use = any(c_across(all_of(cols)) %in% c(1, 2))) %>%
        filter(any_use) %>%
        distinct(EmrID) %>%
        n_distinct()  # Count distinct patients
    }


  }
  # Add the counts for the quarter to the results
  substance_use_results[i, names(substance_mapping)] <- quarter_counts
  # Update start_date for the next quarter
  start_date <- end_date + days(1)
}

# Save the substance use results to memory
substance_use_results_memory <- substance_use_results

# Print the substance use results
View(substance_use_results_memory)
####################################################
# Define the comorbidity mapping
comorbidity_mapping <- list(
  "Cardiovascular" = c("hypertension", "heart failure", "atrial fibrillation"),
  "Diabetes" = c("diabetes mellitus", "type 2 diabetes"),
  "Cancer" = c("lung cancer", "breast cancer", "prostate cancer"),
  "Renal" = c("chronic kidney disease", "acute kidney injury"),
  "Bone" = c("osteoporosis", "arthritis"),
  "Other" = c("asthma", "COPD", "HIV", "hepatitis", "depression", "anxiety")
)

# Create a function to map conditions to comorbidities
map_condition_to_comorbidity <- function(condition) {
  for (comorbidity in names(comorbidity_mapping)) {
    if (condition %in% comorbidity_mapping[[comorbidity]]) {
      return(comorbidity)
    }
  }
  return("Other")
}

truvada <- truvada %>% mutate(Drug = "Truvada")
descovy <- descovy %>% mutate(Drug = "Descovy")
DispensingForm <- DispensingForm %>% mutate(Drug = "TAF study")

drugs <- truvada %>%
  bind_rows(descovy) %>%
  bind_rows(DispensingForm) %>%
  rename(EmrID = `Patient #`)

ptRx <- genericRx %>%
  bind_rows(nonGenericRx) %>%
  bind_rows(ptHPH) %>%
  # Convert Condition to lowercase
  mutate(Condition = tolower(Condition)) %>%
  # Convert Condition to lowercase
  #mutate(Condition = gsub("hiv infection", "hiv", Condition)) %>%
  distinct()

drugs <- drugs %>%
  inner_join(HQPrEP)
ptRx <- ptRx %>%
  inner_join(HQPrEP)

ptRx %>% distinct(EmrID) %>% n_distinct()
drugs %>% distinct(EmrID) %>% n_distinct()

ptRx <- ptRx %>%
  inner_join(drugs)

ptRx <- ptRx %>% filter(Drug == "TAF study")


# Add comorbidity category to ptRx
ptRx <- ptRx %>%
  mutate(Comorbidity = Condition)#sapply(Condition, map_condition_to_comorbidity))


# Define comorbidity columns
# comorbidity_cols <- c("Cardiovascular", "Diabetes", "Cancer", "Renal", "Bone", "Other")
comorbidity_cols <- ptRx$Condition %>%
  table() %>%
  as.data.frame() %>%
  arrange(desc(Freq)) %>%
  head(10) %>%
  pull(".") %>%
  as.character()

# Initialize an empty data frame for the results with all comorbidity columns
comorbidity_results <- data.frame(Quarter = quarter_labels, matrix(0, nrow = length(quarter_labels), ncol = length(comorbidity_cols)))
colnames(comorbidity_results)[2:ncol(comorbidity_results)] <- comorbidity_cols

# Calculate values for each quarter
start_date <- as.Date("2022-04-01")
for (i in 1:length(quarters)) {
  end_date <- quarters[i] + months(3) - days(1)

  # Filter active patients for the quarter
  active_patients <- AllVisitData %>%
    filter(ReasonForVisitAbbrev != "Mental Health Visit") %>%
    filter(EntryDate >= start_date & EntryDate <= end_date) %>%
    distinct(EmrID)

  # Join with demographics to get comorbidity information
  active_comorbidities <- active_patients %>%
    inner_join(ptRx, by = "EmrID")

  # Calculate comorbidity counts
  quarter_counts <- active_comorbidities %>%
    group_by(Comorbidity) %>%
    summarise(Count = n_distinct(EmrID)) %>%
    pivot_wider(names_from = Comorbidity, values_from = Count, values_fill = list(Count = 0))

  # Ensure all comorbidity columns are present in quarter_counts
  for (col in comorbidity_cols) {
    if (!(col %in% names(quarter_counts))) {
      quarter_counts[[col]] <- 0
    }
  }

  # Add the counts for the quarter to the results
  comorbidity_results[i, comorbidity_cols] <- quarter_counts[1, comorbidity_cols]

  # Update start_date for the next quarter
  start_date <- end_date + days(1)
}
# Print the comorbidity results
View(comorbidity_results)
comorbidity_resultsStudy <-  comorbidity_results

writexl::write_xlsx(comorbidity_resultsStudy, "comorbidStudy.xlsx")

# Calculate frequency of comorbidities per drug
comorbidity_freq <- ptRx %>%
  #filter(Drug == "Descovy") %>%
  distinct(EmrID,Condition,.keep_all = T) %>%
  group_by(Drug, Condition) %>%
  summarise(Freq = n(), .groups = 'drop') %>%
  arrange(desc(Freq))
library(tidyverse)
library(knitr)
library(kableExtra)
library(ggplot2)
# Pivot the data to wider format
summary_table <- comorbidity_freq %>%
  pivot_wider(names_from = Drug, values_from = Freq, values_fill = list(Freq = 0))

# Create the formatted table
summary_table %>%
  kable("html", col.names = c("Condition", "Truvada", "TAF Study", "Descovy")) %>%
  kable_styling(bootstrap_options = c("striped", "hover"), full_width = F)

##############################################3333
save.image("PrEPHIVGileadReport.Rdata")

# 1-READ file   ######################| ----
# Load 'OceanPrEP true' CPP data ----
OceanPrEP <- read_tsv("OceanPrEP.txt")
OceanPrEP <- OceanPrEP %>% rename(EmrID = `Patient #`) %>% select(EmrID) %>% mutate(onPrEP = T)

# 2-READ file   ######################| ----
# Load 'OceanHQPrEP Quick Export' CPP data ----
HQPrEP <- read_tsv("OceanHQPrEP.txt")
HQPrEP <- HQPrEP %>% rename(EmrID = `Patient #`) %>% select(EmrID) %>% mutate(PrEPatHQ = T)
HQTEST <- read_tsv("HQTEST_true.txt") %>% rename(EmrID = `Patient #`)
OceanPrEP %>% n_distinct()

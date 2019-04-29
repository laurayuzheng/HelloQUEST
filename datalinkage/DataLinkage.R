library(readxl)
library(xlsx)
library(tidyr)
library(sqldf)
library(tidyverse)
library(lubridate)
library(dplyr)

#load data file
master_sales = read_excel('Master Sales Sheet (2).xlsx')
time_stamps = read.csv("BOB data.csv")
shifts_worked = read_excel('Shifts Worked (2).xlsx')
sales_expense = read_excel('SEL COMBINED.xlsx')

#clean master sales sheet from bob
master_sales = master_sales[master_sales$`Business Name` == "HelloFresh",]
#format dates
master_sales$Date = strptime(as.character(master_sales$Date), "%Y-%m-%d")
master_sales$Date = format(master_sales$Date, "%m/%d/%Y")
#remove columns
master_sales = master_sales[, -c(2,3,5,6,7)]

#clean time stamp data to only keep customer ID and time data
time_stamps = time_stamps[, -c(3:7)]
time_stamps = time_stamps %>% separate('date_sign_up', into = c(NA, "Time"), sep = ' ')
time_stamps$Time = strptime(as.character(time_stamps$Time), "%H:%M")
time_stamps$Time = format(time_stamps$Time, "%H:%M")

#clean shifts worked sheet - format the start day, start time, end time
shifts_worked$start_day = strptime(as.character(shifts_worked$start_day), "%m/%d/%Y")
shifts_worked$start_day = format(shifts_worked$start_day, "%m/%d/%Y")
shifts_worked = shifts_worked[!is.na(shifts_worked$eid) & !is.na(shifts_worked$location) &
                                !is.na(shifts_worked$shift_title),]

# can ignore this
shifts_worked = shifts_worked[!grepl("Meeting", shifts_worked$shift_title) & !grepl("Training", shifts_worked$shift_title) &
                              !grepl("meeting", shifts_worked$shift_title) & !grepl("Workshop", shifts_worked$shift_title) &
                              !grepl("Leadership", shifts_worked$shift_title) & !grepl("Office Hours", shifts_worked$shift_title) &
                              !grapl("Territory", shifts_worked$shift_title) & !grepl("territory", shifts_worked$shift_title),]

shifts_worked$start_time = strptime(as.character(shifts_worked$start_time), "%I:%M %p")
shifts_worked$end_time = strptime(as.character(shifts_worked$end_time), "%I:%M %p")
shifts_worked$start_time = format(shifts_worked$start_time, "%H:%M")
shifts_worked$end_time = format(shifts_worked$end_time, "%H:%M")

#create end day field
shifts_worked$start_day = strptime(as.character(shifts_worked$start_day), "%m/%d/%Y")
shifts_worked$end_day = as.Date(shifts_worked$start_day)

for (i in 1:length(shifts_worked$start_day)) {
  if (shifts_worked$start_time[i] > shifts_worked$end_time[i]) {
    shifts_worked$end_day[i] = shifts_worked$end_day[i] + 1
  }
}

shifts_worked$start_day = as.character(shifts_worked$start_day)
shifts_worked$end_day = as.character(shifts_worked$end_day)

#prepare start and end time for time zone conversions
shifts_worked$DateSTime = paste(shifts_worked$start_day, " ", shifts_worked$start_time)
shifts_worked$DateETime = paste(shifts_worked$end_day, " ", shifts_worked$end_time)
shifts_worked$DateSTime = strptime(as.character(shifts_worked$DateSTime), "%Y-%m-%d %H:%M")
shifts_worked$DateETime = strptime(as.character(shifts_worked$DateETime), "%Y-%m-%d %H:%M")

#timezone conversion code
for (i in 1:length(shifts_worked$DateETime)) {
  if (shifts_worked$location[i] %in% c("Pleasanton, CA", "Santa Ana, CA", "Seattle, WA", "Portland, OR")) {
    pacs = as.POSIXct(shifts_worked$DateSTime[i], tz = "PST8PDT")
    shifts_worked$DateSTime[i] = format(pacs, tz = "EST5EDT")
    pace = as.POSIXct(shifts_worked$DateETime[i], tz = "PST8PDT")
    shifts_worked$DateETime[i] = format(pace, tz = "EST5EDT")
  } else if (shifts_worked$location[i] %in% c("Phoenix, AZ", "Denver, CO")) {
    mtns = as.POSIXct(shifts_worked$DateSTime[i], tz = "MST7MDT")
    shifts_worked$DateSTime[i] = format(mtns, tz = "EST5EDT")
    mtne = as.POSIXct(shifts_worked$DateETime[i], tz = "MST7MDT")
    shifts_worked$DateETime[i] = format(mtne, tz = "EST5EDT")
  } else if (shifts_worked$location[i] %in% c("Houston, TX", "Dallas, TX", "Austin, TX")) {
    cents = as.POSIXct(shifts_worked$DateSTime[i], tz = "CST6CDT")
    shifts_worked$DateSTime[i] = format(cents, tz = "EST5EDT")
    cente = as.POSIXct(shifts_worked$DateETime[i], tz = "CST6CDT")
    shifts_worked$DateETime[i] = format(cente, tz = "EST5EDT")
  }
}

#delete irrelevant shifts worked sheets
shifts_worked = shifts_worked[, -c(4,5)]
#shifts_worked$DateETime = shifts_worked$DateETime + minutes(15)
#shifts_worked$DateSTime = shifts_worked$DateSTime - minutes(15)

# shifts_worked$total_time = difftime(shifts_worked$DateETime, shifts_worked$DateSTime, units = "hours")

#reformat times for range comparison
shifts_worked$DateSTime = as.character(shifts_worked$DateSTime)
shifts_worked$DateETime = as.character(shifts_worked$DateETime)

#merge sales sheet with timestamp data
master_sales = merge(master_sales, time_stamps, by.x = "Customer ID", by.y = "customer_id", 
                     all.x = TRUE, all.y = FALSE)
rm(time_stamps)

#make the time stamp conversion for sales
master_sales$DateTime = paste(master_sales$Date, " ", master_sales$Time)
master_sales$DateTime = strptime(as.character(master_sales$DateTime), "%m/%d/%Y %H:%M")
utc = as.POSIXct(master_sales$DateTime, tz = "PST8PDT")
master_sales$DateTime = format(utc, tz = "EST5EDT")

master_sales$DateTime = as.character(master_sales$DateTime)
master_sales$Date = strptime(as.character(master_sales$Date), "%m/%d/%Y")
master_sales$Date = format(master_sales$Date, "%Y-%m-%d")

#finally merge sales with shiftsworked
combined_2 = merge(master_sales, shifts_worked, by.x = c("Employee ID", "Office Name Formatted", "Date"), 
                 by.y = c("eid", "location", "start_day"), all = TRUE)

rm(master_sales, shifts_worked)

#clean the combined sheet of incomplete entries
# ********* TO DO *********
combined_2[is.na(combined_2)] = 0
combined_2 = combined_2[combined_2$`Employee ID` > 0 & combined_2$`Office Name Formatted` != 0 &
                      combined_2$Date != 0 & combined_2$shift_title != "0", ]

#filter out the shifts that generated 0 sales into a separate sheet for later
nosale = combined_2[combined_2$`# Sales` == 0,]
nosale = nosale %>% distinct(nosale$`Employee ID`, nosale$`Office Name Formatted`, nosale$shift_title, .keep_all = TRUE)

#filter out no sale shifts to prepare for time comparison
combined_2 = combined_2[combined_2$DateTime != 0, ]

#time range fitting
combined_2 = combined_2[!(combined_2$DateTime <= combined_2$DateSTime | combined_2$DateTime >= combined_2$DateETime),]
#remove duplicate sales entries
combined_2 = combined_2 %>% distinct(combined_2$`Customer ID`, combined_2$DateTime, .keep_all = TRUE)
#remove irrelevant columns
combined_2 = combined_2[, -c(3, 8, 11, 14, 15)]

#dupls = combined_2[duplicated(combined_2$`Customer ID`),]

nosale = nosale[, -c(3, 8, 11, 14, 15, 16)]

#rejoin the no sale data with time range filtered sales
saleandnosale = rbind(combined_2, nosale)
rm(combined_2, nosale)

#separating paid events
paid_events = unique(sales_expense$`Vendor Name`) #pull unique vendor names from SEL
paid_event_sales = saleandnosale[saleandnosale$shift_title %in% paid_events,] #extract sales made at events from SEL
free_events = saleandnosale[!(saleandnosale$shift_title %in% paid_events),] #remainder of sales are classified as free event sales

#condense the free event sales by grouping by shift, aggregating number of sales and breakdown of sales per shift
agged = free_events %>% group_by(`Employee ID`, shift_title, `Office Name Formatted`, DateSTime, DateETime) %>%
  summarise(num_sales = sum(`# Sales`),
            twoxtwo = sum(`Product Box Sub Level` == "2x2 Box" | `Product Box Sub Level` == "2x2 box"),
            thrxtwo = sum(`Product Box Sub Level` == "3x2 Box" | `Product Box Sub Level` == "3x2 box"),
            twoxfour = sum(`Product Box Sub Level` == "2x4 Box" | `Product Box Sub Level` == "2x4 box"),
            fourxtwo = sum(`Product Box Sub Level` == "4x2 Box" | `Product Box Sub Level` == "4x2 box"),
            thrxfour = sum(`Product Box Sub Level` == "3x4 Box" | `Product Box Sub Level` == "3x4 box"))
#prepare start and end time data fields for wage cost calculation
agged$DateSTime = strptime(as.character(agged$DateSTime), "%Y-%m-%d %H:%M")
agged$DateETime = strptime(as.character(agged$DateETime), "%Y-%m-%d %H:%M")
agged$total_time = difftime(agged$DateETime, agged$DateSTime, units = "hours")
agged$DateSTime = as.character(agged$DateSTime)
agged$DateETime = as.character(agged$DateETime)
agged$total_time = as.numeric(agged$total_time)
agged$wage_cost = 13.72 * agged$total_time
agged$wage_cost = round(agged$wage_cost, digits = 2)
FREEeventsPERSHIFTSUMMARY = as.data.frame(agged)
rm(free_events, agged)

#write.xlsx(FREEeventsPERSHIFTSUMMARY, file = "FREEeventsPERSHIFTSUMMARY.xlsx", sheetName = "Sheet1", row.names = FALSE)

#summarizing the free events per unique event shift title (combines all shifts for a single event into one to track sales)
FREEeventsPEREVENTSUMMARY = FREEeventsPERSHIFTSUMMARY %>% group_by(shift_title, `Office Name Formatted`) %>% 
  summarise(wage_exp = sum(wage_cost), total_hrs = sum(total_time), num_sales = sum(num_sales),
            twoxtwo = sum(twoxtwo), thrxtwo = sum(thrxtwo), twoxfour = sum(twoxfour), fourxtwo = sum(fourxtwo), thrxfour = sum(thrxfour))
FREEeventsPEREVENTSUMMARY = data.frame(FREEeventsPEREVENTSUMMARY)
#write.xlsx(FREEeventsPEREVENTSUMMARY, file = "FREEeventsPEREVENTSUMMARY.xlsx", sheetName = "Sheet1", row.names = FALSE)

#analysis of paid events
aggedpaid = paid_event_sales %>% group_by(`Employee ID`, shift_title, `Office Name Formatted`, DateSTime, DateETime) %>%
  summarise(num_sales = sum(`# Sales`),
            twoxtwo = sum(`Product Box Sub Level` == "2x2 Box" | `Product Box Sub Level` == "2x2 box"),
            thrxtwo = sum(`Product Box Sub Level` == "3x2 Box" | `Product Box Sub Level` == "3x2 box"),
            twoxfour = sum(`Product Box Sub Level` == "2x4 Box" | `Product Box Sub Level` == "2x4 box"),
            fourxtwo = sum(`Product Box Sub Level` == "4x2 Box" | `Product Box Sub Level` == "4x2 box"),
            thrxfour = sum(`Product Box Sub Level` == "3x4 Box" | `Product Box Sub Level` == "3x4 box"))
#calculation of wage cost for paid events
aggedpaid$DateSTime = strptime(as.character(aggedpaid$DateSTime), "%Y-%m-%d %H:%M")
aggedpaid$DateETime = strptime(as.character(aggedpaid$DateETime), "%Y-%m-%d %H:%M")
aggedpaid$total_time = difftime(aggedpaid$DateETime, aggedpaid$DateSTime, units = "hours")
aggedpaid$DateSTime = as.character(aggedpaid$DateSTime)
aggedpaid$DateETime = as.character(aggedpaid$DateETime)
aggedpaid$total_time = as.numeric(aggedpaid$total_time)
aggedpaid$wage_cost = 13.72 * aggedpaid$total_time
aggedpaid$wage_cost = round(aggedpaid$wage_cost, digits = 2)
PAIDeventsPERSHIFTSUMMARY = as.data.frame(aggedpaid)
rm(aggedpaid, paid_event_sales)
#write.xlsx(PAIDeventsPERSHIFTSUMMARY, file = "PAIDeventsPERSHIFTSUMMARY.xlsx", sheetName = "Sheet1", row.names = FALSE)

#summarizing the paid events per unique shift title (combines all shifts for a single event into one to track sales)
PAIDeventsPEREVENTSUMMARY = PAIDeventsPERSHIFTSUMMARY %>% group_by(shift_title, `Office Name Formatted`) %>% 
  summarise(wage_exp = sum(wage_cost), total_hrs = sum(total_time), num_sales = sum(num_sales),
            twoxtwo = sum(twoxtwo), thrxtwo = sum(thrxtwo), twoxfour = sum(twoxfour), fourxtwo = sum(fourxtwo), thrxfour = sum(thrxfour))
temp = data.frame(PAIDeventsPEREVENTSUMMARY)

#merge the event costs on SEL with the paid events per event summary
PAIDeventsCOSTMERGED = merge(temp, sales_expense, by.x = "shift_title", by.y = "Vendor Name", all.x = TRUE, all.y = FALSE)
PAIDeventsCOSTMERGED = PAIDeventsCOSTMERGED[, -c(11, 14:23, 25:41)]
PAIDeventsCOSTMERGED$total_cost = (PAIDeventsCOSTMERGED$wage_exp + PAIDeventsCOSTMERGED$`Cost ($)`)
write.xlsx(PAIDeventsCOSTMERGED, file = "PAIDeventsCOSTMERGED.xlsx", sheetName = "Sheet1", row.names = FALSE)

#finally merge both the free event sales (and wage costs) with paid event sales (and event costs + wage costs)
PAIDandFREECOMBINED = FREEeventsPEREVENTSUMMARY
PAIDandFREECOMBINED$`Cost ($)` = 0
PAIDandFREECOMBINED$Category = 0
PAIDandFREECOMBINED$Subcategory = 0
PAIDandFREECOMBINED$total_cost = (PAIDandFREECOMBINED$wage_exp + PAIDandFREECOMBINED$`Cost ($)`)
PAIDandFREECOMBINED = rbind(PAIDandFREECOMBINED, PAIDeventsCOSTMERGED)

write.xlsx(PAIDandFREECOMBINED, file = "PAIDandFREECOMBINED.xlsx", sheetName = "Sheet1", row.names = FALSE)

library(readxl)
library(xlsx)
library(tidyr)
library(sqldf)
library(tidyverse)

master_sales = read_excel('Master Sales Sheet (2).xlsx')
time_stamps = read.csv("BOB data.csv")
shifts_worked = read_excel('Shifts Worked (2).xlsx')

master_sales = master_sales[master_sales$`Business Name` == "HelloFresh",] # filtering only HelloFresh sales
master_sales$Date = strptime(as.character(master_sales$Date), "%Y-%m-%d") # converts to date-time object
master_sales$Date = format(master_sales$Date, "%m/%d/%Y") # reformatted time so it matches on both sheets
master_sales = master_sales[, -c(2,3,5,6,7)] # removes columns from master sales

#clean time stamp data to only keep customer ID and time data
time_stamps = time_stamps[, -c(3:7)] # remove initial sales order num and voucher code
time_stamps = time_stamps %>% separate('date_sign_up', into = c(NA, "Time"), sep = ' ') # create time column (?)
time_stamps$Time = strptime(as.character(time_stamps$Time), "%H:%M") # convert time to hours and minutes
time_stamps$Time = format(time_stamps$Time, "%H:%M") # format time column

shifts_worked$start_day = strptime(as.character(shifts_worked$start_day), "%m/%d/%Y")
shifts_worked$start_day = format(shifts_worked$start_day, "%m/%d/%Y")
shifts_worked = shifts_worked[!is.na(shifts_worked$eid) & !is.na(shifts_worked$location) &
                                !is.na(shifts_worked$shift_title),]
shifts_worked$start_time = strptime(as.character(shifts_worked$start_time), "%I:%M %p")
shifts_worked$end_time = strptime(as.character(shifts_worked$end_time), "%I:%M %p")
shifts_worked$start_time = format(shifts_worked$start_time, "%H:%M")
shifts_worked$end_time = format(shifts_worked$end_time, "%H:%M")

master_sales = merge(master_sales, time_stamps, by.x = "Customer ID", by.y = "customer_id", 
                     all.x = TRUE, all.y = FALSE)

master_sales_sql = sqldf('SELECT * FROM master_sales')
shifts_worked_sql = sqldf('SELECT * FROM shifts_worked')

#time filter 
combined = sqldf('SELECT a. start_time, end_time, shift_title, b.* 
                 FROM shifts_worked AS a LEFT JOIN master_sales AS b
                 ON b. `Employee ID` = a. `eid`
                 AND b. `Office Name Formatted` = a. `location`
                 AND b. `Date` = a. `start_day`
                 WHERE b. `Time` >= a. `start_time`
                 AND b. `Time` <= a. `end_time`')

#without time filter
combined_2 = merge(master_sales, shifts_worked, by.x = c("Employee ID", "Office Name Formatted", "Date"), 
                 by.y = c("eid", "location", "start_day"), all = FALSE)

#figure out time stamp thing
#how to implement event ID
#ongoing usage of this

combined_2[is.na(combined_2)] = 0
combined_2 = combined_2[combined_2$`Employee ID` > 0 & combined_2$`Office Name Formatted` != 0 &
                      combined_2$Date != 0 & combined_2$shift_title != "0" & combined_2$Time != 0, ]

nums = unique(combined_2$`Customer ID`)
dupls = combined_2[duplicated(combined_2$`Customer ID`),]

combined_2 = combined_2[!(duplicated(combined_2$`Customer ID`) & (combined_2$Time <= combined_2$start_time 
                          | combined_2$Time >= combined_2$end_time)),]
combined_2 = combined_2[!(duplicated(combined_2$`Customer ID`, fromLast = TRUE) & (combined_2$Time <= combined_2$start_time 
                          | combined_2$Time >= combined_2$end_time)),]
combined_2 = combined_2 %>% distinct(combined_2$`Customer ID`, .keep_all = TRUE)

write.xlsx(combined, file = "combined.xlsx", sheetName = "Sheet1", row.names = FALSE)

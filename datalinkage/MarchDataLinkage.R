library(readxl)
library(xlsx)
library(tidyr)
library(sqldf)
library(tidyverse)
library(lubridate)
library(dplyr)
rm(list = ls())

#load data files
master_sales = read_excel('2019-03 BOB.xlsx', sheet = 1)
voucher_eid = read_excel('2019-03 BOB.xlsx', sheet = 2)
salesrule_office = read_excel('2019-03 BOB.xlsx', sheet = 3)
productsku_boxtype = read_excel('2019-03 BOB.xlsx', sheet = 4)
shifts_worked = read_excel('2019-03 Humanity.xlsx')
sales_expense = read_excel('2019-03 SEL.xlsx')

#format dates
#remove columns
master_sales = master_sales[, -c(3,4,6,8)]

#link eid to voucher
master_sales = merge(master_sales, voucher_eid, by.x = "voucher_code", by.y = "voucher_code", all.x = TRUE, all.y = FALSE)
master_sales = master_sales[, -c(1)]
#link salesrule to office
master_sales = merge(master_sales, salesrule_office, by.x = "sales_rule_set", by.y = "sales_rule_set", all.x = TRUE, all.y = FALSE)
master_sales = master_sales[, -c(1)]
#link productsku to box type
master_sales = merge(master_sales, productsku_boxtype, by.x = "first_product_sku", by.y = "first_product_sku", all.x = TRUE, all.y = FALSE)
master_sales = master_sales[, -c(1)]
write.xlsx(master_sales, file = "mastersales.xlsx", sheetName = "Sheet1", row.names = FALSE)

#clean shifts worked sheet - format the start day, start time, end time
shifts_worked = shifts_worked[!is.na(shifts_worked$eid) & !is.na(shifts_worked$location) &
                                !is.na(shifts_worked$shift_title),]
shifts_worked = shifts_worked %>% separate(`start_time`, into = c(NA, "start_time"), sep = " ")
shifts_worked = shifts_worked %>% separate(`end_time`, into = c(NA, "end_time"), sep = " ")
# shifts_worked$start_time = as.character(shifts_worked$start_time)
# shifts_worked$end_time = as.character(shifts_worked$end_time)

#create end day field
shifts_worked$end_day = shifts_worked$start_day

#prepare start and end time for time zone conversions
shifts_worked$DateSTime = paste(shifts_worked$start_day, " ", shifts_worked$start_time)
shifts_worked$DateETime = paste(shifts_worked$end_day, " ", shifts_worked$end_time)
# shifts_worked$DateSTime = strptime(as.character(shifts_worked$DateSTime), "%Y-%m-%d %H:%M")
# shifts_worked$DateETime = strptime(as.character(shifts_worked$DateETime), "%Y-%m-%d %H:%M")

#timezone conversion code
daylight_savings = as.Date("2019-03-10")
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

for (i in 1:length(shifts_worked$start_day)) {
  if (shifts_worked$start_time[i] > shifts_worked$end_time[i]) {
    shifts_worked$end_day[i] = shifts_worked$end_day[i] + 1
  }
}

shifts_worked$start_day = as.character(shifts_worked$start_day)
shifts_worked$end_day = as.character(shifts_worked$end_day)
shifts_worked = shifts_worked %>% separate(`end_day`, into = c("end_day", NA), sep = " ")

#reformat times for range comparison
shifts_worked$DateSTime = as.character(shifts_worked$DateSTime)
shifts_worked$DateETime = as.character(shifts_worked$DateETime)

#make the time stamp conversion for sales
#utc = as.POSIXct(master_sales$date_sign_up, tz = "PST8PDT")
#master_sales$date_sign_up = format(utc, tz = "EST5EDT")

#master_sales$date_sign_up = as.character(master_sales$date_sign_up)
master_sales$date_sign_up = as.POSIXct(master_sales$date_sign_up)
for (i in 1:length(master_sales$date_sign_up)) {
  if (master_sales$date_sign_up[i] < daylight_savings) {
   master_sales$date_sign_up[i] = master_sales$date_sign_up[i] + 3600
  }
}
master_sales = master_sales %>% separate(`date_sign_up`, into = c("Date", "Time"), sep = " ")
rm(productsku_boxtype)
rm(salesrule_office)
rm(voucher_eid)

#finally merge sales with shiftsworked
combined_2 = merge(master_sales, shifts_worked, by.x = c("Employee ID", "Office", "Date"), 
                   by.y = c("eid", "location", "start_day"), all = TRUE)

rm(master_sales, shifts_worked)

#clean the combined sheet of incomplete entries
combined_2[is.na(combined_2)] = 0
combined_2 = combined_2[combined_2$`Employee ID` > 0 & combined_2$`Office` != 0 &
                          combined_2$Date != 0 & combined_2$shift_title != "0", ]

#create DateTime
combined_2$DateTime = paste(combined_2$Date, " ", combined_2$Time)

#filter out the shifts that generated 0 sales into a separate sheet for later
nosale = combined_2[combined_2$customer_id == 0,]
nosale = nosale %>% distinct(nosale$`Employee ID`, nosale$`Office`, nosale$shift_title, .keep_all = TRUE)

#filter out no sale shifts to prepare for time comparison
combined_2 = combined_2[combined_2$customer_id != 0, ]

#time range fitting
#combined_2 = combined_2[!(combined_2$DateTime <= combined_2$DateSTime | combined_2$DateTime >= combined_2$DateETime),]
#remove duplicate sales entries
#combined_2 = combined_2 %>% distinct(combined_2$`customer_id`, combined_2$DateTime, .keep_all = TRUE)
#remove irrelevant columns
combined_2 = combined_2[, -c(3, 4, 7, 8)]

#dupls = combined_2[duplicated(combined_2$`Customer ID`),]

nosale = nosale[, -c(3, 4, 7, 8, 16, 17, 18)]

#rejoin the no sale data with time range filtered sales
saleandnosale = rbind(combined_2, nosale)
rm(combined_2, nosale)

saleandnosale$numsales = 0
for (i in 1:length(saleandnosale$numsales)) {
  if (saleandnosale$customer_id[i] != 0) {
    saleandnosale$numsales[i] = 1
  }
}

sales_expense = sales_expense[, -c(1:9, 11:16, 18)]

#condense the free event sales by grouping by shift, aggregating number of sales and breakdown of sales per shift
agged = saleandnosale %>% group_by(`Employee ID`, shift_title, `Office`, DateSTime, DateETime, total_time, event_id) %>%
  summarise(num_sales = sum(`numsales`),
            twoxtwo = sum(`Box Type` == "2x2"),
            thrxtwo = sum(`Box Type` == "3x2"),
            twoxfour = sum(`Box Type` == "2x4"),
            fourxtwo = sum(`Box Type` == "4x2"),
            thrxfour = sum(`Box Type` == "3x4"))
agged$wage_cost = 13.72 * agged$total_time
agged$wage_cost = round(agged$wage_cost, digits = 2)
PERSHIFTSUMMARY = as.data.frame(agged)
rm(saleandnosale, agged)

write.xlsx(PERSHIFTSUMMARY, file = "PERSHIFTSUMMARY.xlsx", sheetName = "Sheet1", row.names = FALSE)

#summarizing the free events per unique event shift title (combines all shifts for a single event into one to track sales)
PEREVENTSUMMARY = PERSHIFTSUMMARY %>% group_by(shift_title, `Office`, event_id) %>% 
  summarise(wage_exp = sum(wage_cost), total_hrs = sum(total_time), num_sales = sum(num_sales),
            twoxtwo = sum(twoxtwo), thrxtwo = sum(thrxtwo), twoxfour = sum(twoxfour), fourxtwo = sum(fourxtwo), thrxfour = sum(thrxfour))
PEREVENTSUMMARY = data.frame(PEREVENTSUMMARY)

combinedsheet = merge(PEREVENTSUMMARY, sales_expense, by.x = "event_id", by.y = "Event ID", all.x = TRUE, all.y = FALSE)
combinedsheet$`Amortized Cost`[is.na(combinedsheet$`Amortized Cost`)] = 0
combinedsheet$total_cost = combinedsheet$wage_exp + combinedsheet$`Amortized Cost`
write.xlsx(combinedsheet, file = "combinedsheet.xlsx", sheetName = "Sheet1", row.names = FALSE)

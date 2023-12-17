# install.packages("devtools")
# devtools::install_github("pridiltal/staplr")
library(staplr)
library(openxlsx)
library(janitor)
library(data.table)
library(lubridate)

emp_name = "..."
initial="..."
emp_id = "..."
dept = "..."
program = "..."
comment = "..."

setwd("~/data/smith_timesheet")
timesheet = read.csv("timesheet_temp.csv")
setDT(timesheet)
remove_empty(timesheet,  which = "rows")

timesheet[timesheet==""]<-NA
timesheet<-timesheet[complete.cases(timesheet),]

timesheet[, wk_day := wday(dmy(day))]
timesheet[, wks := substr(wk, 6, 7) ]
timesheet[, wk := NULL]
timesheet[, day_readable := NULL]
 
enum = data.table(
  c("sun_1", "from_sun_1", "to_sun_1", "total_sun_1", "cmt_sun_1", "pg_sun_1"), 
  c("mon_1", "from_mon_1", "to_mon_1", "total_mon_1", "cmt_mon_1", "pg_mon_1"), 
  c("tue_1", "from_tue_1", "to_tue_1", "total_tue_1", "cmt_tue_1", "pg_tue_1"), 
  c("wed_1", "from_wed_1", "to_wed_1", "total_wed_1", "cmt_wed_1", "pg_wed_1"), 
  c("thur_1", "from_thur_1", "to_thur_1", "total_thur_1", "cmt_thur_1", "pg_thur_1"), 
  c("fri_1", "from_fri_1", "to_fri_1", "total_fri_1", "cmt_fri_1", "pg_fri_1"), 
  c("sat_1", "from_sat_1", "to_sat_1", "total_sat_1", "cmt_sat_1", "pg_sat_1"), 
  c("sun_2", "from_sun_2", "to_sun_2", "total_sun_2", "cmt_sun_2", "pg_sun_2"),
  c("mon_2", "from_mon_2", "to_mon_2", "total_mon_2", "cmt_mon_2", "pg_mon_2"),
  c("tue_2", "from_tue_2", "to_tue_2", "total_tue_2", "cmt_tue_2", "pg_tue_2"),
  c("wed_2", "from_wed_2", "to_wed_2", "total_wed_2", "cmt_wed_2", "pg_wed_2"),
  c("thur_2", "from_thur_2", "to_thur_2", "total_thur_2", "cmt_thur_2", "pg_thur_2"),
  c("fri_2", "from_fri_2", "to_fri_2", "total_fri_2", "cmt_fri_2", "pg_fri_2"),
  c("sat_2", "from_sat_2", "to_sat_2", "total_sat_2", "cmt_sat_2", "pg_sat_2")
)

enum=transpose(enum)

# p = "2023-25" 
# period_dt= timesheet[period==p]

each_period = function(period_dt,period_text,cmt, pg, excel=FALSE) {
  
  if (excel) {
    excel_temp = "timesheet_fillable.xlsx"
    wb = loadWorkbook(excel_temp)
    }
  
  enum_cp = enum

  for (i in transpose(period_dt)) {
    # i = transpose(period_dt)[,V1]
    
    if (i[7] == "1") {
      if (i[6] == "1") {
        enum_cp[1] = data.table(i[2], i[3], i[4], i[5], cmt, pg, stringsAsFactors = TRUE)
        if (excel) writeData(wb, sheet=1, 
                             x=matrix(c(i[2], i[3], i[4], i[5], i[3], i[4], i[5], cmt, pg, initial), nrow=1), 
                             startRow = 5, startCol = 2, colNames = FALSE)
      }
      if (i[6] == "2" ){
        enum_cp[2] = data.table(i[2], i[3], i[4], i[5], cmt, pg, stringsAsFactors = TRUE)
        if (excel) writeData(wb, sheet=1, 
                                  x=matrix(c(i[2], i[3], i[4], i[5], i[3], i[4], i[5], cmt, pg, initial), nrow=1), 
                                  startRow = 6, startCol = 2, colNames = FALSE)
      } 
      if (i[6] == "3") {
        enum_cp[3] = data.table(i[2], i[3], i[4], i[5], cmt, pg, stringsAsFactors = TRUE)
        if (excel) writeData(wb, sheet=1, 
                                  x=matrix(c(i[2], i[3], i[4], i[5], i[3], i[4], i[5], cmt, pg, initial), nrow=1), 
                                  startRow = 7, startCol = 2, colNames = FALSE)
      } 
      if (i[6] == "4") {
        enum_cp[4] = data.table(i[2], i[3], i[4], i[5], cmt, pg, stringsAsFactors = TRUE)
        if (excel) writeData(wb, sheet=1, 
                                  x=matrix(c(i[2], i[3], i[4], i[5], i[3], i[4], i[5], cmt, pg, initial), nrow=1), 
                                  startRow = 8, startCol = 2, colNames = FALSE)
      } 
      if( i[6] == "5") {
        enum_cp[5] = data.table(i[2], i[3], i[4], i[5], cmt, pg, stringsAsFactors = TRUE)
        if (excel) writeData(wb, sheet=1, 
                                  x=matrix(c(i[2], i[3], i[4], i[5], i[3], i[4], i[5], cmt, pg, initial), nrow=1), 
                                  startRow = 9, startCol = 2, colNames = FALSE)
      } 
      if( i[6] == "6") {
        enum_cp[6] = data.table(i[2], i[3], i[4], i[5], cmt, pg, stringsAsFactors = TRUE)
        if (excel) writeData(wb, sheet=1, 
                                  x=matrix(c(i[2], i[3], i[4], i[5], i[3], i[4], i[5], cmt, pg, initial), nrow=1), 
                                  startRow = 10, startCol = 2, colNames = FALSE)
      } 
      if( i[6] == "7") {
        enum_cp[7] = data.table(i[2], i[3], i[4], i[5], cmt, pg, stringsAsFactors = TRUE)
        if (excel) writeData(wb, sheet=1, 
                                  x=matrix(c(i[2], i[3], i[4], i[5], i[3], i[4], i[5], cmt, pg, initial), nrow=1), 
                                  startRow = 11, startCol = 2, colNames = FALSE)
      } 
    }
    if (i[7] == "2") {
      if (i[6] == "1") {
        enum_cp[8] = data.table(i[2], i[3], i[4], i[5], cmt, pg, stringsAsFactors = TRUE)
        if (excel) writeData(wb, sheet=1, 
                             x=matrix(c(i[2], i[3], i[4], i[5], i[3], i[4], i[5], cmt, pg, initial), nrow=1), 
                             startRow = 13, startCol = 2, colNames = FALSE)
      } 
      if (i[6] == "2" ){
        enum_cp[9] = data.table(i[2], i[3], i[4], i[5], cmt, pg, stringsAsFactors = TRUE)
        if (excel) writeData(wb, sheet=1, 
                             x=matrix(c(i[2], i[3], i[4], i[5], i[3], i[4], i[5], cmt, pg, initial), nrow=1), 
                             startRow = 14, startCol = 2, colNames = FALSE)
      } 
      if (i[6] == "3") {
        enum_cp[10] = data.table(i[2], i[3], i[4], i[5], cmt, pg, stringsAsFactors = TRUE)
        if (excel) writeData(wb, sheet=1, 
                             x=matrix(c(i[2], i[3], i[4], i[5], i[3], i[4], i[5], cmt, pg, initial), nrow=1), 
                             startRow = 15, startCol = 2, colNames = FALSE)
      } 
      if (i[6] == "4") {
        enum_cp[11] = data.table(i[2], i[3], i[4], i[5], cmt, pg, stringsAsFactors = TRUE)
        if (excel) writeData(wb, sheet=1, 
                             x=matrix(c(i[2], i[3], i[4], i[5], i[3], i[4], i[5], cmt, pg, initial), nrow=1), 
                             startRow = 16, startCol = 2, colNames = FALSE)
      } 
      if( i[6] == "5") {
        enum_cp[12] = data.table(i[2], i[3], i[4], i[5], cmt, pg, stringsAsFactors = TRUE)
        if (excel) writeData(wb, sheet=1, 
                             x=matrix(c(i[2], i[3], i[4], i[5], i[3], i[4], i[5], cmt, pg, initial), nrow=1), 
                             startRow = 17, startCol = 2, colNames = FALSE)
      } 
      if( i[6] == "6") {
        enum_cp[13] = data.table(i[2], i[3], i[4], i[5], cmt, pg, stringsAsFactors = TRUE)
        if (excel) writeData(wb, sheet=1, 
                             x=matrix(c(i[2], i[3], i[4], i[5], i[3], i[4], i[5], cmt, pg, initial), nrow=1), 
                             startRow = 18, startCol = 2, colNames = FALSE)
      } 
      if( i[6] == "7") {
        enum_cp[14] = data.table(i[2], i[3], i[4], i[5], cmt, pg, stringsAsFactors = TRUE)
        if (excel) writeData(wb, sheet=1, 
                             x=matrix(c(i[2], i[3], i[4], i[5], i[3], i[4], i[5], cmt, pg, initial), nrow=1), 
                             startRow = 19, startCol = 2, colNames = FALSE)
      } 
    }
}
  
  
  fields = get_fields("timesheet_fillable.pdf")
  
  for (j in 1:nrow(enum)) {
    if (all(enum[j] == enum_cp[j]) == FALSE) {
      for (k in 1:length(enum[j])) {
        fields[[enum[[k]][j]]]$value = enum_cp[[k]][j]
      }
    }
  }
  
  fields$name$value = emp_name
  fields$id$value = emp_id
  fields$dept$value = dept
  fields$period$value = period_text
  
  wk_1_total = sum(period_dt[wks=="1", hr])
  fields$total_wk_1$value = wk_1_total
  wk_2_total = sum(period_dt[wks=="2", hr])
  fields$total_wk_2$value = wk_2_total
  t_total = wk_1_total+wk_2_total
  fields$total$value = t_total
  
  if (excel) {
    writeData(wb, sheet=1, x=emp_name, startRow = 2, startCol = 2, colNames = FALSE)
    writeData(wb, sheet=1, x=emp_id, startRow = 2, startCol = 5, colNames = FALSE)
    writeData(wb, sheet=1, x=dept, startRow = 2, startCol = 9, colNames = FALSE)
    writeData(wb, sheet=1, x=period_text, startRow = 2, startCol = 12, colNames = FALSE)
    
    writeData(wb, sheet=1, x=wk_1_total, startRow = 12, startCol = 5, colNames = FALSE)
    writeData(wb, sheet=1, x=wk_1_total, startRow = 12, startCol = 8, colNames = FALSE)
    writeData(wb, sheet=1, x=wk_2_total, startRow = 20, startCol = 5, colNames = FALSE)
    writeData(wb, sheet=1, x=wk_2_total, startRow = 20, startCol = 8, colNames = FALSE)
    writeData(wb, sheet=1, x=t_total, startRow = 21, startCol = 5, colNames = FALSE)
    writeData(wb, sheet=1, x=t_total, startRow = 21, startCol = 8, colNames = FALSE)
  }
  
  if (excel) {
    new_file_name = paste0("timesheet_", period_text, ".xlsx")
    saveWorkbook(wb, new_file_name, overwrite = TRUE)
    
  } else {
    new_file_name = paste0("timesheet_", period_text, ".pdf")
    set_fields("timesheet_fillable.pdf", new_file_name, fields)
  }
    
}

for (p in unique(timesheet[, period])) {
  print(timesheet[period == p])
  print("-----")
  each_period(timesheet[period == p], p, comment, program, excel=TRUE)
}


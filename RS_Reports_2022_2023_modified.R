
#************************************************************************************************
#                            
#                         RUN TIME : 20 MINS 
# 
#                       
#
#************************************************************************************************

#########################################################
#  ODBC DRIVERS NEEDED TO RUN PROGRAM 
######################################################### 
#  SEO_MART --> ES00VPADOSQL180,51433
#  MUST HAVE RTOOLS DOWNLOADED
#########################################################
gc()

start<- Sys.time()

library(sqldf)
library(tcltk)
library(xlsx)
library("XLConnect")
library (RODBC)
library (excel.link)
library(SOAR)
library(plyr)
library(data.table)
library(dplyr) 
library(gridExtra)



##################################################################################
# DATA PREP FOR REPORTS
#
#
##################################################################################

seo <- odbcConnect("SEO_MART")
qry_report <- ("SELECT * from [SEO_MART].[dbo].[RPT_RSProvisioning] where SUBSTRING(EnrolledDBN,1,2)<'84'")
comp_report <- ("SELECT * from [SEO_MART].[dbo].[RPT_RSCompliance] where SchoolDistrict<84")
comp_report2 <- ("SELECT * from [SEO_MART].[dbo].[RPT_RSCompliancebyGroup]")
comp_report3 <- ("SELECT * from [SEO_MART].[dbo].[RPT_Locations]")
report_RS <- sqlQuery (seo, qry_report)
report_citywide <- sqlQuery(seo, comp_report)
report_citywide2 <- sqlQuery(seo, comp_report2)
report_lcgms <- sqlQuery(seo,comp_report3)
close(seo)


#convert dates from SQL to m/d/y format
#create copy of table

report_RS_p2 <-report_RS



#--load lubridate for 'is.POSIXct' and 'date' functions

library(lubridate)

is.POSIXct(report_RS$dtimes)

str(report_RS)

#--load dplyr for 'mutate_if' function and 'filter' function

library(dplyr)

report_RS_p2  <- report_RS_p2 %>% mutate_if(is.POSIXct, date)


# replace original table
report_RS <-report_RS_p2
str(report_RS)

end <- Sys.time()
start
end 

gc()

 
report_RS_Temp1 <-sqldf("Select StudentID, LastName||','||FirstName as STUDENTNAME
                    ,round((AttendRate*100),0)||'%' as Attendrate
                        ,ServiceType
                        ,RecommendedGroupSizeNumeric
                        ,RecommendedFrequencyNumeric
                        ,RecommendedDurationNumeric
                        ,RSMandateLanguage
                        ,EnrolledDBN EnrolledDBN
                        ,GradeLevel Grade
                        ,BirthDate
                        ,EffectiveOutcomeDate --IEPConferenceDate
                        ,RecentAuthorizationDate
                        ,PhysicalLocation 
                        ,PhysicalLocationName
                        ,PhysicalLocationZipCode
                        ,MandateType 
                        ,FirstEncounterDate
                        ,PAFirstPartialAttendDate                                 
                        ,SESISFirstPartialEncounterDate
                        ,TotalPartialEncounters
                        ,SESISLastPartialEncounterDate
                        ,SUBSTR(EnrolledDBN,1,2) as district
                        ,FirmName as Agency
                        ,EncounterProvider
                        ,ProcessedDate
                        from report_RS ")


report_RS_Temp <-sqldf("Select StudentId, STUDENTNAME
                    ,AttendRate
                    ,ServiceType 
                    ,RecommendedGroupSizeNumeric
                    ,RecommendedFrequencyNumeric
                    ,RecommendedDurationNumeric
                    ,RSMandateLanguage
                    ,EnrolledDBN
                    ,Grade
                    ,BirthDate
                    ,EffectiveOutcomeDate--,IEPConferenceDate
                    ,RecentAuthorizationDate
                    ,PhysicalLocation 
                    ,PhysicalLocationName
                    ,PhysicalLocationZipCode 
                    ,MandateType 
                    ,FIrstEncounterDate
                    ,PAFirstPartialAttendDate                                 
                    ,SESISFirstPartialEncounterDate
                    ,TotalPartialEncounters 
                    ,SESISLastPartialEncounterDate
                    ,Agency
                    ,EncounterProvider
                    ,ProcessedDate
                    from report_RS_Temp1 where district < 84 ")


report_RS_comp <-sqldf("Select EnrolledDBN
                    ,FieldSupportCenterReportingName
                    ,SuperintendentName
                    ,NotEncounteredCounsel
                    ,EncounteredCounsel
                    ,PercentageEncounteredCounsel
                    ,NotEncounteredMajor
                    ,EncounteredMajor
                    ,PercentageEncounteredMajor
                    ,PercentageEncounteredAllRS
                    from report_citywide ")




####################################################################################################################
### Creating different table for District level, FSC level, Superintendent level and DBN level FOR cONSOLIDATED FILE
####################################################################################################################


report_RS_comp2 <-sqldf("Select EnrolledDBN
                    ,b.SuperintendentDistrict
                    ,a.SuperintendentName
                    ,NotEncounteredCounsel
                    ,EncounteredCounsel
                    ,PercentageEncounteredCounsel
                    ,NotEncounteredMajor
                    ,EncounteredMajor
                    ,PercentageEncounteredMajor
                    ,PercentageEncounteredAllRS
                    from report_citywide a
                       left join report_lcgms b
                       on a.EnrolledDBN=b.SchoolDBN")


report_RS_dbn <-sqldf("Select EnrolledDBN
                    ,SuperintendentDistrict
                    ,SuperintendentName
                    ,NotEncounteredCounsel
                    ,EncounteredCounsel
                    ,PercentageEncounteredCounsel
                    ,NotEncounteredMajor
                    ,EncounteredMajor
                    ,PercentageEncounteredMajor
                    ,PercentageEncounteredAllRS
                    from report_RS_comp2  ")

report_RS_dst <-sqldf("Select 
                     SchoolDistrict
                    ,NotEncounteredCounsel
                    ,EncounteredCounsel
                    ,PercentageEncounteredCounsel
                    ,NotEncounteredMajor
                    ,EncounteredMajor
                    ,PercentageEncounteredMajor
                    ,PercentageEncounteredAllRS
                    from report_citywide2 where  ReportGroupDesc='SchoolDistrict' order by SchoolDistrict ")

report_RS_sup <-sqldf("Select SuperintendentName
                    ,NotEncounteredCounsel
                    ,EncounteredCounsel
                    ,PercentageEncounteredCounsel
                    ,NotEncounteredMajor
                    ,EncounteredMajor
                    ,PercentageEncounteredMajor
                    ,PercentageEncounteredAllRS 
                    from report_citywide2 where ReportGroupDesc='Superintendent' order by SuperintendentName ")


##where SuperintendentName!='' and FIELD_SUPPORT_CENTER_REPORTING_NAME!='' 

#To display the asofdate in the report
report_asofdt <-sqldf("Select ProcessedDate
                    from report_RS LIMIT 1")



dt <- Sys.Date()




###########################################################################
#Output Excel reports - WEEKLY SCHOOL LEVEL WITH 2 TABS
###########################################################################
#rm(wb)
startr <- Sys.time()
library(openxlsx)
Sys.setenv("R_ZIPCMD" = "C:/Rtools/bin/zip.exe")

#rm(pth)
#rm(pth1)
#rm(pth2)
dt <- Sys.Date()
dt <- format(dt, format="%Y/%m/%d")
dt<-gsub("[^A-Za-z0-9]", "", dt)
dt
pth2 <- paste("R:/SEO Analytics/Reporting/Related Services/Output Files/SY 22-23/MandatedServices_", dt, sep="") 
pth2
dir.create(pth2) 


mydata <- report_RS_Temp
mycomp <- report_RS_comp
asofdt <- report_asofdt
gc()
### Output files based on a variable 
dt <- Sys.Date()
dt <- format(dt, format="%m-%d-20%y")
varNames <- unique(mycomp$EnrolledDBN)
#varNames <- '28Q350'
varNames
for(i in varNames) {
  print(i)
  mydata2 <- dplyr::filter(mydata, EnrolledDBN== i)
  mycomp2 <- dplyr::filter(mycomp, EnrolledDBN== i)
  wb <- openxlsx::loadWorkbook("C:/Template/RS_Template_new.xlsx")
  
  openxlsx::writeData(wb, "Data", mydata2, startCol = 1, startRow = 6, xy = NULL,
                      colNames = FALSE, rowNames = FALSE, headerStyle = NULL,  
                      borders = c("none", "surrounding", "rows", "columns", "all"),
                      borderColour = getOption("openxlsx.borderColour", "black"),
                      borderStyle = getOption("openxlsx.borderStyle", "thin"),
                      withFilter = FALSE, keepNA = FALSE)
  openxlsx::writeData(wb, "Data", asofdt, startCol = 2, startRow = 1, xy = NULL,
                      colNames = FALSE, rowNames = FALSE, headerStyle = NULL,  
                      borders = c("none", "surrounding", "rows", "columns", "all"),
                      borderColour = getOption("openxlsx.borderColour", "black"),
                      borderStyle = getOption("openxlsx.borderStyle", "thin"),
                      withFilter = FALSE, keepNA = FALSE)
  openxlsx::writeData(wb, "Completion Reports", mycomp2, startCol = 1, startRow = 5, xy = NULL,
                      colNames = FALSE, rowNames = FALSE, headerStyle = NULL,  
                        borders = c("none", "surrounding", "rows", "columns", "all"),
                      borderColour = getOption("openxlsx.borderColour", "black"),
                      borderStyle = getOption("openxlsx.borderStyle", "thin"),
                      withFilter = FALSE, keepNA = FALSE)
  openxlsx::writeData(wb, "Completion Reports", asofdt, startCol = 3, startRow = 1, xy = NULL,
                      colNames = FALSE, rowNames = FALSE, headerStyle = NULL,  
                      borders = c("none", "surrounding", "rows", "columns", "all"),
                      borderColour = getOption("openxlsx.borderColour", "black"),
                      borderStyle = getOption("openxlsx.borderStyle", "thin"),
                      withFilter = FALSE, keepNA = FALSE)
  pth <- paste(pth2,"/",i, "_MandatedServices_", dt,".xlsx", sep="")
  saveWorkbook(wb, pth , overwrite = TRUE) 
}


end <- Sys.time()
start





###########################################################################
#Output Consolidated Excel reports - WITH 4 TABS
###########################################################################
rm(wb)
startr <- Sys.time()
library(openxlsx)
Sys.setenv("R_ZIPCMD" = "C:/Rtools/bin/zip.exe")

rm(pth)
rm(pth2)
rm(pth3)
rm(pth4)
dt <- Sys.Date()
dt <- format(dt, format="%Y/%m/%d")
dt<-gsub("[^A-Za-z0-9]", "", dt)
dt
pth2 <- paste("R:/SEO Analytics/Share/Related Services/", dt, sep="") 
pth2
dir.create(pth2) 



mycomp1 <- report_RS_dbn
mycomp2 <- report_RS_dst
mycomp3 <- report_RS_sup
asofdt  <- report_asofdt

gc()
 
  wb <- openxlsx::loadWorkbook("C:/Template/RS_Compliance_new.xlsx")
  
  openxlsx::writeData(wb, "RS Supt Citywide Summary", mycomp1, startCol = 1, startRow = 8, xy = NULL,
                      colNames = FALSE, rowNames = FALSE, headerStyle = NULL,  
                      borders = c("none", "surrounding", "rows", "columns", "all"),
                      borderColour = getOption("openxlsx.borderColour", "black"),
                      borderStyle = getOption("openxlsx.borderStyle", "thin"),
                      withFilter = FALSE, keepNA = FALSE)
  openxlsx::writeData(wb, "RS Supt Citywide Summary", asofdt, startCol = 3, startRow = 1, xy = NULL,
                      colNames = FALSE, rowNames = FALSE, headerStyle = NULL,  
                      borders = c("none", "surrounding", "rows", "columns", "all"),
                      borderColour = getOption("openxlsx.borderColour", "black"),
                      borderStyle = getOption("openxlsx.borderStyle", "thin"),
                      withFilter = FALSE, keepNA = FALSE) 
openxlsx::writeData(wb, "RS District Summary", mycomp2, startCol = 1, startRow = 8, xy = NULL,
                    colNames = FALSE, rowNames = FALSE, headerStyle = NULL,  
                    borders = c("none", "surrounding", "rows", "columns", "all"),
                    borderColour = getOption("openxlsx.borderColour", "black"),
                    borderStyle = getOption("openxlsx.borderStyle", "thin"),
                    withFilter = FALSE, keepNA = FALSE)
openxlsx::writeData(wb, "RS District Summary", asofdt, startCol = 3, startRow = 1, xy = NULL,
                    colNames = FALSE, rowNames = FALSE, headerStyle = NULL,  
                    borders = c("none", "surrounding", "rows", "columns", "all"),
                    borderColour = getOption("openxlsx.borderColour", "black"),
                    borderStyle = getOption("openxlsx.borderStyle", "thin"),
                    withFilter = FALSE, keepNA = FALSE)
openxlsx::writeData(wb, "RS Superintendent Summary", mycomp3, startCol = 1, startRow = 8, xy = NULL,
                    colNames = FALSE, rowNames = FALSE, headerStyle = NULL,  
                    borders = c("none", "surrounding", "rows", "columns", "all"),
                    borderColour = getOption("openxlsx.borderColour", "black"),
                    borderStyle = getOption("openxlsx.borderStyle", "thin"),
                    withFilter = FALSE, keepNA = FALSE)
openxlsx::writeData(wb, "RS Superintendent Summary", asofdt, startCol = 3, startRow = 1, xy = NULL,
                    colNames = FALSE, rowNames = FALSE, headerStyle = NULL,  
                    borders = c("none", "surrounding", "rows", "columns", "all"),
                    borderColour = getOption("openxlsx.borderColour", "black"),
                    borderStyle = getOption("openxlsx.borderStyle", "thin"),
                    withFilter = FALSE, keepNA = FALSE)

  pth1 <- paste(pth2,"/","RS Compliance Report_", dt,".xlsx", sep="")
  ##pth4 <- paste(pth3,"/","RS Compliance Report_", dt,".xlsx", sep="")
  print(pth1)
  ##print(pth4)

  saveWorkbook(wb, pth1 , overwrite = TRUE) 
  ##saveWorkbook(wb, pth4 , overwrite = TRUE) 


end <- Sys.time()
start


#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
# RStudio Workbench is strictly for use by Public Health Scotland staff and     
# authorised users only, and is governed by the Acceptable Usage Policy https://github.com/Public-Health-Scotland/R-Resources/blob/master/posit_workbench_acceptable_use_policy.md.
#
# This is a shared resource and is hosted on a pay-as-you-go cloud computing
# platform.  Your usage will incur direct financial cost to Public Health
# Scotland.  As such, please ensure
#
#   1. that this session is appropriately sized with the minimum number of CPUs
#      and memory required for the size and scale of your analysis;
#   2. the code you write in this script is optimal and only writes out the
#      data required, nothing more.
#   3. you close this session when not in use; idle sessions still cost PHS
#      money!
#
# For further guidance, please see https://github.com/Public-Health-Scotland/R-Resources/blob/master/posit_workbench_best_practice_with_r.md.
#
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


library(dplyr)
library(openxlsx)
library(readxl)

### ESTABLISH ODBC CONNECTION TO DVPROD MANUALLY
conn <- odbc::dbConnect(odbc::odbc(),
                        dsn = "DVPROD", 
                        uid = paste(Sys.info()['user']),
                        pwd = .rs.askForPassword("password"))

reporting_start_date = as.Date("2024-09-01")
answer <- svDialogs::dlgInput(paste0("Hello, ",(Sys.info()['user']),". Do you want to report on only the latest 4 weeks of data? y/n"))
if (answer$res == "y" | answer$res == "Y") {
  reporting_start_date = as.Date(Sys.Date()-28)
  answer <- 1
} else {answer <- 0}

### LIST HEALTH BOARD CODES AND NAMES FOR CREATING HB REPORTS
hb_cypher <- c("A","B","Y","F","V","G","N","H","L","S","R","Z","T","W","NWTC","SAS")
hb_name <- c("NHS AYRSHIRE & ARRAN","NHS BORDERS","NHS DUMFRIES & GALLOWAY","NHS FIFE",
             "NHS FORTH VALLEY","NHS GREATER GLASGOW & CLYDE","NHS GRAMPIAN","NHS HIGHLAND",
             "NHS LANARKSHIRE","NHS LOTHIAN","NHS ORKNEY","NHS SHETLAND","NHS TAYSIDE",
             "NHS WESTERN ISLES","NATIONAL FACILITY","SCOTTISH AMBULANCE SERVICE")

########################################################################
#############################  SECTION B1 ##############################
############## VDL SYSTEM EXTRACT AND VACCINATION SUMMARIES ############
########################################################################
Vaxeventtot <- conn %>%
tbl(dbplyr::in_schema("vaccination", "vaccination_event_analysis")) %>%
select(vacc_type_target_disease,vacc_status,vacc_data_source,
vacc_event_created_at) %>%
mutate(vacc_event_created_at = substr(vacc_event_created_at,1,10)) %>%
group_by(vacc_status,vacc_type_target_disease,vacc_data_source,
vacc_event_created_at) %>%
summarise(number_of_vacc_events = n()) %>% collect()
saveRDS(Vaxeventtot,"Outputs/Temp/Vaxeventtot.rds")
# Vaxeventtot <- readRDS("Outputs/Temp/Vaxeventtot.rds")
### CREATE SUMMARY TABLE OF FLU VACC RECORDS
FluSystemSummary <- Vaxeventtot %>%
filter(vacc_type_target_disease == "Influenza (disorder)" &
vacc_event_created_at>=reporting_start_date) %>%
group_by(vacc_status,vacc_type_target_disease,vacc_data_source) %>%
summarise(record_count = sum(number_of_vacc_events))
rm(Vaxeventtot,CovSystemSummary)
gc()
########################################################################
###############################  SECTION C #############################
#############################  PATIENT DATA ############################
########################################################################
### EXTRACT VARIABLES IN vaccination_patient_analysis VIEW
Vaxpatientraw <- odbc::dbGetQuery(conn, "select
source_system_patient_id,
patient_derived_chi_number,
patient_derived_upi_number,
patient_given_name,
patient_family_name,
patient_date_of_birth,
patient_sex,
source_system_gp_practice_code,
source_system_gp_practice_text,
patient_chi_derived_status_desc,
patient_nrs_date_of_death,
patient_chi_temporary_resident_flag,
patient_chi_transfer_out_flag
from vaccination.vaccination_patient_analysis ")
### CLEAN INVALID CHARACTERS FROM FREE-TEXT FIELDS IN PATIENT ANALYSIS DATA
Vaxpatientraw$patient_family_name <- textclean::replace_non_ascii(Vaxpatientraw$patient_family_name,replacement = "")
Vaxpatientraw$patient_given_name <- textclean::replace_non_ascii(Vaxpatientraw$patient_given_name,replacement = "")
########################################################################
###############################  SECTION D2 ############################
############################  FLU VACCINATIONS #########################
########################################################################
### CREATE FLU VACCINATIONS DATAFRAME
#############################################################################################
### EXTRACT VARIABLES FROM COMPLETED FLU VACC RECORDS IN vaccination_event_analysis VIEW
FluVaxData <- odbc::dbGetQuery(conn, "select
source_system_patient_id,
patient_derived_encrypted_upi,
vacc_source_system_event_id,
vacc_type_target_disease,
vacc_event_created_at,
vacc_record_created_at,
vacc_occurence_time,
vacc_location_health_board_name,
vacc_location_name,
vacc_product_name,
vacc_batch_number,
vacc_performer_name,
vacc_dose_number,
vacc_booster,
vacc_data_source,
vacc_data_source_display,
age_at_vacc
from vaccination.vaccination_event_analysis
WHERE vacc_status='completed'
AND vacc_type_target_disease='Influenza (disorder)'
AND vacc_clinical_trial_flag='0' AND vacc_occurence_time>'2024-08-31' ")
### CLEAN INVALID CHARACTERS FROM FREE-TEXT FIELDS IN EVENT ANALYSIS DATA
FluVaxData$vacc_location_name <- textclean::replace_non_ascii(FluVaxData$vacc_location_name,replacement = "")
FluVaxData$vacc_performer_name <- textclean::replace_non_ascii(FluVaxData$vacc_performer_name,replacement = "")
### CREATE NEW COLUMN FOR DAYS BETWEEN VACCINATION AND RECORD CREATION
FluVaxData$vacc_record_date <- as.Date(substr(FluVaxData$vacc_record_created_at,1,10))
FluVaxData$days_between_vacc_and_recording <- as.double(difftime(FluVaxData$vacc_record_date,FluVaxData$vacc_occurence_time,units="days"))
### CREATE BLANK PLACEHOLDER COLUMNS
FluVaxData$dups_same_day_flag <- NA
FluVaxData$VMT_only_dups_flag <- NA
FluVaxData$"cov_booster_interval (days)" <- NA
### CALCULATE INTERVAL BETWEEN VACCINES PER PATIENT
FluVaxData <- FluVaxData %>%
arrange(patient_derived_encrypted_upi,vacc_occurence_time) %>%
group_by(patient_derived_encrypted_upi) %>%
mutate(dose_number = row_number()) %>% mutate(doses = n()) %>% ungroup() %>%
mutate(prev_vacc_date=lag(vacc_occurence_time)) %>%
mutate("flu_interval (days)"=as.integer(difftime(vacc_occurence_time,prev_vacc_date,units = "days")))
FluVaxData$"flu_interval (days)" <- ifelse(FluVaxData$dose_number=="1",NA,FluVaxData$"flu_interval (days)")
FluVaxData$"flu_interval (days)" [is.na(FluVaxData$patient_derived_encrypted_upi)] <- NA
FluVaxData$"hz_interval (days)" <- NA
FluVaxData$"pneum_vacc_interval (weeks)" <- NA
FluVaxData$sort_date <- NA
### CREATE FLU VACC DATAFRAME BY JOINING FLU VACC EVENT RECORDS WITH PATIENT RECORDS
FluVaxData <- FluVaxData %>%
inner_join(Vaxpatientraw, by=("source_system_patient_id")) %>%
arrange(desc(vacc_event_created_at)) %>% # sort data by date amended
select(-patient_derived_encrypted_upi,-vacc_record_date,-dose_number,-doses,-prev_vacc_date) %>%
mutate(Date_Administered = substr(vacc_occurence_time, 1, 10)) %>% # create new vacc date data item for date only (no time)
mutate(CHIcheck = phsmethods::chi_check(patient_derived_chi_number)) # create CHI check data item
### CREATE VACC OCCURENCE PHASE TO LINK COHORT DATA BY COHORT PHASE
FluVaxData$vacc_phase <- NA
FluVaxData$vacc_phase [between(FluVaxData$vacc_occurence_time,
as.Date("2021-09-01"),as.Date("2022-03-31"))] <-
"Tranche2"
FluVaxData$vacc_phase [between(FluVaxData$vacc_occurence_time,
as.Date("2022-09-01"),as.Date("2023-03-31"))] <-
"Autumn Winter 2022_23"
FluVaxData$vacc_phase [between(FluVaxData$vacc_occurence_time,
as.Date("2023-09-01"),as.Date("2024-03-31"))] <-
"Autumn Winter 2023_24"
FluVaxData$vacc_phase [between(FluVaxData$vacc_occurence_time,
as.Date("2024-09-01"),as.Date("2025-03-31"))] <-
"Autumn Winter 2024_25"
table(FluVaxData$vacc_phase,useNA = "ifany")
FluVaxData$vacc_phase [FluVaxData$vacc_occurence_time>=as.Date("2021-04-01") &
is.na(FluVaxData$vacc_phase)] <-
"Outwith Autumn/Winter"
table(FluVaxData$vacc_phase,useNA = "ifany")
table(FluVaxData$vacc_product_name)

######################################

fluenz <- FluVaxData %>% 
  filter(grepl("Fluenz", vacc_product_name))

######################################

### CREATE TABLE OF RECORDS & SUMMARY OF LAIV (FLUENZ) VACCINATION GIVEN AT AGE <2 OR >18
flu_laiv_ageDQ <- FluVaxData %>%
  filter(grepl("Fluenz", vacc_product_name)) %>% 
  filter(age_at_vacc < 2 | age_at_vacc > 18) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "Add01. FLU LAIV given age <2 or >18")

flu_laiv_ageDQSumm <- flu_laiv_ageDQ %>%
  group_by(vacc_location_health_board_name,vacc_location_name,
           vacc_product_name,Date_Administered, age_at_vacc) %>%
  summarise(record_count = n())

flu_laiv_ageDQ <- flu_laiv_ageDQ %>% select(-Date_Administered)

######################################

### CREATE TABLE OF RECORDS & SUMMARY OF AQIV VACCINATION GIVEN AT AGE <50
flu_aqiv_ageDQ <- FluVaxData %>%
  filter(vacc_product_name=="Adjuvanted Quadrivalent Influenza Vaccine Seqirus") %>% 
  filter(age_at_vacc < 50) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "Add02. FLU AQIV given age <50")

flu_aqiv_ageDQSumm <- flu_aqiv_ageDQ %>%
  group_by(vacc_location_health_board_name, vacc_location_name, Date_Administered, age_at_vacc) %>%
  summarise(record_count = n())

flu_aqiv_ageDQ <- flu_aqiv_ageDQ %>% select(-Date_Administered)

######################################

### CREATE TABLE OF RECORDS & SUMMARY OF CQIV VACCINATION GIVEN AT AGE >64
flu_cqiv_ageDQ <- FluVaxData %>%
  filter(vacc_product_name=="Cell-based Quadrivalent Influenza Vaccine Seqirus") %>% 
  filter(age_at_vacc > 64) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "Add03. FLU CQIV given age >64")

flu_cqiv_ageDQSumm <- flu_cqiv_ageDQ %>%
  group_by(vacc_location_health_board_name, vacc_location_name, Date_Administered, age_at_vacc) %>%
  summarise(record_count = n())

flu_cqiv_ageDQ <- flu_cqiv_ageDQ %>% select(-Date_Administered)

######################################

#Collates pivot tables into a single report
flu_add_SummaryReport <-
  list("System Summary" = FluSystemSummary,
       "LAIV aged <2 or >18" = flu_laiv_ageDQSumm,
       "AQIV aged <65" = flu_aqiv_ageDQSumm)

write.xlsx(flu_add_SummaryReport,
           paste("Outputs/DQ Summary Reports/Flu_Vacc_DQ_2024-25_addendum_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
           asTable = TRUE,
           colWidths = "auto")


######################################

flu_vacc <- rbind(flu_laiv_ageDQ,flu_aqiv_ageDQ)

saveRDS(flu_vacc,"Outputs/Temp/flu_vacc_addendum.rds")

######################################

### CREATE UNIQUE DQ ID FOR ALL RECORDS IN flu_vacc TABLE
flu_vacc$DQ_ID <- paste(substr(flu_vacc$QueryName,1,5),
                        flu_vacc$vacc_source_system_event_id,
                        substr(flu_vacc$vacc_event_created_at,1,10),sep = ".")

### TIDY UP flu_vacc DATA TABLE
flu_vacc <- flu_vacc %>% arrange(desc(sort_date),DQ_ID,desc(vacc_event_created_at)) %>% 
  select(QueryName,DQ_ID,patient_derived_chi_number,patient_derived_upi_number,
         everything())

### FINALISE DATA TABLES FOR HB REPORTS
flu_vacc <- flu_vacc %>% select(-source_system_patient_id,-sort_date,-vacc_phase)


######################################

### CREATE DQ OUTPUTS FOR EACH HEALTH BOARD AND SAVE TO FOLDERS

for(i in 1:16) {
  
  df_flu <- flu_vacc %>% filter(vacc_location_health_board_name == hb_name[i])
  
  HBReport <-
    list("Flu Vacc Addendum - Q01-02" = df_flu)
  
  HBReportWB <- buildWorkbook(HBReport, asTable = TRUE)
  setColWidths(HBReportWB,sheet = 1,cols = 1,widths = "auto")
  
  
  saveWorkbook(HBReportWB,paste("Outputs/DQ HB Reports/",hb_cypher[i],"_Vacc_Addendum_DQ_Report_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep=""))
  
  rm(HBReport,HBReportWB)
} 


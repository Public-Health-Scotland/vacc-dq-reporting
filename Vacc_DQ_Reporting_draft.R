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


########################################################################
#################  VACCINATIONS DATA QUALITY REPORTS ###################
########################################################################

### Script to produce weekly PHS vaccination reports based on Covid-19, Flu,
### Herpes Zoster (shingles) & Pneumococcal vaccinations data in the
### Vaccination Data Layer (VDL).
#
### Sections A, B1 and C need to be run before sections D1-D4, which can then 
### be selected to run based on the vaccination type of choice if required to 
### run individually.
#
#   Section A: load libraries; connect to DVPROD; set reporting period; 
#              list health boards for outputs.
#
#   Section B1: extract data for VDL system summary of all 
#              data records; create vacc summaries per vacc type.
#
#   Section B2: create and save out new VDL system summary. 
#
#   Section C: extract and clean data from Patient Analysis view
#
#   Section D1: extract and clean covid-19 vacc data; link to patient data to 
#               create shingles dataframe; create covid-19 DQ queries.
#
#   Section D2: extract and clean flu vacc data; link to patient data
#               to create pneumococcal dataframe; create flu DQ queries.
#
#   Section D3: extract and clean shingles vacc data; link to patient data to 
#               create flu dataframe; create shingles DQ queries.
#
#   Section D4: extract and clean pneumococcal-19 vacc data; link to patient  
#               data to create covid-19 dataframe; create pneumococcal DQ queries.
#
#   Section E:  create unique DQ IDs for each query; sort and order data in each
#               table; collate all queries into new df.
#
#   Section F:  create and save out individual HB reports.
#
#   Section G:  update Query Count spreadsheet to compare query figures with 
#               previous run's figures.

start.time <- Sys.time()

########################################################################
###############################  SECTION A #############################
###  LIBRARIES, CONNECTION, SET REPORTING PERIOD, LIST HEALTH BOARDS ###
########################################################################

### INSTALL PACKAGES - FIRST TIME USERS MAY NEED TO INSTALL PACKAGES
### BEFORE LOADING
#install.packages("svDialogs")

### LOAD PACKAGES TO LIBRARIES - HASHED OUT PACKAGES ARE USED IN SCRIPT BUT WITH
### FUNCTION CALL SO DON'T NEED LOADED AT START EXCEPT AS CHECK FOR 1ST TIME USERS
library(dplyr)
library(openxlsx)
library(readxl)
# library(writexl)
# library(lubridate)
# library(textclean)
# library(odbc)
# library(phsmethods)
# library(svDialogs)

### ESTABLISH ODBC CONNECTION TO DVPROD MANUALLY
conn <- odbc::dbConnect(odbc::odbc(),
                  dsn = "DVPROD", 
                  uid = paste(Sys.info()['user']),
                  pwd = .rs.askForPassword("password"))

reporting_start_date = as.Date("2021-01-01")
answer <- svDialogs::dlgInput(paste0("Hello, ",(Sys.info()['user']),". Do you want to report on only the latest 4 weeks of data? y/n"))
if (answer$res == "y" | answer$res == "Y") {
  reporting_start_date = as.Date(Sys.Date()-28)
  answer <- 1
} else
    {answer <- 0}

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

### CREATE SUMMARY TABLE OF COVID VACC RECORDS
CovSystemSummary <- Vaxeventtot %>%
  filter(vacc_type_target_disease == "COVID-19" &
           vacc_event_created_at>=reporting_start_date) %>% 
  group_by(vacc_status,vacc_type_target_disease,vacc_data_source) %>% 
  summarise(record_count = sum(number_of_vacc_events))

### CREATE SUMMARY TABLE OF FLU VACC RECORDS
FluSystemSummary <- Vaxeventtot %>% 
  filter(vacc_type_target_disease == "Influenza (disorder)" &
           vacc_event_created_at>=reporting_start_date) %>%
  group_by(vacc_status,vacc_type_target_disease,vacc_data_source) %>% 
  summarise(record_count = sum(number_of_vacc_events))

### CREATE SUMMARY TABLE OF SHINGLES VACC RECORDS
HZSystemSummary <- Vaxeventtot %>%
  filter(vacc_type_target_disease == "Herpes zoster (disorder)" &
           vacc_event_created_at>=reporting_start_date) %>% 
  group_by(vacc_status,vacc_type_target_disease,vacc_data_source) %>% 
  summarise(record_count = sum(number_of_vacc_events))

### CREATE SUMMARY TABLE OF PNEUMOCOCCAL VACC RECORDS
PneumSystemSummary <- Vaxeventtot %>%
  filter(vacc_type_target_disease == "Pneumococcal infectious disease (disorder)" &
           vacc_event_created_at>=reporting_start_date) %>%
  group_by(vacc_status,vacc_type_target_disease,vacc_data_source) %>% 
  summarise(record_count = sum(number_of_vacc_events))

### UN-HASH AND RUN THE BELOW 2 LINES IF NOT GOING ON TO RUN SECTION B2
# rm(Vaxeventtot)
# gc()

########################################################################
#############################  SECTION B2 ##############################
#########################  VDL SYSTEM SUMMARY ##########################
########################################################################

### CREATE SUMMARY TABLE OF ALL RECORDS IN VDL BY SYSTEM AND SAVE TO VACCINEDM FOLDER
# Create summary of current week's figures
sheet0 <- Vaxeventtot %>%
  group_by(vacc_status,vacc_type_target_disease,vacc_data_source) %>% 
  summarise(number_of_vacc_events = sum(number_of_vacc_events))

sheet0$summary_date <- Sys.Date() # Adds date the summary is run
sheet0$diff_since_last_run <- as.numeric(0) # Adds blank column for later weekly difference calculation
sheet0 <- sheet0 %>% select(summary_date,everything()) # Reorders columns

vdl_summ <- file.path("VDL_summary.xlsx")

# Read in previous week's figures
sheet1 <- read_excel(vdl_summ, 
                     sheet = "sheet1")

# Prep previous week's table for comparison
sheet1_ed <- sheet1 %>% select(-summary_date,-diff_since_last_run) %>% 
  rename(prev_figs = "number_of_vacc_events")

# Join current and previous week tables, calculate difference of figures & remove previous figures
sheet0 <- sheet0 %>% 
  left_join(sheet1_ed,by = c("vacc_status","vacc_type_target_disease","vacc_data_source")) %>% 
  mutate(diff_since_last_run=(number_of_vacc_events-prev_figs)) %>% 
  select(-prev_figs)

# Read in the other previous 6 previous weeks tables
sheet2 <- read_excel(vdl_summ, 
                     sheet = "sheet2")
sheet3 <- read_excel(vdl_summ, 
                     sheet = "sheet3")
sheet4 <- read_excel(vdl_summ, 
                     sheet = "sheet4")
sheet5 <- read_excel(vdl_summ, 
                     sheet = "sheet5")
sheet6 <- read_excel(vdl_summ, 
                     sheet = "sheet6")
sheet7 <- read_excel(vdl_summ, 
                     sheet = "sheet7")

sheet1$summary_date <- as.Date(sheet1$summary_date,"%Y-%m-%d")
sheet2$summary_date <- as.Date(sheet2$summary_date,"%Y-%m-%d")
sheet3$summary_date <- as.Date(sheet3$summary_date,"%Y-%m-%d")
sheet4$summary_date <- as.Date(sheet4$summary_date,"%Y-%m-%d")
sheet5$summary_date <- as.Date(sheet5$summary_date,"%Y-%m-%d")
sheet6$summary_date <- as.Date(sheet6$summary_date,"%Y-%m-%d")
sheet7$summary_date <- as.Date(sheet7$summary_date,"%Y-%m-%d")

#Collates all week summary tables into a single report for latest rolling 8 weeks
SystemSummaryFull <-
  list("sheet1" = sheet0,
       "sheet2" = sheet1,
       "sheet3" = sheet2,
       "sheet4" = sheet3,
       "sheet5" = sheet4,
       "sheet6" = sheet5,
       "sheet7" = sheet6,
       "sheet8" = sheet7)

options(openxlsx.dateFormat = "yyyy-mm-dd")

# Overwrite summary Excel file with latest 8 weeks
write.xlsx(SystemSummaryFull,
           paste("VDL_summary.xlsx"),
           asTable = TRUE,
           colWidths = "auto")

# save backup of workbook every month
if (lubridate::day(Sys.Date()) < 8) {
  write.xlsx(SystemSummaryFull,
           paste("Archive/VDL_summary_backup.xlsx"),
           asTable = TRUE,
           colWidths = "auto") }

# Save previous run's figures as backup
write.xlsx(sheet1,
           paste("Archive/VDL_summary_sheet1_backup.xlsx"),
           asTable = TRUE,
           colWidths = "auto")

rm(SystemSummaryFull,sheet0,sheet1,sheet1_ed,sheet2,sheet3,
   sheet4,sheet5,sheet6,sheet7,vdl_summ,Vaxeventtot) # Removes summaries from working environment

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
###############################  SECTION D1 ############################
##########################  COVID-19 VACCINATIONS ######################
########################################################################

### CREATE COVID-19 VACCINATIONS DATAFRAME
#############################################################################################

### EXTRACT COMPLETED COVID-19 VACC RECORDS IN vaccination_event_analysis VIEW
CovVaxData <- odbc::dbGetQuery(conn, "select 
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
                          AND vacc_type_target_disease='COVID-19'
                          AND vacc_clinical_trial_flag='0' ")

### CLEAN INVALID CHARACTERS FROM FREE-TEXT FIELDS IN EVENT ANALYSIS DATA
CovVaxData$vacc_location_name <- textclean::replace_non_ascii(CovVaxData$vacc_location_name,replacement = "")
CovVaxData$vacc_performer_name <- textclean::replace_non_ascii(CovVaxData$vacc_performer_name,replacement = "")

### CREATE NEW COLUMN FOR DAYS BETWEEN VACCINATION AND RECORD CREATION
CovVaxData$vacc_record_date <- as.Date(substr(CovVaxData$vacc_record_created_at,1,10))
CovVaxData$days_between_vacc_and_recording <- as.double(difftime(CovVaxData$vacc_record_date,CovVaxData$vacc_occurence_time,units="days"))

### CREATE BLANK PLACEHOLDER COLUMNS
CovVaxData$dups_same_day_flag <- NA
CovVaxData$VMT_only_dups_flag <- NA

### CREATE AN INTERVAL FOR BOOSTER VACCINES PER PATIENT
CovVaxData <- CovVaxData %>%
  arrange(patient_derived_encrypted_upi,vacc_occurence_time) %>% 
  group_by(patient_derived_encrypted_upi) %>%
  mutate(dose_number = row_number()) %>% mutate(doses = n()) %>% ungroup() %>% 
  mutate(prev_vacc_date=lag(vacc_occurence_time)) %>%
  mutate("cov_booster_interval (days)"=as.integer(difftime(vacc_occurence_time,prev_vacc_date,units = "days")))

CovVaxData$"cov_booster_interval (days)" <- ifelse(CovVaxData$dose_number=="1",NA,CovVaxData$"cov_booster_interval (days)")
CovVaxData$"cov_booster_interval (days)" [is.na(CovVaxData$patient_derived_encrypted_upi)] <- NA
CovVaxData$"cov_booster_interval (days)" [CovVaxData$vacc_booster=="FALSE"] <- NA

### CREATE BLANK PLACEHOLDER COLUMNS
CovVaxData$"flu_interval (days)" <- NA
CovVaxData$"hz_interval (days)" <- NA
CovVaxData$"pneum_vacc_interval (weeks)" <- NA
CovVaxData$sort_date <- NA

### CREATE COVID-19 VACC DATAFRAME BY JOINING COVID VACC EVENT RECORDS WITH PATIENT RECORDS
CovVaxData <- CovVaxData %>%
  inner_join(Vaxpatientraw, by=("source_system_patient_id")) %>%
  arrange(desc(vacc_event_created_at)) %>% # sort data by date amended
  select(-patient_derived_encrypted_upi,-vacc_record_date,-dose_number,-doses,-prev_vacc_date) %>% # remove temp items from Interval calculation
  mutate(CHIcheck = phsmethods::chi_check(patient_derived_chi_number)) # create CHI check data item

### CREATE COVID-19 VACCINATIONS DQ QUERIES
#############################################################################################

### CREATE SUMMARY TABLE OF PATIENT CHI STATUS USING PHSMETHODS CHI_CHECK
cov_chi_check <- CovVaxData %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name, vacc_data_source, CHIcheck) %>%
  summarise(record_count = n())

### CREATE SUMMARY TABLE OF VACC PRODUCT TYPE
cov_vacc_prodSumm <- CovVaxData %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name, vacc_data_source, vacc_product_name) %>%
  summarise(record_count = n()) 

### CREATE SUMMARY TABLE OF VACC DOSE NUMBER FROM APRIL 2024
cov_vacc_doseSumm <- CovVaxData %>%
  filter(vacc_event_created_at >= reporting_start_date &
           vacc_occurence_time >= "2024-04-01") %>% 
  group_by(vacc_location_health_board_name, vacc_data_source, vacc_dose_number,
           vacc_booster) %>%
  summarise(record_count = n()) 

### CREATE TABLE AND SUMMARY OF PATIENTS MISSING (NA) CHI NUMBER
cov_chi_na <- CovVaxData %>% filter(grepl("Missing",CHIcheck)) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  select(-CHIcheck) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "01. Missing CHI number")

cov_chi_naSumm <- cov_chi_na %>% group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source) %>% 
  summarise(record_count = n())

### CREATE TABLE OF PATIENTS WITH 2 OR MORE DOSE 1 RECORDS AND SUMMARY
cov_dose1 <- CovVaxData %>% filter(vacc_dose_number == "1" & vacc_booster == "FALSE")

cov_dose1x2IDs <- cov_dose1 %>%
  group_by(patient_derived_upi_number) %>% 
  summarise(count_by_patient_derived_upi_number = n()) %>%
  na.omit(count_by_patient_derived_upi_number) %>%
  filter(count_by_patient_derived_upi_number > 1)

cov_dose1x2 <- cov_dose1 %>%
  filter(patient_derived_upi_number %in% cov_dose1x2IDs$patient_derived_upi_number)

cov_dose1x2sort <- cov_dose1x2 %>% select(vacc_event_created_at,patient_derived_upi_number) %>%
  group_by(patient_derived_upi_number) %>%
  summarise(vacc_event_created_at=max(vacc_event_created_at)) %>%
  rename(latest_date = vacc_event_created_at)

cov_dose1x2 <- cov_dose1x2 %>% left_join(cov_dose1x2sort,by="patient_derived_upi_number") %>%
  arrange(desc(latest_date)) %>% 
  mutate(sort_date = latest_date) %>% 
  mutate(QueryName = "02. COV Two or more dose 1") %>%
  filter(sort_date >= reporting_start_date) %>% 
  select(-latest_date,-CHIcheck)

rm(cov_dose1x2IDs,cov_dose1x2sort)

if (nrow(cov_dose1x2)>0) {
  
    # Create flag for duplicates occuring on same day
    cov_dose1x2_samedateIDs <- cov_dose1x2 %>% 
      group_by(patient_derived_upi_number,vacc_occurence_time) %>% 
      summarise(nn = n()) %>%
      filter(nn > 1)
    
    cov_dose1x2 <- cov_dose1x2 %>% 
      left_join(cov_dose1x2_samedateIDs,join_by(patient_derived_upi_number,vacc_occurence_time))
    
    cov_dose1x2$dups_same_day_flag <- 0
    cov_dose1x2$dups_same_day_flag [cov_dose1x2$nn>0] <- 1
    ###
    
    # Create flag for VMT-only dose 1 dups
    cov_dose1x2_VMT <- cov_dose1x2 %>% filter(vacc_data_source == "TURAS")
    
    cov_dose1x2_VMT_ids <- cov_dose1x2_VMT %>% 
      group_by(patient_derived_upi_number) %>% 
      summarise(nn2=n()) %>% ungroup() %>% 
      filter(nn2>1)
    
    cov_dose1x2 <- cov_dose1x2 %>% 
      left_join(cov_dose1x2_VMT_ids,by="patient_derived_upi_number")
    
    cov_dose1x2$VMT_only_dups_flag <- 0
    cov_dose1x2$VMT_only_dups_flag [cov_dose1x2$nn2>0] <- 1
    ###

cov_dose1x2 <- cov_dose1x2 %>% select(-nn,-nn2)

rm(cov_dose1x2_VMT,cov_dose1x2_VMT_ids,cov_dose1x2_samedateIDs)
}

cov_dose1x2Summ <- cov_dose1x2 %>% group_by(vacc_location_health_board_name) %>% 
  summarise(record_count = n())

### CREATE TABLE OF PATIENTS WITH 2 OR MORE DOSE 2 RECORDS AND SUMMARY
cov_dose2 <- CovVaxData %>% filter(vacc_dose_number == "2" & vacc_booster == "FALSE")

cov_dose2x2IDs <- cov_dose2 %>%
  group_by(patient_derived_upi_number) %>% 
  summarise(count_by_patient_derived_upi_number = n()) %>%
  na.omit(count_by_patient_derived_upi_number) %>%
  filter(count_by_patient_derived_upi_number > 1)

cov_dose2x2 <- cov_dose2 %>%
  filter(patient_derived_upi_number %in% cov_dose2x2IDs$patient_derived_upi_number)

cov_dose2x2sort <- cov_dose2x2 %>% select(vacc_event_created_at,patient_derived_upi_number) %>%
  group_by(patient_derived_upi_number) %>%
  summarise(vacc_event_created_at=max(vacc_event_created_at)) %>%
  rename(latest_date = vacc_event_created_at)

cov_dose2x2 <- cov_dose2x2 %>% left_join(cov_dose2x2sort,by="patient_derived_upi_number") %>%
  arrange(desc(latest_date)) %>% 
  mutate(sort_date = latest_date)%>%
  mutate(QueryName = "03. COV Two or more dose 2") %>%
  filter(sort_date >= reporting_start_date) %>% 
  select(-latest_date,-CHIcheck)

rm(cov_dose2x2IDs,cov_dose2x2sort)

if (nrow(cov_dose2x2)>0) {
  
    # Create flag for duplicates occuring on same day
    cov_dose2x2_samedateIDs <- cov_dose2x2 %>% 
      group_by(patient_derived_upi_number,vacc_occurence_time) %>% 
      summarise(nn = n()) %>%
      filter(nn > 1)
    
    cov_dose2x2 <- cov_dose2x2 %>% 
      left_join(cov_dose2x2_samedateIDs,join_by(patient_derived_upi_number,vacc_occurence_time))
    
    cov_dose2x2$dups_same_day_flag <- 0
    cov_dose2x2$dups_same_day_flag [cov_dose2x2$nn>0] <- 1
    ###
    
    # Create flag for VMT-only duplicates
    cov_dose2x2_VMT <- cov_dose2x2 %>% filter(vacc_data_source == "TURAS")
    
    cov_dose2x2_VMT_ids <- cov_dose2x2_VMT %>% 
      group_by(patient_derived_upi_number) %>% 
      summarise(nn2=n()) %>% ungroup() %>% 
      filter(nn2>1)
    
    cov_dose2x2 <- cov_dose2x2 %>% 
      left_join(cov_dose2x2_VMT_ids,by="patient_derived_upi_number")
    
    cov_dose2x2$VMT_only_dups_flag <- 0
    cov_dose2x2$VMT_only_dups_flag [cov_dose2x2$nn2>0] <- 1
    ###

cov_dose2x2 <- cov_dose2x2 %>% select(-nn,-nn2)

rm(cov_dose2x2_samedateIDs,cov_dose2x2_VMT,cov_dose2x2_VMT_ids)
}

cov_dose2x2Summ <- cov_dose2x2 %>% group_by(vacc_location_health_board_name) %>% 
  summarise(record_count = n())

### CREATE TABLE OF PATIENTS WITH 2 OR MORE DOSE 3 RECORDS AND SUMMARY
cov_dose3 <- CovVaxData %>% filter(vacc_dose_number == "3" & vacc_booster == "FALSE")

cov_dose3x2IDs <- cov_dose3 %>%
  group_by(patient_derived_upi_number) %>% 
  summarise(count_by_patient_derived_upi_number = n()) %>%
  na.omit(count_by_patient_derived_upi_number) %>%
  filter(count_by_patient_derived_upi_number > 1)

cov_dose3x2 <- cov_dose3 %>%
  filter(patient_derived_upi_number %in% cov_dose3x2IDs$patient_derived_upi_number)

cov_dose3x2sort <- cov_dose3x2 %>% select(vacc_event_created_at,patient_derived_upi_number) %>%
  group_by(patient_derived_upi_number) %>%
  summarise(vacc_event_created_at=max(vacc_event_created_at)) %>%
  rename(latest_date = vacc_event_created_at)

cov_dose3x2 <- cov_dose3x2 %>% left_join(cov_dose3x2sort,by="patient_derived_upi_number") %>%
  arrange(desc(latest_date)) %>% 
  mutate(sort_date = latest_date) %>%
  mutate(QueryName = "04. COV Two or more dose 3") %>% 
  filter(sort_date >= reporting_start_date) %>% 
  select(-latest_date,-CHIcheck)

rm(cov_dose3x2IDs,cov_dose3x2sort)

if (nrow(cov_dose3x2)>0) {
  
    # Create flag for duplicates occuring on same day
    cov_dose3x2_samedateIDs <- cov_dose3x2 %>% 
      group_by(patient_derived_upi_number,vacc_occurence_time) %>% 
      summarise(nn = n()) %>%
      filter(nn > 1)
    
    cov_dose3x2 <- cov_dose3x2 %>% 
      left_join(cov_dose3x2_samedateIDs,join_by(patient_derived_upi_number,vacc_occurence_time))
    
    cov_dose3x2$dups_same_day_flag <- 0
    cov_dose3x2$dups_same_day_flag [cov_dose3x2$nn>0] <- 1
    ###
    
    # Create flag for VMT-only duplicates
    cov_dose3x2_VMT <- cov_dose3x2 %>% filter(vacc_data_source == "TURAS")
    
    cov_dose3x2_VMT_ids <- cov_dose3x2_VMT %>% 
      group_by(patient_derived_upi_number) %>% 
      summarise(nn2=n()) %>% ungroup() %>% 
      filter(nn2>1)
    
    cov_dose3x2 <- cov_dose3x2 %>% 
      left_join(cov_dose3x2_VMT_ids,by="patient_derived_upi_number")
    
    cov_dose3x2$VMT_only_dups_flag <- 0
    cov_dose3x2$VMT_only_dups_flag [cov_dose3x2$nn2>0] <- 1
    ###

cov_dose3x2 <- cov_dose3x2 %>% select(-nn,-nn2)

rm(cov_dose3x2_samedateIDs,cov_dose3x2_VMT,cov_dose3x2_VMT_ids)
}

cov_dose3x2Summ <- cov_dose3x2 %>% group_by(vacc_location_health_board_name) %>% 
  summarise(record_count = n())

### CREATE TABLE OF PATIENTS WITH 2 OR MORE DOSE 4 RECORDS AND SUMMARY
cov_dose4 <- CovVaxData %>% filter(vacc_dose_number == "4" & vacc_booster == "FALSE")

cov_dose4x2IDs <- cov_dose4 %>%
  group_by(patient_derived_upi_number) %>% 
  summarise(count_by_patient_derived_upi_number = n()) %>%
  na.omit(count_by_patient_derived_upi_number) %>%
  filter(count_by_patient_derived_upi_number > 1)

cov_dose4x2 <- cov_dose4 %>%
  filter(patient_derived_upi_number %in% cov_dose4x2IDs$patient_derived_upi_number)

cov_dose4x2sort <- cov_dose4x2 %>% select(vacc_event_created_at,patient_derived_upi_number) %>%
  group_by(patient_derived_upi_number) %>%
  summarise(vacc_event_created_at=max(vacc_event_created_at)) %>%
  rename(latest_date = vacc_event_created_at)

cov_dose4x2 <- cov_dose4x2 %>% left_join(cov_dose4x2sort,by="patient_derived_upi_number") %>%
  arrange(desc(latest_date)) %>% 
  mutate(sort_date = latest_date) %>%
  mutate(QueryName = "05. COV Two or more dose 4") %>% 
  filter(sort_date >= reporting_start_date) %>% 
  select(-latest_date,-CHIcheck)

rm(cov_dose4x2IDs,cov_dose4x2sort)

if (nrow(cov_dose4x2)>0) {
  
    # Create flag for duplicates occuring on same day
    cov_dose4x2_samedateIDs <- cov_dose4x2 %>% 
      group_by(patient_derived_upi_number,vacc_occurence_time) %>% 
      summarise(nn = n()) %>%
      filter(nn > 1)
    
    cov_dose4x2 <- cov_dose4x2 %>% 
      left_join(cov_dose4x2_samedateIDs,join_by(patient_derived_upi_number,vacc_occurence_time))
    
    cov_dose4x2$dups_same_day_flag <- 0
    cov_dose4x2$dups_same_day_flag [cov_dose4x2$nn>0] <- 1
    ###
    
    # Create flag for VMT-only duplicates
    cov_dose4x2_VMT <- cov_dose4x2 %>% filter(vacc_data_source == "TURAS")
    
    cov_dose4x2_VMT_ids <- cov_dose4x2_VMT %>% 
      group_by(patient_derived_upi_number) %>% 
      summarise(nn2=n()) %>% ungroup() %>% 
      filter(nn2>1)
    
    cov_dose4x2 <- cov_dose4x2 %>% 
      left_join(cov_dose4x2_VMT_ids,by="patient_derived_upi_number")

    cov_dose4x2$VMT_only_dups_flag <- 0
    cov_dose4x2$VMT_only_dups_flag [cov_dose4x2$nn2>0] <- 1
    ###

cov_dose4x2 <- cov_dose4x2 %>% select(-nn,-nn2)

rm(cov_dose4x2_samedateIDs,cov_dose4x2_VMT,cov_dose4x2_VMT_ids)
}

cov_dose4x2Summ <- cov_dose4x2 %>% group_by(vacc_location_health_board_name) %>% 
  summarise(record_count = n())

### CREATE TABLE OF BOOSTER RECORDS GIVEN AT AN INTERVAL OF <12 WEEKS (84 DAYS)
cov_booster <- CovVaxData %>% filter(vacc_booster=="TRUE")

cov_booster_intervalDQ <- cov_booster %>% 
  filter(`cov_booster_interval (days)` < 84) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "06. COV Booster interval < 12 weeks") %>% 
  select(-CHIcheck)

cov_booster_intervalDQSumm <- cov_booster_intervalDQ %>% 
  group_by(vacc_location_health_board_name,vacc_location_name,`cov_booster_interval (days)`) %>% 
  summarise(record_count = n())

### CREATE SUMMARY TABLE OF RECORDS OF VACCINATIONS GIVEN TO AGE <12
cov_age12Summ <- CovVaxData %>% filter(age_at_vacc <12) %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_product_name) %>% 
  summarise(record_count=n())

### CREATE TABLE & SUMMARY OF PATIENTS AGED UNDER 12 GIVEN FULL DOSE
cov_under12fulldose <- CovVaxData %>%
  filter((age_at_vacc<12 &
           vacc_product_name != "Covid-19 mRNA Vaccine Comirnaty 10mcg (Paed)" &
           vacc_product_name != "Covid-19 mRNA Vaccine Comirnaty 3mcg (Infant)" &
            !grepl("Children",vacc_product_name) )) %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "07. COV Under 12 adult dose") %>% 
  select(-CHIcheck)

cov_under12fulldoseSumm <- cov_under12fulldose %>% 
  group_by(vacc_location_health_board_name,vacc_location_name,vacc_product_name,age_at_vacc) %>%
  summarise(record_count=n())

### CREATE TABLE & SUMMARY OF PATIENTS AGED OVER 11, OR UNDER 5 AFTER 01.04.2023, GIVEN PAEDIATRIC DOSE
# this excludes those aged 12 who received a dose 1 paediatric dose at age 11
cov_over12_under5_childdose <- CovVaxData %>%
  filter((age_at_vacc>12 | (age_at_vacc<5 & vacc_occurence_time >= "2023-04-01")) &
           (vacc_product_name == "Covid-19 mRNA Vaccine Comirnaty 10mcg (Paed)" |
           vacc_product_name == "Comirnaty XBB.1.5 Children 5-11 years COVID-19 mRNA Vaccine 10micrograms/0.3ml dose"))

cov_age11childdose <- CovVaxData %>% filter(age_at_vacc == "11" &
                                              (vacc_product_name == "Covid-19 mRNA Vaccine Comirnaty 10mcg (Paed)" |
                                                 vacc_product_name == "Comirnaty XBB.1.5 Children 5-11 years COVID-19 mRNA Vaccine 10micrograms/0.3ml dose"))
cov_age12childdose <- CovVaxData %>% filter(age_at_vacc == "12" &
                                              (vacc_product_name == "Covid-19 mRNA Vaccine Comirnaty 10mcg (Paed)" |
                                                 vacc_product_name == "Comirnaty XBB.1.5 Children 5-11 years COVID-19 mRNA Vaccine 10micrograms/0.3ml dose"))

cov_over11_under5_childdose <- anti_join(cov_age12childdose,cov_age11childdose,by="patient_derived_upi_number") %>%
  rbind(cov_over12_under5_childdose) %>%
  arrange(desc(vacc_event_created_at)) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "08. COV Over 11 or under 5 paediatric dose") %>% 
  select(-CHIcheck)

cov_over11_under5_childdoseSumm <- cov_over11_under5_childdose %>% 
  group_by(vacc_location_health_board_name,vacc_location_name,vacc_product_name,age_at_vacc) %>%
  summarise(record_count=n())

rm(cov_over12_under5_childdose,cov_age11childdose,cov_age12childdose)

### CREATE TABLE & SUMMARY OF PATIENTS AGED OVER 4 GIVEN INFANT DOSE
# this excludes those aged 5 who received a dose 1 infant dose at age 4
cov_over5_infant_dose <- CovVaxData %>% 
  filter(age_at_vacc>5 & (CovVaxData$vacc_product_name == "Covid-19 mRNA Vaccine Comirnaty 3mcg (Infant)" |
                          CovVaxData$vacc_product_name == "Comirnaty XBB.1.5 Children 6 months - 4 years COVID-19 mRNA Vaccine 3micrograms/0.2ml dose"))

cov_age4_infantdose <- CovVaxData %>% 
  filter(age_at_vacc==4 & (CovVaxData$vacc_product_name == "Covid-19 mRNA Vaccine Comirnaty 3mcg (Infant)" |
                            CovVaxData$vacc_product_name == "Comirnaty XBB.1.5 Children 6 months - 4 years COVID-19 mRNA Vaccine 3micrograms/0.2ml dose"))
  
cov_age5_infantdose <- CovVaxData %>% 
  filter(age_at_vacc==5 & (CovVaxData$vacc_product_name == "Covid-19 mRNA Vaccine Comirnaty 3mcg (Infant)" |
                             CovVaxData$vacc_product_name == "Comirnaty XBB.1.5 Children 6 months - 4 years COVID-19 mRNA Vaccine 3micrograms/0.2ml dose"))

cov_over4_infant_dose <- anti_join(cov_age5_infantdose,cov_age4_infantdose,by="patient_derived_upi_number") %>% 
  rbind(cov_over5_infant_dose) %>% 
  arrange(desc(vacc_event_created_at)) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "09. COV Over 4 infant dose") %>% 
  select(-CHIcheck)

cov_over4_infant_doseSumm <- cov_over4_infant_dose %>% 
  group_by(vacc_location_health_board_name,vacc_location_name,vacc_product_name,age_at_vacc) %>%
  summarise(record_count=n())

rm(cov_age4_infantdose,cov_age5_infantdose,cov_over5_infant_dose)

### CREATE TABLE OF PATIENTS WITH A DOSE 2 RECORD BUT NO DOSE 1
cov_dose2nodose1 <- anti_join(cov_dose2, cov_dose1, by="patient_derived_upi_number") %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(QueryName = "10. COV Dose 2, no dose 1") %>% 
  select(-CHIcheck)

cov_dose2nodose1Summ <- cov_dose2nodose1 %>% group_by(vacc_location_health_board_name) %>% 
  summarise(record_count=n())

### CREATE TABLE & SUMMARY OF PATIENTS RECEIVING DOSE 1 AND BOOSTER BUT NO DOSE 2
cov_dose1andboosterIDs <- inner_join(cov_dose1,cov_booster,by="patient_derived_upi_number") %>%
  anti_join(cov_dose2,by="patient_derived_upi_number")

cov_dose1andbooster <- rbind(cov_dose1,cov_booster) %>%
  filter(patient_derived_upi_number %in% cov_dose1andboosterIDs$patient_derived_upi_number) %>%
  arrange(desc(vacc_event_created_at))

cov_dose1andboostersort <- cov_dose1andbooster %>% select(vacc_event_created_at,patient_derived_upi_number) %>%
  group_by(patient_derived_upi_number) %>%
  slice_max(vacc_event_created_at, with_ties = FALSE) %>%
  rename(latest_date = vacc_event_created_at)

cov_dose1andbooster <- cov_dose1andbooster %>% left_join(cov_dose1andboostersort,by="patient_derived_upi_number") %>%
  arrange(desc(latest_date)) %>% 
  mutate(sort_date = latest_date) %>% 
  mutate(QueryName = "11. COV Dose 1 and booster, no dose 2") %>% 
  filter(sort_date >= reporting_start_date) %>% 
  select(-latest_date,-CHIcheck)

rm(cov_dose1andboostersort,cov_dose1andboosterIDs)

cov_dose1andboosterSumm <- cov_dose1andbooster %>% group_by(vacc_location_health_board_name) %>% 
  summarise(record_count=n())

### CREATE & SAVE OUT COVID-19 VACC SUMMARY REPORT
#############################################################################################

#Collates pivot tables into a single report
cov_SummaryReport <-
  list("System Summary" = CovSystemSummary,
       "CHI Check" = cov_chi_check,
       "Vacc Product Type" = cov_vacc_prodSumm,
       "Dose Given After 01.04.2024" = cov_vacc_doseSumm,
       "Missing CHIs" = cov_chi_naSumm,
       "2 or More First Doses" = cov_dose1x2Summ,
       "2 or More Second Doses" = cov_dose2x2Summ,
       "2 or More Third doses" = cov_dose3x2Summ,
       "2 or More Fourth doses" = cov_dose4x2Summ,
       "Booster Interval < 12 weeks" = cov_booster_intervalDQSumm,
       "Age <12" = cov_age12Summ,
       "Under 12 full dose" = cov_under12fulldoseSumm,
       ">11 or <5 paediatric dose" = cov_over11_under5_childdoseSumm,
       "Over 4 infant dose" = cov_over4_infant_doseSumm,
       "Dose 2, no dose 1" = cov_dose2nodose1Summ,
       "Dose 1 and booster, no dose 2" = cov_dose1andboosterSumm)

rm(CovVaxData,CovSystemSummary,cov_chi_check,cov_vacc_prodSumm,
   cov_vacc_doseSumm,cov_chi_naSumm,
   cov_dose1x2Summ,cov_dose2x2Summ,cov_dose3x2Summ,cov_dose4x2Summ,cov_booster_intervalDQSumm,
   cov_age12Summ,cov_under12fulldoseSumm,cov_over11_under5_childdoseSumm,
   cov_over4_infant_doseSumm,cov_dose2nodose1Summ,cov_dose1andboosterSumm,
   cov_dose1,cov_dose2,cov_dose3,cov_dose4,cov_booster)

#Saves out collated tables into an excel file
if (answer==1) {
  write.xlsx(cov_SummaryReport,
           paste("DQ Summary Reports/Covid-19_Vacc_DQ_4wk_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
           asTable = TRUE,
           colWidths = "auto")
  } else {
  write.xlsx(cov_SummaryReport,
             paste("DQ Summary Reports/Covid-19_Vacc_DQ_Full_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
             asTable = TRUE,
             colWidths = "auto")
  }

rm(cov_SummaryReport)

### COLLATE COVID QUERIES FOR VACCINATIONS DQ REPORT
#############################################################################################

multi_vacc <- rbind(cov_chi_na,cov_dose1x2,cov_dose2x2,cov_dose3x2,cov_dose4x2)
covid_vacc <- rbind(cov_booster_intervalDQ,cov_under12fulldose,cov_over11_under5_childdose,
                    cov_over4_infant_dose,cov_dose2nodose1,cov_dose1andbooster)

rm(cov_chi_na,cov_dose1x2,cov_dose2x2,cov_dose3x2,cov_dose4x2,
   cov_booster_intervalDQ,cov_under12fulldose,cov_over11_under5_childdose,
   cov_over4_infant_dose,cov_dose2nodose1,cov_dose1andbooster)

gc()

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
                          AND vacc_clinical_trial_flag='0'
                          AND vacc_occurence_time > ?",
                         params = reporting_start_date-28)

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

### CREATE FLU VACCINATIONS DQ QUERIES
#############################################################################################

### CREATE SUMMARY TABLE OF PATIENT CHI STATUS USING PHSMETHODS CHI_CHECK
flu_chi_check <- FluVaxData %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name, vacc_data_source, CHIcheck) %>%
  summarise(record_count = n())

### CREATE SUMMARY TABLE OF VACC PRODUCT TYPE
flu_vacc_prodSumm <- FluVaxData %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name, vacc_data_source, vacc_product_name) %>%
  summarise(record_count = n()) 

### CREATE TABLE AND SUMMARY OF PATIENTS MISSING (NA) CHI NUMBER
flu_chi_na <- FluVaxData %>% filter(grepl("Missing",CHIcheck)) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "01. Missing CHI number")

flu_chi_naSumm <- flu_chi_na %>% group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source, Date_Administered) %>% 
  summarise(record_count = n())

flu_chi_na <- flu_chi_na %>% select(-Date_Administered,-CHIcheck)

### CREATE TABLE OF RECORDS & SUMMARY OF VACCINATIONS GIVEN AT INTERVAL OF < 4 WEEKS
flu_interval <- FluVaxData %>% filter(`flu_interval (days)`<28) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "12. FLU interval < 4 weeks")

flu_intervalSumm <- flu_interval %>% group_by(vacc_location_health_board_name,vacc_location_name,`flu_interval (days)`,Date_Administered) %>% 
  summarise(record_count = n())

flu_interval <- flu_interval %>% select(-Date_Administered,-CHIcheck)
  
### CREATE TABLE OF RECORDS & SUMMARY OF VACCINATIONS THAT OCCUR BEFORE 06/09/2021
flu_early <- FluVaxData %>% filter(vacc_occurence_time < "2021-09-06") %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "13. FLU vacc before 06.09.2021")

flu_earlySumm <- flu_early %>% group_by(vacc_location_health_board_name,vacc_location_name,Date_Administered) %>%
  summarise(record_count = n())

flu_early <- flu_early %>% select(-Date_Administered,-CHIcheck)

### CREATE TABLE OF RECORDS & SUMMARY OF VACCINATION GIVEN AT AGE <2 OR >17
flu_fluenz_ageDQ <- FluVaxData %>%
  filter(vacc_product_name == "Fluenz Tetra Vaccine AstraZeneca") %>%
  filter(age_at_vacc < 2 | age_at_vacc > 17) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "14. FLU Fluenz given age <2 or >17")

flu_fluenz_ageDQSumm <- flu_fluenz_ageDQ %>%
  group_by(vacc_location_health_board_name, vacc_location_name, Date_Administered, age_at_vacc) %>%
  summarise(record_count = n())

flu_fluenz_ageDQ <- flu_fluenz_ageDQ %>% select(-Date_Administered,-CHIcheck)

### CREATE & SAVE OUT FLU VACC SUMMARY REPORT
#############################################################################################

#Collates pivot tables into a single report
flu_SummaryReport <-
  list("System Summary" = FluSystemSummary,
       "CHI Check" = flu_chi_check,
       "Vacc Product Type" = flu_vacc_prodSumm,
       "Missing CHIs" = flu_chi_naSumm,
       "Interval < 4 weeks" =  flu_intervalSumm,
       "Vaccine before 06.09.2021" = flu_earlySumm,
       "Fluenz Given to Aged <2 or >17" = flu_fluenz_ageDQSumm)

rm(FluVaxData,FluSystemSummary,flu_chi_check,flu_vacc_prodSumm,flu_chi_naSumm,
   flu_intervalSumm,flu_earlySumm,flu_fluenz_ageDQSumm)

#Saves out collated tables into an excel file
if (answer==1) {
  write.xlsx(flu_SummaryReport,
           paste("DQ Summary Reports/Flu_Vacc_DQ_4wk_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
           asTable = TRUE,
           colWidths = "auto")
  } else {
  write.xlsx(flu_SummaryReport,
             paste("DQ Summary Reports/Flu_Vacc_DQ_Full_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
             asTable = TRUE,
             colWidths = "auto")
  }

rm(flu_SummaryReport)

### COLLATE FLU QUERIES FOR VACCINATIONS DQ REPORT
#############################################################################################

multi_vacc <- multi_vacc %>% rbind(flu_chi_na)
flu_vacc <- rbind(flu_interval,flu_early,flu_fluenz_ageDQ)

rm(flu_chi_na,flu_interval,flu_early,flu_fluenz_ageDQ)

gc()

########################################################################
###############################  SECTION D3 ############################
#################  HERPES ZOSTER (SHINGLES) VACCINATIONS ###############
########################################################################

### CREATE SHINGLES VACCINATIONS DATAFRAME
#############################################################################################

### EXTRACT VARIABLES FROM COMPLETED SHINGLES VACC RECORDS IN vaccination_event_analysis VIEW
HZVaxData <- odbc::dbGetQuery(conn, "select 
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
                          AND vacc_type_target_disease='Herpes zoster (disorder)' ")

### CLEAN INVALID CHARACTERS FROM FREE-TEXT FIELDS IN EVENT ANALYSIS DATA
HZVaxData$vacc_location_name <- textclean::replace_non_ascii(HZVaxData$vacc_location_name,replacement = "")
HZVaxData$vacc_performer_name <- textclean::replace_non_ascii(HZVaxData$vacc_performer_name,replacement = "")

### CREATE NEW COLUMN FOR DAYS BETWEEN VACCINATION AND RECORD CREATION
HZVaxData$vacc_record_date <- as.Date(substr(HZVaxData$vacc_record_created_at,1,10))
HZVaxData$days_between_vacc_and_recording <- as.double(difftime(HZVaxData$vacc_record_date,HZVaxData$vacc_occurence_time,units="days"))

### CREATE BLANK PLACEHOLDER COLUMNS
HZVaxData$dups_same_day_flag <- NA
HZVaxData$VMT_only_dups_flag <- NA
HZVaxData$"cov_booster_interval (days)" <- NA
HZVaxData$"flu_interval (days)" <- NA

### CALCULATE INTERVAL BETWEEN VACCINES PER PATIENT
HZVaxData <- HZVaxData %>%
  arrange(patient_derived_encrypted_upi,vacc_occurence_time) %>% 
  group_by(patient_derived_encrypted_upi) %>%
  mutate(dose_number = row_number()) %>% mutate(doses = n()) %>% ungroup() %>% 
  mutate(prev_vacc_date=lag(vacc_occurence_time)) %>%
  mutate("hz_interval (days)"=as.integer(difftime(vacc_occurence_time,prev_vacc_date,units = "days")))

HZVaxData$"hz_interval (days)" <- ifelse(HZVaxData$dose_number=="1",NA,HZVaxData$"hz_interval (days)")
HZVaxData$"hz_interval (days)" [is.na(HZVaxData$patient_derived_encrypted_upi)] <- NA

HZVaxData$"pneum_vacc_interval (weeks)" <- NA
HZVaxData$sort_date <- NA

### CREATE SHINGLES VACC DATAFRAME BY JOINING SHINGLES VACC EVENT RECORDS WITH PATIENT RECORDS
HZVaxData <- HZVaxData %>%
  inner_join(Vaxpatientraw, by=("source_system_patient_id")) %>%
  arrange(desc(vacc_event_created_at)) %>% # sort data by date amended
  select(-patient_derived_encrypted_upi,-vacc_record_date) %>%
  mutate(Date_Administered = substr(vacc_occurence_time, 1, 10)) %>% # create new vacc date data item for date only (no time)
  mutate(CHIcheck = phsmethods::chi_check(patient_derived_chi_number)) # create CHI check data item

### CREATE SHINGLES VACCINATIONS DQ QUERIES
#############################################################################################

### CREATE SUMMARY TABLE OF PATIENT CHI STATUS USING PHSMETHODS CHI_CHECK
hz_chi_check <- HZVaxData %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name, vacc_data_source, CHIcheck) %>%
  summarise(record_count = n())

### CREATE SUMMARY TABLE OF VACC PRODUCT TYPE
hz_vacc_prodSumm <- HZVaxData %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name, vacc_data_source, vacc_product_name) %>%
  summarise(record_count = n()) 

### CREATE TABLE AND SUMMARY OF PATIENTS MISSING (NA) CHI NUMBER
hz_chi_na  <- HZVaxData %>% filter(grepl("Missing",CHIcheck)) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "01. Missing CHI number")

hz_chi_naSumm <- hz_chi_na %>%
  group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source, Date_Administered) %>% 
  summarise(record_count = n())

hz_chi_na <- hz_chi_na %>% select(-CHIcheck,-Date_Administered)

### CREATE TABLE OF RECORDS & SUMMARY OF 2 OR MORE DOSE 1 VACCINATIONS
hz_dose1 <- HZVaxData %>% filter(vacc_dose_number == "1")

hz_dose1x2IDs <- hz_dose1 %>% group_by(patient_derived_upi_number) %>% 
  summarise(count_by_patient_derived_upi_number = n()) %>%
  na.omit(count_by_patient_derived_upi_number) %>%
  filter(count_by_patient_derived_upi_number > 1)

hz_dose1x2 <- hz_dose1 %>%
  filter(patient_derived_upi_number %in% hz_dose1x2IDs$patient_derived_upi_number)

hz_dose1x2sort <- hz_dose1x2 %>% select(vacc_event_created_at,patient_derived_upi_number) %>%
  group_by(patient_derived_upi_number) %>%
  summarise(vacc_event_created_at=max(vacc_event_created_at)) %>%
  rename(latest_date = vacc_event_created_at)

hz_dose1x2 <- hz_dose1x2 %>% left_join(hz_dose1x2sort,by="patient_derived_upi_number") %>%
  arrange(desc(latest_date)) %>%
  mutate(sort_date = latest_date) %>%
  mutate(QueryName = "02. HZ Two or more dose 1") %>% 
  filter(sort_date >= reporting_start_date) %>% 
  select(-latest_date)

rm(hz_dose1x2IDs,hz_dose1x2sort)

if (nrow(hz_dose1x2)>0) {
    # Create flag for duplicates occuring on same day
    hz_dose1x2_samedateIDs <- hz_dose1x2 %>% 
      group_by(patient_derived_upi_number,vacc_occurence_time) %>% 
      summarise(nn = n()) %>%
      filter(nn > 1)
    
    hz_dose1x2 <- hz_dose1x2 %>% 
      left_join(hz_dose1x2_samedateIDs,join_by(patient_derived_upi_number,vacc_occurence_time))
    
    hz_dose1x2$dups_same_day_flag <- 0
    hz_dose1x2$dups_same_day_flag [hz_dose1x2$nn>0] <- 1
    ###
    
    # Create flag for VMT-only dose 1 dups
    hz_dose1x2_VMT <- hz_dose1x2 %>% filter(vacc_data_source == "TURAS") %>% 
      arrange(desc(vacc_occurence_time))
    
    hz_dose1x2_VMT_ids <- hz_dose1x2_VMT %>% 
      group_by(patient_derived_upi_number) %>% 
      summarise(nn2=n()) %>% ungroup() %>% 
      filter(nn2>1)
    
    hz_dose1x2 <- hz_dose1x2 %>% 
      left_join(hz_dose1x2_VMT_ids,by="patient_derived_upi_number")
    
    hz_dose1x2$VMT_only_dups_flag <- 0
    hz_dose1x2$VMT_only_dups_flag [hz_dose1x2$nn2>0] <- 1
    ###

hz_dose1x2 <- hz_dose1x2 %>% select(-nn,-nn2)
    
rm(hz_dose1x2_samedateIDs,hz_dose1x2_VMT,hz_dose1x2_VMT_ids)
}
    
hz_dose1x2Summ <- hz_dose1x2 %>%
  group_by(vacc_location_health_board_name, vacc_data_source, Date_Administered) %>%
  summarise(record_count = n())

hz_dose1x2 <- hz_dose1x2 %>% select(-Date_Administered,-CHIcheck)

### CREATE TABLE OF RECORDS & SUMMARY OF 2 OR MORE DOSE 2 VACCINATIONS
hz_dose2 <- HZVaxData %>% filter(vacc_dose_number == "2")

hz_dose2x2IDs <- hz_dose2 %>% group_by(patient_derived_upi_number) %>% 
  summarise(count_by_patient_derived_upi_number = n()) %>%
  na.omit(count_by_patient_derived_upi_number) %>%
  filter(count_by_patient_derived_upi_number > 1)

hz_dose2x2 <- hz_dose2 %>%
  filter(patient_derived_upi_number %in% hz_dose2x2IDs$patient_derived_upi_number)

hz_dose2x2sort <- hz_dose2x2 %>% select(vacc_event_created_at,patient_derived_upi_number) %>%
  group_by(patient_derived_upi_number) %>%
  summarise(vacc_event_created_at=max(vacc_event_created_at)) %>%
  rename(latest_date = vacc_event_created_at)

hz_dose2x2 <- hz_dose2x2 %>% left_join(hz_dose2x2sort,by="patient_derived_upi_number") %>%
  arrange(desc(latest_date)) %>%
  mutate(sort_date = latest_date) %>%
  mutate(QueryName = "03. HZ Two or more dose 2") %>% 
  filter(sort_date >= reporting_start_date) %>% 
  select(-latest_date)

rm(hz_dose2x2IDs,hz_dose2x2sort)

if (nrow(hz_dose2x2)>0) {

    # Create flag for duplicates occuring on same day
    hz_dose2x2_samedateIDs <- hz_dose2x2 %>% 
      group_by(patient_derived_upi_number,vacc_occurence_time) %>% 
      summarise(nn = n()) %>%
      filter(nn > 1)
    
    hz_dose2x2 <- hz_dose2x2 %>% 
      left_join(hz_dose2x2_samedateIDs,join_by(patient_derived_upi_number,vacc_occurence_time))
    
    hz_dose2x2$dups_same_day_flag <- 0
    hz_dose2x2$dups_same_day_flag [hz_dose2x2$nn>0] <- 1
    ###
    
    # Create flag for VMT-only dose 2 dups
    hz_dose2x2IDsVMT <- hz_dose2x2 %>% filter(vacc_data_source=="TURAS")
    
    hz_dose2x2IDsVMT_ids <- hz_dose2x2IDsVMT %>% 
      group_by(patient_derived_upi_number) %>% 
      summarise(nn2=n()) %>% ungroup() %>% 
      filter(nn2>1)
    
    hz_dose2x2 <- hz_dose2x2 %>% 
      left_join(hz_dose2x2IDsVMT_ids,by="patient_derived_upi_number")
    
    hz_dose2x2$VMT_only_dups_flag <- 0
    hz_dose2x2$VMT_only_dups_flag [hz_dose2x2$nn2>0] <- 1
    ###
    
hz_dose2x2 <- hz_dose2x2 %>% select(-nn,-nn2)
    
rm(hz_dose2x2_samedateIDs,hz_dose2x2IDsVMT,hz_dose2x2IDsVMT_ids)
}

hz_dose2x2Summ <- hz_dose2x2 %>%
  group_by(vacc_location_health_board_name, vacc_data_source, Date_Administered) %>%
  summarise(record_count = n())

hz_dose2x2 <- hz_dose2x2 %>% select(-Date_Administered,-CHIcheck)

### CREATE TABLE OF RECORDS & SUMMARY OF ZOSTAVAX GIVEN FOR DOSE 2
hz_dose2_Zost <- hz_dose2 %>% 
  filter(vacc_product_name == "Shingles Vaccine Zostavax Merck") %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "15. HZ Dose 2 Zostavax")

hz_dose2_ZostSumm <- hz_dose2_Zost %>%
  group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source, Date_Administered) %>%
  summarise(record_count = n())

hz_dose2_Zost <- hz_dose2_Zost %>% select(-Date_Administered,-CHIcheck)

### CREATE TABLE OF RECORDS & SUMMARY OF SHINGRIX DOSE 2 GIVEN AT AN INTERVAL OF <56 DAYS
hz_dose2_early <- hz_dose2 %>% 
  filter(vacc_product_name == "Shingles Vaccine Shingrix GlaxoSmithKline") %>%
  filter(`hz_interval (days)` < "56") %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "16. HZ Dose 2 interval < 56 days")

    # remove patients with no dose 1
    hz_dose2_early <- hz_dose2_early %>%
      filter(patient_derived_upi_number %in% hz_dose1$patient_derived_upi_number)

hz_dose2_earlySumm <- hz_dose2_early %>%
  group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source, Date_Administered, `hz_interval (days)`) %>%
  summarise(record_count = n())

hz_dose2_early <- hz_dose2_early %>% select(-Date_Administered,-CHIcheck)

### CREATE TABLE OF RECORDS & SUMMARY OF SHINGRIX DOSE 2 GIVEN AFTER ZOSTAVAX DOSE 1
hz_zostd1 <- hz_dose1 %>% 
  filter(vacc_product_name == "Shingles Vaccine Zostavax Merck")

hz_shingrixd1 <- hz_dose1 %>% 
  filter(vacc_product_name == "Shingles Vaccine Shingrix GlaxoSmithKline")

hz_shingrixd2 <- hz_dose2 %>%
  filter(vacc_product_name == "Shingles Vaccine Shingrix GlaxoSmithKline")

# list IDs of patients with D2 Shingrix & D1 Zostavax with no D1 Shingrix
hz_wrongvaxtypeIDs <- hz_shingrixd2 %>%
  filter(patient_derived_upi_number %in% hz_zostd1$patient_derived_upi_number &
           !patient_derived_upi_number %in% hz_shingrixd1$patient_derived_upi_number)

# select relevant records for the ID list created above
hz_wrongvaxtype <- rbind(hz_shingrixd2,hz_zostd1) %>%
  filter(patient_derived_upi_number %in% hz_wrongvaxtypeIDs$patient_derived_upi_number)

hz_wrongvaxtype_sort <- hz_wrongvaxtype %>% select(vacc_event_created_at,patient_derived_upi_number) %>%
  group_by(patient_derived_upi_number) %>%
  summarise(vacc_event_created_at=max(vacc_event_created_at)) %>%
  rename(latest_date = vacc_event_created_at)

hz_wrongvaxtype <- hz_wrongvaxtype %>% left_join(hz_wrongvaxtype_sort,by="patient_derived_upi_number") %>%
  arrange(desc(latest_date)) %>%
  mutate(sort_date = latest_date) %>% 
  mutate(QueryName = "17. HZ Shingrix dose 2, Zostavax dose 1") %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  select(-latest_date)

hz_wrongvaxtypeSumm <- hz_wrongvaxtype %>%
  group_by(vacc_location_health_board_name, vacc_data_source, Date_Administered) %>%
  summarise(record_count = n())

hz_wrongvaxtype <- hz_wrongvaxtype %>% select(-Date_Administered,-CHIcheck)

rm(hz_zostd1,hz_shingrixd1,hz_shingrixd2,hz_wrongvaxtypeIDs,hz_wrongvaxtype_sort)

### CREATE TABLE OF RECORDS & SUMMARY OF PATIENTS WITH A DOSE 2 RECORD BUT NO DOSE 1
hz_dose2_nodose1 <- anti_join(hz_dose2, hz_dose1, by="patient_derived_upi_number") %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "18. HZ Dose 2, no dose 1") 

hz_dose2_nodose1Summ <- hz_dose2_nodose1 %>%
  group_by(vacc_location_health_board_name, vacc_data_source, Date_Administered) %>%
  summarise(record_count = n())

hz_dose2_nodose1 <- hz_dose2_nodose1 %>% select(-Date_Administered,-CHIcheck)

### CREATE TABLE OF RECORDS & SUMMARY OF ZOSTAVAX VACCINE GIVEN AFTER 01.09.2023
hz_zostavax_error <- HZVaxData %>% 
  filter(vacc_product_name == "Shingles Vaccine Zostavax Merck" &
           vacc_occurence_time >= "2023-09-01") %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "19. HZ Zostavax given after 01.09.2023") 

hz_zostavax_errorSumm <- hz_zostavax_error %>%
  group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source, Date_Administered) %>%
  summarise(record_count = n())

hz_zostavax_error <- hz_zostavax_error %>% select(-Date_Administered,-CHIcheck)

### CREATE TABLE OF RECORDS & SUMMARY OF VACC GIVEN OUTWITH AGE GUIDELINES
### NOTE THAT THIS QUERY DOES NOT CONSIDER IMMUNOCOMPROMISED PATIENTS

# patients not aged 70-79 on 1st Sep 2021 and vaccinated 2021-22
hz_ageDQ1 <- HZVaxData %>% 
  filter(between(vacc_occurence_time,as.Date("2021-09-01"),as.Date("2022-08-31")) &
           !between(patient_date_of_birth,as.Date("1941-09-02"),as.Date("1951-09-01")))         

# patients not aged 70-79 on 1st Sep 2022 and vaccinated 2022-23
hz_ageDQ2 <- HZVaxData %>%
  filter(between(vacc_occurence_time,as.Date("2022-09-01"),as.Date("2023-08-31")) &
           !between(patient_date_of_birth,as.Date("1942-09-02"),as.Date("1952-09-01")))         

# patients not aged 65 or 70-79 on 1st Sep 2023 and vaccinated 2023-24
hz_ageDQ3 <- HZVaxData %>%
  filter(between(vacc_occurence_time,as.Date("2023-09-01"),as.Date("2024-08-31")) &
           !between(patient_date_of_birth,as.Date("1957-09-02"),as.Date("1958-09-01")) &
           !between(patient_date_of_birth,as.Date("1943-09-02"),as.Date("1953-09-01")))         

hz_ageDQ <- rbind(hz_ageDQ1,hz_ageDQ2,hz_ageDQ3) %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "20. HZ Vacc given outwith age guidance") 

rm(hz_ageDQ1,hz_ageDQ2,hz_ageDQ3)

hz_ageDQSumm <- hz_ageDQ %>%
  group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source, Date_Administered, age_at_vacc) %>%
  summarise(record_count = n())

hz_ageDQ <- hz_ageDQ %>% select(-Date_Administered,-CHIcheck)

### CREATE & SAVE OUT SHINGLES VACC SUMMARY REPORT
#############################################################################################

#Collates pivot tables into a single report
HZSummaryReport <-
  list("System Summary" = HZSystemSummary,
       "CHI Check" = hz_chi_check,
       "Vacc Product Type" = hz_vacc_prodSumm,
       "Missing CHIs" = hz_chi_naSumm,
       "2 or More First Doses" = hz_dose1x2Summ,
       "2 or More Second Doses" = hz_dose2x2Summ,
       "Dose 2 Zostavax" = hz_dose2_ZostSumm,
       "Dose 2 Shingrix<60days" = hz_dose2_earlySumm,
       "Shingrix D2 Zostavax D1" = hz_wrongvaxtypeSumm,
       "Dose 2, no dose 1" = hz_dose2_nodose1Summ,
       "Zostavax given after 01.09.2023" = hz_zostavax_errorSumm,
       "Vacc outwith age guidance" = hz_ageDQSumm)

rm(HZVaxData,HZSystemSummary,hz_chi_check,hz_vacc_prodSumm,hz_chi_naSumm,
   hz_dose1x2Summ,hz_dose2x2Summ,hz_dose2_ZostSumm,hz_dose2_earlySumm,
   hz_wrongvaxtypeSumm,hz_dose2_nodose1Summ,hz_zostavax_errorSumm,hz_ageDQSumm,
   hz_dose1,hz_dose2)

#Saves out collated tables into an excel file
if (answer==1) {
  write.xlsx(HZSummaryReport,
             paste("DQ Summary Reports/Shingles_Vacc_DQ_4wk_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
             asTable = TRUE,
             colWidths = "auto")
  } else {
    write.xlsx(HZSummaryReport,
            paste("DQ Summary Reports/Shingles_Vacc_DQ_Full_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
            asTable = TRUE,
            colWidths = "auto")
  }

rm(HZSummaryReport)

### COLLATE SHINGLES QUERIES FOR VACCINATIONS DQ REPORT
#############################################################################################

multi_vacc <- multi_vacc %>% rbind(hz_chi_na,hz_dose1x2,hz_dose2x2)
hz_vacc <- rbind(hz_dose2_Zost,hz_dose2_early,hz_wrongvaxtype,hz_dose2_nodose1,
                 hz_zostavax_error,hz_ageDQ)

rm(hz_chi_na,hz_dose1x2,hz_dose2x2,hz_dose2_Zost,hz_dose2_early,hz_wrongvaxtype,
   hz_dose2_nodose1,hz_zostavax_error,hz_ageDQ)

gc()

########################################################################
###############################  SECTION D4 ############################
########################  PNEUMOCOCCAL VACCINATIONS ####################
########################################################################

### CREATE PNEUMOCOCCAL VACCINATIONS DATAFRAME
#############################################################################################

### EXTRACT VARIABLES FROM COMPLETED PNEUMOCOCCAL VACC RECORDS IN vaccination_event_analysis VIEW
PneumVaxData <- odbc::dbGetQuery(conn, "select 
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
                          AND vacc_type_target_disease='Pneumococcal infectious disease (disorder)'
                          AND vacc_occurence_time > ?",
                           params = reporting_start_date-1825)

### CLEAN INVALID CHARACTERS FROM FREE-TEXT FIELDS IN EVENT ANALYSIS DATA
PneumVaxData$vacc_location_name <- textclean::replace_non_ascii(PneumVaxData$vacc_location_name,replacement = "")
PneumVaxData$vacc_performer_name <- textclean::replace_non_ascii(PneumVaxData$vacc_performer_name,replacement = "")

### CALCULATE NEW COLUMN FOR DAYS BETWEEN VACCINATION AND RECORD CREATION
PneumVaxData$vacc_record_date <- as.Date(substr(PneumVaxData$vacc_record_created_at,1,10))
PneumVaxData$days_between_vacc_and_recording <- as.double(difftime(PneumVaxData$vacc_record_date,PneumVaxData$vacc_occurence_time,units="days"))

### CREATE BLANK PLACEHOLDER COLUMNS
PneumVaxData$dups_same_day_flag <- NA
PneumVaxData$VMT_only_dups_flag <- NA
PneumVaxData$"cov_booster_interval (days)" <- NA
PneumVaxData$"flu_interval (days)" <- NA
PneumVaxData$"hz_interval (days)" <- NA

### CREATE AN INTERVAL BETWEEN PNEUMOCOCCAL VACCINES PER PATIENT
PneumVaxData <- PneumVaxData %>%
  arrange(patient_derived_encrypted_upi,PneumVaxData$vacc_occurence_time) %>% 
  group_by(patient_derived_encrypted_upi) %>%
  mutate(dose_number = row_number()) %>% mutate(doses = n()) %>% ungroup() %>% 
  mutate(prev_vacc_date=lag(vacc_occurence_time)) %>%
  mutate("pneum_vacc_interval (weeks)"=as.integer(difftime(vacc_occurence_time,prev_vacc_date,units = "weeks")))

PneumVaxData$"pneum_vacc_interval (weeks)" <- ifelse(PneumVaxData$dose_number=="1",NA,PneumVaxData$"pneum_vacc_interval (weeks)")

### CREATE BLANK PLACEHOLDER COLUMN
PneumVaxData$sort_date <- NA

### CREATE PNEUOMOCOCCAL VACC DATAFRAME BY JOINING PNEUOMOCOCCAL VACC EVENT RECORDS WITH PATIENT RECORDS
PneumVaxData <- PneumVaxData %>%
  inner_join(Vaxpatientraw, by=("source_system_patient_id")) %>%
  arrange(desc(vacc_event_created_at)) %>% # sort data by date amended
  select(-patient_derived_encrypted_upi,-vacc_record_date,-dose_number,-doses,-prev_vacc_date) %>% # remove temp items from Interval calculation
  mutate(Date_Administered = substr(vacc_occurence_time, 1, 10)) %>% # create new vacc date data item for date only (no time)
  mutate(CHIcheck = phsmethods::chi_check(patient_derived_chi_number)) # create CHI check data item

### CREATE PNEUOMOCOCCAL VACCINATIONS DQ QUERIES
#############################################################################################

### CREATE SUMMARY TABLE OF CHI STATUS USING PHSMETHODS CHI_CHECK
pneum_chi_check <- PneumVaxData %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name, vacc_data_source, CHIcheck) %>%
  summarise(record_count = n()) 
 
### CREATE SUMMARY TABLE OF VACC PRODUCT TYPE
pneum_vacc_prodSumm <- PneumVaxData %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name, vacc_data_source, vacc_product_name) %>%
  summarise(record_count = n()) 

### CREATE TABLE AND SUMMARY OF PATIENTS MISSING (NA) CHI NUMBER
pneum_chi_na  <- PneumVaxData %>% filter(grepl("Missing",CHIcheck)) %>%
  filter(vacc_event_created_at >= reporting_start_date) %>%
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "01. Missing CHI number") 

pneum_chi_naSumm <- pneum_chi_na %>% group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source, Date_Administered) %>% 
  summarise(record_count = n())

pneum_chi_na <- pneum_chi_na %>% select(-Date_Administered,-CHIcheck)

### CREATE TABLE OF RECORDS & SUMMARY WITH INTERVAL OF LESS THAN 5YRS (260  WEEKS)
pneum_early <- PneumVaxData %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  filter(`pneum_vacc_interval (weeks)` <260)

pneum_earlySumm <- pneum_early %>%
  group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source, Date_Administered, `pneum_vacc_interval (weeks)`) %>%
  summarise(record_count = n())

pneum_early <- pneum_early %>%
  select(-Date_Administered,-CHIcheck) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "21. PNEUM Vacc interval < 260 weeks") 

### MONITOR MULTIPLE DOSE FOR PATIENTS NOT IN AN AT-RISK GROUP i.e. AGED 65+
pneum_ageDQ_oldSumm <- PneumVaxData %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  filter(vacc_dose_number > 1 ) %>%
  filter(age_at_vacc > 65)  %>%
  group_by(vacc_location_health_board_name, vacc_data_source, Date_Administered, age_at_vacc) %>%
  summarise(record_count = n())

### CREATE TABLE OF RECORDS & SUMMARY OF PATIENTS AGED <2 AT VACCINATION
pneum_ageDQ <- PneumVaxData %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  filter(age_at_vacc < 2)

pneum_ageDQSumm <- pneum_ageDQ %>%
  group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source, Date_Administered, age_at_vacc) %>%
  summarise(record_count = n())

pneum_ageDQ <- pneum_ageDQ %>%
  select(-Date_Administered,-CHIcheck) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "22. PNEUM Age < 2 years") 

### CREATE & SAVE OUT PNEUMOCOCCAL VACC SUMMARY REPORT
#############################################################################################

#Collates pivot tables into a single report
PneumSummaryReport <-
  list("System Summary" = PneumSystemSummary,
       "CHI Check" = pneum_chi_check,
       "Vacc Product Type" = pneum_vacc_prodSumm,
       "Missing CHIs" = pneum_chi_naSumm,
       "Interval <5years" = pneum_earlySumm,
       "Multiples Doses Age>65" = pneum_ageDQ_oldSumm,
       "Vaccine Given to Age<2" = pneum_ageDQSumm)

rm(PneumVaxData,PneumSystemSummary,pneum_chi_check,pneum_vacc_prodSumm,
   pneum_chi_naSumm,pneum_earlySumm,pneum_ageDQ_oldSumm,pneum_ageDQSumm)

#Saves out collated tables into an excel file
if (answer==1) {
  write.xlsx(PneumSummaryReport,
           paste("DQ Summary Reports/Pneumococcal_Vacc_DQ_4wk_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
           asTable = TRUE,
           colWidths = "auto")
  } else {
  write.xlsx(PneumSummaryReport,
             paste("DQ Summary Reports/Pneumococcal_Vacc_DQ_Full_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
             asTable = TRUE,
             colWidths = "auto")
    }

rm(PneumSummaryReport)

### COLLATE PNEUMOCOCCAL QUERIES FOR VACCINATIONS DQ REPORT
#############################################################################################

multi_vacc <- multi_vacc %>% rbind(pneum_chi_na)
pneum_vacc <- rbind(pneum_early,pneum_ageDQ)

rm(pneum_chi_na,pneum_early,pneum_ageDQ)

gc()

########################################################################
############################### SECTION E ##############################
########  UNIQUE DQ IDS, SORT & ORDER DATA, ALL QUERIES TABLE ##########
########################################################################

### CREATE UNIQUE DQ ID FOR ALL RECORDS IN multi_vacc TABLE
multi_vacc$DQ_ID <- paste(substr(multi_vacc$QueryName,1,2),
                  multi_vacc$vacc_source_system_event_id,
                  substr(multi_vacc$vacc_event_created_at,1,10),sep = ".")

### UPDATE DQ ID FOR DUPLICATE DOSE 1 QUERY
d1ID <- multi_vacc %>% filter(grepl("02.",multi_vacc$QueryName)) %>% 
  group_by(patient_derived_upi_number) %>% 
  summarise(DQ_ID = first(DQ_ID)) %>% 
  rename("maind1DQID" = DQ_ID)

multi_vacc <- multi_vacc %>% left_join(d1ID,by="patient_derived_upi_number")

multi_vacc$DQ_ID <- ifelse(grepl("02.",multi_vacc$QueryName),multi_vacc$maind1DQID,multi_vacc$DQ_ID)

### UPDATE DQ ID FOR DUPLICATE DOSE 2 QUERY
d2ID <- multi_vacc %>% filter(grepl("03.",multi_vacc$QueryName)) %>% 
  group_by(patient_derived_upi_number) %>% 
  summarise(DQ_ID = first(DQ_ID)) %>% 
  rename("maind2DQID" = DQ_ID)

multi_vacc <- multi_vacc %>% left_join(d2ID,by="patient_derived_upi_number")

multi_vacc$DQ_ID <- ifelse(grepl("03.",multi_vacc$QueryName),multi_vacc$maind2DQID,multi_vacc$DQ_ID)

### UPDATE DQ ID FOR DUPLICATE DOSE 3 QUERY
d3ID <- multi_vacc %>% filter(grepl("04.",multi_vacc$QueryName)) %>% 
  group_by(patient_derived_upi_number) %>% 
  summarise(DQ_ID = first(DQ_ID)) %>% 
  rename("maind3DQID" = DQ_ID)

multi_vacc <- multi_vacc %>% left_join(d3ID,by="patient_derived_upi_number")

multi_vacc$DQ_ID <- ifelse(grepl("04.",multi_vacc$QueryName),multi_vacc$maind3DQID,multi_vacc$DQ_ID)

### UPDATE DQ ID FOR DUPLICATE DOSE 4 QUERY
d4ID <- multi_vacc %>% filter(grepl("05.",multi_vacc$QueryName)) %>% 
  group_by(patient_derived_upi_number) %>% 
  summarise(DQ_ID = first(DQ_ID)) %>% 
  rename("maind4DQID" = DQ_ID)

multi_vacc <- multi_vacc %>% left_join(d4ID,by="patient_derived_upi_number")

multi_vacc$DQ_ID <- ifelse(grepl("05.",multi_vacc$QueryName),multi_vacc$maind4DQID,multi_vacc$DQ_ID)

### TIDY UP multi_vacc DATA TABLE
multi_vacc <- multi_vacc %>% arrange(desc(sort_date),DQ_ID,desc(vacc_event_created_at)) %>% 
  select(QueryName,DQ_ID,patient_derived_chi_number,patient_derived_upi_number,
         everything(),-maind1DQID,-maind2DQID,
         -maind3DQID,-maind4DQID)

rm(d1ID,d2ID,d3ID,d4ID)

### CREATE UNIQUE DQ ID FOR ALL RECORDS IN covid_vacc TABLE
covid_vacc$DQ_ID <- paste(substr(covid_vacc$QueryName,1,2),
                          covid_vacc$vacc_source_system_event_id,
                          substr(covid_vacc$vacc_event_created_at,1,10),sep = ".")

### UPDATE DQ ID FOR COVID DOSE 1 & BOOSTER QUERY
d1andbID <- covid_vacc %>% filter(grepl("11.",covid_vacc$QueryName)) %>% 
  group_by(patient_derived_upi_number) %>% 
  summarise(DQ_ID = first(DQ_ID)) %>% 
  rename("maind1andbDQID" = DQ_ID)

covid_vacc <- covid_vacc %>% left_join(d1andbID,by="patient_derived_upi_number")

covid_vacc$DQ_ID <- ifelse(grepl("11.",covid_vacc$QueryName),covid_vacc$maind1andbDQID,covid_vacc$DQ_ID)

### TIDY UP covid_vacc DATA TABLE
covid_vacc <- covid_vacc %>% arrange(desc(sort_date),DQ_ID,desc(vacc_event_created_at)) %>% 
  select(QueryName,DQ_ID,patient_derived_chi_number,patient_derived_upi_number,
         everything(),-maind1andbDQID)

rm(d1andbID)

### CREATE UNIQUE DQ ID FOR ALL RECORDS IN flu_vacc TABLE
flu_vacc$DQ_ID <- paste(substr(flu_vacc$QueryName,1,2),
                        flu_vacc$vacc_source_system_event_id,
                          substr(flu_vacc$vacc_event_created_at,1,10),sep = ".")

### TIDY UP flu_vacc DATA TABLE
flu_vacc <- flu_vacc %>% arrange(desc(sort_date),DQ_ID,desc(vacc_event_created_at)) %>% 
  select(QueryName,DQ_ID,patient_derived_chi_number,patient_derived_upi_number,
         everything())


### CREATE UNIQUE DQ ID FOR ALL RECORDS IN hz_vacc TABLE
hz_vacc$DQ_ID <- paste(substr(hz_vacc$QueryName,1,2),
                       hz_vacc$vacc_source_system_event_id,
                          substr(hz_vacc$vacc_event_created_at,1,10),sep = ".")

### UPDATE DQ ID FOR SHINGLES SHINGRIX DOSE 2, ZOSTAVAX DOSE 1 QUERY
Zd1ID <- hz_vacc %>% filter(grepl("17.",hz_vacc$QueryName)) %>% 
  group_by(patient_derived_upi_number) %>% 
  summarise(DQ_ID = first(DQ_ID)) %>% 
  rename("mainZd1ID" = DQ_ID)

hz_vacc <- hz_vacc %>% left_join(Zd1ID,by="patient_derived_upi_number")

hz_vacc$DQ_ID <- ifelse(grepl("17.",hz_vacc$QueryName),hz_vacc$mainZd1ID,hz_vacc$DQ_ID)

### TIDY UP hz_vacc DATA TABLE
hz_vacc <- hz_vacc %>% arrange(desc(sort_date),DQ_ID,desc(vacc_event_created_at)) %>% 
  select(QueryName,DQ_ID,patient_derived_chi_number,patient_derived_upi_number,
         everything(),-mainZd1ID)

rm(Zd1ID)

### CREATE UNIQUE DQ ID FOR ALL RECORDS IN pneum_vacc TABLE
pneum_vacc$DQ_ID <- paste(substr(pneum_vacc$QueryName,1,2),
                          pneum_vacc$vacc_source_system_event_id,
                        substr(pneum_vacc$vacc_event_created_at,1,10),sep = ".")

### TIDY UP pneum_vacc DATA TABLE
pneum_vacc <- pneum_vacc %>% arrange(desc(sort_date),DQ_ID,desc(vacc_event_created_at)) %>% 
  select(QueryName,DQ_ID,patient_derived_chi_number,patient_derived_upi_number,
         everything())

### CREATE A DATA TABLE OF ALL DATA QUERY RECORDS FOR HB REPORTS
all_queries <- rbind(multi_vacc,covid_vacc,flu_vacc,hz_vacc,pneum_vacc) %>% 
  arrange(desc(sort_date),DQ_ID,desc(vacc_event_created_at)) %>% 
  select(-sort_date)

### FINALISE DATA TABLES FOR HB REPORTS
multi_vacc <- multi_vacc %>% select(-source_system_patient_id,-sort_date)
covid_vacc <- covid_vacc %>% select(-source_system_patient_id,-sort_date)
flu_vacc <- flu_vacc %>% select(-source_system_patient_id,-sort_date)
hz_vacc <- hz_vacc %>% select(-source_system_patient_id,-sort_date)
pneum_vacc <- pneum_vacc %>% select(-source_system_patient_id,-sort_date)


########################################################################
################################ SECTION F #############################
##############################  HB REPORTS #############################
########################################################################

### CREATE DQ OUTPUTS FOR EACH HEALTH BOARD AND SAVE TO FOLDERS

for(i in 1:16) {
  
  df_multi <- multi_vacc %>% filter(vacc_location_health_board_name == hb_name[i])
  df_cov <- covid_vacc %>% filter(vacc_location_health_board_name == hb_name[i])
  df_flu <- flu_vacc %>% filter(vacc_location_health_board_name == hb_name[i])
  df_hz <- hz_vacc %>% filter(vacc_location_health_board_name == hb_name[i])
  df_pneum <- pneum_vacc %>% filter(vacc_location_health_board_name == hb_name[i])
  df_all <- all_queries %>% filter(vacc_location_health_board_name == hb_name[i])
  
  df_summ <- df_all %>% group_by(QueryName) %>% 
    summarise(patient_count = n_distinct(patient_derived_upi_number)) 
  
  if (n_distinct(df_all$source_system_patient_id
                 [df_all$QueryName=="01. Missing CHI number"])>0) {
    df_summ[1,2] = n_distinct(df_all$source_system_patient_id [df_all$QueryName=="01. Missing CHI number"])
    }
  
  df_all <- df_all %>% select(-source_system_patient_id)
  
  if (nrow(df_summ)>0) {
  
      HBReport <-
        list("Summary" = df_summ,
             "Multi Vacc - Q01-05" = df_multi,
             "Covid-19 Vacc - Q06-11" = df_cov,
             "Flu Vacc - Q12-14" = df_flu,
             "Herpes Zoster Vacc - Q15-20" =  df_hz,
             "Pneumococcal Vacc - Q21-22" = df_pneum,
             "All query data - Q01-22" = df_all )

    HBReportWB <- buildWorkbook(HBReport, asTable = TRUE)
    setColWidths(HBReportWB,sheet = 1,cols = 1,widths = "auto")
    setColWidths(HBReportWB,sheet = 2,cols = 1,widths = "auto")
    setColWidths(HBReportWB,sheet = 3,cols = 1,widths = "auto")
    setColWidths(HBReportWB,sheet = 4,cols = 1,widths = "auto")
    setColWidths(HBReportWB,sheet = 5,cols = 1,widths = "auto")
    setColWidths(HBReportWB,sheet = 6,cols = 1,widths = "auto")
    setColWidths(HBReportWB,sheet = 7,cols = 1,widths = "auto")
    

    insertImage(HBReportWB,sheet = 1,file = "Vacc HB DQ Summary.jpg",
                  width = 10,height = 10,startRow = (nrow(df_summ) + 3),startCol = 1, units = "in",dpi = 300)

    saveWorkbook(HBReportWB,paste("DQ HB Reports/",hb_cypher[i],"_Vacc_DQ_Report_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep=""))
    
    rm(HBReport,HBReportWB)
  } else {
    writexl::write_xlsx(df_summ,paste("DQ HB Reports/",hb_cypher[i],"_NULL_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""))
  }
}
  

########################################################################
################################ SECTION G #############################
##########################  QUERY COUNT SUMMARY ########################
########################################################################

# read in latest query_count output and save as backup
temp_query_count <- readxl::read_excel("DQ Report Query count.xlsx")

write.xlsx(temp_query_count,
           paste("Archive/DQ Report Query count_backup.xlsx"),
           asTable = TRUE,
           colWidths = "auto")

### CREATE SUMMARY TABLE OF ALL QUERIES PRODUCED IN LATEST RUN
query_count <- all_queries %>% group_by(QueryName) %>% 
  summarise("record_count_latest_run"=n()) %>% 
  mutate(date_latest_run = Sys.Date()) %>% 
  select(QueryName,date_latest_run,record_count_latest_run)

### READ IN AND UPDATE QUERY COUNT FROM PREVIOUS RUN
df <- temp_query_count %>% 
  select(-date_previous_run,-record_count_previous_run) %>% 
  rename("date_previous_run" = date_latest_run) %>% 
  rename("record_count_previous_run" = record_count_latest_run)

# join figures for latest and previous DQ reports for comparison
df2 <- query_count %>% full_join(df, by="QueryName") %>% 
  mutate(diff_since_last_run = record_count_latest_run - record_count_previous_run)

df2$date_latest_run <- as.Date(df2$date_latest_run,"%Y-%m-%d")
df2$date_previous_run <- as.Date(df2$date_previous_run,"%Y-%m-%d")

# Save out collated tables into an excel file
write.xlsx(df2,
           paste("DQ Report Query count.xlsx"),
           asTable = TRUE,
           colWidths = "auto")

rm(df,df2,query_count)

end.time <- Sys.time()

time.taken <- round(end.time - start.time,2)

time.taken



############# END OF SCRIPT ###########################



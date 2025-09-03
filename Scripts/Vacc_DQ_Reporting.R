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

# Save console outputs to Scripts folder
console <- file("Scripts/DQ console output.txt")
sink(console)
sink(file = console, type = "message")

start.time <- Sys.time()
print(paste("Start time: ", start.time))

########################################################################
###############################  SECTION A #############################
###  LIBRARIES, CONNECTION, SET REPORTING PERIOD, LIST HEALTH BOARDS ###
########################################################################

### INSTALL PACKAGES - FIRST TIME USERS MAY NEED TO INSTALL PACKAGES
### BEFORE LOADING
# install.packages("svDialogs")

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

reporting_start_date = as.Date("2020-12-01")
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

### CREATE SUMMARY TABLE OF RSV VACC RECORDS
rsv_system_summary <- Vaxeventtot %>%
  filter(vacc_type_target_disease == "Respiratory syncytial virus infection (disorder)" &
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

vdl_summ <- file.path("Outputs/VDL_summary.xlsx")

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
           paste("Outputs/VDL_summary.xlsx"),
           asTable = TRUE,
           colWidths = "auto")

# save backup of workbook every month
if (lubridate::day(Sys.Date()) < 8) {
  write.xlsx(SystemSummaryFull,
           paste("Outputs/Archive/VDL_summary_backup.xlsx"),
           asTable = TRUE,
           colWidths = "auto") }

# Save previous run's figures as backup
write.xlsx(sheet1,
           paste("Outputs/Archive/VDL_summary_sheet1_backup.xlsx"),
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
                        reporting_location_type,
                        vacc_location_derived_location_type,
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

### EXTRACT COVID-19 COHORT RECORDS FROM vaccination_patient_cohort_analysis VIEW
cov_cohort <- odbc::dbGetQuery(conn, "select source_system_patient_id,
                                      cohort,
                                      # cohort_source,
                                      # cohort_extract_time,
                                      cohort_reporting_label,
                                      cohort_description,
                                      cohort_target_diseases,
                                      # patient_cohort_created_at,
                                      # patient_cohort_updated_at,
                                      # patient_cohort_removal_datetime,
                                      # patient_cohort_removal_status,
                                      cohort_phase
                          from vaccination.vaccination_patient_cohort_analysis_audit
            where cohort_target_diseases like '%840539006 - COVID-19%' ")

table(cov_cohort$cohort_description,useNA = "ifany")
table(cov_cohort$cohort_phase,useNA = "ifany")
table(cov_cohort$cohort [is.na(cov_cohort$cohort_phase)])

cov_cohort$cohort_phase [cov_cohort$cohort=="NHS_STAFF_2022"] <-
  "Autumn Winter 2022_23"
cov_cohort$cohort_phase [cov_cohort$cohort_phase=="Autumn Winter 23_24"] <-
  "Autumn Winter 2023_24"
cov_cohort$cohort_phase <- ifelse((is.na(cov_cohort$cohort_phase) |
                             cov_cohort$cohort_phase=="July_24" |
                             cov_cohort$cohort_phase=="September 2023" |
                             cov_cohort$cohort_phase=="Imported from NCDS"),
                        cov_cohort$cohort_description,cov_cohort$cohort_phase)

table(cov_cohort$cohort_phase,useNA = "ifany")

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
CovVaxData$"rsv_interval (days)" <- NA
CovVaxData$sort_date <- NA

### CREATE COVID-19 VACC DATAFRAME BY JOINING COVID VACC EVENT RECORDS WITH PATIENT & COHORT RECORDS
CovVaxData <- CovVaxData %>%
  inner_join(Vaxpatientraw, by=("source_system_patient_id")) %>%
  arrange(desc(vacc_event_created_at)) %>% # sort data by date amended
  select(-patient_derived_encrypted_upi,-vacc_record_date,-dose_number,-doses,-prev_vacc_date) %>% # remove temp items from Interval calculation
  mutate(CHIcheck = phsmethods::chi_check(patient_derived_chi_number)) # create CHI check data item

### CREATE VACC OCCURENCE PHASE TO LINK COHORT DATA BY COHORT PHASE
CovVaxData$vacc_phase <- NA
CovVaxData$vacc_phase [CovVaxData$vacc_occurence_time>=as.Date("2020-12-01") &
                         CovVaxData$vacc_occurence_time<as.Date("2021-09-01")] <-
  "Tranche1"
CovVaxData$vacc_phase [CovVaxData$vacc_occurence_time>=as.Date("2021-09-01") &
                         CovVaxData$vacc_occurence_time<as.Date("2022-04-01")] <-
  "Tranche2"
CovVaxData$vacc_phase [CovVaxData$vacc_occurence_time>=as.Date("2022-04-01") &
                         CovVaxData$vacc_occurence_time<as.Date("2022-09-01")] <-
  "COVID Spring Booster 2022"
CovVaxData$vacc_phase [CovVaxData$vacc_occurence_time>=as.Date("2022-09-01") &
                         CovVaxData$vacc_occurence_time<as.Date("2023-04-01")] <-
  "Autumn Winter 2022_23"
CovVaxData$vacc_phase [CovVaxData$vacc_occurence_time>=as.Date("2023-04-01") &
                         CovVaxData$vacc_occurence_time<as.Date("2023-09-01")] <-
  "Spring 2023"
CovVaxData$vacc_phase [CovVaxData$vacc_occurence_time>=as.Date("2023-09-01") &
                         CovVaxData$vacc_occurence_time<as.Date("2024-04-01")] <-
  "Autumn Winter 2023_24"
CovVaxData$vacc_phase [CovVaxData$vacc_occurence_time>=as.Date("2024-04-01") &
                         CovVaxData$vacc_occurence_time<as.Date("2024-07-01")] <-
  "Spring 2024"
CovVaxData$vacc_phase [CovVaxData$vacc_occurence_time>=as.Date("2024-09-23") &
                         CovVaxData$vacc_occurence_time<as.Date("2025-02-01")] <-
  "Autumn Winter 2024_25"
CovVaxData$vacc_phase [CovVaxData$vacc_occurence_time>=as.Date("2025-03-31") &
                         CovVaxData$vacc_occurence_time<as.Date("2025-07-01")] <-
  "Spring 2025"

table(CovVaxData$vacc_phase,useNA = "ifany")

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

### CREATE SUMMARY TABLE OF VACC COHORTS
cov_vacc_cohorts <- CovVaxData %>% 
  left_join(cov_cohort, by=(c("source_system_patient_id","vacc_phase"="cohort_phase"))) %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name, cohort, vacc_phase) %>%
  summarise(record_count = n()) 

### CREATE TABLE AND SUMMARY OF PATIENTS WITH MISSING (NA) OR INVALID CHI NUMBER
cov_chi_inv <- CovVaxData %>% filter(CHIcheck!="Valid CHI") %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>%
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = paste0("01. Missing/Invalid CHI number - ",CHIcheck))

cov_chi_invSumm <- cov_chi_inv %>% group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source,CHIcheck) %>% 
  summarise(record_count = n())

cov_chi_inv <- cov_chi_inv %>% select(-CHIcheck)

CovVaxData <- CovVaxData %>% select(-CHIcheck)

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
  select(-latest_date)

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
  select(-latest_date)

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
  select(-latest_date)

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
  select(-latest_date)

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
  mutate(QueryName = "06. COV Booster interval < 12 weeks")

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
  filter(age_at_vacc<12 & !grepl("child",vacc_product_name) &
           !grepl("Child",vacc_product_name) &
           !grepl("inf",vacc_product_name) &
           !grepl("Inf",vacc_product_name) &
           !grepl("paed",vacc_product_name) &
           !grepl("Paed",vacc_product_name)) %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "07. COV Under 12 adult dose")

cov_under12fulldoseSumm <- cov_under12fulldose %>% 
  group_by(vacc_location_health_board_name,vacc_location_name,vacc_product_name,age_at_vacc) %>%
  summarise(record_count=n())

### CREATE TABLE & SUMMARY OF PATIENTS AGED OVER 11, OR UNDER 5 AFTER 01.04.2023, GIVEN PAEDIATRIC DOSE
# this excludes those aged 12 who received a dose 1 paediatric dose at age 11
cov_over12_under5_childdose <- CovVaxData %>%
  filter((age_at_vacc>12 | (age_at_vacc<5 & vacc_occurence_time >= "2023-04-01")) &
           (grepl("Paed",vacc_product_name) | grepl("paed",vacc_product_name) |
              grepl("5-11",vacc_product_name)))

cov_age11childdose <- CovVaxData %>% filter(age_at_vacc == "11" &
                                              (grepl("Paed",vacc_product_name) | grepl("paed",vacc_product_name) |
                                                 grepl("5-11",vacc_product_name)))
cov_age12childdose <- CovVaxData %>% filter(age_at_vacc == "12" &
                                              (grepl("Paed",vacc_product_name) | grepl("paed",vacc_product_name) |
                                                 grepl("5-11",vacc_product_name)))

cov_over11_under5_childdose <- anti_join(cov_age12childdose,cov_age11childdose,by="patient_derived_upi_number") %>%
  rbind(cov_over12_under5_childdose) %>%
  arrange(desc(vacc_event_created_at)) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "08. COV Over 11 or under 5 paediatric dose")

cov_over11_under5_childdoseSumm <- cov_over11_under5_childdose %>% 
  group_by(vacc_location_health_board_name,vacc_location_name,vacc_product_name,age_at_vacc) %>%
  summarise(record_count=n())

rm(cov_over12_under5_childdose,cov_age11childdose,cov_age12childdose)

### CREATE TABLE & SUMMARY OF PATIENTS AGED OVER 4 GIVEN INFANT DOSE
# this excludes those aged 5 who received a dose 1 infant dose at age 4
cov_over5_infant_dose <- CovVaxData %>% 
  filter(age_at_vacc>5 & (CovVaxData$vacc_product_name == "Covid-19 mRNA Vaccine Comirnaty 3mcg (Infant)" |
                          CovVaxData$vacc_product_name == "Comirnaty XBB.1.5 Children 6 months - 4 years COVID-19 mRNA Vaccine 3micrograms/0.2ml dose"))

cov_over5_infant_dose <- CovVaxData %>% 
  filter(age_at_vacc>5 & (grepl("inf",vacc_product_name) | grepl("Inf",vacc_product_name) |
                            grepl("6 months",vacc_product_name)))

cov_age4_infantdose <- CovVaxData %>% 
  filter(age_at_vacc==4 & (grepl("inf",vacc_product_name) | grepl("Inf",vacc_product_name) |
                             grepl("6 months",vacc_product_name)))
  
cov_age5_infantdose <- CovVaxData %>% 
  filter(age_at_vacc==5 & (grepl("inf",vacc_product_name) | grepl("Inf",vacc_product_name) |
                             grepl("6 months",vacc_product_name)))

cov_over4_infant_dose <- anti_join(cov_age5_infantdose,cov_age4_infantdose,by="patient_derived_upi_number") %>% 
  rbind(cov_over5_infant_dose) %>% 
  arrange(desc(vacc_event_created_at)) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "09. COV Over 4 infant dose")

cov_over4_infant_doseSumm <- cov_over4_infant_dose %>% 
  group_by(vacc_location_health_board_name,vacc_location_name,vacc_product_name,age_at_vacc) %>%
  summarise(record_count=n())

rm(cov_age4_infantdose,cov_age5_infantdose,cov_over5_infant_dose)

### CREATE TABLE OF PATIENTS WITH A DOSE 2 RECORD BUT NO DOSE 1
cov_dose2nodose1 <- anti_join(cov_dose2, cov_dose1, by="patient_derived_upi_number") %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(QueryName = "10. COV Dose 2, no dose 1")

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
  select(-latest_date)

rm(cov_dose1andboostersort,cov_dose1andboosterIDs)

cov_dose1andboosterSumm <- cov_dose1andbooster %>% group_by(vacc_location_health_board_name) %>% 
  summarise(record_count=n())

### CREATE TABLE OF PATIENTS WITH 2 OR MORE BOOSTERS IN SINGLE CAMPAIGN
# ONLY ON DATA FROM SPRING 2024 PROG ONWARDS

cov_booster <- cov_booster %>% 
  filter(vacc_occurence_time>=as.Date("2024-04-01"))

cov_boosterx2ids <- cov_booster %>%
  group_by(patient_derived_upi_number,vacc_phase) %>% 
  summarise(count_by_patient_derived_upi_number = n()) %>%
  na.omit(count_by_patient_derived_upi_number) %>%
  filter(count_by_patient_derived_upi_number > 1)

cov_boosterx2 <- cov_booster %>%
  left_join(cov_boosterx2ids,by=c("patient_derived_upi_number","vacc_phase")) %>% 
  filter(count_by_patient_derived_upi_number > 1) %>% 
  select(-count_by_patient_derived_upi_number)

cov_boosterx2sort <- cov_boosterx2 %>% select(vacc_event_created_at,patient_derived_upi_number,vacc_phase) %>%
  group_by(patient_derived_upi_number,vacc_phase) %>%
  summarise(vacc_event_created_at=max(vacc_event_created_at)) %>%
  rename(latest_date = vacc_event_created_at)

cov_boosterx2 <- cov_boosterx2 %>% 
  left_join(cov_boosterx2sort,by=c("patient_derived_upi_number","vacc_phase")) %>%
  arrange(desc(latest_date)) %>% 
  mutate(sort_date = latest_date) %>%
  mutate(QueryName = "25. COV Two or more boosters in campaign") %>% 
  filter(sort_date >= reporting_start_date) %>% 
  select(-c(latest_date))

rm(cov_boosterx2ids,cov_boosterx2sort)

if (nrow(cov_boosterx2)>0) {
  
  # Create flag for duplicates occuring on same day
  cov_boosterx2_samedateIDs <- cov_boosterx2 %>% 
    group_by(patient_derived_upi_number,vacc_occurence_time) %>% 
    summarise(nn = n()) %>%
    filter(nn > 1)
  
  cov_boosterx2 <- cov_boosterx2 %>% 
    left_join(cov_boosterx2_samedateIDs,join_by(patient_derived_upi_number,vacc_occurence_time))
  
  cov_boosterx2$dups_same_day_flag <- 0
  cov_boosterx2$dups_same_day_flag [cov_boosterx2$nn>0] <- 1
  ###
  
  # Create flag for VMT-only duplicates
  cov_boosterx2_VMT <- cov_boosterx2 %>% filter(vacc_data_source == "TURAS")
  
  cov_boosterx2_VMT_ids <- cov_boosterx2_VMT %>% 
    group_by(patient_derived_upi_number,vacc_phase) %>% 
    summarise(nn2=n()) %>% ungroup() %>% 
    filter(nn2>1)
  
  cov_boosterx2 <- cov_boosterx2 %>% 
    left_join(cov_boosterx2_VMT_ids,by=c("patient_derived_upi_number","vacc_phase"))
  
  cov_boosterx2$VMT_only_dups_flag <- 0
  cov_boosterx2$VMT_only_dups_flag [cov_boosterx2$nn2>0] <- 1
  ###
  
  cov_boosterx2 <- cov_boosterx2 %>% select(-nn,-nn2)
  
  rm(cov_boosterx2_samedateIDs,cov_boosterx2_VMT,cov_boosterx2_VMT_ids)
}

cov_boosterx2Summ <- cov_boosterx2 %>% group_by(vacc_location_health_board_name,
                                                vacc_phase) %>% 
  summarise(record_count = n())

### CREATE TABLE & SUMMARY OF VACCINATIONS OUTWITH VERIFIED VACC PROGRAMME (vacc_phase)
# ONLY ON DATA FROM SPRING 2024 ONWARDS

cov_outwith_prog <- CovVaxData %>% 
  filter(vacc_occurence_time>=as.Date("2024-04-01") &
           is.na(vacc_phase)) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "26. COV Outwith vacc programme")

cov_outwith_progSumm <- cov_outwith_prog %>% 
  group_by(vacc_location_health_board_name,vacc_location_name,vacc_occurence_time) %>% 
  summarise(record_count = n())

### CREATE TABLE OF RECORDS & SUMMARY OF VACC GIVEN OUTWITH AGE GUIDELINES
### THAT ARE NOT IN SIS ELIGIBILITY COHORT OR IN A CARE HOME
# ONLY ON DATA FROM SPRING 2025 PROG ONWARDS

# patients aged <75 on 30th Jun 2025 and non-cohort, or <6months on 31st March 2025,
# and vaccinated Spring 2025
cov_ageDQ <- CovVaxData %>% 
  filter(vacc_occurence_time>=as.Date("2025-03-31") &
           reporting_location_type!="CARE HOME CODE" &
           vacc_location_derived_location_type!="Home for the Elderly" &
           vacc_location_derived_location_type!="Private Nursing Home, Private Hospital etc" &
           !grepl("Care Home",vacc_location_name) &
           !grepl("CARE HOME",vacc_location_name) &
           !grepl("Carehome",vacc_location_name) &
           !grepl("Nursing Home",vacc_location_name) &
           !grepl("Residential",vacc_location_name) &
           !grepl("Residents",vacc_location_name)) %>% 
left_join(cov_cohort, by=(c("source_system_patient_id","vacc_phase"="cohort_phase"))) %>%
  filter(patient_date_of_birth>as.Date("1950-06-30") & is.na(cohort) |
           patient_date_of_birth>as.Date("2024-09-30")) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "27. COV Vacc given outwith age guidance") 

cov_ageDQSumm <- cov_ageDQ %>%
  group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source, age_at_vacc) %>%
  summarise(record_count = n())

cov_ageDQ <- cov_ageDQ %>% select(-c(cohort:cohort_target_diseases))


### CREATE & SAVE OUT COVID-19 VACC SUMMARY REPORT
#############################################################################################

#Collates pivot tables into a single report
cov_SummaryReport <-
  list("System Summary" = CovSystemSummary,
       "CHI Check" = cov_chi_check,
       "Vacc Product Type" = cov_vacc_prodSumm,
       "Dose Given After 01.04.2024" = cov_vacc_doseSumm,
       "Vacc Cohorts" = cov_vacc_cohorts,
       "Missing or Invalid CHIs" = cov_chi_invSumm,
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
       "Dose 1 and booster, no dose 2" = cov_dose1andboosterSumm,
       "2 or more boosters in campaign" = cov_boosterx2Summ,
       "Outwith vacc programme" = cov_outwith_progSumm,
       "Vacc given outwith age guidance" = cov_ageDQSumm)

rm(CovVaxData,CovSystemSummary,cov_chi_check,cov_vacc_prodSumm,
   cov_vacc_doseSumm,cov_vacc_cohorts,cov_chi_invSumm,
   cov_dose1x2Summ,cov_dose2x2Summ,cov_dose3x2Summ,cov_dose4x2Summ,cov_booster_intervalDQSumm,
   cov_age12Summ,cov_under12fulldoseSumm,cov_over11_under5_childdoseSumm,
   cov_over4_infant_doseSumm,cov_dose2nodose1Summ,cov_dose1andboosterSumm,
   cov_dose1,cov_dose2,cov_dose3,cov_dose4,cov_booster,cov_cohort,
   cov_boosterx2Summ,cov_outwith_progSumm,cov_ageDQSumm)

#Saves out collated tables into an excel file
if (answer==1) {
  write.xlsx(cov_SummaryReport,
           paste("Outputs/DQ Summary Reports/Covid-19_Vacc_DQ_4wk_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
           asTable = TRUE,
           colWidths = "auto")
  } else {
  write.xlsx(cov_SummaryReport,
             paste("Outputs/DQ Summary Reports/Covid-19_Vacc_DQ_Full_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
             asTable = TRUE,
             colWidths = "auto")
  }

rm(cov_SummaryReport)

### COLLATE COVID QUERIES FOR VACCINATIONS DQ REPORT
#############################################################################################

multi_vacc <- rbind(cov_chi_inv,cov_dose1x2,cov_dose2x2,cov_dose3x2,cov_dose4x2) %>% 
  select(-c(reporting_location_type,vacc_location_derived_location_type))
covid_vacc <- rbind(cov_booster_intervalDQ,cov_under12fulldose,cov_over11_under5_childdose,
                    cov_over4_infant_dose,cov_dose2nodose1,cov_dose1andbooster,
                    cov_boosterx2,cov_outwith_prog,cov_ageDQ)%>% 
  select(-c(reporting_location_type,vacc_location_derived_location_type))

rm(cov_chi_inv,cov_dose1x2,cov_dose2x2,cov_dose3x2,cov_dose4x2,
   cov_booster_intervalDQ,cov_under12fulldose,cov_over11_under5_childdose,
   cov_over4_infant_dose,cov_dose2nodose1,cov_dose1andbooster,
   cov_boosterx2,cov_outwith_prog,cov_ageDQ)

gc()

saveRDS(multi_vacc,"Outputs/Temp/multi_vacc.rds")
# multi_vacc <- readRDS("Outputs/Temp/multi_vacc.rds")
saveRDS(covid_vacc,"Outputs/Temp/covid_vacc.rds")
# covid_vacc <- readRDS("Outputs/Temp/covid_vacc.rds")

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
                          AND vacc_clinical_trial_flag='0' ")

### CLEAN INVALID CHARACTERS FROM FREE-TEXT FIELDS IN EVENT ANALYSIS DATA
FluVaxData$vacc_location_name <- textclean::replace_non_ascii(FluVaxData$vacc_location_name,replacement = "")
FluVaxData$vacc_performer_name <- textclean::replace_non_ascii(FluVaxData$vacc_performer_name,replacement = "")

### EXTRACT SHINGLES COHORT RECORDS FROM vaccination_patient_cohort_analysis VIEW
flu_cohort <- odbc::dbGetQuery(conn, "select source_system_patient_id,
                                      cohort,
                                      # cohort_reporting_label,
                                      cohort_description,
                                      cohort_target_diseases,
                                      # patient_cohort_created_at,
                                      # patient_cohort_updated_at,
                                      cohort_phase
                          from vaccination.vaccination_patient_cohort_analysis_audit
            where cohort_target_diseases like '%Influenza%' ")

table(flu_cohort$cohort_target_diseases,useNA = "ifany")
table(flu_cohort$cohort_description,useNA = "ifany")
table(flu_cohort$cohort_phase,useNA = "ifany")
table(flu_cohort$cohort [is.na(flu_cohort$cohort_phase)])

flu_cohort <- flu_cohort %>% filter(!grepl("Spring", cohort_phase))

flu_cohort$cohort_phase [flu_cohort$cohort=="NHS_STAFF_2022"] <-
  "Autumn Winter 2022_23"
flu_cohort$cohort_phase [flu_cohort$cohort_phase==c("Autumn Winter 23_24")] <-
  "Autumn Winter 2023_24"
flu_cohort$cohort_phase [flu_cohort$cohort %in% c("NHS_STAFF","NHS_STAFF_2") &
                           is.na(flu_cohort$cohort_phase)] <-
  "Autumn Winter 2024_25"
flu_cohort$cohort_phase [flu_cohort$cohort_phase %in%
                           c("Imported from NCDS","July_24","September 2023")] <-
  "Autumn Winter 2024_25"

table(flu_cohort$cohort_phase,useNA = "ifany")

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
FluVaxData$"rsv_interval (days)" <- NA
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
FluVaxData$vacc_phase [FluVaxData$vacc_occurence_time>=as.Date("2021-09-01") &
                         FluVaxData$vacc_occurence_time<as.Date("2022-04-01")] <-
  "Tranche2"
FluVaxData$vacc_phase [FluVaxData$vacc_occurence_time>=as.Date("2022-09-01") &
                         FluVaxData$vacc_occurence_time<as.Date("2023-04-01")] <-
  "Autumn Winter 2022_23"
FluVaxData$vacc_phase [FluVaxData$vacc_occurence_time>=as.Date("2023-09-01") &
                         FluVaxData$vacc_occurence_time<as.Date("2024-04-01")] <-
  "Autumn Winter 2023_24"
FluVaxData$vacc_phase [FluVaxData$vacc_occurence_time>=as.Date("2024-09-01") &
                         FluVaxData$vacc_occurence_time<as.Date("2025-04-01")] <-
  "Autumn Winter 2024_25"
FluVaxData$vacc_phase [FluVaxData$vacc_occurence_time>=as.Date("2025-09-01") &
                         FluVaxData$vacc_occurence_time<as.Date("2026-03-29")] <-
  "Autumn Winter 2025_26"

table(FluVaxData$vacc_phase,useNA = "ifany")
FluVaxData$vacc_phase [FluVaxData$vacc_occurence_time>=as.Date("2021-04-01") &
                         is.na(FluVaxData$vacc_phase)] <-
  "Outwith Autumn/Winter"
table(FluVaxData$vacc_phase,useNA = "ifany")


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

### CREATE SUMMARY TABLE OF VACC COHORTS
flu_vacc_cohorts <- FluVaxData %>% 
  left_join(flu_cohort, by=(c("source_system_patient_id","vacc_phase"="cohort_phase"))) %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name, cohort, vacc_phase) %>%
  summarise(record_count = n()) 

### CREATE TABLE AND SUMMARY OF PATIENTS WITH MISSING (NA) OR INVALID CHI NUMBER
flu_chi_inv <- FluVaxData %>% filter(CHIcheck!="Valid CHI") %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = paste0("01. Missing/Invalid CHI number - ",CHIcheck))

flu_chi_invSumm <- flu_chi_inv %>% group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source, CHIcheck,Date_Administered) %>% 
  summarise(record_count = n())

flu_chi_inv <- flu_chi_inv %>% select(-Date_Administered,-CHIcheck)

FluVaxData <- FluVaxData %>% select(-CHIcheck)

### CREATE TABLE OF PATIENTS WITH 2 OR MORE dose 1 RECORDS AND SUMMARY
flu_dose1 <- FluVaxData |>
  filter(vacc_dose_number == "1" &
           !(is.na(vacc_phase)) &
           vacc_phase != "Outwith Autumn/Winter") |> 
  arrange(desc(patient_derived_upi_number))

table(flu_dose1$vacc_phase)

## Modified to use vacc_phase
flu_multidose1IDs <- flu_dose1 |> 
  group_by(patient_derived_upi_number, vacc_phase) |> 
  summarise(count_ID = n()) |> 
  na.omit(count_ID) 

flu_multidose1 <- flu_dose1 |> 
  left_join(flu_multidose1IDs,by=c("patient_derived_upi_number","vacc_phase")) |> 
  filter(count_ID>1) |> 
  select(-count_ID)

flu_multidose1sort <- flu_multidose1 %>% 
  select(vacc_event_created_at,patient_derived_upi_number) %>%
  group_by(patient_derived_upi_number) %>% 
  summarise(vacc_event_created_at=max(vacc_event_created_at)) %>%
  rename(latest_date = vacc_event_created_at)

flu_multidose1 <- flu_multidose1 %>% 
  left_join(flu_multidose1sort,by="patient_derived_upi_number") %>% 
  arrange(desc(latest_date)) %>% 
  mutate(sort_date = latest_date) %>% 
  mutate(QueryName = "02. FLU Two or more dose 1") %>%
  filter(sort_date >= reporting_start_date) %>% 
  select(-latest_date, -Date_Administered) 

if (nrow(flu_multidose1)>0) {
  
  # Create flag for duplicates on the same day
  flu_multidose1_samedayIDs <- flu_multidose1 %>% 
    group_by(patient_derived_upi_number,vacc_occurence_time) %>% 
    summarise(nn = n()) %>%
    filter(nn > 1)
  
  flu_multidose1 <- flu_multidose1 %>% 
    left_join(flu_multidose1_samedayIDs,join_by(patient_derived_upi_number,vacc_occurence_time))
  
  flu_multidose1$dups_same_day_flag <- 0
  flu_multidose1$dups_same_day_flag [flu_multidose1$nn>0] <- 1
  
  # Create flag for VMT-only doses dups
  flu_multidose1_VMT <- flu_multidose1 %>% filter(vacc_data_source == "TURAS")
  
  flu_multidose1_VMT_ids <- flu_multidose1_VMT %>% 
    group_by(patient_derived_upi_number) %>% 
    summarise(nn2=n()) %>% ungroup() %>% 
    filter(nn2>1)
  
  flu_multidose1 <- flu_multidose1 %>% 
    left_join(flu_multidose1_VMT_ids,by="patient_derived_upi_number")
  
  flu_multidose1$VMT_only_dups_flag <- 0
  flu_multidose1$VMT_only_dups_flag [flu_multidose1$nn2>0] <- 1
  
  flu_multidose1 <- flu_multidose1 %>% select(-nn,-nn2)
  
  rm(flu_multidose1_VMT,flu_multidose1_VMT_ids,flu_multidose1_samedayIDs)
}

flu_multidose1Summ <- flu_multidose1 %>% group_by(vacc_location_health_board_name) %>% 
  summarise(record_count = n())

rm(flu_multidose1IDs,flu_multidose1sort)

### CREATE TABLE OF PATIENTS WITH 2 OR MORE dose 2 RECORDS AND SUMMARY
flu_dose2 <- FluVaxData |> 
  filter(vacc_dose_number == "2" &
           !(is.na(vacc_phase)) &
           vacc_phase != "Outwith Autumn/Winter") |> 
  arrange(desc(patient_derived_upi_number))

## Modified to use vacc_phase
flu_multidose2IDs <- flu_dose2 |> 
  group_by(patient_derived_upi_number, vacc_phase) |> 
  summarise(count_ID = n()) |> 
  na.omit(count_ID) 

flu_multidose2 <- flu_dose2 |> 
  left_join(flu_multidose2IDs,by=c("patient_derived_upi_number","vacc_phase")) |> 
  filter(count_ID>1) |> 
  select(-count_ID)

flu_multidose2sort <- flu_multidose2 %>% 
  select(vacc_event_created_at,patient_derived_upi_number) %>% 
  group_by(patient_derived_upi_number) %>% 
  summarise(vacc_event_created_at=max(vacc_event_created_at)) %>%
  rename(latest_date = vacc_event_created_at)

flu_multidose2 <- flu_multidose2 %>% 
  left_join(flu_multidose2sort,by="patient_derived_upi_number") %>% 
  arrange(desc(latest_date)) %>% 
  mutate(sort_date = latest_date) %>% 
  mutate(QueryName = "03. FLU Two or more dose 2") %>%
  filter(sort_date >= reporting_start_date) %>% 
  select(-latest_date, -Date_Administered) 

if (nrow(flu_multidose2)>0) {
  
  # Create flag for duplicates on the same day
  flu_multidose2_samedayIDs <- flu_multidose2 %>% 
    group_by(patient_derived_upi_number,vacc_occurence_time) %>% 
    summarise(nn = n()) %>%
    filter(nn > 1)
  
  flu_multidose2 <- flu_multidose2 %>% 
    left_join(flu_multidose2_samedayIDs,join_by(patient_derived_upi_number,vacc_occurence_time))
  
  flu_multidose2$dups_same_day_flag <- 0
  flu_multidose2$dups_same_day_flag [flu_multidose2$nn>0] <- 1
  
  # Create flag for VMT-only doses dups
  flu_multidose2_VMT <- flu_multidose2 %>% filter(vacc_data_source == "TURAS")
  
  flu_multidose2_VMT_ids <- flu_multidose2_VMT %>% 
    group_by(patient_derived_upi_number) %>% 
    summarise(nn2=n()) %>% ungroup() %>% 
    filter(nn2>1)
  
  flu_multidose2 <- flu_multidose2 %>% 
    left_join(flu_multidose2_VMT_ids,by="patient_derived_upi_number")
  
  flu_multidose2$VMT_only_dups_flag <- 0
  flu_multidose2$VMT_only_dups_flag [flu_multidose2$nn2>0] <- 1
  
  flu_multidose2 <- flu_multidose2 %>% select(-nn,-nn2)
  
  rm(flu_multidose2_VMT,flu_multidose2_VMT_ids,flu_multidose2_samedayIDs)
}

flu_multidose2Summ <- flu_multidose2 %>% group_by(vacc_location_health_board_name) %>% 
  summarise(record_count = n())

rm(flu_multidose2IDs,flu_multidose2sort)

### CREATE TABLE OF RECORDS & SUMMARY OF VACCINATIONS GIVEN AT INTERVAL OF < 4 WEEKS
flu_interval <- FluVaxData %>% filter(`flu_interval (days)`<28) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "12. FLU interval < 4 weeks")

flu_intervalSumm <- flu_interval %>% group_by(vacc_location_health_board_name,vacc_location_name,`flu_interval (days)`,Date_Administered) %>% 
  summarise(record_count = n())

flu_interval <- flu_interval %>% select(-Date_Administered)
  
### CREATE TABLE OF RECORDS & SUMMARY OF VACCINATIONS THAT OCCUR BEFORE 06/09/2021
flu_early <- FluVaxData %>% filter(vacc_occurence_time < "2021-09-06") %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "13. FLU vacc before 06.09.2021")

flu_earlySumm <- flu_early %>% group_by(vacc_location_health_board_name,vacc_location_name,Date_Administered) %>%
  summarise(record_count = n())

flu_early <- flu_early %>% select(-Date_Administered)

### CREATE TABLE OF RECORDS & SUMMARY OF LAIV (FLUENZ) VACCINATION GIVEN AT AGE <2 OR >18
flu_laiv_ageDQ <- FluVaxData %>%
  filter(grepl("Fluenz", vacc_product_name)) %>% 
  filter(age_at_vacc < 2 | age_at_vacc > 18) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "14. FLU LAIV given age <2 or >18")

flu_laiv_ageDQSumm <- flu_laiv_ageDQ %>%
  group_by(vacc_location_health_board_name,vacc_location_name,
           vacc_product_name,Date_Administered, age_at_vacc) %>%
  summarise(record_count = n())

flu_laiv_ageDQ <- flu_laiv_ageDQ %>% select(-Date_Administered)

### CREATE TABLE OF RECORDS & SUMMARY OF AQIV VACCINATION GIVEN AT AGE <50
flu_aqiv_ageDQ <- FluVaxData %>%
  filter(vacc_product_name=="Adjuvanted Quadrivalent Influenza Vaccine Seqirus" &
           vacc_occurence_time>="2024-09-01" &
           age_at_vacc < 50) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "28. FLU AQIV given age <50")

flu_aqiv_ageDQSumm <- flu_aqiv_ageDQ %>%
  group_by(vacc_location_health_board_name, vacc_location_name, Date_Administered, age_at_vacc) %>%
  summarise(record_count = n())

flu_aqiv_ageDQ <- flu_aqiv_ageDQ %>% select(-Date_Administered)

### CREATE TABLE AND SUMMARY OF PATIENTS OF VACCINATIONS GIVEN OUTWITH AUTUMN/WINTER CAMPAIGNS
flu_outwith <- FluVaxData %>% filter(vacc_phase == "Outwith Autumn/Winter") %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "29. FLU Outwith Autumn/Winter campaigns")

flu_outwithSumm <- flu_outwith %>% 
  group_by(vacc_location_health_board_name,vacc_location_name,vacc_phase,Date_Administered) %>% 
  summarise(record_count = n())

flu_outwith <- flu_outwith %>% select(-Date_Administered)

### CREATE & SAVE OUT FLU VACC SUMMARY REPORT
#############################################################################################

#Collates pivot tables into a single report
flu_SummaryReport <-
  list("System Summary" = FluSystemSummary,
       "CHI Check" = flu_chi_check,
       "Vacc Product Type" = flu_vacc_prodSumm,
       "Vacc Cohorts" = flu_vacc_cohorts,
       "Missing or Invalid CHIs" = flu_chi_invSumm,
       "2 or more dose 1" = flu_multidose1Summ,
       "2 or more dose 2" = flu_multidose2Summ,
       "Interval < 4 weeks" =  flu_intervalSumm,
       "Vaccine before 06.09.2021" = flu_earlySumm,
       "LAIV given age <2 or >18" = flu_laiv_ageDQSumm,
       "AQIV given age <50" = flu_aqiv_ageDQSumm,
       "Outwith Autumn-Winter Campaigns" = flu_outwithSumm)

rm(FluVaxData,FluSystemSummary,flu_chi_check,flu_vacc_prodSumm,flu_vacc_cohorts,
   flu_chi_invSumm,flu_intervalSumm,flu_earlySumm,flu_cohort,
   flu_multidose1Summ,flu_multidose2Summ,flu_laiv_ageDQSumm,flu_aqiv_ageDQSumm,
   flu_outwithSumm,flu_dose1,flu_dose2)

#Saves out collated tables into an excel file
if (answer==1) {
  write.xlsx(flu_SummaryReport,
           paste("Outputs/DQ Summary Reports/Flu_Vacc_DQ_4wk_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
           asTable = TRUE,
           colWidths = "auto")
  } else {
  write.xlsx(flu_SummaryReport,
             paste("Outputs/DQ Summary Reports/Flu_Vacc_DQ_Full_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
             asTable = TRUE,
             colWidths = "auto")
  }

rm(flu_SummaryReport)

### COLLATE FLU QUERIES FOR VACCINATIONS DQ REPORT
#############################################################################################

multi_vacc <- multi_vacc %>% rbind(flu_chi_inv,flu_multidose1,flu_multidose2)
flu_vacc <- rbind(flu_interval,flu_early,flu_laiv_ageDQ,flu_aqiv_ageDQ,flu_outwith)

rm(flu_chi_inv,flu_multidose1,flu_multidose2,flu_interval,flu_early,
   flu_laiv_ageDQ,flu_aqiv_ageDQ,flu_outwith)

gc()

saveRDS(multi_vacc,"Outputs/Temp/multi_vacc.rds")
# multi_vacc <- readRDS("Outputs/Temp/multi_vacc.rds")
saveRDS(flu_vacc,"Outputs/Temp/flu_vacc.rds")
# flu_vacc <- readRDS("Outputs/Temp/flu_vacc.rds")

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

### EXTRACT SHINGLES COHORT RECORDS FROM vaccination_patient_cohort_analysis VIEW
hz_cohort <- odbc::dbGetQuery(conn, "select source_system_patient_id,
                                      cohort,
                                      # cohort_reporting_label,
                                      cohort_description,
                                      cohort_target_diseases,
                                      # patient_cohort_created_at,
                                      # patient_cohort_updated_at,
                                      cohort_phase
                          from vaccination.vaccination_patient_cohort_analysis_audit
                          where (cohort like '%SHINGLES%' OR
                          cohort_target_diseases='4740000 - Herpes zoster (disorder)') ")

table(hz_cohort$cohort_target_diseases, useNA = "ifany")
table(hz_cohort$cohort_phase, useNA = "ifany")
table(hz_cohort$cohort, useNA = "ifany")

hz_cohort$cohort_phase [is.na(hz_cohort$cohort_target_diseases) &
                          hz_cohort$cohort_phase=="Scottish Immunisation Programme"] <-
  "Sept21_Aug22"
hz_cohort$cohort_phase [hz_cohort$cohort_phase=="Scottish Immunisation Programme"] <-
  "Sept22_Aug23"

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
HZVaxData$"rsv_interval (days)" <- NA
HZVaxData$sort_date <- NA

### CREATE SHINGLES VACC DATAFRAME BY JOINING SHINGLES VACC EVENT RECORDS WITH PATIENT RECORDS
HZVaxData <- HZVaxData %>%
  inner_join(Vaxpatientraw, by=("source_system_patient_id")) %>%
  arrange(desc(vacc_event_created_at)) %>% # sort data by date amended
  select(-patient_derived_encrypted_upi,-vacc_record_date,-dose_number,-doses,-prev_vacc_date) %>%
  mutate(Date_Administered = substr(vacc_occurence_time, 1, 10)) %>% # create new vacc date data item for date only (no time)
  mutate(CHIcheck = phsmethods::chi_check(patient_derived_chi_number)) # create CHI check data item

### CREATE VACC OCCURENCE PHASE TO LINK COHORT DATA BY COHORT PHASE
HZVaxData$vacc_phase <- NA 
HZVaxData$vacc_phase [between(HZVaxData$vacc_occurence_time,
                              as.Date("2021-09-01"),as.Date("2022-08-31"))] <-
  "Sept21_Aug22"
HZVaxData$vacc_phase [between(HZVaxData$vacc_occurence_time,
                              as.Date("2022-09-01"),as.Date("2023-08-31"))] <-
  "Sept22_Aug23"
HZVaxData$vacc_phase [between(HZVaxData$vacc_occurence_time,
                              as.Date("2023-09-01"),as.Date("2024-08-31"))] <-
  "Sept23_Aug24"
HZVaxData$vacc_phase [between(HZVaxData$vacc_occurence_time,
                              as.Date("2024-09-01"),as.Date("2025-08-31"))] <-
  "Sept24_Aug25"

table(HZVaxData$vacc_phase, useNA = "ifany")

### CREATE A COHORT LIST OF SEVERELY IMMUNOSUPPRESSED (SIS)
hz_cohort_sis <- hz_cohort %>% filter(grepl("IMMUNO_SUPP",cohort)) %>% 
  distinct()

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

### CREATE SUMMARY TABLE OF VACC COHORTS
hz_vacc_cohorts <- HZVaxData %>% 
  left_join(hz_cohort, by=(c("source_system_patient_id","vacc_phase"="cohort_phase"))) %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name, cohort, vacc_phase) %>%
  summarise(record_count = n()) 

### CREATE TABLE AND SUMMARY OF PATIENTS WITH MISSING (NA) OR INVALID CHI NUMBER
hz_chi_inv  <- HZVaxData %>% filter(CHIcheck!="Valid CHI") %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = paste0("01. Missing/Invalid CHI number - ",CHIcheck))

hz_chi_invSumm <- hz_chi_inv %>%
  group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source, CHIcheck, Date_Administered) %>% 
  summarise(record_count = n())

hz_chi_inv <- hz_chi_inv %>% select(-CHIcheck,-Date_Administered)

HZVaxData <- HZVaxData %>% select(-CHIcheck)

### CREATE TABLE OF RECORDS & SUMMARY OF 2 OR MORE DOSE 1 VACCINATIONS (OF SAME
### VACC PRODUCT - THIS EXCLUDES PERSONS NEW TO WIS LIST WHO HAVE PREVIOUSLY
### RECEIVED ZOSTAVAX - ALSO EXCLUDES SIS COHORT)
hz_dose1 <- HZVaxData %>% filter(vacc_dose_number == "1")

hz_dose1x2IDs <- hz_dose1 %>% group_by(patient_derived_upi_number,vacc_product_name) %>% 
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

hz_dose1x2 <- hz_dose1x2 %>%
  filter(!source_system_patient_id %in% hz_cohort_sis$source_system_patient_id)

hz_dose1x2Summ <- hz_dose1x2 %>%
  group_by(vacc_location_health_board_name, vacc_data_source, Date_Administered) %>%
  summarise(record_count = n())

hz_dose1x2 <- hz_dose1x2 %>% select(-Date_Administered)

### CREATE TABLE OF RECORDS & SUMMARY OF 2 OR MORE DOSE 2 VACCINATIONS (OF SAME
### VACC PRODUCT - THIS EXCLUDES PERSONS NEW TO WIS LIST WHO HAVE PREVIOUSLY
### RECEIVED ZOSTAVAX - ALSO EXCLUDES SIS COHORT)
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

hz_dose2x2 <- hz_dose2x2 %>%
  filter(!source_system_patient_id %in% hz_cohort_sis$source_system_patient_id)

hz_dose2x2Summ <- hz_dose2x2 %>%
  group_by(vacc_location_health_board_name, vacc_data_source, Date_Administered) %>%
  summarise(record_count = n())

hz_dose2x2 <- hz_dose2x2 %>% select(-Date_Administered)

### CREATE TABLE OF RECORDS & SUMMARY OF ZOSTAVAX GIVEN FOR DOSE 2
hz_dose2_Zost <- hz_dose2 %>% 
  filter(vacc_product_name == "Shingles Vaccine Zostavax Merck") %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "15. HZ Dose 2 Zostavax")

hz_dose2_ZostSumm <- hz_dose2_Zost %>%
  group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source, Date_Administered) %>%
  summarise(record_count = n())

hz_dose2_Zost <- hz_dose2_Zost %>% select(-Date_Administered)

### CREATE TABLE OF RECORDS & SUMMARY OF SHINGRIX DOSE 2 GIVEN AT AN INTERVAL OF <56 DAYS
hz_dose2_early <- hz_dose2 %>% 
  filter(vacc_product_name == "Shingles Vaccine Shingrix GlaxoSmithKline") %>%
  filter(`hz_interval (days)` < 56) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "16. HZ Dose 2 interval < 56 days")

    # remove patients with no dose 1
    hz_dose2_early <- hz_dose2_early %>%
      filter(patient_derived_upi_number %in% hz_dose1$patient_derived_upi_number)

hz_dose2_earlySumm <- hz_dose2_early %>%
  group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source, Date_Administered, `hz_interval (days)`) %>%
  summarise(record_count = n())

hz_dose2_early <- hz_dose2_early %>% select(-Date_Administered)

### CREATE TABLE OF RECORDS & SUMMARY OF SHINGRIX DOSE 2 GIVEN AFTER ZOSTAVAX DOSE 1
### AND NO DOSE 1 SHINGRIX
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

hz_wrongvaxtype <- hz_wrongvaxtype %>% select(-Date_Administered)

rm(hz_zostd1,hz_shingrixd1,hz_shingrixd2,hz_wrongvaxtypeIDs,hz_wrongvaxtype_sort)

### CREATE TABLE OF RECORDS & SUMMARY OF PATIENTS WITH A DOSE 2 RECORD BUT NO DOSE 1
hz_dose2_nodose1 <- anti_join(hz_dose2, hz_dose1, by="patient_derived_upi_number") %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "18. HZ Dose 2, no dose 1") 

hz_dose2_nodose1Summ <- hz_dose2_nodose1 %>%
  group_by(vacc_location_health_board_name, vacc_data_source, Date_Administered) %>%
  summarise(record_count = n())

hz_dose2_nodose1 <- hz_dose2_nodose1 %>% select(-Date_Administered)

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

hz_zostavax_error <- hz_zostavax_error %>% select(-Date_Administered)

### CREATE TABLE OF RECORDS & SUMMARY OF VACC GIVEN OUTWITH AGE GUIDELINES
### THAT ARE NOT IN SIS ELIGIBILITY COHORT

# patients not aged 70-79 on 1st Sep 2021 and vaccinated 2021-22
hz_ageDQ1 <- HZVaxData %>% 
  filter(vacc_phase=="Sept21_Aug22" &
           !between(patient_date_of_birth,as.Date("1941-09-02"),as.Date("1951-09-01")))         

# patients not aged 70-79 on 1st Sep 2022 and vaccinated 2022-23
hz_ageDQ2 <- HZVaxData %>%
  filter(vacc_phase=="Sept22_Aug23" &
           !between(patient_date_of_birth,as.Date("1942-09-02"),as.Date("1952-09-01")))

# patients not aged 65 or 70-79 on 1st Sep 2023 and vaccinated 2023-24
hz_ageDQ3 <- HZVaxData %>%
  filter(vacc_phase=="Sept23_Aug24" &
           !between(patient_date_of_birth,as.Date("1957-09-02"),as.Date("1958-09-01")) &
           !between(patient_date_of_birth,as.Date("1943-09-02"),as.Date("1953-09-01")))         

# patients not aged 65,66 or 70-79 on 1st Sep 2024 and vaccinated 2024-25
hz_ageDQ4 <- HZVaxData %>%
  filter(vacc_phase=="Sept24_Aug25" &
           !between(patient_date_of_birth,as.Date("1958-09-02"),as.Date("1959-09-01")) &
           !between(patient_date_of_birth,as.Date("1957-09-02"),as.Date("1958-09-01")) &
           !between(patient_date_of_birth,as.Date("1944-09-02"),as.Date("1954-09-01")))         

# combine 4 ageDQ dfs and remove anyone in eligibility cohorts
hz_ageDQ <- rbind(hz_ageDQ1,hz_ageDQ2,hz_ageDQ3,hz_ageDQ4) %>%
  left_join(hz_cohort, by=(c("source_system_patient_id","vacc_phase"="cohort_phase"))) %>% 
  filter(vacc_event_created_at >= reporting_start_date &
           is.na(cohort)) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "20. HZ Vacc given outwith age guidance") 

rm(hz_ageDQ1,hz_ageDQ2,hz_ageDQ3,hz_ageDQ4)

hz_ageDQSumm <- hz_ageDQ %>%
  group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source, Date_Administered, age_at_vacc) %>%
  summarise(record_count = n())

hz_ageDQ <- hz_ageDQ %>% select(-c(Date_Administered,cohort:cohort_target_diseases))

### CREATE & SAVE OUT SHINGLES VACC SUMMARY REPORT
#############################################################################################

#Collates pivot tables into a single report
HZSummaryReport <-
  list("System Summary" = HZSystemSummary,
       "CHI Check" = hz_chi_check,
       "Vacc Product Type" = hz_vacc_prodSumm,
       "Vacc Cohorts" = hz_vacc_cohorts,
       "Missing or Invalid CHIs" = hz_chi_invSumm,
       "2 or More First Doses" = hz_dose1x2Summ,
       "2 or More Second Doses" = hz_dose2x2Summ,
       "Dose 2 Zostavax" = hz_dose2_ZostSumm,
       "Dose 2 Shingrix<56days" = hz_dose2_earlySumm,
       "Shingrix D2 Zostavax D1" = hz_wrongvaxtypeSumm,
       "Dose 2, no dose 1" = hz_dose2_nodose1Summ,
       "Zostavax given after 01.09.2023" = hz_zostavax_errorSumm,
       "Vacc outwith age guidance" = hz_ageDQSumm)

rm(HZVaxData,HZSystemSummary,hz_chi_check,hz_vacc_prodSumm,hz_vacc_cohorts,
   hz_chi_invSumm,hz_dose1x2Summ,hz_dose2x2Summ,hz_dose2_ZostSumm,
   hz_dose2_earlySumm,hz_wrongvaxtypeSumm,hz_dose2_nodose1Summ,
   hz_zostavax_errorSumm,hz_ageDQSumm,hz_dose1,hz_dose2,hz_cohort,hz_cohort_sis)

#Saves out collated tables into an excel file
if (answer==1) {
  write.xlsx(HZSummaryReport,
             paste("Outputs/DQ Summary Reports/Shingles_Vacc_DQ_4wk_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
             asTable = TRUE,
             colWidths = "auto")
  } else {
    write.xlsx(HZSummaryReport,
            paste("Outputs/DQ Summary Reports/Shingles_Vacc_DQ_Full_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
            asTable = TRUE,
            colWidths = "auto")
  }

rm(HZSummaryReport)

### COLLATE SHINGLES QUERIES FOR VACCINATIONS DQ REPORT
#############################################################################################

multi_vacc <- multi_vacc %>% rbind(hz_chi_inv,hz_dose1x2,hz_dose2x2)
hz_vacc <- rbind(hz_dose2_Zost,hz_dose2_early,hz_wrongvaxtype,hz_dose2_nodose1,
                 hz_zostavax_error,hz_ageDQ)

rm(hz_chi_inv,hz_dose1x2,hz_dose2x2,hz_dose2_Zost,hz_dose2_early,hz_wrongvaxtype,
   hz_dose2_nodose1,hz_zostavax_error,hz_ageDQ)

gc()

saveRDS(multi_vacc,"Outputs/Temp/multi_vacc.rds")
# multi_vacc <- readRDS("Outputs/Temp/multi_vacc.rds")
saveRDS(hz_vacc,"Outputs/Temp/hz_vacc.rds")
# hz_vacc <- readRDS("Outputs/Temp/hz_vacc.rds")


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
                          AND vacc_type_target_disease='Pneumococcal infectious disease (disorder)' ")

### CLEAN INVALID CHARACTERS FROM FREE-TEXT FIELDS IN EVENT ANALYSIS DATA
PneumVaxData$vacc_location_name <- textclean::replace_non_ascii(PneumVaxData$vacc_location_name,replacement = "")
PneumVaxData$vacc_performer_name <- textclean::replace_non_ascii(PneumVaxData$vacc_performer_name,replacement = "")

### EXTRACT PNEUMOCOCCAL COHORT RECORDS FROM vaccination_patient_cohort_analysis VIEW
pneum_cohort <- odbc::dbGetQuery(conn, "select source_system_patient_id,
                                      cohort,
                                      cohort_extract_time,
                                      cohort_reporting_label,
                                      cohort_description,
                                      cohort_target_diseases,
                                      # patient_cohort_created_at,
                                      # patient_cohort_updated_at,
                                      cohort_phase
                          from vaccination.vaccination_patient_cohort_analysis_audit
            where cohort like '%PNEUMOCOCCAL%' ")

table(pneum_cohort$cohort_target_diseases, useNA = "ifany")
table(pneum_cohort$cohort_phase, useNA = "ifany")
table(pneum_cohort$cohort, useNA = "ifany")

pneum_cohort$cohort_phase [pneum_cohort$cohort_phase=="Scottish Immunisation Programme"] <- 
  "Apr22_Mar23"

table(pneum_cohort$cohort_phase, useNA = "ifany")

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

PneumVaxData$"rsv_interval (days)" <- NA

### CREATE BLANK PLACEHOLDER COLUMN
PneumVaxData$sort_date <- NA

### CREATE PNEUOMOCOCCAL VACC DATAFRAME BY JOINING PNEUOMOCOCCAL VACC EVENT RECORDS WITH PATIENT RECORDS
PneumVaxData <- PneumVaxData %>%
  inner_join(Vaxpatientraw, by=("source_system_patient_id")) %>%
  arrange(desc(vacc_event_created_at)) %>% # sort data by date amended
  select(-patient_derived_encrypted_upi,-vacc_record_date,-dose_number,-doses,-prev_vacc_date) %>% # remove temp items from Interval calculation
  mutate(Date_Administered = substr(vacc_occurence_time, 1, 10)) %>% # create new vacc date data item for date only (no time)
  mutate(CHIcheck = phsmethods::chi_check(patient_derived_chi_number)) # create CHI check data item

### CREATE VACC OCCURENCE PHASE TO LINK COHORT DATA BY COHORT PHASE
PneumVaxData$vacc_phase <- NA
PneumVaxData$vacc_phase [between(PneumVaxData$vacc_occurence_time,
                                 as.Date("2022-04-01"),as.Date("2023-03-31"))] <-
  "Apr22_Mar23"
PneumVaxData$vacc_phase [between(PneumVaxData$vacc_occurence_time,
                                 as.Date("2023-04-01"),as.Date("2024-03-31"))] <-
  "Apr23_Mar24"
PneumVaxData$vacc_phase [between(PneumVaxData$vacc_occurence_time,
                                 as.Date("2024-04-01"),as.Date("2025-03-31"))] <-
  "Apr24_Mar25"
PneumVaxData$vacc_phase [between(PneumVaxData$vacc_occurence_time,
                                 as.Date("2025-04-01"),as.Date("2026-03-31"))] <-
  "Apr25_Mar26"

table(PneumVaxData$vacc_phase,useNA = "ifany")

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

### CREATE SUMMARY TABLE OF VACC COHORTS
pneum_vacc_cohorts <- PneumVaxData %>% 
  left_join(pneum_cohort, by=(c("source_system_patient_id","vacc_phase"="cohort_phase"))) %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name, cohort, vacc_phase) %>%
  summarise(record_count = n()) 

### CREATE TABLE AND SUMMARY OF PATIENTS WITH MISSING (NA) OR INVALID CHI NUMBER
pneum_chi_inv  <- PneumVaxData %>% filter(CHIcheck!="Valid CHI") %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>%
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = paste0("01. Missing/Invalid CHI number - ",CHIcheck)) 

pneum_chi_invSumm <- pneum_chi_inv %>% group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source,CHIcheck,Date_Administered) %>% 
  summarise(record_count = n())

pneum_chi_inv <- pneum_chi_inv %>% select(-Date_Administered,-CHIcheck)

PneumVaxData <- PneumVaxData %>% select(-CHIcheck)

### CREATE TABLE OF RECORDS & SUMMARY WITH INTERVAL OF LESS THAN 5YRS (260  WEEKS)
pneum_early <- PneumVaxData %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  filter(`pneum_vacc_interval (weeks)` <260)

pneum_earlySumm <- pneum_early %>%
  group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source, Date_Administered, `pneum_vacc_interval (weeks)`) %>%
  summarise(record_count = n())

pneum_early <- pneum_early %>%
  select(-Date_Administered) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "21. PNEUM Vacc interval < 260 weeks") 

### MONITOR MULTIPLE DOSE FOR PATIENTS NOT IN AN AT-RISK GROUP i.e. AGED 65+
pneum_ageDQ_oldSumm <- PneumVaxData %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  filter(vacc_dose_number > 1 ) %>%
  filter(age_at_vacc > 65)  %>%
  group_by(vacc_location_health_board_name, vacc_data_source, Date_Administered, age_at_vacc) %>%
  summarise(record_count = n())

### CREATE TABLE OF RECORDS & SUMMARY OF VACC GIVEN AGE < 65
### THAT ARE NOT IN ELIGIBILITY COHORT
pneum_ageDQ <- PneumVaxData %>%
  #filter(vacc_occurence_time<"2025-04-01") %>% #temp end point until new cohorts for 2025/26 added
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  filter(age_at_vacc < 65 & !is.na(vacc_phase)) %>%
  left_join(pneum_cohort, by=(c("source_system_patient_id","vacc_phase"="cohort_phase"))) %>% 
  filter(is.na(cohort)) %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "22. PNEUM Vacc given outwith age guidance") 

pneum_ageDQSumm <- pneum_ageDQ %>% 
  group_by(vacc_location_health_board_name,vacc_location_name,vacc_data_source,
           Date_Administered,age_at_vacc,vacc_phase) %>%
  summarise(record_count = n())

pneum_ageDQ <- pneum_ageDQ %>% select(-c(Date_Administered,cohort:cohort_target_diseases))

### CREATE & SAVE OUT PNEUMOCOCCAL VACC SUMMARY REPORT
#############################################################################################

#Collates pivot tables into a single report
PneumSummaryReport <-
  list("System Summary" = PneumSystemSummary,
       "CHI Check" = pneum_chi_check,
       "Vacc Product Type" = pneum_vacc_prodSumm,
       "Vacc Cohorts" = pneum_vacc_cohorts,
       "Missing or Invalid CHIs" = pneum_chi_invSumm,
       "Interval <5years" = pneum_earlySumm,
       "Multiples Doses Age>65" = pneum_ageDQ_oldSumm,
       "Vacc Given Outwith Age Guidance" = pneum_ageDQSumm)

rm(PneumVaxData,PneumSystemSummary,pneum_chi_check,pneum_vacc_prodSumm,
   pneum_vacc_cohorts,pneum_chi_invSumm,pneum_earlySumm,pneum_ageDQ_oldSumm,
   pneum_ageDQSumm,pneum_cohort)

#Saves out collated tables into an excel file
if (answer==1) {
  write.xlsx(PneumSummaryReport,
           paste("Outputs/DQ Summary Reports/Pneumococcal_Vacc_DQ_4wk_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
           asTable = TRUE,
           colWidths = "auto")
  } else {
  write.xlsx(PneumSummaryReport,
             paste("Outputs/DQ Summary Reports/Pneumococcal_Vacc_DQ_Full_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
             asTable = TRUE,
             colWidths = "auto")
    }

rm(PneumSummaryReport)

### COLLATE PNEUMOCOCCAL QUERIES FOR VACCINATIONS DQ REPORT
#############################################################################################

multi_vacc <- multi_vacc %>% rbind(pneum_chi_inv)
pneum_vacc <- rbind(pneum_early,pneum_ageDQ)

rm(pneum_chi_inv,pneum_early,pneum_ageDQ)

gc()

saveRDS(multi_vacc,"Outputs/Temp/multi_vacc.rds")
# multi_vacc <- readRDS("Outputs/Temp/multi_vacc.rds")
saveRDS(pneum_vacc,"Outputs/Temp/pneum_vacc.rds")
# pneum_vacc <- readRDS("Outputs/Temp/pneum_vacc.rds")

########################################################################
###############################  SECTION D5 ############################
############################  RSV VACCINATIONS #########################
########################################################################

### CREATE RSV VACCINATIONS DATAFRAME
##############################################################################

### EXTRACT COMPLETED RSV VACC RECORDS IN vaccination_event_analysis VIEW
rsv_vacc <- odbc::dbGetQuery(conn, "select 
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
                          AND vacc_type_target_disease='Respiratory syncytial virus infection (disorder)'
                          AND vacc_clinical_trial_flag='0' ")

### CLEAN INVALID CHARACTERS FROM FREE-TEXT FIELDS IN EVENT ANALYSIS DATA
rsv_vacc$vacc_location_name <- textclean::replace_non_ascii(rsv_vacc$vacc_location_name,replacement = "")
rsv_vacc$vacc_performer_name <- textclean::replace_non_ascii(rsv_vacc$vacc_performer_name,replacement = "")

### EXTRACT RSV COHORT DATA FROM VDL
rsv_cohort <- odbc::dbGetQuery(conn, "select source_system_patient_id,
                                      cohort,
                                      cohort_reporting_label,
                                      cohort_description,
                                      cohort_target_diseases,
                                      # patient_cohort_created_at,
                                      # patient_cohort_updated_at,
                                      cohort_phase
                          from vaccination.vaccination_patient_cohort_analysis_audit
where cohort_target_diseases like '%55735004 - Respiratory syncytial virus infection (disorder)%' ")

table(rsv_cohort$cohort, useNA = "ifany")
table(rsv_cohort$cohort_reporting_label, useNA = "ifany")
table(rsv_cohort$cohort_description, useNA = "ifany")
table(rsv_cohort$cohort_target_diseases, useNA = "ifany")
# table(rsv_cohort$patient_cohort_created_at, useNA = "ifany")
# table(rsv_cohort$patient_cohort_updated_at, useNA = "ifany")
table(rsv_cohort$cohort_phase, useNA = "ifany")

### CREATE NEW COLUMN FOR DAYS BETWEEN VACCINATION AND RECORD CREATION
rsv_vacc$vacc_record_date <- as.Date(substr(rsv_vacc$vacc_record_created_at,1,10))
rsv_vacc$days_between_vacc_and_recording <- as.double(difftime(rsv_vacc$vacc_record_date,rsv_vacc$vacc_occurence_time,units="days"))

### CREATE BLANK PLACEHOLDER COLUMNS
rsv_vacc$dups_same_day_flag <- NA
rsv_vacc$VMT_only_dups_flag <- NA
rsv_vacc$"cov_booster_interval (days)" <- NA
rsv_vacc$"flu_interval (days)" <- NA
rsv_vacc$"hz_interval (days)" <- NA
rsv_vacc$"pneum_vacc_interval (weeks)" <- NA

### CALCULATE INTERVAL BETWEEN VACCINES PER PATIENT
rsv_vacc <- rsv_vacc %>%
  arrange(patient_derived_encrypted_upi,vacc_occurence_time) %>% 
  group_by(patient_derived_encrypted_upi) %>%
  mutate(dose_number = row_number()) %>% mutate(doses = n()) %>% ungroup() %>% 
  mutate(prev_vacc_date=lag(vacc_occurence_time)) %>%
  mutate("rsv_interval (days)"=as.integer(difftime(vacc_occurence_time,prev_vacc_date,units = "days")))

rsv_vacc$"rsv_interval (days)" <- ifelse(rsv_vacc$dose_number=="1",NA,rsv_vacc$"rsv_interval (days)")
rsv_vacc$"rsv_interval (days)" [is.na(rsv_vacc$patient_derived_encrypted_upi)] <- NA

rsv_vacc$sort_date <- NA

### CREATE RSV VACC DATAFRAME BY JOINING RSV VACC EVENT RECORDS WITH PATIENT RECORDS
rsv_vacc <- rsv_vacc %>%
  inner_join(Vaxpatientraw, by=("source_system_patient_id")) %>%
  arrange(desc(vacc_event_created_at)) %>% # sort data by date amended
  # select(-patient_derived_encrypted_upi,-vacc_record_date) %>% # remove temp items from Interval calculation
  select(-patient_derived_encrypted_upi,-vacc_record_date,-dose_number,-doses,-prev_vacc_date) %>% # remove temp items from Interval calculation
  mutate(Date_Administered = substr(vacc_occurence_time, 1, 10)) %>% # create new vacc date data item for date only (no time)
  mutate(CHIcheck = phsmethods::chi_check(patient_derived_chi_number)) # create CHI check data item

rsv_vacc$vacc_phase <- NA
rsv_vacc$vacc_phase [between(rsv_vacc$vacc_occurence_time,
                             as.Date("2024-08-01"),as.Date("2025-05-31"))] <-
  "Aug24_Jul25"
rsv_vacc$vacc_phase [between(rsv_vacc$vacc_occurence_time,
                             as.Date("2025-06-01"),as.Date("2026-05-31"))] <-
  "Aug25_Jul26"


### CREATE RSV VACCINATIONS DQ QUERIES
#############################################################################################

### CREATE SUMMARY TABLE OF PATIENT CHI STATUS USING PHSMETHODS CHI_CHECK
rsv_chi_check <- rsv_vacc %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name, vacc_data_source, CHIcheck) %>%
  summarise(record_count = n())

### CREATE SUMMARY TABLE OF VACC PRODUCT TYPE
rsv_vacc_prodSumm <- rsv_vacc %>%
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name, vacc_data_source, vacc_product_name) %>%
  summarise(record_count = n()) 

### CREATE SUMMARY TABLE OF VACC COHORTS - TRY TO INDICATE IF ONE COHORT OR MORE PER PERSON
rsv_vacc_cohorts <- rsv_vacc %>% 
  left_join(rsv_cohort,by=c("source_system_patient_id","vacc_phase"="cohort_phase")) %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name, cohort, vacc_phase) %>%
  summarise(record_count = n()) 

### CREATE SUMMARY TABLE OF DOSE NUMBERS
rsv_dose_number <- rsv_vacc %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name,vacc_dose_number,vacc_booster) %>%
  summarise(record_count = n()) 

### CREATE SUMMARY TABLE OF VACC OCCURENCE DATES
rsv_vacc_date <- rsv_vacc %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name, vacc_occurence_time) %>%
  summarise(record_count = n()) 

### CREATE SUMMARY TABLE OF VACC OCCURENCE ISO WEEKS
rsv_vacc_iso <- rsv_vacc %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  mutate(iso_week = lubridate::isoweek(vacc_occurence_time)) %>% 
  group_by(vacc_location_health_board_name, iso_week) %>%
  summarise(record_count = n()) 

### CREATE TABLE AND SUMMARY OF PATIENTS WITH MISSING (NA) OR INVALID CHI NUMBER
rsv_chi_inv  <- rsv_vacc %>% filter(CHIcheck!="Valid CHI") %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>%
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = paste0("01. Missing/Invalid CHI number - ",CHIcheck)) 

rsv_chi_invSumm <- rsv_chi_inv %>%
  group_by(vacc_location_health_board_name, vacc_location_name,
           vacc_data_source,CHIcheck,Date_Administered) %>%
  summarise(record_count = n())

rsv_chi_inv <- rsv_chi_inv %>% select(-Date_Administered,-CHIcheck)

rsv_vacc <- rsv_vacc %>% select(-CHIcheck)

### CREATE TABLE OF RECORDS & SUMMARY OF 2 OR MORE DOSE 1 VACCINATIONS
### FOR PEOPLE AGED >55, OR <55 AND INTERVAL < 28 WEEKS
rsv_dose1 <- rsv_vacc %>% filter(vacc_dose_number == "1")

rsv_dose1$`rsv_interval (days)` [is.na(rsv_dose1$`rsv_interval (days)`)] <- 
  "0.0" # numbers required for 'max' function below

rsv_dose1x2IDs <- rsv_dose1 %>% group_by(patient_derived_upi_number) %>% 
  summarise(count_by_patient_derived_upi_number = n(),
            min_age=as.integer(min(age_at_vacc)),
            interval=as.integer(max(`rsv_interval (days)`)) ) %>%
  na.omit(count_by_patient_derived_upi_number) %>%
  filter(count_by_patient_derived_upi_number > 1 &
           (min_age>55 | interval<195) )

rsv_dose1$`rsv_interval (days)` [rsv_dose1$`rsv_interval (days)`=="0.0"] <- 
  NA # convert 0.0s back to nulls

rsv_dose1x2 <- rsv_dose1 %>%
  filter(patient_derived_upi_number %in% rsv_dose1x2IDs$patient_derived_upi_number)

rsv_dose1x2sort <- rsv_dose1x2 %>%
  select(vacc_event_created_at,patient_derived_upi_number) %>%
  group_by(patient_derived_upi_number) %>%
  summarise(vacc_event_created_at=max(vacc_event_created_at)) %>%
  rename(latest_date = vacc_event_created_at)

rsv_dose1x2 <- rsv_dose1x2 %>%
  left_join(rsv_dose1x2sort,by="patient_derived_upi_number") %>%
  arrange(desc(latest_date)) %>%
  mutate(sort_date = latest_date) %>%
  mutate(QueryName = "02. RSV Two or more dose 1") %>% 
  filter(sort_date >= reporting_start_date) %>% 
  select(-latest_date)

rm(rsv_dose1x2IDs,rsv_dose1x2sort)

if (nrow(rsv_dose1x2)>0) {
  # Create flag for duplicates occuring on same day
  rsv_dose1x2_samedateIDs <- rsv_dose1x2 %>% 
    group_by(patient_derived_upi_number,vacc_occurence_time) %>% 
    summarise(nn = n()) %>%
    filter(nn > 1)
  
  rsv_dose1x2 <- rsv_dose1x2 %>% 
    left_join(rsv_dose1x2_samedateIDs,join_by(patient_derived_upi_number,vacc_occurence_time))
  
  rsv_dose1x2$dups_same_day_flag <- 0
  rsv_dose1x2$dups_same_day_flag [rsv_dose1x2$nn>0] <- 1
  ###
  
  # Create flag for VMT-only dose 1 dups
  rsv_dose1x2_VMT <- rsv_dose1x2 %>% filter(vacc_data_source == "TURAS") %>% 
    arrange(desc(vacc_occurence_time))
  
  rsv_dose1x2_VMT_ids <- rsv_dose1x2_VMT %>% 
    group_by(patient_derived_upi_number) %>% 
    summarise(nn2=n()) %>% ungroup() %>% 
    filter(nn2>1)
  
  rsv_dose1x2 <- rsv_dose1x2 %>% 
    left_join(rsv_dose1x2_VMT_ids,by="patient_derived_upi_number")
  
  rsv_dose1x2$VMT_only_dups_flag <- 0
  rsv_dose1x2$VMT_only_dups_flag [rsv_dose1x2$nn2>0] <- 1
  ###
  
  rsv_dose1x2 <- rsv_dose1x2 %>% select(-nn,-nn2)
  
  rm(rsv_dose1x2_samedateIDs,rsv_dose1x2_VMT,rsv_dose1x2_VMT_ids)
}

rsv_dose1x2Summ <- rsv_dose1x2 %>%
  group_by(vacc_location_health_board_name, vacc_data_source, Date_Administered) %>%
  summarise(record_count = n())

rsv_dose1x2 <- rsv_dose1x2 %>% select(-Date_Administered)

### CREATE SUMMARY TABLE OF BOOSTER DOSES
rsv_booster <- rsv_vacc %>% 
  filter(vacc_event_created_at >= reporting_start_date &
           vacc_booster=="TRUE") %>% 
  mutate(sort_date = vacc_event_created_at) %>% 
  mutate(QueryName = "23. RSV Booster Recorded") %>% 
  select(-Date_Administered)

rsv_boosterSumm <- rsv_booster %>% 
  group_by(vacc_location_health_board_name,vacc_location_name,
           vacc_dose_number,vacc_booster,vacc_occurence_time) %>%
  summarise(record_count = n()) 

### CREATE TABLE OF RECORDS & SUMMARY OF VACC GIVEN OUTWITH AGE GUIDELINES
### THAT ARE NOT IN AGE-RELATED ELIGIBILITY COHORT

# patients vaccinated 2024-25 and not aged 74 on 31/07/2024 or aged 75-79 on 01/08/2024
rsv_non_cohort2425 <- rsv_vacc %>%
  filter(vacc_phase=="Aug24_Jul25" &
           !between(patient_date_of_birth,as.Date("1944-08-02"),as.Date("1950-07-31")))

# patients vaccinated 2025-26 and not aged 74 on 31/07/2025 or eligible in previous year
rsv_non_cohort2526 <- rsv_vacc %>%
  filter(vacc_phase=="Aug25_Jul26" &
           !between(patient_date_of_birth,as.Date("1944-08-02"),as.Date("1951-07-31")))

rsv_non_cohort <- rbind(rsv_non_cohort2425,rsv_non_cohort2526) %>% 
  left_join(rsv_cohort, by=(c("source_system_patient_id","vacc_phase"="cohort_phase"))) %>% 
  filter(vacc_event_created_at >= reporting_start_date &
           is.na(cohort)) %>%  
  mutate(sort_date = vacc_event_created_at) %>%
  mutate(QueryName = "24. RSV Non-cohort vacc age>55")

rsv_non_cohortSumm <- rsv_non_cohort %>% 
  group_by(vacc_location_health_board_name, age_at_vacc, patient_sex) %>%
  summarise(record_count = n())

rsv_non_cohort <- rsv_non_cohort %>%
  filter(age_at_vacc>55 | patient_sex=="MALE") %>% 
  select(-c(Date_Administered,cohort:cohort_target_diseases))

rm(rsv_non_cohort2425,rsv_non_cohort2526)

### CREATE & SAVE OUT RSV VACC SUMMARY REPORT
#############################################################################################

#Collates pivot tables into a single report
rsv_summary_report <-
  list("System Summary" = rsv_system_summary,
       "CHI Check" = rsv_chi_check,
       "Vacc Product Type" = rsv_vacc_prodSumm,
       "Vacc Cohorts" = rsv_vacc_cohorts,
       "Vacc Dose Numbers" = rsv_dose_number,
       "Vacc Boosters" = rsv_boosterSumm,
       "Vacc Dates" = rsv_vacc_date,
       "Vacc ISO Weeks" = rsv_vacc_iso,
       "Missing or Invalid CHIs" = rsv_chi_invSumm,
       "2 or More First Doses" = rsv_dose1x2Summ,
       "Non-Cohort by age, sex" = rsv_non_cohortSumm)

rm(rsv_system_summary,rsv_chi_check,rsv_vacc_prodSumm,rsv_vacc_cohorts,
   rsv_dose_number,rsv_boosterSumm,rsv_vacc_date,rsv_vacc_iso,rsv_chi_invSumm,
   rsv_dose1x2Summ,rsv_dose1,rsv_cohort,rsv_non_cohortSumm)

#Saves out collated tables into an excel file
if (answer==1) {
  openxlsx::write.xlsx(rsv_summary_report,
                       paste("Outputs/DQ Summary Reports/RSV_Vacc_DQ_4wk_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
                       asTable = TRUE,
                       colWidths = "auto")
} else {
  openxlsx::write.xlsx(rsv_summary_report,
                       paste("Outputs/DQ Summary Reports/RSV_Vacc_DQ_Full_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
                       asTable = TRUE,
                       colWidths = "auto")
}

rm(rsv_summary_report)

### COLLATE RSV QUERIES FOR VACCINATIONS DQ REPORT
#############################################################################################

multi_vacc <- multi_vacc %>% rbind(rsv_chi_inv,rsv_dose1x2)
rsv_vacc <- rbind(rsv_booster,rsv_non_cohort)

rm(rsv_chi_inv,rsv_dose1x2,rsv_booster,rsv_non_cohort)

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

### CREATE UNIQUE DQ ID FOR ALL RECORDS IN multi_vacc TABLE
rsv_vacc$DQ_ID <- paste(substr(rsv_vacc$QueryName,1,2),
                        rsv_vacc$vacc_source_system_event_id,
                        substr(rsv_vacc$vacc_event_created_at,1,10),sep = ".")

### UPDATE DQ ID FOR DUPLICATE DOSE 1 QUERY
d1ID <- rsv_vacc %>% filter(grepl("02.",rsv_vacc$QueryName)) %>% 
  group_by(patient_derived_upi_number) %>% 
  summarise(DQ_ID = first(DQ_ID)) %>% 
  rename("maind1DQID" = DQ_ID)

rsv_vacc <- rsv_vacc %>% left_join(d1ID,by="patient_derived_upi_number")

rsv_vacc$DQ_ID <- ifelse(grepl("02.",rsv_vacc$QueryName),rsv_vacc$maind1DQID,rsv_vacc$DQ_ID)

### TIDY UP rsv_vacc DATA TABLE
rsv_vacc <- rsv_vacc %>% arrange(desc(sort_date),DQ_ID,desc(vacc_event_created_at)) %>% 
  select(QueryName,DQ_ID,patient_derived_chi_number,patient_derived_upi_number,
         everything(),-maind1DQID)

rm(d1ID)

### CREATE A DATA TABLE OF ALL DATA QUERY RECORDS FOR HB REPORTS
all_queries <- rbind(multi_vacc,covid_vacc,flu_vacc,hz_vacc,
                     pneum_vacc,rsv_vacc) %>% 
  arrange(desc(sort_date),DQ_ID,desc(vacc_event_created_at)) %>% 
  select(-sort_date,-vacc_phase)

### FINALISE DATA TABLES FOR HB REPORTS
multi_vacc <- multi_vacc %>% select(-source_system_patient_id,-sort_date,-vacc_phase)
covid_vacc <- covid_vacc %>% select(-source_system_patient_id,-sort_date,-vacc_phase)
flu_vacc <- flu_vacc %>% select(-source_system_patient_id,-sort_date,-vacc_phase)
hz_vacc <- hz_vacc %>% select(-source_system_patient_id,-sort_date,-vacc_phase)
pneum_vacc <- pneum_vacc %>% select(-source_system_patient_id,-sort_date,-vacc_phase)
rsv_vacc <- rsv_vacc %>% select(-sort_date,-vacc_phase)

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
  df_rsv <- rsv_vacc %>% filter(vacc_location_health_board_name == hb_name[i])
  df_all <- all_queries %>% filter(vacc_location_health_board_name == hb_name[i])
  
  df_summ <- df_all %>% group_by(QueryName) %>% 
    summarise(patient_count = n_distinct(patient_derived_upi_number)) 
  
  if (n_distinct(df_all$source_system_patient_id
                 [grepl("01. ",df_all$QueryName)])>0) {
    df_summ[1,1] = "01. Missing/Invalid CHI number"
    df_summ[1,2] = n_distinct(df_all$source_system_patient_id [grepl("01. ",df_all$QueryName)])
    }
  if (grepl("01.",df_summ[2,1])) {df_summ <- df_summ[-2,]}
  
  df_all <- df_all %>% select(-source_system_patient_id)
  
  if (nrow(df_summ)>0) {
  
      HBReport <-
        list("Summary" = df_summ,
             "Multi Vacc - Q01-05" = df_multi,
             "Covid-19 Vacc - Q06-11,25-27" = df_cov,
             "Flu Vacc - Q12-14,28" = df_flu,
             "Herpes Zoster Vacc - Q15-20" =  df_hz,
             "Pneumococcal Vacc - Q21-22" = df_pneum,
             "RSV Vacc - Q23-24" = df_rsv,
             "All query data - Q01-28" = df_all )

    HBReportWB <- buildWorkbook(HBReport, asTable = TRUE)
    setColWidths(HBReportWB,sheet = 1,cols = 1,widths = "auto")
    setColWidths(HBReportWB,sheet = 2,cols = 1,widths = "auto")
    setColWidths(HBReportWB,sheet = 3,cols = 1,widths = "auto")
    setColWidths(HBReportWB,sheet = 4,cols = 1,widths = "auto")
    setColWidths(HBReportWB,sheet = 5,cols = 1,widths = "auto")
    setColWidths(HBReportWB,sheet = 6,cols = 1,widths = "auto")
    setColWidths(HBReportWB,sheet = 7,cols = 1,widths = "auto")
    

    insertImage(HBReportWB,sheet = 1,file = "Scripts/Vacc HB DQ Summary.jpg",
                  width = 10,height = 11,startRow = (nrow(df_summ) + 3),startCol = 1, units = "in",dpi = 300)

    saveWorkbook(HBReportWB,paste("Outputs/DQ HB Reports/",hb_cypher[i],"_Vacc_DQ_Report_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep=""))
    
    rm(HBReport,HBReportWB)
  } else {
    writexl::write_xlsx(df_summ,paste("Outputs/DQ HB Reports/",hb_cypher[i],"_NULL_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""))
  }
}
  

########################################################################
################################ SECTION G #############################
##########################  QUERY COUNT SUMMARY ########################
########################################################################

# read in latest query_count output and save as backup
temp_query_count <- readxl::read_excel("Outputs/DQ Report Query count.xlsx")

write.xlsx(temp_query_count,
           paste("Outputs/Archive/DQ Report Query count_backup.xlsx"),
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
           paste("Outputs/DQ Report Query count.xlsx"),
           asTable = TRUE,
           colWidths = "auto")

rm(df,df2,query_count)

end.time <- Sys.time()
print(paste("End time: ", end.time))

time.taken <- round(end.time - start.time,2)

time.taken

sink(NULL)

############# END OF SCRIPT ###########################



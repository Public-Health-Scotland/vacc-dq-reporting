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

reporting_start_date = as.Date("2021-01-01")
answer <- 0

### ESTABLISH ODBC CONNECTION TO DVPROD MANUALLY
conn <- odbc::dbConnect(odbc::odbc(),
                        dsn = "DVPROD", 
                        uid = paste(Sys.info()['user']),
                        pwd = .rs.askForPassword("password"))

# ### LOOK UP RSV COHORTS IN VDL
# cohort_summ <- conn %>%
#   tbl(dbplyr::in_schema("vaccination", "vaccination_patient_cohort_analysis")) %>% 
#   select(cohort,
#          cohort_description,
#          cohort_extract_time,
#          cohort_reporting_label,
#          cohort_target_diseases,
#          cohort_phase) %>%
#   group_by(cohort,
#            cohort_description,
#            cohort_extract_time,
#            cohort_reporting_label,
#            cohort_target_diseases,
#            cohort_phase) %>%
#   summarise(number_of_people = n()) %>% collect()
# 
# table(cohort_summ$cohort_target_diseases)
# 
### EXTRACT PATIENT DATA VARIABLES IN vaccination_patient_analysis VIEW
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

### CHECK VACC TARGET DISEASES IN VDL FOR EXACT RSV NAME
Vaxeventtot <- conn %>%
  tbl(dbplyr::in_schema("vaccination", "vaccination_event_analysis")) %>%
  # select(vacc_type_target_disease,vacc_status,vacc_data_source,vacc_event_created_at) %>%
  # group_by(vacc_status,vacc_type_target_disease,vacc_data_source,vacc_event_created_at) %>%
  select(vacc_type_target_disease,vacc_status,vacc_data_source) %>%
  group_by(vacc_status,vacc_type_target_disease,vacc_data_source) %>%
  summarise(number_of_vacc_events = n()) %>% collect() %>%
  arrange(vacc_status,vacc_type_target_disease,vacc_data_source)

# unique(Vaxeventtot$vacc_type_target_disease)
# 
### ADD vacc_type_target_disease FROM TABLE PRODUCED ABOVE TO CREATE SUMMARY
rsv_system_summary <- Vaxeventtot %>%
  filter(vacc_type_target_disease == "Respiratory syncytial virus infection (disorder)") %>% # &
           # vacc_event_created_at>=reporting_start_date) %>% 
  group_by(vacc_status,vacc_type_target_disease,vacc_data_source) %>% 
  summarise(record_count = sum(number_of_vacc_events))

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
                                      # cohort_source,
                                      # cohort_extract_time,
                                      cohort_reporting_label,
                                      cohort_description,
                                      cohort_target_diseases,
                                      patient_cohort_created_at,
                                      patient_cohort_updated_at,
                                      # patient_cohort_removal_datetime,
                                      # patient_cohort_removal_status,
                                      cohort_phase
                          from vaccination.vaccination_patient_cohort_analysis
where cohort_target_diseases like '%55735004 - Respiratory syncytial virus infection (disorder)%' ")

table(rsv_cohort$cohort, useNA = "ifany")
# table(rsv_cohort$cohort_source, useNA = "ifany")
# table(rsv_cohort$cohort_extract_time, useNA = "ifany")
table(rsv_cohort$cohort_reporting_label, useNA = "ifany")
table(rsv_cohort$cohort_description, useNA = "ifany")
table(rsv_cohort$cohort_target_diseases, useNA = "ifany")
table(rsv_cohort$patient_cohort_created_at, useNA = "ifany")
table(rsv_cohort$patient_cohort_updated_at, useNA = "ifany")
# table(rsv_cohort$patient_cohort_removal_datetime) # - 0
# table(rsv_cohort$patient_cohort_removal_status) # - 0
table(rsv_cohort$cohort_phase, useNA = "ifany")

### CREATE NEW COLUMN FOR DAYS BETWEEN VACCINATION AND RECORD CREATION
rsv_vacc$vacc_record_date <- as.Date(substr(rsv_vacc$vacc_record_created_at,1,10))
rsv_vacc$days_between_vacc_and_recording <- as.double(difftime(rsv_vacc$vacc_record_date,rsv_vacc$vacc_occurence_time,units="days"))

### CREATE BLANK PLACEHOLDER COLUMNS
rsv_vacc$dups_same_day_flag <- NA
rsv_vacc$VMT_only_dups_flag <- NA

# ### CREATE AN INTERVAL FOR BOOSTER VACCINES PER PATIENT
# rsv_vacc <- rsv_vacc %>%
#   arrange(patient_derived_encrypted_upi,vacc_occurence_time) %>% 
#   group_by(patient_derived_encrypted_upi) %>%
#   mutate(dose_number = row_number()) %>% mutate(doses = n()) %>% ungroup() %>% 
#   mutate(prev_vacc_date=lag(vacc_occurence_time)) %>%
#   mutate("cov_booster_interval (days)"=as.integer(difftime(vacc_occurence_time,prev_vacc_date,units = "days")))
# 
# rsv_vacc$"cov_booster_interval (days)" <- ifelse(rsv_vacc$dose_number=="1",NA,rsv_vacc$"cov_booster_interval (days)")
# rsv_vacc$"cov_booster_interval (days)" [is.na(rsv_vacc$patient_derived_encrypted_upi)] <- NA
# rsv_vacc$"cov_booster_interval (days)" [rsv_vacc$vacc_booster=="FALSE"] <- NA

# ### CREATE BLANK PLACEHOLDER COLUMNS
# rsv_vacc$"flu_interval (days)" <- NA
# rsv_vacc$"hz_interval (days)" <- NA
# rsv_vacc$"pneum_vacc_interval (weeks)" <- NA
rsv_vacc$sort_date <- NA

### CREATE RSV VACC DATAFRAME BY JOINING RSV VACC EVENT RECORDS WITH PATIENT RECORDS
rsv_vacc <- rsv_vacc %>%
  inner_join(Vaxpatientraw, by=("source_system_patient_id")) %>%
  arrange(desc(vacc_event_created_at)) %>% # sort data by date amended
  select(-patient_derived_encrypted_upi,-vacc_record_date) %>% # remove temp items from Interval calculation
  # select(-patient_derived_encrypted_upi,-vacc_record_date,-dose_number,-doses,-prev_vacc_date) %>% # remove temp items from Interval calculation
  mutate(Date_Administered = substr(vacc_occurence_time, 1, 10)) %>% # create new vacc date data item for date only (no time)
  mutate(CHIcheck = phsmethods::chi_check(patient_derived_chi_number)) # create CHI check data item

rsv_vacc$vacc_phase <- NA
rsv_vacc$vacc_phase [between(rsv_vacc$vacc_occurence_time,
                              as.Date("2024-08-01"),as.Date("2025-07-31"))] <-
  "Aug24_Jul25"


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
  left_join(rsv_cohort,by="source_system_patient_id") %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name, cohort, cohort_phase) %>%
  summarise(record_count = n()) 

### CREATE SUMMARY TABLE OF DOSE NUMBERS
rsv_dose_number <- rsv_vacc %>% 
  filter(vacc_event_created_at >= reporting_start_date) %>% 
  group_by(vacc_location_health_board_name,vacc_dose_number,vacc_booster) %>%
  summarise(record_count = n()) 

### CREATE SUMMARY TABLE OF BOOSTER DOSES
rsv_booster <- rsv_vacc %>% 
  filter(vacc_event_created_at >= reporting_start_date &
           vacc_booster=="TRUE")

rsv_boosterSumm <- rsv_booster %>% 
  group_by(vacc_location_health_board_name,vacc_location_name,
           vacc_dose_number,vacc_booster,vacc_occurence_time) %>%
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
rsv_dose1 <- rsv_vacc %>% filter(vacc_dose_number == "1")

rsv_dose1x2IDs <- rsv_dose1 %>% group_by(patient_derived_upi_number) %>% 
  summarise(count_by_patient_derived_upi_number = n()) %>%
  na.omit(count_by_patient_derived_upi_number) %>%
  filter(count_by_patient_derived_upi_number > 1)

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

### CREATE SUMMARY TABLE OF AGE & SEX OF NON-COHORT VACCINATIONS
rsv_age_sex <- rsv_vacc %>%
  left_join(rsv_cohort,by="source_system_patient_id") %>% 
  filter(vacc_event_created_at >= reporting_start_date &
           is.na(cohort))

rsv_age_sexSumm <- rsv_age_sex %>% 
  group_by(vacc_location_health_board_name, age_at_vacc, patient_sex) %>%
  summarise(record_count = n())

### EXTRACT COMPLETED RSV VACC RECORDS IN vaccination_event_analysis VIEW
rsv_vacc_mat <- odbc::dbGetQuery(conn, "select *
                        from vaccination.vaccination_event_analysis
                        WHERE vacc_status='completed' 
                          AND vacc_type_target_disease='Respiratory syncytial virus infection (disorder)'
                          AND vacc_location_name='V6800 - FVRH W&C Maternity' ")

# ### CREATE TABLE OF RECORDS & SUMMARY OF VACC GIVEN OUTWITH AGE GUIDELINES
# ### THAT ARE NOT IN AGE-RELATED ELIGIBILITY COHORT
# 
# # patients not turning 75 from 1st Aug 2024 to 31st July 2025 and vaccinated 2024-25
# rsv_ageDQ1 <- rsv_vacc %>% 
#   filter(between(vacc_occurence_time,as.Date("2024-08-01"),as.Date("2025-07-31")) &
#            !between(patient_date_of_birth,as.Date("1949-08-01"),as.Date("1950-07-31")))         
# 
# # patients not aged 75-79 on 1st Aug 2024 and vaccinated 2024-25
# rsv_ageDQ2 <- rsv_vacc %>%
#   filter(between(vacc_occurence_time,as.Date("2024-08-01"),as.Date("2025-07-31")) &
#            !between(patient_date_of_birth,as.Date("1944-08-02"),as.Date("1949-08-01")))
# 
# # patients not turning 80 from 1st Aug 2024 to 31st July 2025 and vaccinated 2024-25
# rsv_ageDQ3 <- rsv_vacc %>% 
#   filter(between(vacc_occurence_time,as.Date("2024-08-01"),as.Date("2025-07-31")) &
#            !between(patient_date_of_birth,as.Date("1944-08-01"),as.Date("1945-07-31")))         
# 
# # combine 3 ageDQ dfs and remove anyone in eligibility cohorts
# rsv_ageDQ <- rbind(rsv_ageDQ1,rsv_ageDQ2,rsvz_ageDQ3) %>%
#   left_join(cohort, by=(c("source_system_patient_id","vacc_phase"="cohort_phase"))) %>% 
#   filter(vacc_event_created_at >= reporting_start_date &
#            is.na(cohort) &
#            (!between(age_at_vacc,15,55) | patient_sex=="MALE")) %>% 
#   mutate(sort_date = vacc_event_created_at) %>% 
#   mutate(QueryName = "03. RSV Vacc given outwith cohorts") 
# 
# rm(rsv_ageDQ1,rsv_ageDQ2,rsv_ageDQ3)
# 
# rsv_ageDQSumm <- rsv_ageDQ %>%
#   group_by(vacc_location_health_board_name, vacc_location_name, vacc_data_source, Date_Administered, age_at_vacc) %>%
#   summarise(record_count = n())
# 
# rsv_ageDQ <- rsv_ageDQ %>% select(-c(Date_Administered:cohort_target_diseases))

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
       "Patient Age, Sex - Non-Cohort" = rsv_age_sex)

# rm()

#Saves out collated tables into an excel file
if (answer==1) {
  openxlsx::write.xlsx(rsv_summary_report,
             paste("RSV_Vacc_DQ_4wk_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
             asTable = TRUE,
             colWidths = "auto")
} else {
  openxlsx::write.xlsx(rsv_summary_report,
             paste("RSV_Vacc_DQ_Full_Summary_",format(as.Date(Sys.Date()),"%Y-%m-%d"),".xlsx",sep = ""),
             asTable = TRUE,
             colWidths = "auto")
}

rm(rsv_summary_report)


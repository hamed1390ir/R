# Initialization ----------------------------------------------------------

# clear the console
cat("\014")
# clear all of the objects in the workspace
rm(list = ls())
# set the beginning time
time1 <- Sys.time()
# set the working directory
setwd('Q:/shared/Hamed/Customer_360_Report/R_360/Hamed')
# loading necessary packages

packages <- 
        c(
                "wordcloud", "RColorBrewer", "tidytext", 
                "readr", "lubridate", "RODBC", "bizdays",
                "timeDate", "outliers", "leaflet", "leaflet.extras", "geojson", 
                "geojsonio", "purrr", "Hmisc", "readxl", "dplyr", "stringr"
        )
check.packages <- function(pkg){
        new.pkg <- pkg[!(pkg %in% installed.packages()[, "Package"])]
        if (length(new.pkg)) 
                install.packages(new.pkg, dependencies = TRUE)
        sapply(pkg, require, character.only = TRUE)
}
check.packages(packages)
rm("packages", "check.packages")

# calander ----------------------------------------------------------------

Beg.Year <- 2010
End.Year <- year(now())

# find NSD holidays between Beg.Year and current year
holidays<-
        date(
                c(
                        USNewYearsDay(Beg.Year:End.Year),
                        USMemorialDay(Beg.Year:End.Year),
                        USIndependenceDay(Beg.Year:End.Year),
                        USLaborDay(Beg.Year:End.Year),
                        USThanksgivingDay(Beg.Year:End.Year),
                        USChristmasDay(Beg.Year:End.Year)
                )
        )

# create the NSD calander
create.calendar(
        name = "NSD_Calander", 
        holidays = holidays, 
        weekdays = c("sunday", "saturday"),
        adjust.from = adjust.next, 
        adjust.to = adjust.previous
)

rm("holidays", "Beg.Year", "End.Year")

# Import_Orders and_drivers_Table -----------------------------------------------------

channel <- odbcConnect("nsdserv26")
imported.RockHopper <- 
        sqlQuery(
                channel,
                "SELECT job_date, nsd_num, agent, agent_state, account_type, svc_type, po_num, pro_num, return_designation, weight, items, ltl_pick_up, ltl_pod_datetime, pend_to_dock, first_console_update_date, scheduled_by_agent, sched_time1, pod_datetime, pod_name, mileage, completed_datetime, status, driver1, origin_zip, agent_zip, dest_zip FROM orders;",
                as.is = TRUE
        )
RH_Columns <- sqlColumns(channel, "orders", errors = FALSE, as.is = TRUE,
                                 special = FALSE, catalog = NULL, schema = NULL,
                                 literal = FALSE)
imported.RockHopper <- 
        imported.RockHopper %>% 
        dplyr::rename(
                Job.Date = job_date,
                NSD.Num = nsd_num,
                Agent = agent,
                State = agent_state,
                Account.Type = account_type,
                Service.Type = svc_type,
                PO.Num = po_num,
                Pro.Num = pro_num,
                Return.Designation = return_designation,
                Weight = weight,
                Piece.Count = items,
                LTL.Pick.Up = ltl_pick_up,
                LTL.POD = ltl_pod_datetime,
                Pend.to.Dock = pend_to_dock,
                First.Update = first_console_update_date,
                Scheduled.by.Agent = scheduled_by_agent,
                Schedule.Time = sched_time1,
                POD.Date = pod_datetime,
                POD.Name = pod_name,
                Mileage = mileage,
                Completed.Date = completed_datetime,
                Status = status,
                Driver1 = driver1,
                Origin.Zip = origin_zip, 
                Agent.Zip = agent_zip,
                Dest.Zip = dest_zip
        )

agents_pay <-
        sqlQuery(channel, "SELECT nsd_num, agent_pay FROM totals;", as.is = TRUE) %>% 
        mutate(agent_pay = as.numeric(agent_pay)) %>% 
        dplyr::rename(
                NSD.Num = nsd_num, 
                Agent.Pay = agent_pay
        )

drivers <- 
        sqlQuery(
                channel,
                "SELECT code, firstname FROM rockhopper.driver;",
                as.is = TRUE
        ) %>% 
        dplyr::rename(
                Driver1 = code, 
                agent2 = firstname
        )

Accounts <- 
        sqlQuery(
                channel,
                "SELECT id, description FROM rockhopper.actype;",
                as.is = TRUE
        )%>% 
        dplyr::rename(
                Account.Type = id, 
                Account.Name = description
        )

close(channel)
rm("channel")


# TM import ---------------------------------------------------------------

channel <- odbcConnect("TMLooker", uid = "NSDuser", pwd = "myN5DL0Gin")

TM_Columns <- sqlColumns(channel, "dbo.TM_ORDER_ROLLUP", errors = FALSE, as.is = TRUE,
                         special = FALSE, catalog = NULL, schema = NULL,
                         literal = FALSE)

TMOrders <-
        sqlQuery(
                channel,
                "SELECT JOB_DATE, NSD_NBR, CLIENT, AGENT, AGENT_STATE, ACCOUNT_MANAGER, SVC_TYPE, PO_NBR, PRO_NUM, RETURN_DESIGNATION, WEIGHT, TENDERED_PIECES, AGENT_PAY, LTL_PICKUP_DATE, LTL_POD_DATETIME, PEND_TO_DOCK, FIRST_CONSOLE_UPDATE_DATE, SCHEDULED_BY_AGENT, SCHEDULED_BEGIN, POD_DATETIME, POD_NAME, MILEAGE, COMPLETED_DATETIME, STATUS, ORIGIN_ZIP, AGENT_ZIP, DEST_ZIP, HD_GAP, NSD_GAP, ORDERCANCELLED, CANCELLEDREASON, ATTEMPTS, MODE, DISPOSAL_SI, SHIPPER, CLIENT, OSD_REPORTED, OSD_TYPE, RETURN_PICKED_UP FROM dbo.TM_ORDER_ROLLUP;",
                as.is = TRUE
        ) %>%
        select(
                Job.Date = JOB_DATE,
                NSD.Num = NSD_NBR,
                Agent = AGENT,
                State = AGENT_STATE,
                Service.Type = SVC_TYPE,
                Account.Type = ACCOUNT_MANAGER,
                PO.Num = PO_NBR,
                Pro.Num = PRO_NUM,
                Return.Designation = RETURN_DESIGNATION,
                Weight = WEIGHT,
                Piece.Count = TENDERED_PIECES,
                Agent.Pay = AGENT_PAY,
                LTL.Pick.Up = LTL_PICKUP_DATE,
                LTL.POD = LTL_POD_DATETIME,
                Pend.to.Dock = PEND_TO_DOCK,
                First.Update = FIRST_CONSOLE_UPDATE_DATE,
                Scheduled.by.Agent = SCHEDULED_BY_AGENT,
                Schedule.Time = SCHEDULED_BEGIN,
                POD.Date = POD_DATETIME,
                POD.Name = POD_NAME,
                Mileage = MILEAGE,
                Completed.Date = COMPLETED_DATETIME,
                Status = STATUS,
                Origin.Zip = ORIGIN_ZIP,
                Agent.Zip = AGENT_ZIP,
                Dest.Zip = DEST_ZIP,
                HD_Gap = HD_GAP,
                NSD_Gap = NSD_GAP,
                ORDERCANCELLED,
                CANCELLEDREASON,
                ATTEMPTS,
                MODE,
                DISPOSAL_SI,
                SHIPPER,
                CLIENT, 
                OSD_REPORTED, 
                OSD_TYPE,
                RETURN_PICKED_UP
        ) %>%
        mutate(
                Job.Date = Job.Date %>% as.Date,
                NSD.Num = NSD.Num %>% str_trim,
                LTL.POD = 
                        case_when(
                                Return.Designation == "RETURN" ~  POD.Date,
                                T ~ LTL.POD
                        ),
                POD.Date = 
                        case_when(
                                Return.Designation == "RETURN" ~  RETURN_PICKED_UP,
                                T ~ POD.Date
                        )
                
        ) %>% 
        select(-RETURN_PICKED_UP)

close(channel)
rm("channel")

# Import_Settings -----------------------------------------------------------
info_filename <- "info.xlsx"
for (sheet in info_filename %>% excel_sheets) assign(sheet, read_excel(info_filename, sheet))

Agents <-
        Agents %>% 
        merge(LMRegions_State, by = "State") %>% 
        merge(RRegions_State, by = "State") %>% 
        merge(LMRegions_LMC, by = "LM_Region") %>% 
        merge(RRegions_RC, by = "R_Region") %>% 
        select(-State_Name.y) %>% 
        dplyr::rename(State_Name = State_Name.x)

datetime_process <- function(x) {
        x %>% 
                str_trim %>% 
                ifelse(. == "0000-00-00 00:00:00", NA, .) %>% 
                ymd_hms
}

Hold.no.LTL.Variations <- names(table(imported.RockHopper$Pro.Num[grepl("^H", imported.RockHopper$Pro.Num)]))



accounts_RH_std <- function(x){
        needed_accounts <- 
                c(
                        "Home Depot Vendor", 
                        "Home Depot Direct", 
                        "Home Depot Stores", 
                        "Purchasing Power", 
                        "Wayfair"
                )
        x[!x %in% needed_accounts] <- "Other"
        x %>% 
                str_replace("^Home Depot Vendor$", "Home Depot") %>% 
                str_replace("^Home Depot Direct$", "Home Depot") %>% 
                str_replace("^Home Depot Stores$", "Home Depot")
}

first_update_path <- 
        if (list.files("Q:/rockhopper_reports") %>% grepl("first_console_update_since_2013", .) %>% sum()) {
                "Q:/rockhopper_reports/"
        } else {
                "Q:/rockhopper_reports/OTHER PEOPLES files/"
        }

first_update <- 
        first_update_path %>% 
        paste0(
                dir(first_update_path) %>% 
                        subset(grepl("first_console_update_since_2013", .)) %>% 
                        sort %>% 
                        last
                ) %>% 
        read.delim(stringsAsFactors = FALSE) %>% 
        select(
                NSD.Num = id,
                first_console_update
        ) %>% 
        mutate(
                NSD.Num = NSD.Num %>% as.numeric,
                first_console_update = first_console_update %>% mdy
        ) %>% 
        filter(!is.na(first_console_update))

RockHopper <-
        imported.RockHopper %>% 
        left_join(first_update, by = "NSD.Num") %>% 
        # Merge Service Types table with orders table
        merge(Service_Types_Table, by = "Service.Type", all.x = TRUE) %>% 
        # Import HD Gap table and merge it with orders table
        merge(HD_Gap_Table, by = "Dest.Zip", all.x = TRUE) %>% 
        # Import NSD Gap table and merge it with orders table
        merge(NSD_Gap_Table %>% select(Dest.Zip, NSD_Gap), by = "Dest.Zip", all.x = TRUE) %>% 
        # Merge Agent.Pay
        merge(agents_pay, by = "NSD.Num",all.x = TRUE) %>% 
        # Some of the orders have UYSN as their Agent. substitute UYSN with the real agent name
        merge(drivers, by = "Driver1",all.x = TRUE) %>% 
        # Merge account names
        merge(Accounts, by = "Account.Type",all.x = TRUE) %>% 
        mutate(
                # Removing spaces from Agent
                Agent = str_trim(Agent),
                # Keep only the accounts that we need. The rest is called other.
                Account.Name = accounts_RH_std(Account.Name),
                # inja
                # Substitute UYSN with the real agent name
                agent2 = as.character(agent2), 
                Agent = ifelse(Agent == "UYSN", agent2, Agent),
                # Date-time columns of orders table
                LTL.Pick.Up = LTL.Pick.Up %>% datetime_process,
                LTL.POD = LTL.POD %>% datetime_process,
                Pend.to.Dock = Pend.to.Dock %>% datetime_process,
                First.Update = First.Update %>% datetime_process() %>% date(),
                First.Update = 
                        case_when(
                                Return.Designation == "Return" ~ first_console_update,
                                T ~ First.Update
                        ),
                Schedule.Time = Schedule.Time %>% datetime_process,
                POD.Date = POD.Date %>% datetime_process,
                Completed.Date = Completed.Date %>% datetime_process,
                # change the Job.Date column's format to ymd
                Job.Date = ymd(Job.Date),
                # change the NSD.Num column's format to character
                NSD.Num = as.character(NSD.Num),
                # put NA in the cells containing "0000-00-00" as their value and change its format to ymd
                Scheduled.by.Agent = ymd(ifelse(Scheduled.by.Agent == "0000-00-00", NA, Scheduled.by.Agent)),
                # Replace empty POD.Names with NA
                POD.Name = ifelse(POD.Name == "", NA, POD.Name),
                # create Exception column
                Exception = ifelse(
                        # POD.Names NA <-NA
                        is.na(POD.Name), NA, 
                        # POD.Names starting with **## <- ***
                        ifelse(grepl("^\\*\\*[[:digit:]][[:digit:]]$",POD.Name), TRUE, 
                               FALSE
                        )
                ),
                # destination zip codes should be a 5 digit number
                Dest.Zip = ifelse(Dest.Zip == 0, NA, sprintf("%05d", as.numeric(Dest.Zip))),
                # Disposals
                Disposal = ifelse(Return.Designation=="Return" & Pro.Num %in% Hold.no.LTL.Variations, "Disposal", NA),
                NSD.System = "RockHopper"
        ) %>% 
        filter(
                # remove the orders that have a completed_date before Jan 01, 2013
                Completed.Date >= "2013-01-01 00:00:00" | is.na(Completed.Date),
                # Exclude orders with service_type=LTL, ending with A, and pod-names starting with **41
                Service.Type!="LTL",
                !grepl("A$",Service.Type),
                !grepl("^[*][*]41",POD.Name)
        ) %>%
        select(-Driver1, -agent2, -Account.Type, -Service.Type, -first_console_update) %>% 
        arrange(Job.Date)

rm("Hold.no.LTL.Variations", "drivers")

# TMdata ----------------------------------------------------------------

SERVICE_TYPES <- c("BASIC DELIVERY", "BASICPLUS DELIVERY", "THRESHOLD DELIVERY", "THRESHOLD RETURN", "WHITEGLOVE DELIVERY", "WHITEGLOVE RETURN", "LINEHAUL RETURN")

std_svc_typ <- function(x){
        x %>% 
                str_replace("^BASIC DELIVERY$", "Basic Delivery") %>% 
                str_replace("^BASICPLUS DELIVERY$", "Basic Plus Delivery") %>% 
                str_replace("^THRESHOLD DELIVERY$", "Threshold Delivery") %>% 
                str_replace("^THRESHOLD RETURN$", "Threshold Return") %>% 
                str_replace("^WHITEGLOVE DELIVERY$", "White Glove Delivery") %>% 
                str_replace("^WHITEGLOVE RETURN$", "White Glove Return") %>% 
                str_replace("^LINEHAUL RETURN$", "Linehaul Return")
}

accounts_TM_std <- function(x){
        needed_accounts <- 
                c(
                        "J.C. PENNEY CORPORATION, INC.", 
                        "PURCHASING POWER, LLC", 
                        "AMAZON", 
                        "WAYFAIR, LLC",
                        "THE HOME DEPOT"
                )
        x[!x %in% needed_accounts] <- "Other"
        x %>% 
                str_replace("^PURCHASING POWER, LLC$", "Purchasing Power") %>% 
                str_replace("^WAYFAIR, LLC$", "Wayfair") %>% 
                str_replace("^J.C. PENNEY CORPORATION, INC.$", "J.C. Penny") %>% 
                str_replace("^AMAZON$", "Amazon") %>% 
                str_replace("^THE HOME DEPOT$", "Home Depot")
}



TMdata <- 
        TMOrders %>% 
        mutate(
                # Destination zip codes should be a 5 digit number
                Dest.Zip = Dest.Zip %>% str_trim,
                Dest.Zip = ifelse(!grepl("^\\d{5}", Dest.Zip), NA, str_sub(Dest.Zip, 1, 5)),
                # Making Description like in orders table
                Description = paste(Service.Type, Return.Designation),
                Description = if_else(MODE=="LHR","LINEHAUL RETURN", Description),
                Description = ifelse(Description %in% SERVICE_TYPES, Description, NA),
                Description = std_svc_typ(Description),
                # Removing spaces from Date-time columns of TMdata and transforming characters to ymd_hms
                LTL.Pick.Up = LTL.Pick.Up %>% datetime_process,
                LTL.POD = LTL.POD %>% datetime_process,
                Pend.to.Dock = Pend.to.Dock %>% datetime_process,
                First.Update = First.Update %>% datetime_process,
                Scheduled.by.Agent = Scheduled.by.Agent %>% datetime_process,
                Schedule.Time = Schedule.Time %>% datetime_process,
                POD.Date = POD.Date %>% datetime_process,
                Completed.Date = Completed.Date %>% datetime_process,
                # Making Return.Designation like in orders table
                Return.Designation = Return.Designation %>% tolower %>% capitalize,
                # identify cancellations in POD.Name
                ORDERCANCELLED = if_else(ORDERCANCELLED == "Yes" & Return.Designation == "Delivery", "**20", ORDERCANCELLED),
                ORDERCANCELLED = if_else(ORDERCANCELLED == "Yes" & Return.Designation == "Return", "**19", ORDERCANCELLED),
                POD.Name = if_else(ORDERCANCELLED == "**20", "**20", POD.Name),
                POD.Name = if_else(ORDERCANCELLED == "**19", "**19", POD.Name),
                # Exception column
                Exception = (grepl("^[A-z]", OSD_TYPE) & OSD_TYPE != "NON_OSD") | ORDERCANCELLED == "Yes" | Status == "COMPLETED",
                # Major clients
                Account.Name = CLIENT %>% str_trim %>% accounts_TM_std,
                # inja
                Disposal = ifelse(Return.Designation == "Return" & grepl("^[A-z]", DISPOSAL_SI), "Disposal", NA),
                NSD.System = "TruckMate",
                Pro.Num = if_else(MODE %in% c("DD", "LHR"), "NSD____", "____")
        ) %>% 
        select(names(RockHopper))

orders <- 
        rbind(RockHopper, TMdata) %>% 
        mutate(State = str_trim(State)) %>% 
        filter(State %in% LMRegions_State$State) %>% 
        mutate(
                Weight = as.numeric(Weight),
                Piece.Count = as.numeric(Piece.Count),
                Mileage = as.numeric(Mileage),
                Agent = Agent %>% 
                        toupper %>% 
                        str_trim %>% 
                        str_replace_all("^PHILLY$","EMPIRE") %>% 
                        str_replace_all("^DNAIAD$","CLAUDIO")
        )


# difdat ------------------------------------------------------------------

# This is a function to calculte the difference between first.date and second.date.
# If second.date < first.date or (second.date - first.date) > 120, returns NA
difdat<-function(first.date,second.date){
        
        # bizdays
        result <- bizdays(first.date, second.date, "NSD_Calander")
        # same date but -1 results
        result[is.bizday(first.date, "NSD_Calander") == FALSE & is.bizday(second.date, "NSD_Calander") == FALSE & result == -1] <- 0
        # different dates and negative results
        result[(date(first.date) != date(second.date)) & result < 0] <- NA
        # > 120 days
        result[result > 120]<- NA
        result
}

# Delivery Operations -----------------------------------------------------


Deliveries <-
        orders %>% 
        filter(Return.Designation == "Delivery") %>%
        # Merge Last-Mile states
        merge(LMRegions_State %>% select(LM_Region, State), by = "State", all.x = TRUE) %>%
        # Merge Last-Mile coordinator
        merge(LMRegions_LMC %>% select(LM_Region, LM_ID), by = "LM_Region", all.x = TRUE) %>% 
        dplyr::rename(Region = LM_Region, Coordinator = LM_ID)
        
Delivery.Colnames.before.Metrics <- colnames(Deliveries)

# Delivery Metrics --------------------------------------------------------


Deliveries <-
        Deliveries %>% 
        mutate(
                # Same_Day_Receiving_Compliance
                Delivery.Same.Day.Receiving.compliance = case_when(
                        !is.bizday(LTL.POD,"NSD_Calander") ~ Pend.to.Dock <= (adjust.next(LTL.POD, "NSD_Calander") %>% update(hour=23, minute=59, second=59)),
                        is.bizday(LTL.POD,"NSD_Calander") & (hour(LTL.POD) < 14) ~ Pend.to.Dock <= (LTL.POD %>% update(hour=23,minute=59,second=59)),
                        is.bizday(LTL.POD,"NSD_Calander") & (hour(LTL.POD) >= 14) ~ Pend.to.Dock <= (offset(LTL.POD, 1, "NSD_Calander") %>% update(hour = hour(LTL.POD), minute = minute(LTL.POD), second = second(LTL.POD)))
                ),
                # Receiving_LeadTime
                Delivery.Receiving.LeadTime = difdat(LTL.POD, Pend.to.Dock),
                # Delivery_First_Note_Compliance
                Delivery.First.Note.Compliance = difdat(Pend.to.Dock, First.Update) <= 1,
                # Delivery_First_Note
                Delivery.First.Note = difdat(Pend.to.Dock, First.Update),
                # Delivery_Scheduling_Lead_time
                Delivery.Scheduling.Lead.Time = difdat(Pend.to.Dock, Scheduled.by.Agent),
                # Delivery_Lead_Time_from_Scheduling
                Delivery.Lead.Time.from.Scheduling = ifelse(Exception == TRUE, NA, difdat(Scheduled.by.Agent, POD.Date)),
                # Delivery_Window_Compliance
                Delivery.Window.Compliance = ifelse(Exception == TRUE, NA, as.integer(difftime(POD.Date, Schedule.Time)/3600) %in% 0:4),
                # Delivery_Days_Late
                Delivery.Days.Late = ifelse(Exception == TRUE, NA, difdat(Schedule.Time, POD.Date)),
                # Delivery_Date_Compliance
                Delivery.Date.compliance = ifelse(Exception == TRUE, NA, date(Schedule.Time) == date(POD.Date)),
                # Delivery_D2D_Transit
                Delivery.D2D.Transit = ifelse(Exception == TRUE, NA, difdat(LTL.Pick.Up, POD.Date)),
                # Delivery_LM_Transit
                Delivery.LM.Transit = ifelse(Exception == TRUE, NA, difdat(Pend.to.Dock, POD.Date)),
                # Delivery_HD_Gap_Compliance
                Delivery.HD.Gap.Compliance = ifelse(is.na(HD_Gap), Delivery.LM.Transit <= 3, Delivery.LM.Transit <= HD_Gap),
                # Delivery_NSD_Gap_Compliance
                Delivery.NSD.Gap.Compliance = ifelse(is.na(NSD_Gap), Delivery.LM.Transit <= 3, Delivery.LM.Transit <= NSD_Gap),
                # Delivery_POD_to_Completed
                Delivery.POD.to.Completed = ifelse(Exception == TRUE, NA, difdat(POD.Date, Completed.Date)),
                # Delivery_Same_Day_POD_Compliance
                Delivery.Same.Day.POD.compliance = ifelse(Exception == TRUE, NA, 
                                                          ifelse(Description == "Basic Delivery" & Delivery.POD.to.Completed == 1, FALSE,
                                                                 Delivery.POD.to.Completed %in% c(0,1)))
                 
        ) %>% 
        select(-Disposal)

Delivery.Metrics.Columns <- colnames(Deliveries)[!colnames(Deliveries) %in% Delivery.Colnames.before.Metrics]

# Delivery Counts ---------------------------------------------------------

Deliveries <-
        Deliveries %>%
        mutate(
                Created = Job.Date,
                Cancelled = if_else(grepl("^\\*\\*20$", POD.Name), date(POD.Date), as.Date(NA)),
                LTLPedU = date(LTL.Pick.Up),
                LTLPODed = date(LTL.POD),
                Docked = date(Pend.to.Dock),
                Updated = First.Update,
                Scheduled = Scheduled.by.Agent,
                PODed = if_else(Exception == TRUE, as.Date(NA), date(POD.Date)),
                NotMetSchedule = if_else(Exception == TRUE | (date(POD.Date) == date(Schedule.Time)), as.Date(NA), date(POD.Date)),
                Completed = if_else(Exception == TRUE, as.Date(NA), date(Completed.Date))
        )

Delivery.Counts.Columns <- colnames(Deliveries)[!colnames(Deliveries) %in% union(Delivery.Colnames.before.Metrics, Delivery.Metrics.Columns)]

# Return Operations -----------------------------------------------------

Returns <-
        orders %>% 
        filter(Return.Designation == "Return") %>% 
        # Merge Return states
        merge(RRegions_State %>% select(R_Region, State), by = "State", all.x = TRUE) %>%
        # Merge Return coordinator
        merge(RRegions_RC %>% select(R_Region, RC_ID), by = "R_Region", all.x = TRUE) %>% 
        dplyr::rename(Region = R_Region, Coordinator = RC_ID)

Return.Colnames.before.Metrics <- colnames(Returns)

# Return Metrics --------------------------------------------------------

Returns <-
        Returns %>%
        mutate(
                # Return_POD_to_LTL.Pick.Up
                Return.POD.to.LTL.Pick.UP = ifelse(!grepl("^NSD", Pro.Num) | Exception == TRUE, NA, difdat(POD.Date, LTL.Pick.Up)),
                # Return_Scheduling_Lead_Time
                Return.Scheduling.Lead.Time = ifelse(Description == "Linehaul Return" | Account.Name == "J.C. Penny", NA, difdat(Job.Date, Scheduled.by.Agent)),
                # Return_Pick_Up_Days_Late
                Return.Pick.Up.days.Late = ifelse(Exception == TRUE | Account.Name == "J.C. Penny", NA, difdat(Schedule.Time, POD.Date)),
                # Return_Job.Date_to_POD
                Return.First.Note = ifelse(Description == "Linehaul Return" | Account.Name == "J.C. Penny", NA, difdat(Job.Date, First.Update)),
                # Return_Job.Date_to_POD
                Return.Job.Date.to.POD = ifelse(Exception == TRUE | Description == "Linehaul Return" | Account.Name == "J.C. Penny", NA, difdat(Job.Date, POD.Date)),
                # Return_POD_to_LTL_POD
                Return.LTL.Pick.UP.to.LTL.POD = ifelse(!grepl("^NSD", Pro.Num) | Exception == TRUE, NA, difdat(LTL.Pick.Up, LTL.POD)),
                # Return_Job.Date_to_LTL_POD
                Return.Job.Date.to.LTL.POD = ifelse(!grepl("^NSD", Pro.Num) | Exception == TRUE, NA, difdat(Job.Date, LTL.POD)),
                # Return_Total_Transit_Compliance
                Return.Total.Transit.Compliance = Return.Job.Date.to.LTL.POD <= 14,
                # Return_Dock_Dwell_Compliance
                Return.Dock.Dwell.Compliance = Return.POD.to.LTL.Pick.UP <= 5,
                # Return_Scheduling_Compliance
                Return.Scheduling.Compliance = Return.Scheduling.Lead.Time <= 3,
                # Return_Window_Compliance
                Return.Window.Compliance = ifelse(Description == "Linehaul Return" | Account.Name == "J.C. Penny", NA, as.integer(difftime(POD.Date,Schedule.Time)/3600) %in% 0:4)
        )

Returns.Metrics.Columns <- colnames(Returns)[!colnames(Returns) %in% Return.Colnames.before.Metrics]

# Return Counts ---------------------------------------------------------

Returns <-
        Returns %>% 
        mutate(
                Created = date(Job.Date),
                Cancelled = if_else(grepl("^\\*\\*19$",POD.Name), date(POD.Date), as.Date(NA)),
                Updated = First.Update,
                Scheduled = date(Scheduled.by.Agent),
                PickedUp = if_else(Exception == TRUE, as.Date(NA), date(POD.Date)),
                Exceptions = if_else(Exception == TRUE, date(POD.Date), as.Date(NA)),
                PODed = date(POD.Date),
                Disposed = if_else(Disposal == "Disposal", date(POD.Date), as.Date(NA)),
                LTLPedU = date(LTL.Pick.Up),
                LTLPODed = date(LTL.POD),
                NotMetSchedule = if_else(Exception == TRUE | (date(POD.Date) == date(Schedule.Time)), as.Date(NA), date(POD.Date)),
                Completed = if_else(Exception == TRUE, as.Date(NA), date(Completed.Date))
        )

Returns.Counts.Columns <- colnames(Returns)[!colnames(Returns) %in% union(Return.Colnames.before.Metrics, Returns.Metrics.Columns)]

# Timeline_AG_BEG_END -----------------------------------------------------

AG_BEG_END <- orders %>%
        group_by(Agent) %>%
        dplyr::summarize(
                start = as.Date(min(Job.Date, na.rm = TRUE)),
                end = as.Date(max(Job.Date, na.rm = TRUE))
        ) %>%
        mutate(
                Status = ifelse(Agent %in% Agents$Agent, "Active", "Inactive"),
                start = if_else(start < as.Date("2013-01-01"), as.Date("2013-01-01"), start),
                end = if_else(end > date(today()) | Agent %in% Agents$Agent, date(today()), end)
        ) %>% 
        filter(!is.na(Agent), Agent != "UYSN", start != end, end >= as.Date("2013-01-01"))


Accounts <- c(unique(orders$Account.Name) %>% setdiff("Other") %>% sort, "Other")
Servive_Types <- unique(Returns$Description)

rm("TMdata", "TMOrders", "imported.RockHopper", "orders", "RockHopper")


data(stop_words)

wordcloud_create <- function(path, filename_key_word, column_name) {
        
        allwords <- 
                path %>% 
                paste0(
                        dir(path) %>% 
                                subset(grepl(filename_key_word, .)) %>% 
                                sort %>% 
                                last
                ) %>% 
                read_csv %>% 
                
                select(`OP STATUS COMMENT`) %>% 
                mutate(line = 1:nrow(.)) %>% 
                select(2,1) %>% 
                as_tibble %>% 
                unnest_tokens(word, `OP STATUS COMMENT`)%>%
                anti_join(stop_words)%>%
                count(word, sort = TRUE) %>% 
                filter(!grepl("[[:digit:]]", word), nchar(word) >2, !word %in% c("edi", "code"))
        
        wordcloud(words = allwords$word, freq = allwords$n, min.freq = 5,
                  max.words=100, random.order = F, rot.per=0.35, 
                  colors=brewer.pal(8, "Dark2"))
}
png("Q:/shared/NSDComp/www/wordcloud_picture.png", width=640,height=400)
wordcloud_create("Q:/rockhopper_reports/TMW All Orders Report/RETURNS/", "INTERNAL REPORT - DAILY ACTIVE ORDERS - RETURNS_")
dev.off()






# save.image(file='//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/myEnvironment.rda')
# load('//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/myEnvironment.rda')

saveRDS(Deliveries, "//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/Deliveries.rds")
saveRDS(Returns, "//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/Returns.rds")
saveRDS(Accounts, "//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/Accounts.rds")
saveRDS(Delivery.Counts.Columns, "//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/Delivery.Counts.Columns.rds")
saveRDS(Delivery.Metrics.Columns, "//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/Delivery.Metrics.Columns.rds")
saveRDS(Returns.Counts.Columns, "//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/Returns.Counts.Columns.rds")
saveRDS(Returns.Metrics.Columns, "//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/Return.Metrics.Columns.rds")

# Time_Calculation --------------------------------------------------------

time2<-Sys.time()
print(time2-time1)

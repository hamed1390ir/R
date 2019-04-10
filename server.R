check.packages <- function(pkg){
        new.pkg <- pkg[!(pkg %in% installed.packages()[, "Package"])]
        if (length(new.pkg)) 
                install.packages(new.pkg, dependencies = TRUE)
        sapply(pkg, require, character.only = TRUE)
}

packages <- c("datasets", "zipcode", "rlist", "kableExtra", "knitr", "purrr", "formattable", "tidyr", "tools", "lubridate", "tcltk", "shiny", "readxl", "plotly", "RDCOMClient", "DT", "dplyr", "sendmailR","htmlTable", "shinyalert")
check.packages(packages)

# library("devtools")
# install_github('omegahat/RDCOMClient')

picked_color <- "#e95420"
data("zipcode")

source("//nsdstor1/SHARED/NSDComp/init - 3.R")
options(shiny.maxRequestSize=30*1024^2) 
options(device.ask.default = FALSE)

function(input, output, session) {
        
        
        
        # Imports -----------------------------------------------------------------
        
        Deliveries <- reactive({
                readRDS('//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/Deliveries.rds')
        })
        Returns <- reactive({
                readRDS('//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/Returns.rds')
        })
        
        
        # Functions ---------------------------------------------------------------
        
        filter_data <- function(delivery_or_return, svctyp, counts_or_metrics, item, system, acc, date1, date2) {
                Deliveries <- Deliveries()
                base_col_names <- names(Deliveries)[!names(Deliveries) %in% union(Delivery.Metrics.Columns, Delivery.Counts.Columns)]
                appended_columns <-
                        switch (counts_or_metrics,
                                "counts" = item,
                                "metrics" = 
                                        Metrics %>% 
                                        filter(category == delivery_or_return, caption == item) %>% 
                                        select(column, date_column) %>%
                                        t %>% 
                                        as.vector
                        )
                col_names <- c(appended_columns, base_col_names)
                
                data <- 
                        switch(delivery_or_return,
                               "Delivery" = Deliveries(),
                               "Return" = Returns()
                        ) %>% 
                        {if (length(system) != 2) filter(., NSD.System %in% system) else .} %>% 
                        {if (acc != "ALL") filter(., Account.Name == acc) else .} %>% 
                        {if (!svctyp %in% c("All of the Deliveries", "All of the Returns")) filter(., Description == svctyp) else .} %>% 
                        select(col_names) %>% 
                        {if (counts_or_metrics == "counts") rename(., Date = 1) else rename(., Value = 1, Date = 2)} %>% 
                        mutate(Date = Date %>% as.Date) %>%
                        filter(!is.na(Date), Date >= date1, Date <= date2) %>% 
                        {if (counts_or_metrics == "counts") . else filter(., !is.na(Value))}
        }
        
        # TM ------------------------------------------------------------------
        
        
        Return_TM <- reactive({
                
                TM_report <- input$button_TM_update_data
                
                if (is.null(TM_report)) {
                        return(NULL)
                } else {
                        Return_TM_path <- TM_report$datapath
                } 
                
                Return_TM <- 
                        Return_TM_path %>% 
                        read_csv(
                                col_types = cols(
                                        `ADMIN STATUS CHANGED` = col_datetime(format = "%m/%d/%Y %H:%M:%S %p"), 
                                        `AGENT PHONE` = col_character(), 
                                        `CARRIER PRO` = col_character(), 
                                        `CLIENT ID` = col_character(), CREATED = col_date(format = "%m/%d/%Y"), 
                                        `LAST COMMENT DATE` = col_datetime(format = "%m/%d/%Y %H:%M:%S %p"), 
                                        `LAST STATUS UPDATE` = col_datetime(format = "%m/%d/%Y %H:%M:%S %p"), 
                                        `LTL P/U DATE` = col_datetime(format = "%m/%d/%Y %H:%M:%S %p"), 
                                        `LTL READY` = col_date(format = "%m/%d/%Y"), 
                                        `OP STATUS CHANGED` = col_datetime(format = "%m/%d/%Y %H:%M:%S %p"), 
                                        `PICKUP DATE` = col_date(format = "%m/%d/%Y"), 
                                        SCHEDULED = col_date(format = "%m/%d/%Y"), 
                                        SHIPPER_PHONE = col_character(), 
                                        SHIPPER_ZIP = col_character()
                                )
                        ) %>% 
                        left_join(RRegions_State %>% select(`AGENT STATE` = State, R_Region), by = "AGENT STATE") %>% 
                        left_join(RRegions_RC %>% select(R_Region, RC_Name), by = "R_Region") %>% 
                        arrange(AGENT, CREATED)  
                
                
                
        })
        
        # Region ------------------------------------------------------------------
        
        Return_TM_Region <- reactive({
                
                Return_TM <- Return_TM()
                region <- input$region
                if (is.null(Return_TM) | is.null(region)) return(NULL)
                
                Return_TM_Region <- 
                        if ( !"ALL" %in% region) {
                                Return_TM %>%
                                        filter(R_Region %in% region)
                        } else {
                                Return_TM
                        }
                
        })
        
        # Categorizing --------------------------------------------------------------
        
        Return_TM_Region_Categorised <- reactive({
                Return_TM_Region <- Return_TM_Region()
                if (is.null(Return_TM_Region)) return(NULL)
                
                JCP_removed <- 
                        Return_TM_Region %>% 
                        filter(CLIENT != "J.C. PENNEY CORPORATION. INC.")
                
                if (nrow(JCP_removed) == 0) return(NULL)
                
                WGTH_JCP_removed <- 
                        JCP_removed %>% 
                        filter(`SERVICE LEVEL` %in% c("THRESHOLD", "WHITEGLOVE"))
                
                
                if (nrow(WGTH_JCP_removed) == 0) return(NULL)
                
                Return_TM_Region_Categorised <- 
                        WGTH_JCP_removed %>% 
                        mutate(
                                `SERVICE LEVEL` = 
                                        case_when(
                                                `SERVICE LEVEL` == "THRESHOLD" ~ "T",
                                                `SERVICE LEVEL` == "WHITEGLOVE" ~ "WH",
                                                T ~ `SERVICE LEVEL`
                                        ),
                                MAN = 
                                        case_when(
                                                MAN == "1-MAN" ~ "1",
                                                MAN == "2-MAN" ~ "2",
                                                T ~ MAN
                                        )
                        ) %>% 
                        mutate(SVC = paste0(`SERVICE LEVEL`, MAN, "R")) %>% 
                        mutate(
                                SVC = 
                                        case_when(
                                                MODE == "LHR" ~ "RT1",
                                                T ~ SVC
                                        ),
                                
                                `Dock Date` = NA_character_,
                                age = today() - CREATED
                        ) %>% 
                        mutate(`Notes / Instructions` =
                                       case_when(
                                               MODE == "LHR" ~ "RT1",
                                               `ADMIN STATUS` %in% c("NCF", "RCF") ~ "AM",
                                               `ADMIN STATUS` == "PCF" ~ 
                                                       case_when(
                                                               `OP REASON CODE` == "CANCELLED" ~ "Cancel",
                                                               `OP STATUS` %in% c("LTLAVAIL", "PICKEDUP") ~ "WPU",
                                                               (`OP REASON CODE` %in% c("CANTREACH", "CONTATMPT")) & (grepl("@", `LAST STATUS COMMENT`) | length(gsub("[- .)(+/']|[a-zA-Z]*:?","", `LAST STATUS COMMENT`)) >=10) ~ "Update",
                                                               T ~ "AM"
                                                       ),
                                               `OP STATUS` == "AVAIL-R" ~
                                                       case_when(
                                                               grepl("ESCALAT", toupper(`LAST STATUS COMMENT`)) ~ "Escalation",
                                                               is.na(`OP REASON CODE`) ~
                                                                       case_when(
                                                                               is.na(`LAST STATUS COMMENT`) ~ 
                                                                                       case_when(
                                                                                               age > 14 ~ "2 Weeks",
                                                                                               age <= 14 ~ "Update",
                                                                                               T ~ "?"  
                                                                                       ),
                                                                               grepl("CANCEL|DUPLICATE", toupper(`LAST STATUS COMMENT`)) ~ "Cancel",
                                                                               grepl("^PRINTED$", toupper(`LAST STATUS COMMENT`)) ~ "Update"
                                                                       ),
                                                               `OP REASON CODE` == "STATUPDTE" ~ "Update",
                                                               `OP REASON CODE` == "CONTATMPT" ~ "Update",
                                                               `OP REASON CODE` == "CANTREACH" ~ "Bad #",
                                                               `OP REASON CODE` == "ADDRSCHNG" ~ "Bad #",
                                                               `OP REASON CODE` == "RPLCMNT" ~ "Rplcmnt",
                                                               `OP REASON CODE` == "STORAGE" ~ "Storage",
                                                               `OP REASON CODE` == "ATTEMPT" ~ "Attempt",
                                                               `OP REASON CODE` == "CANCELLED" ~ "Cancel"
                                                       ),
                                               `OP STATUS` == "SHPTACK-T" ~
                                                       case_when(
                                                               grepl("ESCALAT", toupper(`LAST STATUS COMMENT`)) ~ "Escalation",
                                                               `OP REASON CODE` == "STATUPDTE" ~ "Update",
                                                               `OP REASON CODE` == "CONTATMPT" ~ "Update",
                                                               `OP REASON CODE` == "CANTREACH" ~ "Bad #",
                                                               `OP REASON CODE` == "ADDRSCHNG" ~ "Bad #",
                                                               `OP REASON CODE` == "RPLCMNT" ~ "Rplcmnt",
                                                               `OP REASON CODE` == "STORAGE" ~ "Storage",
                                                               `OP REASON CODE` == "ATTEMPT" ~ "Attempt"
                                                       ),
                                               `OP STATUS` == "SCHEDULED" ~
                                                       case_when(
                                                               grepl("ESCALAT", toupper(`LAST STATUS COMMENT`)) ~ "Escalation",
                                                               today() > SCHEDULED ~ "Scheduled",
                                                               today() <= SCHEDULED ~ "OK",
                                                               T ~ "?"
                                                       ),
                                               `OP STATUS` == "OFP" ~
                                                       case_when(
                                                               today() > (as.Date(`OP STATUS CHANGED`) + 2) ~ "Active",
                                                               today() <= (as.Date(`OP STATUS CHANGED`) + 2) ~ "OK",
                                                               T ~ "?"
                                                       ),
                                               `OP STATUS` == "PICKEDUP" ~
                                                       case_when(
                                                               TRANSLATIONS == "WPU" ~ 
                                                                       case_when(
                                                                               MODE == "LM" ~ "AM",
                                                                               MODE == "DD" ~ "WPU",
                                                                               T ~ "?"
                                                                       ),
                                                               T ~ "?"
                                                       ),
                                               `OP STATUS` %in% c("LTLAVAIL", "SHPTACK", "TENDER") ~
                                                       case_when(
                                                               TRANSLATIONS == "WPU" ~ 
                                                                       case_when(
                                                                               MODE == "LM" ~ "AM",
                                                                               MODE == "DD" ~ "WPU",
                                                                               T ~ "?"
                                                                       ),
                                                               T ~ "?"
                                                       ),
                                               `OP STATUS` %in% c("INTRANSIT", "AVAIL", "DOCKED") ~ "Exception",
                                               `OP STATUS` %in% c("ARRDLVTERM", "ARRSHIP", "ARRTERM", "DEPSHIP", "DEPTERM", "LHAPPTCONS", "LHAPPTSHIP", "LHETA") ~ "Unyson",
                                               T ~ NA_character_
                                       )
                        ) %>% 
                        select(
                                `Job Date` = CREATED,
                                ACCT = `CLIENT ID`,
                                `Account Name` = CLIENT,
                                RC = RC_Name,
                                `A.M.` = AM,
                                `NSD #`,
                                `PO#`,
                                AGENT,
                                SVC,
                                `Customer Name` = SHIPPER,
                                City = AGENT_CITY,
                                State = `AGENT STATE`,
                                WGT = WEIGHT,
                                Mile = MILEAGE,
                                STAT = `OP STATUS`,
                                `Ad St` = `ADMIN STATUS`,
                                `Reason Code` = `OP REASON CODE`,
                                `Last Message` = `LAST STATUS COMMENT`,
                                `Note Date` = `LAST COMMENT DATE`,
                                `Sch. Date` = SCHEDULED,
                                `Notes / Instructions`,
                                Email = `AGENT EMAIL`,
                                `Tel #` = `AGENT PHONE`
                        ) %>% 
                        UTF8_compatible_dataframe
        })
        
        
        # Download Excel ------------------------------------------------------------------

        output$TM_Returns_Daily_Report_download <- downloadHandler(
                filename = function() {
                        paste("TM_Return_Daily_Report.xlsx-", Sys.Date(), '.xlsx', sep='')
                },
                content = function(con) {
                        if (is.null(authorization.check(isolate(Return_Coordinators)))) return(NULL)
                        Return_TM_Region_Categorised <- Return_TM_Region_Categorised()
                        
                        Return_TM_Region_Categorised %>%
                                create_return_daily_worksheet("Return_Daily_Report") %>%
                                openxlsx::saveWorkbook(con)
                }
        )
        
        
        # Upload Modified ------------------------------------------------------------------


        Return_TM_Modified <- reactive({

                if (is.null(input$TM_Returns_Daily_Report_upload)) return(NULL)
                TM_modified <- input$TM_Returns_Daily_Report_upload

                if (!grepl("TM_Return_Daily_Report", TM_modified$name)) {
                        showModal(modalDialog(
                                title = "CAUTION!",
                                "Incorrect input file!",
                                easyClose = TRUE,
                                footer = NULL
                        ))
                        return(NULL)
                }

                Return_TM_Modified <- read_excel(TM_modified$datapath)

        })

        # All Agents Data ---------------------------------------------------------


        TM_return_data_All_Agents <- reactive({
                
                if (is.null(Return_TM_Modified())) {
                        if(is.null(Return_TM_Region_Categorised())) {
                                return(NULL)
                        } else {
                                Return_TM_Region_Categorised() %>% 
                                        mutate(`Notes / Instructions` = `Notes / Instructions` %>% replace_na("Not-Assigned"))
                        }
                } else {
                        Return_TM_Modified() %>% 
                                mutate(`Notes / Instructions` = `Notes / Instructions` %>% replace_na("Not-Assigned"))
                }
        })


        # Returns Escalation Emails -----------------------------------------------

        observeEvent(input$Returns_send_escalation_emails, {

                if (is.null(authorization.check(isolate(Return_Coordinators)))) return(NULL)

                TM <- isolate(TM_return_data_All_Agents())

                if (is.null(TM)) {
                        showModal(modalDialog(
                                title = "CAUTION!",
                                "Warning: No Data imported!",
                                easyClose = TRUE,
                                footer = NULL
                        ))
                        return(NULL)
                }

                foldername <- paste0(file.path(Sys.getenv("USERPROFILE"),"Desktop"), "/Escalations ", gsub(":","-",Sys.time()), "/")
                dir.create(foldername)

                if (!is.null(TM)) {

                        WB_list <- TM %>%
                                filter(`Notes / Instructions` == "Escalation") %>%
                                select(`Job Date`:STAT) %>%
                                split(.$AGENT)


                        withProgress(message = 'Creating TM attachments in progress', value = 0, {
                                for (agent_name in names(WB_list)) {
                                        incProgress(1/length(WB_list))

                                        agent_foldername <- paste0(foldername, agent_name, "/")
                                        dir.create(paste0(foldername, agent_name, "/"))

                                        WB_list[[agent_name]] %>%
                                                create_return_daily_worksheet("Return_Daily_Report") %>%
                                                saveWorkbook(
                                                        paste0(agent_foldername, agent_name, "_TM.xlsx"),
                                                        overwrite = TRUE
                                                )
                                }
                        })
                }

                all_of_the_agents <- list.dirs(path = foldername, full.names = F, recursive = TRUE)
                all_of_the_agents <- all_of_the_agents[all_of_the_agents != ""]
                
                Count_of_agents <- length(all_of_the_agents)
                
                if (Count_of_agents > 0) {
                        shinyalert(
                                paste0(Count_of_agents, " emails are going to be sent to ", Count_of_agents, " agents"),
                                type = "info",
                                showCancelButton = T,
                                showConfirmButton = T,
                                callbackR =
                                        function(x) {
                                                if(x != FALSE) {
                                                        
                                                        OutApp <- COMCreate("Outlook.Application")
                                                        withProgress(message = 'Sending E-mails in progress', value = 0, {
                                                                for (agent in all_of_the_agents) {
                                                                        incProgress(1/Count_of_agents)
                                                                        
                                                                        outMail = OutApp$CreateItem(0)
                                                                        outMail[["To"]] = email_address(
                                                                                paste0("AG_", agent,"@nonstopdelivery.com"),
                                                                                isolate(input$Test_Active)
                                                                        )
                                                                        outMail[["subject"]] = paste0("Escalated/LIVE Returns/CODE Red/Immediate Action Needed - ", agent)
                                                                        outMail[["HTMLBody"]] =
                                                                                paste0(
                                                                                        "<p>Good day - </p>",
                                                                                        "<br>",
                                                                                        "<p>These customers MUST be contacted today. These are escalated orders and we need them to be called ASAP. Please, see attached orders, call customers and update the system with scheduled date.</p>",
                                                                                        "<p>Please confirm email was received.</p>",
                                                                                        nsd_signature(Sys.info()["user"])
                                                                                )
                                                                        
                                                                        attachment_TM <- paste0(foldername, agent, "/", agent, "_TM.xlsx")
                                                                        if (file.exists(attachment_TM)) {
                                                                                outMail[["attachments"]]$Add(attachment_TM)
                                                                        }
                                                                        
                                                                        outMail$Send()
                                                                        
                                                                }
                                                        })
                                                        
                                                        shinyalert(
                                                                title = sample(good_job_words,1),
                                                                text = "All of the Escalation emails were sent to the corresponding agents",
                                                                closeOnEsc = TRUE,
                                                                closeOnClickOutside = TRUE,
                                                                html = FALSE,
                                                                type = "success",
                                                                showConfirmButton = TRUE,
                                                                showCancelButton = FALSE,
                                                                confirmButtonText = "OK",
                                                                confirmButtonCol = "#a2a8ab",
                                                                timer = 0,
                                                                imageUrl = "",
                                                                animation = TRUE
                                                        )
                                                        
                                                }
                                        }
                        )
                } else {
                        shinyalert(
                                title = "Oops!",
                                text = "There are no Escalations to be sent!",
                                closeOnEsc = TRUE,
                                closeOnClickOutside = TRUE,
                                html = FALSE,
                                type = "error",
                                showConfirmButton = TRUE,
                                showCancelButton = FALSE,
                                confirmButtonText = "OK",
                                confirmButtonCol = "#a2a8ab",
                                timer = 0,
                                imageUrl = "",
                                animation = TRUE
                        )
                }

        })


        # Returns 2 Week Emails ---------------------------------------------------

        observeEvent(input$Returns_send_2_week_emails, {

                if (is.null(authorization.check(isolate(Return_Coordinators)))) return(NULL)

                TM <- isolate(TM_return_data_All_Agents())

                if (is.null(TM)) {
                        showModal(modalDialog(
                                title = "CAUTION!",
                                "Warning: No Data imported!",
                                easyClose = TRUE,
                                footer = NULL
                        ))
                        return(NULL)
                }

                foldername <- paste0(file.path(Sys.getenv("USERPROFILE"),"Desktop"), "/2 Weeks ", gsub(":","-",Sys.time()), "/")
                dir.create(foldername)

                if (!is.null(TM)) {

                        WB_list <- TM %>%
                                filter(`Notes / Instructions` == "2 Weeks") %>%
                                select(`Job Date`:STAT) %>%
                                split(.$AGENT)


                        withProgress(message = 'Creating TM attachments in progress', value = 0, {
                                for (agent_name in names(WB_list)) {
                                        incProgress(1/length(WB_list))

                                        agent_foldername <- paste0(foldername, agent_name, "/")
                                        dir.create(paste0(foldername, agent_name, "/"))

                                        WB_list[[agent_name]] %>%
                                                create_return_daily_worksheet("Return_Daily_Report") %>%
                                                saveWorkbook(
                                                        paste0(agent_foldername, agent_name, "_TM.xlsx"),
                                                        overwrite = TRUE
                                                )
                                }
                        })
                }

                all_of_the_agents <- list.dirs(path = foldername, full.names = F, recursive = TRUE)
                all_of_the_agents <- all_of_the_agents[all_of_the_agents != ""]

                Count_of_agents <- length(all_of_the_agents)
                
                if (Count_of_agents > 0) {
                        shinyalert(
                                paste0(Count_of_agents, " emails are going to be sent to ", Count_of_agents, " agents"),
                                type = "info",
                                showCancelButton = T,
                                showConfirmButton = T,
                                callbackR =
                                        function(x) {
                                                if(x != FALSE) {
                                                        
                                                        OutApp <- COMCreate("Outlook.Application")
                                                        withProgress(message = 'Sending E-mails in progress', value = 0, {
                                                                for (agent in all_of_the_agents) {
                                                                        incProgress(1/length(all_of_the_agents))
                                                                        
                                                                        outMail = OutApp$CreateItem(0)
                                                                        outMail[["To"]] = email_address(
                                                                                paste0("AG_", agent,"@nonstopdelivery.com"),
                                                                                isolate(input$Test_Active)
                                                                        )
                                                                        outMail[["subject"]] = paste0("Critical/ LIVE Returns/CODE Red/Immediate Action Needed - ", agent)
                                                                        outMail[["HTMLBody"]] =
                                                                                paste0(
                                                                                        "<p>Good day - </p>",
                                                                                        "<br>",
                                                                                        "<p>These orders have been in the system for more than 2 weeks and still not scheduled for pickup.</p>",
                                                                                        "<p>Please, update them by COB today.</p>",
                                                                                        "<p>Please let me know if you have any questions.</p>",
                                                                                        "<p>Please confirm email was received.</p>",
                                                                                        nsd_signature(Sys.info()["user"])
                                                                                )
                                                                        
                                                                        attachment_TM <- paste0(foldername, agent, "/", agent, "_TM.xlsx")
                                                                        if (file.exists(attachment_TM)) {
                                                                                outMail[["attachments"]]$Add(attachment_TM)
                                                                        }
                                                                        
                                                                        outMail$Send()
                                                                        
                                                                }
                                                        })
                                                        
                                                        shinyalert(
                                                                title = sample(good_job_words,1),
                                                                text = "All of the 2 Weeks emails were sent to the corresponding agents",
                                                                closeOnEsc = TRUE,
                                                                closeOnClickOutside = TRUE,
                                                                html = FALSE,
                                                                type = "success",
                                                                showConfirmButton = TRUE,
                                                                showCancelButton = FALSE,
                                                                confirmButtonText = "OK",
                                                                confirmButtonCol = "#a2a8ab",
                                                                timer = 0,
                                                                imageUrl = "",
                                                                animation = TRUE
                                                        )
                                                        
                                                }
                                        }
                        )
                } else {
                        shinyalert(
                                title = "Oops!",
                                text = "There are no 2 Week emails to be sent!",
                                closeOnEsc = TRUE,
                                closeOnClickOutside = TRUE,
                                html = FALSE,
                                type = "error",
                                showConfirmButton = TRUE,
                                showCancelButton = FALSE,
                                confirmButtonText = "OK",
                                confirmButtonCol = "#a2a8ab",
                                timer = 0,
                                imageUrl = "",
                                animation = TRUE
                        )
                }

        })


        # Returns Update Emails -------------------------------------------


        observeEvent(input$Returns_send_update_emails, {

                if (is.null(authorization.check(isolate(Return_Coordinators)))) return(NULL)

                TM <- isolate(TM_return_data_All_Agents())

                if (is.null(TM)) {
                        showModal(modalDialog(
                                title = "CAUTION!",
                                "Warning: No Data imported!",
                                easyClose = TRUE,
                                footer = NULL
                        ))
                        return(NULL)
                }

                foldername <- paste0(file.path(Sys.getenv("USERPROFILE"),"Desktop"), "/Updates ", gsub(":","-",Sys.time()), "/")
                dir.create(foldername)

                if (!is.null(TM)) {

                        WB_list <- TM %>%
                                filter(`Notes / Instructions` == "Update") %>%
                                select(`Job Date`:STAT) %>%
                                split(.$AGENT)


                        withProgress(message = 'Creating TM attachments in progress', value = 0, {
                                for (agent_name in names(WB_list)) {
                                        incProgress(1/length(WB_list))

                                        agent_foldername <- paste0(foldername, agent_name, "/")
                                        dir.create(paste0(foldername, agent_name, "/"))

                                        WB_list[[agent_name]] %>%
                                                create_return_daily_worksheet("Return_Daily_Report") %>%
                                                saveWorkbook(
                                                        paste0(agent_foldername, agent_name, "_TM.xlsx"),
                                                        overwrite = TRUE
                                                )
                                }
                        })
                }

                all_of_the_agents <- list.dirs(path = foldername, full.names = F, recursive = TRUE)
                all_of_the_agents <- all_of_the_agents[all_of_the_agents != ""]

                Count_of_agents <- length(all_of_the_agents)
                
                if (Count_of_agents > 0) {
                        shinyalert(
                                paste0(Count_of_agents, " emails are going to be sent to ", Count_of_agents, " agents"),
                                type = "info",
                                showCancelButton = T,
                                showConfirmButton = T,
                                callbackR =
                                        function(x) {
                                                if(x != FALSE) {
                                                        
                                                        OutApp <- COMCreate("Outlook.Application")
                                                        withProgress(message = 'Sending E-mails in progress', value = 0, {
                                                                for (agent in all_of_the_agents) {
                                                                        incProgress(1/length(all_of_the_agents))
                                                                        
                                                                        outMail = OutApp$CreateItem(0)
                                                                        outMail[["To"]] = email_address(
                                                                                paste0("AG_", agent,"@nonstopdelivery.com"),
                                                                                isolate(input$Test_Active)
                                                                        )
                                                                        outMail[["subject"]] = paste0("Threshold / White Glove Returns - ", agent)
                                                                        outMail[["HTMLBody"]] =
                                                                                paste0(
                                                                                        "<p>Good day - </p>",
                                                                                        "<br>",
                                                                                        "<p>The attached spreadsheets are Threshold / White Glove returns that are in your NSD console. Please call these customers today to schedule the pickups accordingly, let me know if you have any questions. Please confirm email was received.</p>",
                                                                                        nsd_signature(Sys.info()["user"])
                                                                                )
                                                                        
                                                                        attachment_TM <- paste0(foldername, agent, "/", agent, "_TM.xlsx")
                                                                        if (file.exists(attachment_TM)) {
                                                                                outMail[["attachments"]]$Add(attachment_TM)
                                                                        }
                                                                        
                                                                        outMail$Send()
                                                                        
                                                                }
                                                        })
                                                        
                                                        shinyalert(
                                                                title = sample(good_job_words,1),
                                                                text = "All of the Update emails were sent to the corresponding agents",
                                                                closeOnEsc = TRUE,
                                                                closeOnClickOutside = TRUE,
                                                                html = FALSE,
                                                                type = "success",
                                                                showConfirmButton = TRUE,
                                                                showCancelButton = FALSE,
                                                                confirmButtonText = "OK",
                                                                confirmButtonCol = "#a2a8ab",
                                                                timer = 0,
                                                                imageUrl = "",
                                                                animation = TRUE
                                                        )
                                                        
                                                }
                                        }
                        )
                } else {
                        shinyalert(
                                title = "Oops!",
                                text = "There are no Update emails to be sent!",
                                closeOnEsc = TRUE,
                                closeOnClickOutside = TRUE,
                                html = FALSE,
                                type = "error",
                                showConfirmButton = TRUE,
                                showCancelButton = FALSE,
                                confirmButtonText = "OK",
                                confirmButtonCol = "#a2a8ab",
                                timer = 0,
                                imageUrl = "",
                                animation = TRUE
                        )
                }

        })


        # Returns WPU Emails ------------------------------------------------------


        observeEvent(input$Returns_send_WPU_emails, {

                if (is.null(authorization.check(isolate(Return_Coordinators)))) return(NULL)

                TM <- isolate(TM_return_data_All_Agents())

                if (is.null(TM)) {
                        showModal(modalDialog(
                                title = "CAUTION!",
                                "Warning: No Data imported!",
                                easyClose = TRUE,
                                footer = NULL
                        ))
                        return(NULL)
                }

                foldername <- paste0(file.path(Sys.getenv("USERPROFILE"),"Desktop"), "/WPU ", gsub(":","-",Sys.time()), "/")
                dir.create(foldername)

                if (!is.null(TM)) {

                        WB_list <- TM %>%
                                filter(`Notes / Instructions` == "WPU") %>%
                                select(`Job Date`:STAT) %>%
                                split(.$AGENT)


                        withProgress(message = 'Creating TM attachments in progress', value = 0, {
                                for (agent_name in names(WB_list)) {
                                        incProgress(1/length(WB_list))

                                        agent_foldername <- paste0(foldername, agent_name, "/")
                                        dir.create(paste0(foldername, agent_name, "/"))

                                        WB_list[[agent_name]] %>%
                                                create_return_daily_worksheet("Return_Daily_Report") %>%
                                                saveWorkbook(
                                                        paste0(agent_foldername, agent_name, "_TM.xlsx"),
                                                        overwrite = TRUE
                                                )
                                }
                        })
                }

                all_of_the_agents <- list.dirs(path = foldername, full.names = F, recursive = TRUE)
                all_of_the_agents <- all_of_the_agents[all_of_the_agents != ""]
                
                Count_of_agents <- length(all_of_the_agents)
                
                if (Count_of_agents > 0) {
                        shinyalert(
                                paste0(Count_of_agents, " emails are going to be sent to ", Count_of_agents, " agents"),
                                type = "info",
                                showCancelButton = T,
                                showConfirmButton = T,
                                callbackR =
                                        function(x) {
                                                if(x != FALSE) {
                                                        
                                                        OutApp <- COMCreate("Outlook.Application")
                                                        withProgress(message = 'Sending E-mails in progress', value = 0, {
                                                                for (agent in all_of_the_agents) {
                                                                        incProgress(1/length(all_of_the_agents))
                                                                        
                                                                        outMail = OutApp$CreateItem(0)
                                                                        outMail[["To"]] = email_address(
                                                                                paste0("AG_", agent,"@nonstopdelivery.com"),
                                                                                isolate(input$Test_Active)
                                                                        )
                                                                        outMail[["subject"]] = paste0("WPU LTL - ", agent)
                                                                        outMail[["HTMLBody"]] =
                                                                                paste0(
                                                                                        "<p>Good day - </p>",
                                                                                        "<br>",
                                                                                        "<p>The following return orders are not tracking via the LTL carrier. Please see attached and verify if these are on your dock, or released to the LTL carrier. If still at your warehouse, please verify they are ready for pickup so I can schedule a pickup for the day after tomorrow. Please let me know if you need the outbound BOL. Otherwise, I will just reschedule for pickup 2nd business day. If already picked up, please email me a copy of the LTL paperwork showing the carrier pro number so I can update the system.</p>",
                                                                                        "<p>Please let me know if you have any questions.</p>",
                                                                                        nsd_signature(Sys.info()["user"])
                                                                                )
                                                                        
                                                                        attachment_TM <- paste0(foldername, agent, "/", agent, "_TM.xlsx")
                                                                        if (file.exists(attachment_TM)) {
                                                                                outMail[["attachments"]]$Add(attachment_TM)
                                                                        }
                                                                        
                                                                        outMail$Send()
                                                                        
                                                                }
                                                        })
                                                        
                                                        shinyalert(
                                                                title = sample(good_job_words,1),
                                                                text = "All of the WPU emails were sent to the corresponding agents",
                                                                closeOnEsc = TRUE,
                                                                closeOnClickOutside = TRUE,
                                                                html = FALSE,
                                                                type = "success",
                                                                showConfirmButton = TRUE,
                                                                showCancelButton = FALSE,
                                                                confirmButtonText = "OK",
                                                                confirmButtonCol = "#a2a8ab",
                                                                timer = 0,
                                                                imageUrl = "",
                                                                animation = TRUE
                                                        )
                                                        
                                                }
                                        }
                        )
                } else {
                        shinyalert(
                                title = "Oops!",
                                text = "There are no WPU emails to be sent!",
                                closeOnEsc = TRUE,
                                closeOnClickOutside = TRUE,
                                html = FALSE,
                                type = "error",
                                showConfirmButton = TRUE,
                                showCancelButton = FALSE,
                                confirmButtonText = "OK",
                                confirmButtonCol = "#a2a8ab",
                                timer = 0,
                                imageUrl = "",
                                animation = TRUE
                        )
                }
        })


        # Returns Active Emails ---------------------------------------------------

        observeEvent(input$Returns_send_active_emails, {

                if (is.null(authorization.check(isolate(Return_Coordinators)))) return(NULL)

                TM <- isolate(TM_return_data_All_Agents())

                if (is.null(TM)) {
                        showModal(modalDialog(
                                title = "CAUTION!",
                                "Warning: No Data imported!",
                                easyClose = TRUE,
                                footer = NULL
                        ))
                        return(NULL)
                }

                foldername <- paste0(file.path(Sys.getenv("USERPROFILE"),"Desktop"), "/Actives ", gsub(":","-",Sys.time()), "/")
                dir.create(foldername)

                if (!is.null(TM)) {

                        WB_list <- TM %>%
                                filter(`Notes / Instructions` == "Active") %>%
                                select(`Job Date`:STAT) %>%
                                split(.$AGENT)


                        withProgress(message = 'Creating TM attachments in progress', value = 0, {
                                for (agent_name in names(WB_list)) {
                                        incProgress(1/length(WB_list))

                                        agent_foldername <- paste0(foldername, agent_name, "/")
                                        dir.create(paste0(foldername, agent_name, "/"))

                                        WB_list[[agent_name]] %>%
                                                create_return_daily_worksheet("Return_Daily_Report") %>%
                                                saveWorkbook(
                                                        paste0(agent_foldername, agent_name, "_TM.xlsx"),
                                                        overwrite = TRUE
                                                )
                                }
                        })
                }

                all_of_the_agents <- list.dirs(path = foldername, full.names = F, recursive = TRUE)
                all_of_the_agents <- all_of_the_agents[all_of_the_agents != ""]
                
                Count_of_agents <- length(all_of_the_agents)
                
                if (Count_of_agents > 0) {
                        shinyalert(
                                paste0(Count_of_agents, " emails are going to be sent to ", Count_of_agents, " agents"),
                                type = "info",
                                showCancelButton = T,
                                showConfirmButton = T,
                                callbackR =
                                        function(x) {
                                                if(x != FALSE) {
                                                        
                                                        OutApp <- COMCreate("Outlook.Application")
                                                        withProgress(message = 'Sending E-mails in progress', value = 0, {
                                                                for (agent in all_of_the_agents) {
                                                                        incProgress(1/length(all_of_the_agents))
                                                                        
                                                                        outMail = OutApp$CreateItem(0)
                                                                        outMail[["To"]] = email_address(
                                                                                paste0("AG_", agent,"@nonstopdelivery.com"),
                                                                                isolate(input$Test_Active)
                                                                        )
                                                                        outMail[["subject"]] = paste0("!!!Active Orders!!! - ", agent)
                                                                        outMail[["HTMLBody"]] =
                                                                                paste0(
                                                                                        "<p>Good day - </p>",
                                                                                        "<br>",
                                                                                        "<p>The attached RETURN orders are ACTIVE and need to be updated in order to avoid LATE POD FEES.</p>",
                                                                                        "<p>Please, provide an update on either orders were picked up or attempted.</p>",
                                                                                        nsd_signature(Sys.info()["user"])
                                                                                )
                                                                        
                                                                        attachment_TM <- paste0(foldername, agent, "/", agent, "_TM.xlsx")
                                                                        if (file.exists(attachment_TM)) {
                                                                                outMail[["attachments"]]$Add(attachment_TM)
                                                                        }
                                                                        
                                                                        outMail$Send()
                                                                        
                                                                }
                                                        })
                                                        
                                                        shinyalert(
                                                                title = sample(good_job_words,1),
                                                                text = "All of the Active emails were sent to the corresponding agents",
                                                                closeOnEsc = TRUE,
                                                                closeOnClickOutside = TRUE,
                                                                html = FALSE,
                                                                type = "success",
                                                                showConfirmButton = TRUE,
                                                                showCancelButton = FALSE,
                                                                confirmButtonText = "OK",
                                                                confirmButtonCol = "#a2a8ab",
                                                                timer = 0,
                                                                imageUrl = "",
                                                                animation = TRUE
                                                        )
                                                        
                                                }
                                        }
                        )
                } else {
                        shinyalert(
                                title = "Oops!",
                                text = "There are no Active emails to be sent!",
                                closeOnEsc = TRUE,
                                closeOnClickOutside = TRUE,
                                html = FALSE,
                                type = "error",
                                showConfirmButton = TRUE,
                                showCancelButton = FALSE,
                                confirmButtonText = "OK",
                                confirmButtonCol = "#a2a8ab",
                                timer = 0,
                                imageUrl = "",
                                animation = TRUE
                        )
                }
        })

        # Unyson Worksheet --------------------------------------------------------

        output$Returns_Create_Unyson_worksheet <-
                downloadHandler(
                        filename = function(){
                                paste(
                                        "UNYSON worksheet at",
                                        gsub(":","-",Sys.time()), ".xlsx", sep="_")
                        },
                        content = function(file) {

                                job_numbers <-
                                        read_excel(choose.files("UNYSON Input"), col_types = c("text"), col_names = FALSE) %>%
                                        select(`NSD #` = 1) %>%
                                        filter(!is.na(`NSD #`)) %>%
                                        mutate(`NSD #` = str_trim(`NSD #`)) %>%
                                        distinct

                                Return_TM <-
                                        Return_TM() %>%
                                        left_join(
                                                Agents %>%
                                                        select(
                                                                AGENT = Agent,
                                                                Full_Name
                                                        ) %>%
                                                        mutate(AGENT = AGENT %>% toupper),
                                                by = "AGENT"
                                        ) %>%
                                        mutate(`Orders ready for pickup/ Notes` = "") %>%
                                        select(
                                                `NSD #`,
                                                `UNYSON SHIPMENT#` = `CARRIER PRO`,
                                                `REF-1 PO` = `PO#`,
                                                `AGENT FULL NAME` = Full_Name,
                                                AGENT,
                                                `Customer Name` = SHIPPER,
                                                `Agent City` = AGENT_CITY,
                                                `Agent State` = `AGENT STATE`,
                                                `Orders ready for pickup/ Notes`,
                                                `STOP 3` = CONSIGNEE,
                                                `Vendor City` = CONSIGNEE_CITY,
                                                `Vendor State` = CONSIGNEE_STATE,
                                                STATUS = `OP STATUS`,
                                                Items = ITEM_COUNT,
                                                Weight = WEIGHT
                                        ) %>%
                                        inner_join(job_numbers, by = "NSD #")

                                output <-
                                        job_numbers %>%
                                        left_join(Return_TM, by = "NSD #")



                                wb <- openxlsx::createWorkbook()
                                openxlsx::addWorksheet(wb, "UNYSON", zoom = 85)
                                openxlsx::writeData(wb, 1, output, startRow = 1, startCol = 1)
                                setColWidths(wb, 1, cols = 1:ncol(output), widths = "auto")

                                generalStyle <-
                                        createStyle(
                                                fontName = "Arial" ,
                                                fontSize = 11,
                                                fontColour = "#000000",
                                                halign = "left",
                                                valign = "bottom",
                                                fgFill = "#ecf0f1",
                                                border="TopBottomLeftRight",
                                                borderColour = "#000000",
                                                wrapText = FALSE,
                                                textDecoration = NULL
                                        )
                                headerStyle <-
                                        createStyle(
                                                fontName = "Arial" ,
                                                fontSize = 10,
                                                fontColour = "#000000",
                                                halign = "left",
                                                valign = "bottom",
                                                fgFill = "#8c9ac2",
                                                border="TopBottomLeftRight",
                                                borderColour = "#000000",
                                                wrapText = FALSE,
                                                textDecoration = "bold"
                                        )

                                addStyle(wb, sheet = 1, generalStyle, rows = 1:nrow(output) + 1, cols = 1:ncol(output), gridExpand = TRUE)
                                addStyle(wb, sheet = 1, headerStyle, rows = 1, cols = 1:ncol(output), gridExpand = TRUE)

                                wb %>% openxlsx::saveWorkbook(file, overwrite = TRUE)
                        })


        # Home Depot Emails -------------------------------------------------------

        observeEvent(input$Returns_send_HD_emails, {

                if (is.null(authorization.check(isolate(Return_Coordinators)))) return(NULL)

                TM <- isolate(Return_TM_Region())

                if (is.null(TM)) {
                        showModal(modalDialog(
                                title = "CAUTION!",
                                "Warning: No Data imported!",
                                easyClose = TRUE,
                                footer = NULL
                        ))
                        return(NULL)
                }

                already_sent_job_numbers <-
                        file("//nsdstor1/SHARED/NSDComp/Logs/Home Depot Customer Emails.txt") %>%
                        readLines(skipNul = T) %>%
                        unique

                email_orders <-
                        TM %>%
                        filter(CLIENT %in% c("THE HOME DEPOT", "THE HOMEDEPOT"),
                               `OP STATUS` %in% c("AVAIL-R", "SHPTACK-T"),
                               `CONTACT ATTEMPTS` %in% c(0, 1),
                               !is.na(SHIPPER_EMAIL),
                               !is.na(`NSD #`),
                               !`NSD #` %in% already_sent_job_numbers,
                               grepl("^[[:alnum:].-_]+@[[:alnum:].-]+$", SHIPPER_EMAIL)
                        ) %>%
                        select(`NSD #`, SHIPPER_EMAIL)

                Count_of_orders <- nrow(email_orders)

                if (Count_of_orders > 0) {
                        shinyalert(
                                paste0(Count_of_orders, " emails are going to be sent to the customers"),
                                type = "info",
                                showCancelButton = T,
                                showConfirmButton = T,
                                callbackR =
                                        function(x) {
                                                if(x != FALSE) {

                                                        OutApp <- COMCreate("Outlook.Application")
                                                        withProgress(message = 'Sending E-mails in progress', value = 0, {
                                                                for (i in 1:Count_of_orders) {
                                                                        incProgress(1/Count_of_orders)

                                                                        job_number <- email_orders[i, "NSD #"] %>% as.character
                                                                        email <- email_orders[i, "SHIPPER_EMAIL"] %>% as.character

                                                                        outMail = OutApp$CreateItem(0)
                                                                        outMail[["To"]] = email_address(email, isolate(input$Test_Active))
                                                                        outMail[["subject"]] = paste0("Home Depot Return/ Tracking # ", job_number, " - PLEASE DO NOT REPLY")
                                                                        outMail[["HTMLBody"]] =
                                                                                paste0(
                                                                                        "<p>Hello,</p>",
                                                                                        "<br>",
                                                                                        "<p>We are reaching out to you to schedule a pickup of the order to be returned to Home Depot. Please, advise us if you need to schedule a pickup and provide us with a <b><u>good contact number</u></b> we can reach you at to schedule a pickup date and time.</p>",
                                                                                        "<p>Please contact our customer service at 800.956.7212  to schedule or confirm your current pick-up order</p>",
                                                                                        "<br>",
                                                                                        "<p>Thanks,</p>",
                                                                                        "<br>",
                                                                                        "<p>Reverse Logistics</p>",
                                                                                        "<p>-Cooperate Office</p>",
                                                                                        "<p>NSD, Inc</p>",
                                                                                        "<p>www.nonstopdelivery.com</p>"

                                                                                )
                                                                        outMail$Send()

                                                                        if (!isolate(input$Test_Active)) {
                                                                                job_number %>%
                                                                                        write(file="//nsdstor1/SHARED/NSDComp/Logs/Home Depot Customer Emails.txt",append=TRUE)
                                                                        }
                                                                }
                                                        })

                                                        shinyalert(
                                                                title = sample(good_job_words,1),
                                                                text = "All of the Home Depot Customer emails were sent to the corresponding recipients",
                                                                closeOnEsc = TRUE,
                                                                closeOnClickOutside = TRUE,
                                                                html = FALSE,
                                                                type = "success",
                                                                showConfirmButton = TRUE,
                                                                showCancelButton = FALSE,
                                                                confirmButtonText = "OK",
                                                                confirmButtonCol = "#a2a8ab",
                                                                timer = 0,
                                                                imageUrl = "",
                                                                animation = TRUE
                                                        )
                                                }
                                        }
                        )
                } else {
                        shinyalert(
                                title = "Oops!",
                                text = "Something went wrong! There were no emails to be sent!",
                                closeOnEsc = TRUE,
                                closeOnClickOutside = TRUE,
                                html = FALSE,
                                type = "error",
                                showConfirmButton = TRUE,
                                showCancelButton = FALSE,
                                confirmButtonText = "OK",
                                confirmButtonCol = "#a2a8ab",
                                timer = 0,
                                imageUrl = "",
                                animation = TRUE
                        )
                }

        })




        # Wayfair Emails -------------------------------------------------------

        observeEvent(input$Returns_send_wayfair_emails, {

                if (is.null(authorization.check(isolate(Return_Coordinators)))) return(NULL)

                TM <- isolate(Return_TM_Region())

                if (is.null(TM)) {
                        showModal(modalDialog(
                                title = "CAUTION!",
                                "Warning: No Data imported!",
                                easyClose = TRUE,
                                footer = NULL
                        ))
                        return(NULL)
                }

                already_sent_job_numbers <-
                        file("//nsdstor1/SHARED/NSDComp/Logs/Wayfair Customer Emails.txt") %>%
                        readLines(skipNul = T) %>%
                        unique

                email_orders <-
                        TM %>%
                        filter(CLIENT %in% c("WAYFAIR. LLC"),
                               `OP STATUS` %in% c("AVAIL-R", "SHPTACK-T"),
                               `CONTACT ATTEMPTS` %in% c(0, 1),
                               !is.na(SHIPPER_EMAIL),
                               !is.na(`NSD #`),
                               !`NSD #` %in% already_sent_job_numbers,
                               grepl("^[[:alnum:].-_]+@[[:alnum:].-]+$", SHIPPER_EMAIL)
                        ) %>%
                        select(`NSD #`, SHIPPER_EMAIL)

                Count_of_orders <- nrow(email_orders)

                if (Count_of_orders > 0) {
                        shinyalert(
                                paste0(Count_of_orders, " emails are going to be sent to the customers"),
                                type = "info",
                                showCancelButton = T,
                                showConfirmButton = T,
                                callbackR =
                                        function(x) {
                                                if(x != FALSE) {


                                                        OutApp <- COMCreate("Outlook.Application")
                                                        withProgress(message = 'Sending E-mails in progress', value = 0, {
                                                                for (i in 1:Count_of_orders) {
                                                                        incProgress(1/Count_of_orders)

                                                                        job_number <- email_orders[i, "NSD #"] %>% as.character
                                                                        email <- email_orders[i, "SHIPPER_EMAIL"] %>% as.character

                                                                        outMail = OutApp$CreateItem(0)
                                                                        outMail[["To"]] = email_address(email, isolate(input$Test_Active))
                                                                        outMail[["subject"]] = paste0("Wayfair Return/ Tracking # ", job_number, " - PLEASE DO NOT REPLY")
                                                                        outMail[["HTMLBody"]] =
                                                                                paste0(
                                                                                        "<p>Hello,</p>",
                                                                                        "<br>",
                                                                                        "<p>We are reaching out to you to schedule a pickup of the order to be returned to Wayfair. Please, advise us if you need to schedule a pickup and provide us with a <b><u>good contact number</u></b> we can reach you at to schedule a pickup date and time.</p>",
                                                                                        "<p>Please contact our customer service at 800.956.7212  to schedule or confirm your current pick-up order</p>",
                                                                                        "<br>",
                                                                                        "<p>Thanks,</p>",
                                                                                        "<br>",
                                                                                        "<p>Reverse Logistics</p>",
                                                                                        "<p>-Cooperate Office</p>",
                                                                                        "<p>NSD, Inc</p>",
                                                                                        "<p>www.nonstopdelivery.com</p>"

                                                                                )
                                                                        outMail$Send()

                                                                        if (!isolate(input$Test_Active)) {
                                                                                job_number %>%
                                                                                        write(file="//nsdstor1/SHARED/NSDComp/Logs/Wayfair Customer Emails.txt",append=TRUE)
                                                                        }
                                                                }
                                                        })

                                                        shinyalert(
                                                                title = sample(good_job_words,1),
                                                                text = "All of the Wayfair Customer emails were sent to the corresponding recipients",
                                                                closeOnEsc = TRUE,
                                                                closeOnClickOutside = TRUE,
                                                                html = FALSE,
                                                                type = "success",
                                                                showConfirmButton = TRUE,
                                                                showCancelButton = FALSE,
                                                                confirmButtonText = "OK",
                                                                confirmButtonCol = "#a2a8ab",
                                                                timer = 0,
                                                                imageUrl = "",
                                                                animation = TRUE
                                                        )
                                                }
                                        }
                        )
                } else {
                        shinyalert(
                                title = "Oops!",
                                text = "Something went wrong! There were no emails to be sent!",
                                closeOnEsc = TRUE,
                                closeOnClickOutside = TRUE,
                                html = FALSE,
                                type = "error",
                                showConfirmButton = TRUE,
                                showCancelButton = FALSE,
                                confirmButtonText = "OK",
                                confirmButtonCol = "#a2a8ab",
                                timer = 0,
                                imageUrl = "",
                                animation = TRUE
                        )
                }

        })





        # Agent Sidebar -----------------------------------------------------------

        total_agents <- reactive({
                region <- input$region
                TM_return_data_All_Agents <- TM_return_data_All_Agents()

                if (is.null(TM_return_data_All_Agents) | is.null(region)) return(NULL)

                total_agents <-
                        TM_return_data_All_Agents %>%
                        {if ("ALL" %in% region) . else filter(., RC %in% RC_from_Region(region), AGENT != "NA")} %>%
                        pull(AGENT) %>%
                        unique

        })

        Agent_not_in_database <- reactive({

                if (is.null(total_agents())) return(NULL)
                total_agents <- total_agents()

                Agent_not_in_database <- total_agents[!total_agents %in% Agents$Agent]
                if (length(Agent_not_in_database) == 0) return(NULL)
                Agent_not_in_database
        })
        observeEvent(input$return_agent_database_check, {

                if (is.null(authorization.check(isolate(Return_Coordinators)))) return(NULL)

                Agent_not_in_database <- isolate(Agent_not_in_database())
                if (is.null(isolate(total_agents()))) return(NULL)

                if (is.null(Agent_not_in_database))  {
                        shinyalert(
                                title = "Checked!",
                                text = "All of the agents in 6Am and TM reports exist in the database",
                                closeOnEsc = TRUE,
                                closeOnClickOutside = TRUE,
                                html = FALSE,
                                type = "success",
                                showConfirmButton = TRUE,
                                showCancelButton = FALSE,
                                confirmButtonText = "OK",
                                confirmButtonCol = "#a2a8ab",
                                timer = 0,
                                imageUrl = "",
                                animation = TRUE
                        )
                } else {
                        shinyalert(
                                title = "Oops!",
                                text = paste(c("The following agents do not exist in the database:", Agent_not_in_database), collapse = " "),
                                closeOnEsc = TRUE,
                                closeOnClickOutside = TRUE,
                                html = FALSE,
                                type = "error",
                                showConfirmButton = TRUE,
                                showCancelButton = FALSE,
                                confirmButtonText = "OK",
                                confirmButtonCol = "#a2a8ab",
                                timer = 0,
                                imageUrl = "",
                                animation = TRUE
                        )
                }
        })

        observe({updateSelectInput(session, "return_agent_code", choices = c("ALL", total_agents()))})

        output$return_agent_info = renderPrint({

                agent_code <- input$return_agent_code

                if (agent_code %in% Agents$Agent) {
                        name <- agent_code
                        city <- Agents %>% filter(Agent == agent_code) %>% pull(City)
                        state <- Agents %>% filter(Agent == agent_code) %>% pull(State)
                        zip <- Agents %>% filter(Agent == agent_code) %>% pull(zip)
                        Tel <- Agents %>% filter(Agent == agent_code) %>% pull(Tel)
                        email <- Agents %>% filter(Agent == agent_code) %>% pull(Email)

                        cat("**Agent Info**\n\n")

                        if (!is.na(name) & !name %in% c("", "N/A")) cat(paste0(name, "\n"))
                        if (!is.na(city) & !city %in% c("", "N/A")) cat(paste0(city, ", "))
                        if (!is.na(state) & !state %in% c("", "N/A")) cat(paste0(state, " "))
                        if (!is.na(zip) & !zip %in% c("", "N/A")) cat(paste0(zip, "\n"))
                        if (!is.na(Tel) & !Tel %in% c("", "N/A")) cat(paste0("Phone: ", Tel, "\n"))
                        if (!is.na(email) & !email %in% c("", "N/A")) cat(paste0("Email: ", email, "\n"))
                }
        })
        # Return Pie --------------------------------------------------------------
        
        TM_return_data_Agent_Filtered <- reactive({
                
                agent <- input$return_agent_code
                if (is.null(TM_return_data_All_Agents())) return(NULL)
                
                result <-
                        TM_return_data_All_Agents() %>%
                        {if (agent == "ALL") . else filter(., AGENT == agent)}
                if (nrow(result) == 0) return(NULL)
                TM_return_data_Agent_Filtered <- result
        })
        
        output$TM_return_category_pie <- renderPlotly({
                
                TM_return_data_Agent_Filtered <- TM_return_data_Agent_Filtered()
                if (is.null(TM_return_data_Agent_Filtered)) return(NULL)
                
                TM_return_data_Agent_Filtered %>%
                        group_by(`Notes / Instructions`) %>%
                        dplyr::summarize(Count = n()) %>%
                        arrange(Count) %>%
                        plot_ly(labels = ~`Notes / Instructions`, values = ~Count, type = 'pie', source = "A",
                                textinfo = 'label',
                                textfont = list(size = 15),
                                marker = list(line = list(color = "#000000", width = 0.5))
                        ) %>%
                        layout(
                                paper_bgcolor='rgb(255,255,255)',
                                font = list(
                                        family = "Arial, Helvetica, sans-serif",
                                        size = 14,
                                        color = 'black')
                        )
        })
        
        return_pie_hiddenlabels <- reactiveVal(list()) 
        
        observeEvent(input$return_agent_code, {
                return_pie_hiddenlabels(list())
        })
        
        observeEvent(event_data("plotly_relayout", source = "A")$hiddenlabels, {
                d <- event_data("plotly_relayout", source = "A")$hiddenlabels
                return_pie_hiddenlabels(if (is.null(d)) "Hover events appear here (unhover to clear)" else d)
        })
        
        # Returns Calendar --------------------------------------------------------

        TM_Return_Calendar_Data <- reactive({
                
                TM_return_data_Agent_Filtered <- TM_return_data_Agent_Filtered()
                if (is.null(TM_return_data_Agent_Filtered)) return(NULL)
                
                TM_Return_Calendar_Data <- 
                        TM_return_data_Agent_Filtered %>%
                        filter(!`Notes / Instructions` %in% return_pie_hiddenlabels())
                
        })


        output$TM_returns_returns_calendarplot <- renderGvis({
                TM_Return_Calendar_Data <- TM_Return_Calendar_Data()
                if (is.null(TM_Return_Calendar_Data)) return(NULL)
                if (nrow(TM_Return_Calendar_Data) == 0) return(NULL)
                
                TM_Return_Calendar_Data() %>% 
                        pull(`Job Date`) %>% 
                        table %>% 
                        as.data.frame %>% 
                        rename(Date = 1, Count = 2) %>% 
                        mutate(Date = as.Date(Date)) %>%
                        gvisCalendar(
                                datevar = "Date", 
                                numvar = "Count",
                                options = list(
                                        title = "JobDate Distribution",
                                        noDataPattern =  "{backgroundColor: '#000000', color: '#000000'}",
                                        calendar = 
                                                "{
                                                        yearLabel: {fontName: 'Arial', fontSize: 32, color: '#000000', bold: true},
                                                        cellSize: 14,
                                                        focusedCellColor: {stroke:'red'},
                                                        dayOfWeekLabel: {fontName: 'Arial', fontSize: 12, color: '#000000', bold: true},
                                                        dayOfWeekRightSpace: 10,
                                                        cellColor: {stroke: 'red', strokeOpacity: 1, strokeWidth: 1},
                                                        colorAxis: {colors:['red','#004411']}
                                                }",
                                        width="800px",
                                        gvis.listener.jscode = 
                                                paste0(
                                                        "var selected_date = data.getValue(chart.getSelection()[0].row,0);
                                                        var parsed_date = 'TM'+'-'+selected_date.getFullYear()+'-'+(selected_date.getMonth()+1)+'-'+selected_date.getDate();
                                                        Shiny.onInputChange('selected_date',parsed_date)"
                                                )
                                )
                        )
        })

        output$returns_explore_TM_count <- renderText({

                TM_Return_Calendar_Data <- TM_Return_Calendar_Data()
                if (is.null(TM_Return_Calendar_Data)) return(NULL)
                
                paste0("Total: ",
                       TM_Return_Calendar_Data %>%
                               count %>%
                               as.character
                )
        })

        output$returns_explore_table_modal <- renderDT({

                string <- input$selected_date
                if (is.na(string)) return(NULL)

                date_selected <- str_sub(string, 4,str_length(string)) %>% as.Date

                TM_Return_Calendar_Data() %>%
                        filter(`Job Date` == date_selected) %>%
                        select(
                                `NSD #`,
                                Agent = AGENT,
                                `Customer Name`,
                                `Service Type` = SVC,
                                Weight = WGT,
                                Miles = Mile,
                                Status = STAT,
                                `Notes / Instructions`
                        ) %>%
                        arrange(Agent, `Notes / Instructions`) %>%
                        datatable(
                                options = list(
                                        dom = "tp",
                                        pageLength = 10
                                ),
                                rownames = FALSE
                        )
        })

        observeEvent(input$selected_date,{

                string <- input$selected_date
                if (is.na(string)) return(NULL)

                date_selected <- str_sub(string, 4,str_length(string)) %>% as.Date

                showModal(
                        modalDialog(
                                easyClose = T,
                                size = "l",
                                title = h2(date_selected),
                                DTOutput("returns_explore_table_modal"),
                                footer = modalButton("Dismiss")
                        )
                )


        })

# Return Orders -----------------------------------------------------------

        observe({updateSelectInput(session, "return_orders_selectinput", choices = c(Return_TM()$`NSD #`) %>% sort)})
        order_info <- reactive({
                
                jobno <- input$return_orders_selectinput
                Return_TM <- Return_TM()
                if (jobno == "-" | is.null(Return_TM)) return(NULL)
                
                order_info <- Return_TM %>% filter(`NSD #` == jobno)
                
        })
        returns_orders_shipper_data <- reactive({
                
                order_info <- order_info()
                if (is.null(order_info)) return(NULL)
                
                zipa <- order_info %>% pull(SHIPPER_ZIP)
                
                returns_orders_shipper_data <- 
                        list(
                                name = order_info %>% pull(SHIPPER),
                                jobdate = order_info %>% pull(CREATED),
                                phone = order_info %>% pull(SHIPPER_PHONE),
                                email = order_info %>% pull(SHIPPER_EMAIL),
                                city = order_info %>% pull(SHIPPER_CITY),
                                zip = zipa,
                                state = zipcode %>% filter(zip == zipa) %>% pull(state)
                        )
        })
        output$returns_orders_shipper_info <- renderText({
                
                shipper <- returns_orders_shipper_data()
                if (is.null(shipper)) return(NULL)
                
                paste0(
                        "<b>Name: </b>", shipper[["name"]], "<br/>",
                        "<b>Created Date: </b>", shipper[["jobdate"]], "<br/>",
                        "<b>Address: </b>", shipper[["city"]], ", ", shipper[["state"]], " ", shipper[["zip"]], "<br/>",
                        "<b>Phone: </b>", shipper[["phone"]], "<br/>",
                        "<b>Email: </b>", shipper[["email"]], "<br/>"
                        
                ) %>% HTML
        })
        returns_orders_agent_data <- reactive({
                
                order_info <- order_info()
                if (is.null(order_info)) return(NULL)
                
                agent_name <- order_info %>% pull(AGENT)
                
                returns_orders_agent_data <- 
                        list(
                                name = agent_name,
                                full_name = order_info %>% pull(AGENT_NAME),
                                phone = order_info %>% pull(`AGENT PHONE`),
                                email = order_info %>% pull(`AGENT EMAIL`),
                                city = order_info %>% pull(AGENT_CITY),
                                zip = if (agent_name %in% Agents$Agent) Agents %>% filter(Agent == agent_name) %>% pull(zip),
                                state = order_info %>% pull(`AGENT STATE`)
                        )
        })
        output$returns_orders_agent_info <- renderText({
                
                agent <- returns_orders_agent_data()
                if (is.null(agent)) return(NULL)
                
                paste0(
                        "<b>Code: </b>", agent[["name"]], "<br/>",
                        "<b>Name: </b>", agent[["full_name"]], "<br/>",
                        "<b>Address: </b>", agent[["city"]], ", ", agent[["state"]], " ", agent[["zip"]], "<br/>",
                        "<b>Phone: </b>", agent[["phone"]], "<br/>",
                        "<b>Email: </b>", agent[["email"]]
                ) %>% HTML
                
        })
        returns_orders_carrier_data <- reactive({
                
                order_info <- order_info()
                if (is.null(order_info)) return(NULL)
                
                returns_orders_carrier_data <- 
                        list(
                                name = order_info %>% pull(`CARRIER NAME`),
                                carrier_pro = order_info %>% pull(`CARRIER PRO`),
                                unyson_pro = order_info %>% pull(`UNYSON PRO`),
                                ltl_ready = order_info %>% pull(`LTL READY`),
                                ltl_Pickup = order_info %>% pull(`LTL P/U DATE`)
                                
                        )
        })
        output$returns_orders_carrier_info <- renderText({
                
                carrier <- returns_orders_carrier_data()
                if (is.null(carrier)) return(NULL)
                
                paste0(
                        "<b>Carrier: </b>", carrier[["name"]], "<br/>",
                        "<b>Carrier Pro: </b>", carrier[["carrier_pro"]], "<br/>",
                        "<b>Unyson Pro: </b>", carrier[["unyson_pro"]], "<br/>",
                        "<b>LTL Ready: </b>", carrier[["ltl_ready"]], "<br/>",
                        "<b>LTL Pick-Up: </b>", carrier[["ltl_Pickup"]]
                ) %>% HTML
                
        })
        returns_orders_consignee_data <- reactive({
                
                order_info <- order_info()
                if (is.null(order_info)) return(NULL)
                
                returns_orders_consignee_data <- 
                        list(
                                client = order_info %>% pull(CLIENT),
                                client_id = order_info %>% pull(`CLIENT ID`),
                                name = order_info %>% pull(CONSIGNEE),
                                po = order_info %>% pull(`PO#`),
                                city = order_info %>% pull(CONSIGNEE_CITY),
                                state = order_info %>% pull(CONSIGNEE_STATE),
                                zip = order_info %>% pull(CONSIGNEE_ZIP)
                        )
        })
        output$returns_orders_consignee_info <- renderText({
                
                consignee <- returns_orders_consignee_data()
                if (is.null(consignee)) return(NULL)
                
                paste0(
                        "<b>Client: </b>", consignee[["client"]], " (id: ", consignee[["client_id"]], ")", "<br/>",
                        "<b>Name: </b>", consignee[["name"]], "<br/>",
                        "<b>PO number: </b>", consignee[["po"]], "<br/>",
                        "<b>Address: </b>", consignee[["city"]], ", ", consignee[["state"]], " ", consignee[["zip"]], "<br/>",
                        " <br/>"
                ) %>% HTML
                
        })
        output$return_orders_map <- renderLeaflet({
                
                order_info <- order_info()
                if (is.null(order_info)) return(leaflet() %>% addTiles() %>% setView(-96, 37.8, 4))
                
                shipper <- returns_orders_shipper_data()
                agent <- returns_orders_agent_data()
                consignee <- returns_orders_consignee_data()
                
                shipper_zip <- shipper[["zip"]]
                agent_zip <- agent[["zip"]]
                consignee_zip <- consignee[["zip"]]
                
                
                if(!agent[["name"]] %in% Agents$Agent | is.null(agent_zip) | !shipper_zip %in% zipcode$zip) {
                        shinyalert(
                                title = "Oops!",
                                text = "Shipper zipcode does not exist in the zipcodes source",
                                closeOnEsc = TRUE,
                                closeOnClickOutside = TRUE,
                                html = FALSE,
                                type = "error",
                                showConfirmButton = TRUE,
                                showCancelButton = FALSE,
                                confirmButtonText = "OK",
                                confirmButtonCol = "#a2a8ab",
                                timer = 0,
                                imageUrl = "",
                                animation = TRUE
                        )
                        return(NULL)
                }
                if(is.null(agent_zip) | !agent_zip %in% zipcode$zip) {
                        shinyalert(
                                title = "Oops!",
                                text = "Agent zipcode does not exist in the zipcodes source",
                                closeOnEsc = TRUE,
                                closeOnClickOutside = TRUE,
                                html = FALSE,
                                type = "error",
                                showConfirmButton = TRUE,
                                showCancelButton = FALSE,
                                confirmButtonText = "OK",
                                confirmButtonCol = "#a2a8ab",
                                timer = 0,
                                imageUrl = "",
                                animation = TRUE
                        )
                        return(NULL)
                }
                if(is.null(consignee_zip) | !consignee_zip %in% zipcode$zip) {
                        shinyalert(
                                title = "Oops!",
                                text = "Consignee zipcode does not exist in the zipcodes source",
                                closeOnEsc = TRUE,
                                closeOnClickOutside = TRUE,
                                html = FALSE,
                                type = "error",
                                showConfirmButton = TRUE,
                                showCancelButton = FALSE,
                                confirmButtonText = "OK",
                                confirmButtonCol = "#a2a8ab",
                                timer = 0,
                                imageUrl = "",
                                animation = TRUE
                        )
                        return(NULL)
                }
                
                zip_lat <- function(x) {if (x %in% zipcode$zip) zipcode %>% filter(zip == x) %>% pull(latitude)}
                zip_lon <- function(x) {if (x %in% zipcode$zip) zipcode %>% filter(zip == x) %>% pull(longitude)}
                
                shipper_lat <- zip_lat(shipper_zip)
                agent_lat <- zip_lat(agent_zip)
                consignee_lat <- zip_lat(consignee_zip)
                
                shipper_long <- zip_lon(shipper_zip)
                agent_long <- zip_lon(agent_zip)
                consignee_long <- zip_lon(consignee_zip)
                
                shipper_label <- 
                        paste0(
                                "<b>Shipper</b><br/>",
                                shipper[["name"]], "<br/>",
                                shipper[["city"]], ", ", shipper[["state"]], " ", shipper[["zip"]]
                        )
                
                agent_label <- 
                        paste0(
                                "<b>Agent</b><br/>",
                                agent[["name"]], "<br/>",
                                agent[["full_name"]], "<br/>",
                                agent[["city"]], ", ", agent[["state"]], " ", agent[["zip"]]
                        )
                
                consignee_label <- 
                        paste0(
                                "<b>Consignee</b><br/>",
                                consignee[["name"]], "<br/>",
                                consignee[["city"]], ", ", consignee[["state"]], " ", consignee[["zip"]]
                        )
                
                
                map_data <- 
                        data.frame(
                                stop_title = c("Shipper", "Agent", "Consignee"),
                                stop_name = c(shipper[["name"]], agent[["name"]], consignee[["name"]]),
                                stop_num = c(1, 2, 3),
                                latitude = c(shipper_lat, agent_lat, consignee_lat),
                                longitude = c(shipper_long, agent_long, consignee_long)
                                
                        )
                map <- 
                        leaflet() %>% 
                        addTiles() %>%
                        addMarkers(
                                data = map_data %>% filter(stop_title == "Shipper"), 
                                lng = ~longitude, 
                                lat = ~latitude, 
                                group = "Shipper", 
                                popup = shipper_label,
                                popupOptions = popupOptions(closeOnClick = FALSE)
                        ) %>% 
                        addMarkers(
                                data = map_data %>% filter(stop_title == "Agent"), 
                                lng = ~longitude, 
                                lat = ~latitude, 
                                group = "Agent", 
                                popup = agent_label,
                                popupOptions = popupOptions(closeOnClick = FALSE)
                        ) %>% 
                        addMarkers(
                                data = map_data %>% filter(stop_title == "Consignee"), 
                                lng = ~longitude, 
                                lat = ~latitude, 
                                group = "Consignee", 
                                popup = consignee_label,
                                popupOptions = popupOptions(closeOnClick = FALSE)
                        )
        })
        output$return_orders_men <- renderImage({
                order_info <- order_info()
                if (is.null(order_info)) return(list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50))
                
                men <- order_info %>% pull(MAN)
                mode <- order_info %>% pull(MODE)
                
                if (is.na(men)) {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                } else if (mode == "LHR") {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                } else if (men == "2-MAN") {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/twoman.png", height = 50, width = 50)
                } else if (men == "1-MAN") {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/oneman.png", height = 50, width = 50)
                } else {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                }
                
        }, deleteFile = FALSE)
        output$return_orders_servicetype <- renderImage({
                order_info <- order_info()
                if (is.null(order_info)) return(list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50))
                
                service <- order_info %>% pull(`SERVICE LEVEL`)
                mode <- order_info %>% pull(MODE)
                
                if (is.na(service)) {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                } else if (mode == "LHR") {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                } else if (service == "WHITEGLOVE") {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/whiteglove.png", height = 50, width = 50)
                } else if (service == "THRESHOLD") {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/threshold.png", height = 50, width = 50)
                } else {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                }
                
                
        }, deleteFile = FALSE)
        output$return_orders_disposal <- renderImage({
                order_info <- order_info()
                if (is.null(order_info)) return(list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50))
                
                disposal <- order_info %>% pull(`DISPOSAL_SI`)
                
                if (is.na(disposal)) {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                } else if (disposal == "DISPOSAL AUTHORIZED") {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/dispose.png", height = 50, width = 50)
                } else if (disposal == "DISPOSAL AUTHORIZED/HOLD PENDING CLAIM DISPOSITION") {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/dispose-hold.png", height = 50, width = 50)
                } else if (disposal == "HOLD PENDING CLAIM DISPOSITION") {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/dispose-hold.png", height = 50, width = 50)
                } else {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                }
                
                
        }, deleteFile = FALSE)
        output$return_orders_replacement <- renderImage({
                order_info <- order_info()
                if (is.null(order_info)) return(list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50))
                
                replacement <- order_info %>% pull(`REPLACEMENT_ORDER`)
                
                if (is.na(replacement)) {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                } else if (replacement == "YES") {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/replacement.png", height = 50, width = 50)
                } else {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                }
                
                
        }, deleteFile = FALSE)
        output$return_orders_contact_attempts <- renderImage({
                order_info <- order_info()
                if (is.null(order_info)) return(list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50))
                
                attempts <- 
                        order_info %>% 
                        pull(`CONTACT ATTEMPTS`) %>% 
                        as.character
                mode <- order_info %>% pull(MODE)
                
                outfile <- 
                        "//nsdstor1/SHARED/NSDComp/www/phone.png" %>% 
                        image_read %>% 
                        image_annotate(attempts, size = 200, gravity = "southwest", location = "+350-25") %>%
                        image_write(tempfile(fileext='png'), format = 'png')
                
                if (is.na(attempts)) {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                } else if (mode == "LHR") {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                } else {
                        list(src = outfile, height = 50, width = 50)
                }
                
                
                
        }, deleteFile = FALSE)
        output$returns_orders_items <- renderImage({
                order_info <- order_info()
                if (is.null(order_info)) return(list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50))
                
                items <-
                        order_info %>%
                        pull(`ITEM_COUNT`) %>%
                        as.character
                
                sizes = list("1" = "140", "2" = "110", "3" = "80", "4" = "40")
                
                outfile <-
                        "//nsdstor1/SHARED/NSDComp/www/item.png" %>%
                        image_read %>%
                        image_annotate(items, size = 100, gravity = "southwest", location = paste0("+", sizes[[nchar(items)]] %>% as.character, "+100")) %>%
                        image_write(tempfile(fileext='png'), format = 'png')
                
                if (is.na(items)) {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                } else {
                        list(src = outfile, height = 50, width = 50)
                }
                
                
                
        }, deleteFile = FALSE)
        output$returns_orders_weight <- renderImage({
                order_info <- order_info()
                if (is.null(order_info)) return(list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50))
                
                weight <- 
                        order_info %>% 
                        pull(WEIGHT) %>% 
                        {if (as.numeric(.) > 10) round(.) else round(., 2)} %>% 
                        as.character
                
                sizes = list("1" = "190", "2" = "150", "3" = "100", "4" = "60")
                
                
                outfile <-
                        "//nsdstor1/SHARED/NSDComp/www/weight.png" %>%
                        image_read %>%
                        image_annotate(weight, size = 150, gravity = "southwest", location = paste0("+", sizes[[nchar(weight)]] %>% as.character, "+90")) %>%
                        image_write(tempfile(fileext='png'), format = 'png')
                
                if (is.na(weight)) {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                } else {
                        list(src = outfile, height = 50, width = 50)
                }
                
                
                
        }, deleteFile = FALSE)
        output$returns_orders_miles <- renderImage({
                order_info <- order_info()
                if (is.null(order_info)) return(list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50))
                
                miles <- order_info %>% pull(MILEAGE) %>% as.character
                mode <- order_info %>% pull(MODE)
                
                outfile <-
                        "//nsdstor1/SHARED/NSDComp/www/mileage.png" %>%
                        image_read %>%
                        image_annotate(miles, size = 100, gravity = "southwest", location = paste0("+150+20")) %>%
                        image_write(tempfile(fileext='png'), format = 'png')
                
                if (is.na(miles)) {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                } else if (mode == "LHR") {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/LH.png", height = 50, width = 50)
                } else {
                        list(src = outfile, height = 50, width = 50)
                }
                
                
                
        }, deleteFile = FALSE)
        output$returns_orders_op <- renderText({
                
                order_info <- order_info()
                if (is.null(order_info)) return(NULL)
                
                status <- order_info %>% pull(`OP STATUS`)
                reason_code <- order_info %>% pull(`OP REASON CODE`)
                date <- order_info %>% pull(`OP STATUS CHANGED`)
                comment <- order_info %>% pull(`OP STATUS COMMENT`)
                
                paste0(
                        "<b>", if (!is.na(status)) status, "</b><br/>",
                        if (!is.na(date)) date, "<br/>",
                        if (!is.na(reason_code)) reason_code, "<br/>",
                        "<br/>",
                        if (!is.na(comment)) comment
                        
                ) %>% HTML
        })
        output$returns_orders_admin <- renderText({
                
                order_info <- order_info()
                if (is.null(order_info)) return(NULL)
                
                status <- order_info %>% pull(`ADMIN STATUS`)
                reason_code <- order_info %>% pull(`ADMIN REASON CODE`)
                date <- order_info %>% pull(`ADMIN STATUS CHANGED`)
                comment <- order_info %>% pull(`ADMIN STATUS COMMENT`)
                
                paste0(
                        "<b>", if (!is.na(status)) status, "</b><br/>",
                        if (!is.na(date)) date, "<br/>",
                        if (!is.na(reason_code)) reason_code, "<br/>",
                        "<br/>",
                        if (!is.na(comment)) comment
                        
                ) %>% HTML
        })
        output$returns_orders_last <- renderText({
                
                order_info <- order_info()
                if (is.null(order_info)) return(NULL)
                
                date <- order_info %>% pull(`LAST COMMENT DATE`)
                comment <- order_info %>% pull(`LAST STATUS COMMENT`)
                
                paste0(
                        if (!is.na(date)) date, "<br/>",
                        "<br/>",
                        if (!is.na(comment)) comment
                        
                ) %>% HTML
        })
        output$returns_orders_description <- renderDT({
                
                order_info <- order_info()
                if (is.null(order_info)) return(NULL)
                
                order_info %>% 
                        pull(`FREIGHT DESCRIPTION`) %>% 
                        str_split(" / ") %>% 
                        unlist %>% 
                        table %>% 
                        as.data.frame %>% 
                        arrange(Freq) %>% 
                        datatable(
                                options = list(
                                        dom = "tp",
                                        pageLength = 4
                                ),
                                rownames = FALSE,
                                colnames = c('Description', '')
                        )
        })

# Return Consolidation ----------------------------------------------------
        
        return_consolidation_data <- reactive({
                Return_TM <- Return_TM()
                if (is.null(Return_TM)) return(NULL)
                
                return_consolidation_data <- 
                        Return_TM %>% 
                        filter(
                                `OP STATUS` %in% c("LTLAVAIL", "PICKEDUP", "TENDER"),
                                MODE != "LM"
                        ) %>% 
                        left_join(zipcode %>% select(zip, consignee_lat = latitude, consignee_lon = longitude), by = c("CONSIGNEE_ZIP" = "zip")) %>% 
                        left_join(Agents %>% select(Agent, agent_lat = Lat, agent_lon = Long, AGENT_ZIP = zip), by = c("AGENT" = "Agent"))
        })
        
        output$return_consolidation_map_destinations_circle <- renderGvis({
                
                return_consolidation_data <- return_consolidation_data()
                if (is.null(return_consolidation_data)) {
                        return(
                                gvisGeoChart(data.frame(state.name, 1), "state.name", "X1",
                                             options=list(region="US", 
                                                          displayMode="regions", 
                                                          resolution="provinces",
                                                          height = 347*map_scale*0.75,
                                                          width = 556*map_scale*0.75,
                                                          keepAspectRatio = F,
                                                          colorAxis = "{colors:['#f5f5f5', '#f5f5f5']}",
                                                          legend = "none"
                                             )
                                )
                        )
                }
                
                data <- 
                        return_consolidation_data %>%
                        group_by(CONSIGNEE, CONSIGNEE_ZIP, consignee_lat, consignee_lon) %>%
                        summarize(
                                Volume = n()
                        ) %>%
                        mutate(
                                latlong = paste(consignee_lat, consignee_lon, sep = ":"),
                                consignee_id = paste0(CONSIGNEE, "*******", CONSIGNEE_ZIP)
                        ) %>% 
                        arrange(Volume)
                
                Total = data %>% pull(Volume) %>% sum
                
                jscode = sprintf("var returns_consolidation_destination_picked = data.getValue(chart.getSelection()[0].row,2);
                                 Shiny.onInputChange('%s', returns_consolidation_destination_picked.toString())",
                                 session$ns('returns_consolidation_destination_picked'))
                
                data %>%
                        mutate(`Percentage of Total` = (Volume/Total)*100 %>% round(1)) %>% 
                        gvisGeoChart("CONSIGNEE", locationvar = "latlong", sizevar = "Percentage of Total", colorvar = "Volume", hovervar = "consignee_id",
                                     options =
                                             list(
                                                     region = "US",
                                                     resolution = "provinces",
                                                     displayMode = "markers",
                                                     colorAxis = "{colors:['yellow', 'red']}",
                                                     height = 347*map_scale*0.75,
                                                     width = 556*map_scale*0.75,
                                                     keepAspectRatio = F,
                                                     gvis.listener.jscode = jscode
                                             )
                        )
        })
        
        output$return_consolidation_map_origins_circle <- renderGvis({
                
                return_consolidation_data <- return_consolidation_data()
                zip_input <- input$returns_consolidation_destination_picked
                if (is.null(zip_input) | is.null(return_consolidation_data)) {
                        return(
                                gvisGeoChart(data.frame(state.name, 1), "state.name", "X1",
                                             options=list(region="US", 
                                                          displayMode="regions", 
                                                          resolution="provinces",
                                                          height = 347*map_scale*0.75,
                                                          width = 556*map_scale*0.75,
                                                          keepAspectRatio = F,
                                                          colorAxis = "{colors:['#f5f5f5', '#f5f5f5']}",
                                                          legend = "none"
                                             )
                                )
                        )
                }
                
                data <- 
                        return_consolidation_data %>%
                        filter(
                                CONSIGNEE_ZIP ==  strsplit(zip_input, "\\*\\*\\*\\*\\*\\*\\*")[[1]][2],
                                CONSIGNEE ==  strsplit(zip_input, "\\*\\*\\*\\*\\*\\*\\*")[[1]][1]
                        ) %>%
                        group_by(AGENT, agent_lat, agent_lon, AGENT_ZIP) %>%
                        summarize(
                                Volume = n()
                        ) %>%
                        mutate(
                                latlong = paste(agent_lat, agent_lon, sep = ":"),
                                agent_id = paste0(AGENT, "*******", AGENT_ZIP)
                        ) %>% 
                        arrange(Volume)
                
                Total = data %>% pull(Volume) %>% sum
                
                jscode = sprintf("var returns_consolidation_origin_picked = data.getValue(chart.getSelection()[0].row,2);
                                 Shiny.onInputChange('%s', returns_consolidation_origin_picked.toString())",
                                 session$ns('returns_consolidation_origin_picked'))
                
                data %>%
                        mutate(`Percentage of Total` = (Volume/Total)*100 %>% round(1)) %>% 
                        gvisGeoChart("AGENT", locationvar = "latlong", sizevar = "Percentage of Total", colorvar = "Volume", hovervar = "agent_id",
                                     options =
                                             list(
                                                     region = "US",
                                                     resolution = "provinces",
                                                     displayMode = "markers",
                                                     colorAxis = "{colors:['yellow', 'red']}",
                                                     height = 347*map_scale*0.75,
                                                     width = 556*map_scale*0.75,
                                                     keepAspectRatio = F,
                                                     gvis.listener.jscode = jscode
                                             )
                        )
        })
        return_consolidation_map_data <- reactive({
                
                return_consolidation_data <- return_consolidation_data()
                destination_name_and_zip <- input$returns_consolidation_destination_picked
                agent_name_and_zip <- input$returns_consolidation_origin_picked
                
                if(is.null(return_consolidation_data)) return(NULL)
                
                if (!is.null(destination_name_and_zip)) {
                        destination_zip <- strsplit(destination_name_and_zip, "\\*\\*\\*\\*\\*\\*\\*")[[1]][2]
                        destination_name <- strsplit(destination_name_and_zip, "\\*\\*\\*\\*\\*\\*\\*")[[1]][1]
                }
                if (!is.null(agent_name_and_zip)) {
                        agentzip <- strsplit(agent_name_and_zip, "\\*\\*\\*\\*\\*\\*\\*")[[1]][2]
                        agent_name <- strsplit(agent_name_and_zip, "\\*\\*\\*\\*\\*\\*\\*")[[1]][1]
                }
                
                return_consolidation_map_data <- 
                        return_consolidation_data %>% 
                        {if (!is.null(destination_name_and_zip)) {filter(., CONSIGNEE == destination_name, CONSIGNEE_ZIP == destination_zip)} else .} %>% 
                        {if (!is.null(agent_name_and_zip)) {filter(., AGENT == agent_name, AGENT_ZIP == agentzip)} else .}
                        
                        
        })
        output$returns_consolidation_origin_info <- renderText({
                
                if (is.null(observeEvent(input$returns_consolidation_destination_picked, return(NULL)))) return(NULL)
                
                agent_name_and_zip <- input$returns_consolidation_origin_picked
                
                if (!is.null(agent_name_and_zip)) {
                        
                        agentzip <- strsplit(agent_name_and_zip, "\\*\\*\\*\\*\\*\\*\\*")[[1]][2]
                        agent_name <- strsplit(agent_name_and_zip, "\\*\\*\\*\\*\\*\\*\\*")[[1]][1]
                        
                        paste0(
                                "<h3>", agent_name, "</h3><br/>",
                                agentzip
                        ) %>% HTML
                }
                
                
                
                
        })
        output$returns_consolidation_destination_info <- renderText({
                
                destination_name_and_zip <- input$returns_consolidation_destination_picked
                
                if (!is.null(destination_name_and_zip)) {
                        
                        destination_zip <- strsplit(destination_name_and_zip, "\\*\\*\\*\\*\\*\\*\\*")[[1]][2]
                        destination_name <- strsplit(destination_name_and_zip, "\\*\\*\\*\\*\\*\\*\\*")[[1]][1]
                        
                        paste0(
                                "<h3>", destination_name, "</h3><br/>",
                                destination_zip
                        ) %>% HTML
                }
                
                
                
        })
        
        output$return_consolidation_map_table <- renderDT({
                
                return_consolidation_map_data <- return_consolidation_map_data()
                if (is.null(return_consolidation_map_data)) return(NULL)
                
                return_consolidation_map_data %>% 
                        select(CREATED, `NSD #`, CLIENT, ITEM_COUNT, WEIGHT, `LTL READY`) %>% 
                        datatable(
                                options = list(
                                        dom = "tp",
                                        pageLength = 10,
                                        searchHighlight = TRUE
                                ),
                                filter = 'top', 
                                rownames = FALSE
                        )
                
        })
        
        
        
        
        # ICON --------------------------------------------------------------------
        
        return_icon_emails_data <- reactive({
                TMdata <- TM()
                if (is.null(TMdata)) return(NULL)
                if (sum(c("NSD #", "PO#", "SHIPPER", "CONSIGNEE", "RMA/RGA_NUMBER", "AGENT") %in% names(TMdata)) != 6) {
                        showModal(modalDialog(
                                title = "CAUTION!",
                                "Incorrect input file!",
                                easyClose = TRUE,
                                footer = NULL
                        ))
                        return(NULL)
                }
                return_icon_emails_data <-
                        TMdata %>% 
                        filter(AM == "MMASON", 
                               `OP STATUS` %in% c("LTLAVAIL", "PICKEDUP")
                        ) %>% 
                        select(`NSD #`, `PO#`, SHIPPER, CONSIGNEE, `RMA/RGA_NUMBER`, AGENT)
        })
        
        observe({
                if(is.null(input$return_icon_emails_send_button) || input$return_icon_emails_send_button == 0) return(NULL)
                if (is.null(authorization.check(ICON_emails_authorized_people))) return(NULL)
                
                data <- isolate(return_icon_emails_data())
                if (is.null(data)) {
                        showModal(modalDialog(
                                title = "CAUTION!",
                                "No data imported!",
                                easyClose = TRUE,
                                footer = NULL
                        ))
                        return(NULL)
                }
                
                agents_names <- (data %>% pull(AGENT) %>% unique)[1:2]
                
                signature <- nsd_signature(ICON_emails_authorized_people[1])
                
                OutApp <- COMCreate("Outlook.Application")
                
                withProgress(message = 'Sending E-mails in progress', value = 0, {
                        for (agent in agents_names) {
                                incProgress(1/length(agents_names))
                                data <- data %>% filter(AGENT==agent) %>% select(-AGENT)
                                
                                outMail = OutApp$CreateItem(0)
                                outMail[["To"]] = email_address(
                                        paste0("AG_", agent,"@nonstopdelivery.com"),
                                        isolate(input$Test_Active)
                                )
                                outMail[["subject"]] = paste0("ICON Orders - ", agent)
                                outMail[["HTMLBody"]] = 
                                        paste0(
                                                "<p>Hello,</p>",
                                                "<p>Below Icon orders are showing as on your dock and we would like to recover them as soon as possible. Please advise which exercise machines are in original box and which ones are out of box. Once information is provided I will send you BOLs and advise you of the day when carrier comes in to recover. If any of orders are disposal, I will request disposition form and it will be provided to you by our OS&D department within 48 hours.</p>",
                                                data %>% 
                                                        htmlTable(
                                                                align="l", 
                                                                css.cell = "padding-left: 1em; padding-right: 3em;",
                                                                css.tspanner = "font-weight: 900; text-align: left;"
                                                        ), 
                                                signature
                                        )
                                
                                
                                
                                
                                
                                
                                
                                
                                # outMail[["attachments"]]$Add("C:/Path/To/The/Attachment/File.ext")
                                outMail$Send()  
                                
                        }
                })
                
                
        })
        
        # PDD ---------------------------------------------------------------------
        
        
        pdd_input_data <- reactive({
                
                pdd_input <- input$delivery_amazon_pdd_input
                if (is.null(pdd_input)) return(NULL)
                
                pdd_input_data <- 
                        read_excel(pdd_input$datapath) %>% 
                        rename(`NSD#` = `Basic Order Info NSD Number`, 
                               AGENT = `Basic Order Info Agent Code`
                        )
        })
        
        observe({
                template <- input$delivery_amazon_select_template
                updateTextInput(
                        session, 
                        "delivery_amazon_pdd_subject", 
                        value = amazon_email_item(template, "subject")
                )
                updateTextInput(
                        session, 
                        "delivery_amazon_pdd_body", 
                        value = amazon_email_item(template, "body")
                )
        })
        
        observeEvent(input$delivery_amazon_pdd_subject_save, {
                
                if (is.null(authorization.check(Amazon_team))) return(NULL)
                
                template <- isolate(input$delivery_amazon_select_template)
                
                writeChar(
                        isolate(input$delivery_amazon_pdd_subject), 
                        amazon_email_item_path(template, "subject")
                )
                
                shinyalert(
                        title = "Saved",
                        text = "Subject",
                        closeOnEsc = TRUE,
                        closeOnClickOutside = TRUE,
                        html = FALSE,
                        type = "success",
                        showConfirmButton = TRUE,
                        showCancelButton = FALSE,
                        confirmButtonText = "OK",
                        confirmButtonCol = "#a2a8ab",
                        timer = 0,
                        imageUrl = "",
                        animation = TRUE
                )
        })
        
        observeEvent(input$delivery_amazon_pdd_body_save, {
                
                if (is.null(authorization.check(Amazon_team))) return(NULL)
                
                template <- isolate(input$delivery_amazon_select_template)
                
                
                writeChar(
                        isolate(input$delivery_amazon_pdd_body), 
                        amazon_email_item_path(template, "body")
                )
                
                shinyalert(
                        title = "Saved",
                        text = "Body",
                        closeOnEsc = TRUE,
                        closeOnClickOutside = TRUE,
                        html = FALSE,
                        type = "success",
                        showConfirmButton = TRUE,
                        showCancelButton = FALSE,
                        confirmButtonText = "OK",
                        confirmButtonCol = "#a2a8ab",
                        timer = 0,
                        imageUrl = "",
                        animation = TRUE
                )
                
        })
        
        
        observeEvent(input$delivery_amazon_send_button, {
                
                if (is.null(authorization.check(Amazon_team))) return(NULL)
                
                data_all_agents <- isolate(pdd_input_data())
                if (is.null(data_all_agents)) {
                        showModal(modalDialog(
                                title = "CAUTION!",
                                "No data imported!",
                                easyClose = TRUE,
                                footer = NULL
                        ))
                        return(NULL)
                }
                
                agents_names <- 
                        data_all_agents %>% 
                        pull(AGENT) %>% 
                        unique
                
                pdd_subject <- isolate(input$delivery_amazon_pdd_subject) %>% paste0()
                pdd_body <- paste(paste0('<p>', strsplit(isolate(input$delivery_amazon_pdd_body),"\n")[[1]], "</p>"), collapse="")
                
                OutApp <- COMCreate("Outlook.Application")
                withProgress(message = 'Sending E-mails in progress', value = 0, {
                        for (agent in agents_names) {
                                incProgress(1/length(agents_names))
                                data <- data_all_agents %>% filter(AGENT == agent) %>% select(`NSD#`)
                                pdd_subject <- isolate(input$delivery_amazon_pdd_subject) %>% paste0(agent)
                                
                                outMail = OutApp$CreateItem(0)
                                outMail[["To"]] = email_address(
                                        paste0(
                                                if (agent == "CLAUDIO") {""} else paste0("AG_", agent,"@nonstopdelivery.com;"), 
                                                Amazon_Additiona_Email_from_Agent(agent)
                                        ),
                                        isolate(input$Test_Active)
                                )
                                outMail[["CC"]] = email_address(
                                        "nsdamazon@shipnsd.com",
                                        isolate(input$Test_Active)
                                )
                                outMail[["subject"]] = pdd_subject
                                
                                outMail[["HTMLBody"]] = 
                                        paste0(
                                                pdd_body,
                                                data %>% 
                                                        htmlTable(
                                                                align="l",
                                                                align.header = "l",
                                                                css.cell = "padding-left: 1em; padding-right: 1em;",
                                                                css.tspanner = "font-weight: 900; text-align: left;",
                                                                css.table = "margin-top: .5em; margin-bottom: .5em; margin: auto; padding: 0px; width: 100%"
                                                        ), 
                                                nsd_signature(Sys.info()["user"][[1]])
                                        )
                                # outMail[["attachments"]]$Add("C:/Path/To/The/Attachment/File.ext")
                                outMail$Send()  
                                
                        }
                })
                
                
        })
        
        # Daily Amazon ------------------------------------------------------------
        output$delivery_amazon_create_east_coast_button <- downloadHandler(
                filename = function() {
                        paste("AMAZON_EAST_COAST.csv", sep = "")
                },
                content = function(file) {
                        
                        TM_input <- input$delivery_amazon_daily_TM_input
                        files_input <- input$delivery_amazon_daily_raw_data_input
                        
                        if(is.null(TM_input)) {
                                shinyalert(
                                        title = "Oops!",
                                        text = "Missing TM file",
                                        closeOnEsc = TRUE,
                                        closeOnClickOutside = TRUE,
                                        html = FALSE,
                                        type = "error",
                                        showConfirmButton = TRUE,
                                        showCancelButton = FALSE,
                                        confirmButtonText = "OK",
                                        confirmButtonCol = "#a2a8ab",
                                        timer = 0,
                                        imageUrl = "",
                                        animation = TRUE
                                )
                                return(NULL)
                        }
                        
                        if(nrow(TM_input) > 1) {
                                shinyalert(
                                        title = "Oops!",
                                        text = "Multiple TM files not allowed!",
                                        closeOnEsc = TRUE,
                                        closeOnClickOutside = TRUE,
                                        html = FALSE,
                                        type = "error",
                                        showConfirmButton = TRUE,
                                        showCancelButton = FALSE,
                                        confirmButtonText = "OK",
                                        confirmButtonCol = "#a2a8ab",
                                        timer = 0,
                                        imageUrl = "",
                                        animation = TRUE
                                )
                                return(NULL)
                        }
                        
                        
                        TMdata <- read.csv(TM_input %>% pull(datapath), stringsAsFactors = FALSE)
                        
                        if(is.null(files_input) ) {
                                shinyalert(
                                        title = "Oops!",
                                        text = "Missing input file(s)",
                                        closeOnEsc = TRUE,
                                        closeOnClickOutside = TRUE,
                                        html = FALSE,
                                        type = "error",
                                        showConfirmButton = TRUE,
                                        showCancelButton = FALSE,
                                        confirmButtonText = "OK",
                                        confirmButtonCol = "#a2a8ab",
                                        timer = 0,
                                        imageUrl = "",
                                        animation = TRUE
                                )
                                return(NULL)
                        }
                        
                        if (sum((files_input$name) %in% paste0(amazon_east_coast, ".csv")) == 0) {
                                
                                shinyalert(
                                        title = "Oops!",
                                        text = "No East Coast terminal found in the inputs",
                                        closeOnEsc = TRUE,
                                        closeOnClickOutside = TRUE,
                                        html = FALSE,
                                        type = "error",
                                        showConfirmButton = TRUE,
                                        showCancelButton = FALSE,
                                        confirmButtonText = "OK",
                                        confirmButtonCol = "#a2a8ab",
                                        timer = 0,
                                        imageUrl = "",
                                        animation = TRUE
                                )
                                
                                return(NULL)
                                
                        } 
                        
                        if (sum((files_input$name) %in% paste0(amazon_east_coast, ".csv")) < length(files_input$name)) {
                                wrong_files <- setdiff(files_input$name, paste0(amazon_east_coast, ".csv"))
                                shinyalert(
                                        title = "Oops!",
                                        text = "The following files are being ignored because they are not in the list of the East-Coast terminals or in csv format: " %>% paste0(paste(wrong_files, collapse = ",")),
                                        closeOnEsc = TRUE,
                                        closeOnClickOutside = TRUE,
                                        html = FALSE,
                                        type = "error",
                                        showConfirmButton = TRUE,
                                        showCancelButton = FALSE,
                                        confirmButtonText = "OK",
                                        confirmButtonCol = "#a2a8ab",
                                        timer = 0,
                                        imageUrl = "",
                                        animation = TRUE
                                )
                                
                        }
                        
                        
                        
                        inputs <-
                                files_input %>%
                                filter(name %in% paste0(amazon_east_coast, ".csv")) %>%
                                mutate(., name = name %>% str_replace(".csv$", ""))
                        
                        output <- data.frame(Date = as.Date(character()),
                                             Location = character(),
                                             `Tracking ID` = character(),
                                             `Delivery Associate` = character(),
                                             `Accurate Conf#` = character(),
                                             `DA Job` = character(),
                                             stringsAsFactors=FALSE)
                        
                        
                        
                        
                        ############                      EAST COAST                ###############
                        
                        
                        
                        
                        if (sum(grepl("MIAMI_FL", inputs$name)) == 1){
                                output <-
                                        read.csv(inputs %>% filter(name == "MIAMI_FL") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F) %>%
                                        select(DueTimeFrom, Reference2, `Driver NAme`, UserData1, UserData2, UserData3) %>%
                                        dplyr::rename(`Tracking ID` = Reference2,
                                                      Date = DueTimeFrom,
                                                      Driver = `Driver NAme`,
                                                      `Driver Accurate Conf#` = UserData1,
                                                      Helper = UserData2,
                                                      `Helper Accurate Conf#` = UserData3) %>%
                                        mutate(Date = date(mdy_hm(Date)),
                                               Driver = gsub(", ", " ", Driver)
                                        ) %>%
                                        Assistant_Process %>%
                                        Finalize("MIAMI,FL") %>%
                                        rbind(output)
                                
                        }
                        if (sum(grepl("CHANTILLY_VA", inputs$name))==1){
                                output <-
                                        read.csv(inputs %>% filter(name == "CHANTILLY_VA") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F) %>%
                                        as_tibble %>%
                                        select(Date, Account, Driver, `Sale Order`) %>%
                                        filter(Account == "AMAZON") %>%
                                        select(-Account) %>%
                                        cbind(str_split_fixed(.$Driver, ", ", 2) %>%
                                                      as_tibble %>%
                                                      dplyr::rename(Driver1 = V1, `Helper1` = V2)
                                        ) %>%
                                        cbind(str_split_fixed(.$Driver1, " - ", 2) %>%
                                                      as_tibble %>%
                                                      dplyr::rename(Driver2 = V1, `Driver Accurate Conf#` = V2)
                                        ) %>%
                                        select(-Driver, -Driver1) %>%
                                        dplyr::rename(Driver = Driver2, `Tracking ID` = `Sale Order`) %>%
                                        cbind(str_split_fixed(.$Helper1, " - ", 2) %>%
                                                      as_tibble %>%
                                                      dplyr::rename(Helper = V1, `Helper Accurate Conf#` = V2)
                                        ) %>%
                                        select(-Helper1) %>%
                                        mutate(Date = mdy(Date)) %>%
                                        Assistant_Process %>%
                                        Finalize("CHANTILLY,VA") %>%
                                        rbind(output)
                        }
                        if (sum(grepl("PHILADELPHIA_PA", inputs$name))==1){
                                output <-
                                        read.csv(inputs %>% filter(name == "PHILADELPHIA_PA") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F) %>%
                                        select(Date, Driver, Helper, `Sale Order`) %>%
                                        cbind(str_split_fixed(.$Driver, " - ", 2) %>%
                                                      as.data.frame %>%
                                                      dplyr::rename(Driver1 = V1, `Driver Accurate Conf#` = V2)
                                        ) %>%
                                        cbind(str_split_fixed(.$Helper, " - ", 2) %>%
                                                      as.data.frame %>%
                                                      dplyr::rename(Helper1 = V1, `Helper Accurate Conf#` = V2)
                                        ) %>%
                                        dplyr::rename(`Tracking ID` = `Sale Order`) %>%
                                        select(Date, `Tracking ID`, Driver1, `Driver Accurate Conf#`, Helper1, `Helper Accurate Conf#`) %>%
                                        dplyr::rename(Driver = Driver1, Helper = Helper1) %>%
                                        mutate(Date = mdy(Date)) %>%
                                        Assistant_Process %>%
                                        Finalize("PHILADELPHIA,PA") %>%
                                        rbind(output)
                        }
                        if (sum(grepl("CHICAGO_IL", inputs$name))==1){
                                output <-
                                        read.csv(inputs %>% filter(name == "CHICAGO_IL") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F) %>%
                                        select(Date, `NSD Job #`, `Driver Name/Background Check #`, `Helper Name / Background Check #`) %>%
                                        cbind(str_split_fixed(.$`Driver Name/Background Check #`, "/", 2) %>%
                                                      as.data.frame %>%
                                                      dplyr::rename(Driver1 = V1, `Driver Accurate Conf#` = V2)
                                        ) %>%
                                        cbind(str_split_fixed(.$`Helper Name / Background Check #`, "/", 2) %>%
                                                      as.data.frame %>%
                                                      dplyr::rename(Helper1 = V1, `Helper Accurate Conf#` = V2)
                                        ) %>%
                                        dplyr::rename(`Tracking ID` = `NSD Job #`) %>%
                                        select(Date, `Tracking ID`, Driver1, `Driver Accurate Conf#`, Helper1, `Helper Accurate Conf#`) %>%
                                        dplyr::rename(Driver = Driver1, Helper = Helper1) %>%
                                        Assistant_Process %>%
                                        Finalize("CHICAGO,IL") %>%
                                        rbind(output)
                        }
                        if (sum(grepl("SALT LAKE CITY_UT", inputs$name))==1){
                                output <-
                                        read.csv(inputs %>% filter(name == "SALT LAKE CITY_UT") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F) %>%
                                        select(Date, `Sale Order`, Driver, `Driver Number`) %>%
                                        cbind(str_split_fixed(.$`Driver Number`, ", ", 2) %>%
                                                      as.data.frame %>%
                                                      dplyr::rename(`Driver Accurate Conf#` = V1, `Helper Accurate Conf#` = V2)
                                        ) %>%
                                        cbind(str_split_fixed(.$Driver, ", ", 2) %>%
                                                      as.data.frame %>%
                                                      dplyr::rename(Driver1 = V1, Helper1 = V2)
                                        ) %>%
                                        cbind(str_split_fixed(.$Helper1, " [(]", 2) %>%
                                                      as.data.frame %>%
                                                      dplyr::rename(Helper = V1, Helper2 = V2)
                                        ) %>%
                                        mutate(`Tracking ID` = stringr::str_extract(`Sale Order`, "\\d{7}")) %>%
                                        select(Date, `Tracking ID`, Driver1, `Driver Accurate Conf#`, Helper, `Helper Accurate Conf#`) %>%
                                        rename(Driver = Driver1) %>%
                                        Assistant_Process %>%
                                        Finalize("SALT LAKE CITY,UT") %>%
                                        rbind(output)
                        }
                        if (sum(grepl("BELTSVILLE_MD", inputs$name)) == 1){
                                Temp <- read.csv(inputs %>% filter(name == "BELTSVILLE_MD") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F)
                                names(Temp) <- c("Date", "Location", "Tracking ID", "Delivery Associate_1", "Accurate Conf#_1", "DA Job_1", "Delivery Associate_2", "Accurate Conf#_2", "DA Job_2")
                                Temp <- Temp %>% filter(`Tracking ID` != "", Date != "Date")
                                Driver <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_1", `Accurate Conf#` = "Accurate Conf#_1", `DA Job` = "DA Job_1")
                                Helper <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_2", `Accurate Conf#` = "Accurate Conf#_2", `DA Job` = "DA Job_2")
                                output <- rbind(Driver, Helper) %>%
                                        mutate(Date = mdy(Date)) %>%
                                        Finalize("BELTSVILLE,MD") %>%
                                        rbind(output)
                                rm("Driver", "Helper", "Temp")
                        }
                        if (sum(grepl("LORIS_SC", inputs$name))==1){
                                Temp <- read.csv(inputs %>% filter(name == "LORIS_SC") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F)
                                names(Temp) <- c("Date", "Location", "Tracking ID", "Time Frame", "Delivery Associate_1", "Accurate Conf#_1", "DA Job_1", "Delivery Associate_2", "Accurate Conf#_2", "DA Job_2")
                                Temp <- Temp %>% filter(`Tracking ID` != "", Date != "Date")
                                Driver <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_1", `Accurate Conf#` = "Accurate Conf#_1", `DA Job` = "DA Job_1")
                                Helper <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_2", `Accurate Conf#` = "Accurate Conf#_2", `DA Job` = "DA Job_2")
                                output <- rbind(Driver, Helper) %>%
                                        mutate(Date = mdy(Date)) %>%
                                        Finalize("LORIS,SC") %>%
                                        rbind(output)
                                rm("Driver", "Helper", "Temp")
                        }
                        if (sum(grepl("ORLANDO_FL", inputs$name))==1){
                                Temp <- read.csv(inputs %>% filter(name == "ORLANDO_FL") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F)
                                names(Temp) <- c("Date", "Location", "Tracking ID", "Delivery Associate_1", "Accurate Conf#_1", "DA Job_1", "Delivery Associate_2", "Accurate Conf#_2", "DA Job_2")
                                Temp <- Temp %>% filter(`Tracking ID` != "", Date != "Date")
                                Driver <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_1", `Accurate Conf#` = "Accurate Conf#_1", `DA Job` = "DA Job_1")
                                Helper <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_2", `Accurate Conf#` = "Accurate Conf#_2", `DA Job` = "DA Job_2")
                                output <- rbind(Driver, Helper) %>%
                                        mutate(Date = mdy(Date)) %>%
                                        Finalize("ORLANDO,FL") %>%
                                        rbind(output)
                                rm("Driver", "Helper", "Temp")
                        }
                        if (sum(grepl("TAMPA_FL", inputs$name))==1){
                                Temp <- read.csv(inputs %>% filter(name == "TAMPA_FL") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F)
                                names(Temp) <- c("Date", "Location", "Tracking ID", "Delivery Associate_1", "Accurate Conf#_1", "DA Job_1", "Delivery Associate_2", "Accurate Conf#_2", "DA Job_2")
                                Temp <- Temp %>% filter(`Tracking ID` != "", Date != "Date")
                                Driver <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_1", `Accurate Conf#` = "Accurate Conf#_1", `DA Job` = "DA Job_1")
                                Helper <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_2", `Accurate Conf#` = "Accurate Conf#_2", `DA Job` = "DA Job_2")
                                output <- rbind(Driver, Helper) %>%
                                        mutate(Date = mdy(Date)) %>%
                                        cbind(str_split_fixed(.$`Delivery Associate`, ", ", 2) %>%
                                                      as_tibble %>%
                                                      dplyr::rename(`First_Name` = V1, `Last_Name` = V2)
                                        ) %>%
                                        mutate(`Delivery Associate` = paste(First_Name, Last_Name)) %>%
                                        select(-First_Name, -Last_Name) %>%
                                        Finalize("TAMPA,FL") %>%
                                        rbind(output)
                                rm("Driver", "Helper", "Temp")
                        }
                        if (sum(grepl("CLINTON_PA", inputs$name))==1){
                                Temp <- read.csv(inputs %>% filter(name == "CLINTON_PA") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F)
                                names(Temp) <- c("Date", "Location", "Tracking ID", "Delivery Associate_1", "Accurate Conf#_1", "DA Job_1", "Delivery Associate_2", "Accurate Conf#_2", "DA Job_2")
                                Temp <- Temp %>% filter(`Tracking ID` != "", Date != "Date")
                                Driver <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_1", `Accurate Conf#` = "Accurate Conf#_1", `DA Job` = "DA Job_1")
                                Helper <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_2", `Accurate Conf#` = "Accurate Conf#_2", `DA Job` = "DA Job_2")
                                output <- rbind(Driver, Helper) %>%
                                        mutate(Date = mdy(Date)) %>%
                                        cbind(str_split_fixed(.$`Delivery Associate`, ", ", 2) %>%
                                                      as_tibble %>%
                                                      dplyr::rename(`First_Name` = V1, `Last_Name` = V2)
                                        ) %>%
                                        mutate(`Delivery Associate` = paste(First_Name, Last_Name)) %>%
                                        select(-First_Name, -Last_Name) %>%
                                        Finalize("CLINTON,PA") %>%
                                        rbind(output)
                                rm("Driver", "Helper", "Temp")
                        }
                        if (sum(grepl("DURHAM_NC", inputs$name))==1){
                                Temp <- read.csv(inputs %>% filter(name == "DURHAM_NC") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F)
                                names(Temp) <- c("Date", "Location", "Tracking ID", "Delivery Associate_1", "Accurate Conf#_1", "DA Job_1", "Delivery Associate_2", "Accurate Conf#_2", "DA Job_2")
                                Temp <- Temp %>% filter(`Tracking ID` != "", Date != "Date")
                                Driver <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_1", `Accurate Conf#` = "Accurate Conf#_1", `DA Job` = "DA Job_1")
                                Helper <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_2", `Accurate Conf#` = "Accurate Conf#_2", `DA Job` = "DA Job_2")
                                output <- rbind(Driver, Helper) %>%
                                        mutate(Date = mdy(Date)) %>%
                                        cbind(str_split_fixed(.$`Delivery Associate`, ", ", 2) %>%
                                                      as_tibble %>%
                                                      dplyr::rename(`First_Name` = V1, `Last_Name` = V2)
                                        ) %>%
                                        mutate(`Delivery Associate` = paste(First_Name, Last_Name)) %>%
                                        select(-First_Name, -Last_Name) %>%
                                        Finalize("DURHAM,NC") %>%
                                        rbind(output)
                                rm("Driver", "Helper", "Temp")
                        }
                        if (sum(grepl("BIRMINGHAM_AL", inputs$name))==1){
                                Temp <- read.csv(inputs %>% filter(name == "BIRMINGHAM_AL") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F)
                                names(Temp) <- c("Date", "Location", "Tracking ID", "Delivery Associate_1", "Accurate Conf#_1", "DA Job_1", "Delivery Associate_2", "Accurate Conf#_2", "DA Job_2")
                                Temp <- Temp %>% filter(`Tracking ID` != "", Date != "Date")
                                Driver <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_1", `Accurate Conf#` = "Accurate Conf#_1", `DA Job` = "DA Job_1")
                                Helper <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_2", `Accurate Conf#` = "Accurate Conf#_2", `DA Job` = "DA Job_2")
                                output <- rbind(Driver, Helper) %>%
                                        mutate(Date = mdy(Date)) %>%
                                        cbind(str_split_fixed(.$`Delivery Associate`, ", ", 2) %>%
                                                      as_tibble %>%
                                                      dplyr::rename(`First_Name` = V1, `Last_Name` = V2)
                                        ) %>%
                                        mutate(`Delivery Associate` = paste(First_Name, Last_Name)) %>%
                                        select(-First_Name, -Last_Name) %>%
                                        Finalize("BIRMINGHAM,AL") %>%
                                        rbind(output)
                                rm("Driver", "Helper", "Temp")
                        }
                        if (sum(grepl("JACKSONVILLE_FL", inputs$name))==1){
                                Temp <- read.csv(inputs %>% filter(name == "JACKSONVILLE_FL") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F)
                                names(Temp) <- c("Date", "Location", "Tracking ID", "Delivery Associate_1", "Accurate Conf#_1", "DA Job_1", "Delivery Associate_2", "Accurate Conf#_2", "DA Job_2")
                                Temp <- Temp %>% filter(`Tracking ID` != "", Date != "Date")
                                Driver <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_1", `Accurate Conf#` = "Accurate Conf#_1", `DA Job` = "DA Job_1")
                                Helper <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_2", `Accurate Conf#` = "Accurate Conf#_2", `DA Job` = "DA Job_2")
                                output <- rbind(Driver, Helper) %>%
                                        mutate(Date = mdy(Date)) %>%
                                        cbind(str_split_fixed(.$`Delivery Associate`, ", ", 2) %>%
                                                      as_tibble %>%
                                                      dplyr::rename(`First_Name` = V1, `Last_Name` = V2)
                                        ) %>%
                                        mutate(`Delivery Associate` = paste(First_Name, Last_Name)) %>%
                                        select(-First_Name, -Last_Name) %>%
                                        Finalize("JACKSONVILLE,FL") %>%
                                        rbind(output)
                                rm("Driver", "Helper", "Temp")
                        }
                        if (sum(grepl("OWENSBORO_KY", inputs$name))==1){
                                Temp <- read.csv(inputs %>% filter(name == "OWENSBORO_KY") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F)
                                names(Temp) <- c("Date", "Location", "Tracking ID", "Delivery Associate_1", "Accurate Conf#_1", "DA Job_1", "Delivery Associate_2", "Accurate Conf#_2", "DA Job_2")
                                Temp <- Temp %>% filter(`Tracking ID` != "", Date != "Date")
                                Driver <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_1", `Accurate Conf#` = "Accurate Conf#_1", `DA Job` = "DA Job_1")
                                Helper <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_2", `Accurate Conf#` = "Accurate Conf#_2", `DA Job` = "DA Job_2")
                                output <- rbind(Driver, Helper) %>%
                                        mutate(Date = mdy(Date)) %>%
                                        cbind(str_split_fixed(.$`Delivery Associate`, ", ", 2) %>%
                                                      as_tibble %>%
                                                      dplyr::rename(`First_Name` = V1, `Last_Name` = V2)
                                        ) %>%
                                        mutate(`Delivery Associate` = paste(First_Name, Last_Name)) %>%
                                        select(-First_Name, -Last_Name) %>%
                                        Finalize("OWENSBORO,KY") %>%
                                        rbind(output)
                                rm("Driver", "Helper", "Temp")
                        }
                        if (sum(grepl("WILLISTON_VT", inputs$name))==1){
                                Temp <- read.csv(inputs %>% filter(name == "WILLISTON_VT") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F)
                                names(Temp) <- c("Date", "Location", "Tracking ID", "Delivery Associate_1", "Accurate Conf#_1", "DA Job_1", "Delivery Associate_2", "Accurate Conf#_2", "DA Job_2")
                                Temp <- Temp %>% filter(`Tracking ID` != "", Date != "Date")
                                Driver <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_1", `Accurate Conf#` = "Accurate Conf#_1", `DA Job` = "DA Job_1")
                                Helper <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_2", `Accurate Conf#` = "Accurate Conf#_2", `DA Job` = "DA Job_2")
                                output <- rbind(Driver, Helper) %>%
                                        mutate(Date = mdy(Date)) %>%
                                        cbind(str_split_fixed(.$`Delivery Associate`, ", ", 2) %>%
                                                      as_tibble %>%
                                                      dplyr::rename(`First_Name` = V1, `Last_Name` = V2)
                                        ) %>%
                                        mutate(`Delivery Associate` = paste(First_Name, Last_Name)) %>%
                                        select(-First_Name, -Last_Name) %>%
                                        Finalize("WILLISTON,VT") %>%
                                        rbind(output)
                                rm("Driver", "Helper", "Temp")
                        }
                        
                        if (nrow(output) > 0) {
                                output <-
                                        output %>%
                                        merge(TMdata %>% select(NSD.., PO.), all.x = T, by.x = "Tracking ID", by.y = "NSD..") %>%
                                        mutate(
                                                `Delivery Associate` = as.character(`Delivery Associate`),
                                                `Accurate Conf#` = as.character(`Accurate Conf#`),
                                                `Delivery Associate` = ifelse(`Delivery Associate` == "Claudio10", "CLAUDIO MONASTERIO", `Delivery Associate`),
                                                `Accurate Conf#` = ifelse(`Delivery Associate` %in% c("Claudio10", "CLAUDIO MONASTERIO"), "82698909", `Accurate Conf#`),
                                                `Delivery Associate` = ifelse(`Delivery Associate` == "Adrian Suarez  TRK 10 - 24' BOX", "Adrian Suarez", `Delivery Associate`),
                                                `Accurate Conf#` = ifelse(`Delivery Associate` %in% c("Adrian Suarez  TRK 10 - 24' BOX", "Adrian Suarez"), "98025517", `Accurate Conf#`)
                                        ) %>%
                                        rename(`Encrypted ID` = PO.) %>%
                                        mutate(
                                                `Route ID` = NA,
                                                `Accurate Account Owner` = NA,
                                                `Delivery Associate` = toupper(`Delivery Associate`),
                                                `DA Job` = toupper(`DA Job`)
                                        ) %>%
                                        select(
                                                `accurate_account_owner` = `Accurate Account Owner`,
                                                `confirmation_number` = `Accurate Conf#`,
                                                `delivery_associate` = `Delivery Associate`,
                                                `delivery_associate_role` = `DA Job`,
                                                `delivery_date` = Date,
                                                `encrypted_id` = `Encrypted ID`,
                                                location = Location,
                                                route_id = `Route ID`,
                                                tracking_id = `Tracking ID`
                                        ) %>%
                                        arrange(desc(delivery_date), location) %>%
                                        filter(!is.na(confirmation_number), confirmation_number != "") %>%
                                        unique %>% 
                                        write.csv(file, row.names = F)
                        }
                        
                }
        )
        
        
        
        output$delivery_amazon_create_west_coast_button <- downloadHandler(
                filename = function() {
                        paste("AMAZON_WEST_COAST.csv", sep = "")
                },
                content = function(file) {
                        
                        TM_input <- input$delivery_amazon_daily_TM_input
                        files_input <- input$delivery_amazon_daily_raw_data_input
                        
                        if(is.null(TM_input)) {
                                shinyalert(
                                        title = "Oops!",
                                        text = "Missing TM file",
                                        closeOnEsc = TRUE,
                                        closeOnClickOutside = TRUE,
                                        html = FALSE,
                                        type = "error",
                                        showConfirmButton = TRUE,
                                        showCancelButton = FALSE,
                                        confirmButtonText = "OK",
                                        confirmButtonCol = "#a2a8ab",
                                        timer = 0,
                                        imageUrl = "",
                                        animation = TRUE
                                )
                                return(NULL)
                        }
                        
                        if(nrow(TM_input) > 1) {
                                shinyalert(
                                        title = "Oops!",
                                        text = "Multiple TM files not allowed!",
                                        closeOnEsc = TRUE,
                                        closeOnClickOutside = TRUE,
                                        html = FALSE,
                                        type = "error",
                                        showConfirmButton = TRUE,
                                        showCancelButton = FALSE,
                                        confirmButtonText = "OK",
                                        confirmButtonCol = "#a2a8ab",
                                        timer = 0,
                                        imageUrl = "",
                                        animation = TRUE
                                )
                                return(NULL)
                        }
                        
                        
                        TMdata <- read.csv(TM_input %>% pull(datapath), stringsAsFactors = FALSE)
                        
                        if(is.null(files_input) ) {
                                shinyalert(
                                        title = "Oops!",
                                        text = "Missing input file(s)",
                                        closeOnEsc = TRUE,
                                        closeOnClickOutside = TRUE,
                                        html = FALSE,
                                        type = "error",
                                        showConfirmButton = TRUE,
                                        showCancelButton = FALSE,
                                        confirmButtonText = "OK",
                                        confirmButtonCol = "#a2a8ab",
                                        timer = 0,
                                        imageUrl = "",
                                        animation = TRUE
                                )
                                return(NULL)
                        }
                        
                        if (sum((files_input$name) %in% paste0(amazon_west_coast, ".csv")) == 0) {
                                
                                shinyalert(
                                        title = "Oops!",
                                        text = "No West Coast terminal found in the inputs",
                                        closeOnEsc = TRUE,
                                        closeOnClickOutside = TRUE,
                                        html = FALSE,
                                        type = "error",
                                        showConfirmButton = TRUE,
                                        showCancelButton = FALSE,
                                        confirmButtonText = "OK",
                                        confirmButtonCol = "#a2a8ab",
                                        timer = 0,
                                        imageUrl = "",
                                        animation = TRUE
                                )
                                
                                return(NULL)
                                
                        } 
                        
                        if (sum((files_input$name) %in% paste0(amazon_west_coast, ".csv")) < length(files_input$name)) {
                                wrong_files <- setdiff(files_input$name, paste0(amazon_west_coast, ".csv"))
                                shinyalert(
                                        title = "Oops!",
                                        text = "The following files are being ignored because they are not in the list of the West-Coast terminals or in csv format: " %>% paste0(paste(wrong_files, collapse = ",")),
                                        closeOnEsc = TRUE,
                                        closeOnClickOutside = TRUE,
                                        html = FALSE,
                                        type = "error",
                                        showConfirmButton = TRUE,
                                        showCancelButton = FALSE,
                                        confirmButtonText = "OK",
                                        confirmButtonCol = "#a2a8ab",
                                        timer = 0,
                                        imageUrl = "",
                                        animation = TRUE
                                )
                                
                        }
                        
                        
                        
                        inputs <-
                                files_input %>%
                                filter(name %in% paste0(amazon_west_coast, ".csv")) %>%
                                mutate(., name = name %>% str_replace(".csv$", ""))
                        
                        output <- data.frame(Date = as.Date(character()),
                                             Location = character(),
                                             `Tracking ID` = character(),
                                             `Delivery Associate` = character(),
                                             `Accurate Conf#` = character(),
                                             `DA Job` = character(),
                                             stringsAsFactors=FALSE)
                        
                        
                        
                        
                        ############                      WEST COAST                ###############
                        
                        output <- data.frame(Date = as.Date(character()),
                                             Location = character(), 
                                             `Tracking ID` = character(),
                                             `Delivery Associate` = character(),
                                             `Accurate Conf#` = character(),
                                             `DA Job` = character(),
                                             stringsAsFactors=FALSE) 
                        
                        
                        
                        if (sum(grepl("DES MOINES_IA", inputs$name))==1){
                                Temp <- read.csv(inputs %>% filter(name == "DES MOINES_IA") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F)
                                Temp <- Temp[,1:6]
                                output <- 
                                        Temp %>% 
                                        filter(!is.na(`NSD Freight Bill #`)) %>%
                                        cbind(str_split_fixed(.$`Driver Name`, ", ", 2) %>%
                                                      as.data.frame %>%
                                                      dplyr::rename(`Driver Last Name` = V1, `Driver First Name` = V2) %>%
                                                      mutate(Driver = paste0(`Driver First Name`, " ", `Driver Last Name`)) %>%
                                                      select(Driver)
                                        ) %>%
                                        cbind(str_split_fixed(.$`Driver/ Helper Name`, ", ", 2) %>%
                                                      as.data.frame %>%
                                                      dplyr::rename(`Helper Last Name` = V1, `Helper First Name` = V2) %>%
                                                      mutate(Helper = paste0(`Helper First Name`, " ", `Helper Last Name`)) %>%
                                                      select(Helper)
                                        ) %>%
                                        select(Date,
                                               `Tracking ID` = `NSD Freight Bill #`,
                                               Driver,
                                               `Driver Accurate Conf#` = `Integral Accurate Driver #`,
                                               Helper,
                                               `Helper Accurate Conf#` = `Integral Accurate Driver/ Helper #`) %>%
                                        mutate(
                                                Date = mdy(Date),
                                                `Tracking ID` = as.character(`Tracking ID`),
                                                `Driver Accurate Conf#` = as.character(`Driver Accurate Conf#`),
                                                `Helper Accurate Conf#` = as.character(`Helper Accurate Conf#`)
                                        ) %>%
                                        Assistant_Process %>%
                                        Finalize("DES MOINES,IA") %>%
                                        rbind(output)
                        }
                        if (sum(grepl("PORTLAND_OR", inputs$name))==1){
                                output <-
                                        read.csv(inputs %>% filter(name == "PORTLAND_OR") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE) %>%
                                        dplyr::rename(Date = `Date.`, `Tracking ID` = `Order...`, Driver = `Driver.`, `Driver Accurate Conf#` = `Confirmation...`, Helper = `Helper.`, `Helper Accurate Conf#` = `Confirmation....1`, `Window:` = `Window.`) %>% 
                                        mutate(
                                                Date = mdy(Date),
                                                `Tracking ID` = as.character(`Tracking ID`),
                                                `Driver Accurate Conf#` = as.character(`Driver Accurate Conf#`),
                                                `Helper Accurate Conf#` = as.character(`Helper Accurate Conf#`)
                                        ) %>% 
                                        select(-`Window:`) %>% 
                                        Assistant_Process %>% 
                                        Finalize("PORTLAND,OR") %>% 
                                        rbind(output)
                        }
                        if (sum(grepl("COON RAPIDS_MN", inputs$name))==1){
                                output <- 
                                        read.csv(inputs %>% filter(name == "COON RAPIDS_MN") %>% pull(datapath), stringsAsFactors=FALSE, check.names = T) %>%
                                        select(Date, Driver, X, Sale.Order) %>% 
                                        cbind(str_split_fixed(.$Driver, " DR ", 2) %>% 
                                                      as_tibble %>%
                                                      dplyr::rename(Driver1 = V1, `Driver#` = V2)
                                        ) %>% 
                                        cbind(str_split_fixed(.$Driver1, "[(]", 2) %>% as_tibble %>% select(-V1)) %>% 
                                        cbind(str_split_fixed(.$`Driver#`, "[)]", 2) %>% as_tibble %>% select(-V2)) %>% 
                                        select(Date, X, Sale.Order, V2, V1) %>% 
                                        dplyr::rename(Driver = V2, `Driver Accurate Conf#` = V1) %>% 
                                        cbind(str_split_fixed(.$X, " ASST ", 2) %>% 
                                                      as_tibble %>%
                                                      dplyr::rename(Helper = V1, `Helper#` = V2)
                                        ) %>% 
                                        cbind(str_split_fixed(.$Helper, "[(]", 2) %>% as_tibble %>% select(-V1)) %>% 
                                        cbind(str_split_fixed(.$`Helper#`, "[)]", 2) %>% as_tibble %>% select(-V2)) %>% 
                                        select(Date, Sale.Order, Driver, `Driver Accurate Conf#`, V2, V1) %>% 
                                        dplyr::rename(Helper = V2, `Helper Accurate Conf#` = V1) %>% 
                                        mutate(`Tracking ID` = stringr::str_extract(Sale.Order, "\\d{7}"),
                                               Date = mdy(Date)) %>%
                                        select(Date, `Tracking ID`, Driver, `Driver Accurate Conf#`, Helper, `Helper Accurate Conf#`) %>% 
                                        Assistant_Process %>% 
                                        Finalize("COON RAPIDS,MN") %>% 
                                        rbind(output)
                        }
                        if (sum(grepl("HILLSIDE_IL", inputs$name))==1){
                                output <- 
                                        read.csv(inputs %>% filter(name == "HILLSIDE_IL") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F) %>% 
                                        mutate(
                                                Date = mdy(Date), 
                                                `NSD Job #` = as.character(`NSD Job #`)
                                        ) %>% 
                                        cbind(str_split_fixed(.$`Driver Name/Background Check #`, "/", 2) %>% 
                                                      as_tibble %>%
                                                      dplyr::rename(Driver1 = V1, `Driver#` = V2)
                                        ) %>% 
                                        cbind(str_split_fixed(.$`Helper Name / Background Check #`, "/", 2) %>% 
                                                      as_tibble %>%
                                                      dplyr::rename(Helper1 = V1, `Helper#` = V2)
                                        ) %>% 
                                        select(Date, 
                                               `Tracking ID` = `NSD Job #`, 
                                               Driver = Driver1, 
                                               `Driver Accurate Conf#` = `Driver#`, 
                                               Helper = Helper1, 
                                               `Helper Accurate Conf#` = `Helper#`) %>% 
                                        Assistant_Process %>% 
                                        Finalize("HILLSIDE,IL") %>% 
                                        rbind(output)
                        }
                        if (sum(grepl("WEST VALLEY CITY_UT", inputs$name))==1){
                                output <- 
                                        read.csv(inputs %>% filter(name == "WEST VALLEY CITY_UT") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F) %>% 
                                        cbind(str_split_fixed(.$Driver, ", ", 2) %>% as_tibble) %>% 
                                        select(-Driver) %>% 
                                        dplyr::rename(Driver = V1, Helper = V2) %>% 
                                        cbind(str_split_fixed(.$`Driver Number`, ", ", 2) %>% as_tibble) %>% 
                                        dplyr::rename(`Driver Accurate Conf#` = V1, `Helper Accurate Conf#` = V2) %>% 
                                        select(Date, 
                                               `Tracking ID` = `Sale Order`, 
                                               Driver, 
                                               `Driver Accurate Conf#`, 
                                               Helper, 
                                               `Helper Accurate Conf#`) %>% 
                                        mutate(`Tracking ID` = stringr::str_extract(`Tracking ID`, "\\d{7}"),
                                               Date = mdy(Date)) %>% 
                                        Assistant_Process %>% 
                                        Finalize("WEST VALLEY CITY,UT") %>% 
                                        rbind(output)
                        }
                        if (sum(grepl("FARMERS BRANCH_TX", inputs$name))==1){
                                
                                output <- 
                                        read.csv(inputs %>% filter(name == "FARMERS BRANCH_TX") %>% pull(datapath), 
                                                 fileEncoding="UTF-8-BOM", 
                                                 stringsAsFactors=FALSE, 
                                                 check.names = T
                                        ) %>% 
                                        rename(
                                                `Tracking ID` = Order...,
                                                Driver = Driver.,
                                                `Driver Accurate Conf#` = Confirmation...,
                                                Helper = Helper.,
                                                `Helper Accurate Conf#` = Confirmation....1
                                        ) %>% 
                                        mutate(Date = date(mdy_hm(Start))) %>% 
                                        Assistant_Process %>% 
                                        Finalize("FARMERS BRANCH,TX") %>% 
                                        rbind(output)
                        }
                        if (sum(grepl("LAS VEGAS_NV", inputs$name))==1){
                                
                                output <- 
                                        read.csv(inputs %>% filter(name == "LAS VEGAS_NV") %>% pull(datapath), fileEncoding="UTF-8-BOM", header = F, stringsAsFactors=FALSE, check.names = F) %>% 
                                        filter(V3 != "", V1 != "Date") %>% 
                                        rename("Date" = V1, "Location" = V2, "Tracking ID" = V3, "Driver" = V4, "Driver Accurate Conf#" = V5, "DA Job_1" = V6, "Helper" = V7, "Helper Accurate Conf#" = V8, "DA Job_2" = V9) %>% 
                                        select(-`DA Job_1`, -`DA Job_2`) %>% 
                                        Assistant_Process %>% 
                                        mutate(Date = mdy(Date)) %>% 
                                        Finalize("LAS VEGAS,NV") %>%
                                        rbind(output)
                        }
                        if (sum(grepl("SAINT LOUIS_MO", inputs$name))==1){
                                Temp <- read.csv(inputs %>% filter(name == "SAINT LOUIS_MO") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F)
                                names(Temp) <- c("Date", "Location", "Tracking ID", "Delivery Associate_1", "Accurate Conf#_1", "DA Job_1", "Delivery Associate_2", "Accurate Conf#_2", "DA Job_2")
                                Temp <- Temp %>% filter(`Tracking ID` != "", Date != "Date")
                                Driver <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_1", `Accurate Conf#` = "Accurate Conf#_1", `DA Job` = "DA Job_1")
                                Helper <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_2", `Accurate Conf#` = "Accurate Conf#_2", `DA Job` = "DA Job_2")
                                output <- rbind(Driver, Helper) %>% 
                                        mutate(Date = mdy(Date)) %>% 
                                        cbind(str_split_fixed(.$`Delivery Associate`, ", ", 2) %>% 
                                                      as_tibble %>%
                                                      dplyr::rename(`First_Name` = V1, `Last_Name` = V2)
                                        ) %>% 
                                        mutate(`Delivery Associate` = paste(First_Name, Last_Name)) %>% 
                                        select(-First_Name, -Last_Name) %>% 
                                        Finalize("SAINT LOUIS,MO") %>%
                                        rbind(output)
                                rm("Driver", "Helper", "Temp")
                        }
                        if (sum(grepl("SAN ANTONIO_TX", inputs$name))==1){
                                Temp <- read.csv(inputs %>% filter(name == "SAN ANTONIO_TX") %>% pull(datapath), fileEncoding="UTF-8-BOM", stringsAsFactors=FALSE, check.names = F)
                                names(Temp) <- c("Date", "Location", "Tracking ID", "Delivery Associate_1", "Accurate Conf#_1", "DA Job_1", "Delivery Associate_2", "Accurate Conf#_2", "DA Job_2")
                                Temp <- Temp %>% filter(`Tracking ID` != "", Date != "Date")
                                Driver <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_1", `Accurate Conf#` = "Accurate Conf#_1", `DA Job` = "DA Job_1")
                                Helper <- Temp %>% select(Date, Location, `Tracking ID`, `Delivery Associate` = "Delivery Associate_2", `Accurate Conf#` = "Accurate Conf#_2", `DA Job` = "DA Job_2")
                                output <- rbind(Driver, Helper) %>% 
                                        mutate(Date = mdy(Date)) %>% 
                                        cbind(str_split_fixed(.$`Delivery Associate`, ", ", 2) %>% 
                                                      as_tibble %>%
                                                      dplyr::rename(`First_Name` = V1, `Last_Name` = V2)
                                        ) %>% 
                                        mutate(`Delivery Associate` = paste(First_Name, Last_Name)) %>% 
                                        select(-First_Name, -Last_Name) %>% 
                                        Finalize("SAN ANTONIO,TX") %>%
                                        rbind(output)
                                rm("Driver", "Helper", "Temp")
                        }
                        
                        if (nrow(output) > 0) {
                                output <-
                                        output %>% 
                                        mutate(
                                                `Delivery Associate` = as.character(`Delivery Associate`),
                                                `Accurate Conf#` = as.character(`Accurate Conf#`)
                                        ) %>% 
                                        merge(TMdata %>% select(NSD.., PO.), all.x = T, by.x = "Tracking ID", by.y = "NSD..") %>% 
                                        mutate(
                                                `Delivery Associate` = ifelse(`Delivery Associate` == "Claudio10", "CLAUDIO MONASTERIO", `Delivery Associate`),
                                                `Accurate Conf#` = ifelse(`Delivery Associate` %in% c("Claudio10", "CLAUDIO MONASTERIO"), "82698909", `Accurate Conf#`),
                                                `Delivery Associate` = ifelse(`Delivery Associate` == "Adrian Suarez  TRK 10 - 24' BOX", "Adrian Suarez", `Delivery Associate`),
                                                `Accurate Conf#` = ifelse(`Delivery Associate` %in% c("Adrian Suarez  TRK 10 - 24' BOX", "Adrian Suarez"), "98025517", `Accurate Conf#`),
                                                `Encrypted ID` = PO.,
                                                `Route ID` = NA,
                                                `Accurate Account Owner` = NA,
                                                `Delivery Associate` = toupper(`Delivery Associate`),
                                                `DA Job` = toupper(`DA Job`)
                                        ) %>%
                                        select(                         
                                                `accurate_account_owner` = `Accurate Account Owner`,                         
                                                `confirmation_number` = `Accurate Conf#`,                         
                                                `delivery_associate` = `Delivery Associate`,                         
                                                `delivery_associate_role` = `DA Job`,                         
                                                `delivery_date` = Date,                         
                                                `encrypted_id` = `Encrypted ID`,                         
                                                location = Location,                         
                                                route_id = `Route ID`,                         
                                                tracking_id = `Tracking ID`                 
                                        ) %>% 
                                        arrange(desc(delivery_date), location) %>% 
                                        filter(!is.na(confirmation_number), confirmation_number != "",
                                               !is.na(delivery_associate), delivery_associate != "") %>% 
                                        unique %>% 
                                        write.csv(file, row.names = F)
                        }
                })
        
        # Estes -------------------------------------------------------------------
        
        estes_input_data <- reactive({
                estes_input <- input$delivery_amazon_estes_raw_data_input
                if (is.null(estes_input)) return(NULL)
                estes_input_data <- 
                        read_excel(estes_input$datapath) %>% 
                        # read_excel("C:/Users/hyazarloo/Desktop/Amazon Estes OFD Report 2019-01-09T1208 (1).xlsx") %>% 
                        rename(
                                `Job Date` = `Basic Order Info Job Date`,
                                `Agent Code` = `Basic Order Info Agent Code`,
                                `NSD Number` = `Basic Order Info NSD Number`,
                                `Prescheduled Delivery Date` = `Delivery Details Prescheduled Delivery Date`,
                                `Consignee Name` = `Basic Order Info Consignee Name`,
                                `Current Status` = `Basic Order Info Current Status`,
                                `PO Numbers` = `Basic Order Info PO Numbers`
                        ) %>% 
                        mutate(
                                `Job Date` = `Job Date` %>% date,
                                `Prescheduled Delivery Date` = `Prescheduled Delivery Date` %>% date,
                                date = `Prescheduled Delivery Date`,
                                `Prescheduled Delivery Date` = `Prescheduled Delivery Date` %>% as.character %>% replace_na("NO PDD")
                        )
        })
        
        observeEvent(input$delivery_amazon_estes_subject_save, {
                
                if (is.null(authorization.check(Amazon_team))) return(NULL)
                
                writeChar(
                        isolate(input$delivery_amazon_estes_subject), 
                        amazon_email_item_path("Estes", "subject")
                )
                shinyalert(
                        title = "Saved",
                        text = "Subject",
                        closeOnEsc = TRUE,
                        closeOnClickOutside = TRUE,
                        html = FALSE,
                        type = "success",
                        showConfirmButton = TRUE,
                        showCancelButton = FALSE,
                        confirmButtonText = "OK",
                        confirmButtonCol = "#a2a8ab",
                        timer = 0,
                        imageUrl = "",
                        animation = TRUE
                )
        })
        
        observeEvent(input$delivery_amazon_estes_body_save,{
                
                if (is.null(authorization.check(Amazon_team))) return(NULL)
                
                writeChar(
                        isolate(input$delivery_amazon_estes_body), 
                        amazon_email_item_path("Estes", "body")
                )
                
                shinyalert(
                        title = "Saved",
                        text = "Body",
                        closeOnEsc = TRUE,
                        closeOnClickOutside = TRUE,
                        html = FALSE,
                        type = "success",
                        showConfirmButton = TRUE,
                        showCancelButton = FALSE,
                        confirmButtonText = "OK",
                        confirmButtonCol = "#a2a8ab",
                        timer = 0,
                        imageUrl = "",
                        animation = TRUE
                )
                
        })
        
        
        observeEvent(input$delivery_amazon_estes_send_button, {
                
                if (is.null(authorization.check(Amazon_team))) return(NULL)
                
                data_all_agents <- isolate(estes_input_data())
                if (is.null(data_all_agents)) {
                        showModal(modalDialog(
                                title = "CAUTION!",
                                "No data imported!",
                                easyClose = TRUE,
                                footer = NULL
                        ))
                        return(NULL)
                }
                
                agents_names <- 
                        data_all_agents %>% 
                        pull(`Agent Code`) %>% 
                        unique
                
                estes_subject <- isolate(input$delivery_amazon_estes_subject) %>% paste0()
                estes_body <- paste(paste0('<p>', strsplit(isolate(input$delivery_amazon_estes_body),"\n")[[1]], "</p>"), collapse="")
                
                OutApp <- COMCreate("Outlook.Application")
                withProgress(message = 'Sending E-mails in progress', value = 0, {
                        for (agent in agents_names) {
                                incProgress(1/length(agents_names))
                                data <- data_all_agents %>% filter(`Agent Code` == agent)
                                estes_subject <- isolate(input$delivery_amazon_estes_subject) %>% paste0(agent)
                                
                                outMail = OutApp$CreateItem(0)
                                outMail[["To"]] = email_address(
                                        paste0(
                                                if (agent == "CLAUDIO") {""} else paste0("AG_", agent,"@nonstopdelivery.com;"), 
                                                Amazon_Additiona_Email_from_Agent(agent)),
                                        isolate(input$Test_Active)
                                )
                                outMail[["CC"]] = email_address(
                                        "nsdamazon@shipnsd.com",
                                        isolate(input$Test_Active)
                                )
                                outMail[["subject"]] = estes_subject
                                
                                red_rows <- which(data$date < today())
                                orange_rows <- which(data$date == today() + 1)
                                green_rows <- which(is.na(data$date))
                                blue_rows <- which(data$`Current Status` %in% c("APPRVD", "CANCELLED", "COMPLETED", "PRINTED"))
                                
                                table <- 
                                        data %>% select(1:7) %>% 
                                        kable(align = "c") %>%
                                        kable_styling(full_width = F) %>% 
                                        # column_spec(1, width = "14%") %>% 
                                        row_spec(red_rows, color = "white", background = "#D7261E") %>%
                                        row_spec(orange_rows, color = "white", background = "#f2a710") %>%
                                        row_spec(green_rows, color = "white", background = "#33ad2e") %>%
                                        row_spec(blue_rows, color = "white", background = "#4286f4")
                                
                                outMail[["HTMLBody"]] = 
                                        paste0(
                                                estes_body,
                                                table %>% HTML, 
                                                nsd_signature(Sys.info()["user"][[1]])
                                        )
                                # outMail[["attachments"]]$Add("C:/Path/To/The/Attachment/File.ext")
                                outMail$Send()
                                
                                
                                
                        }
                })
                
                
        })
        
        # Agent Scorecards --------------------------------------------------------
        
        observe({
                if(is.null(input$delivery_scorecards_emails_button) || input$delivery_scorecards_emails_button == 0) return(NULL)
                if (is.null(authorization.check(Agent_Scorecards_emails_authorized_people))) return(NULL)
                
                file_addresses <- choose.files()
                if (is.na(file_addresses)) return(NULL)
                
                agents_names <- tools::file_path_sans_ext(basename(file_addresses))
                
                OutApp <- COMCreate("Outlook.Application")
                withProgress(message = 'Sending E-mails in progress', value = 0, {
                        for (agent in agents_names) {
                                incProgress(1/length(agents_names))
                                agent = "AIC"
                                outMail = OutApp$CreateItem(0)
                                outMail[["To"]] = email_address(
                                        paste0("AG_", agent,"@nonstopdelivery.com"),
                                        isolate(input$Test_Active)
                                )
                                outMail[["subject"]] = paste0("NSD Agent Scorecard - ", agent)
                                outMail[["HTMLBody"]] = 
                                        paste0(
                                                "<p>Valued Partners,</p>",
                                                "<p>In an effort to drive continues improvements I have attached a copy of your scorecard for the last four weeks.</p>",
                                                "<p>Here is the breakdown of the categories we are measuring:</p>",
                                                data.frame(
                                                        Metric = c(
                                                                "Receiving Leadtime", 
                                                                "Last Mile Transit", 
                                                                "Window Compliance", 
                                                                "Date Compliance", 
                                                                "First Update Compliance"
                                                        ),
                                                        Description = c("Percent orders moved to  DOCKED Status same day",
                                                                        "Average Transit time from DOCKED to DELIVERED",
                                                                        "Percent delivered within 4 hour window",
                                                                        "Percent delivered on scheduled delivery date",
                                                                        "Percent customers contacted within 24 hours from DOCKED"),
                                                        `Target Metric` = c("95%", "2 days", "95%",  "99%", "95%")
                                                ) %>% 
                                                        htmlTable(
                                                                align="l",
                                                                align.header = "l",
                                                                css.cell = "padding-left: 1em; padding-right: 1em;",
                                                                css.tspanner = "font-weight: 900; text-align: left;",
                                                                css.table = "margin-top: .5em; margin-bottom: .5em; margin: auto; padding: 0px; width: 100%"
                                                        ), 
                                                "<p>Please contact your last mile coordinator if you have any questions.</p>",
                                                nsd_signature(Sys.info()["user"])
                                        )
                                
                                outMail[["attachments"]]$Add(file_addresses[agents_names == agent])
                                outMail$Send()  
                                
                        }
                })
                
                
                
        })
        
        
        # TR Inbound Report -------------------------------------------------------
        
        TRansportation_TM <- reactive({
                TM_report <- input$button_Transportation_TM_Input
                if (is.null(TM_report)) return(NULL)
                
                TRansportation_TM <- read_csv(TM_report$datapath,
                                              col_types = cols(
                                                      `ADMIN STATUS CHANGED` = col_datetime(format = "%m/%d/%Y %H:%M:%S %p"), 
                                                      `NSD #` = col_character(),
                                                      `AGENT PHONE` = col_character(), 
                                                      `CARRIER PRO` = col_character(), 
                                                      `CLIENT ID` = col_character(), CREATED = col_date(format = "%m/%d/%Y"), 
                                                      `LAST COMMENT DATE` = col_datetime(format = "%m/%d/%Y %H:%M:%S %p"), 
                                                      `LAST STATUS UPDATE` = col_datetime(format = "%m/%d/%Y %H:%M:%S %p"), 
                                                      `LTL P/U DATE` = col_datetime(format = "%m/%d/%Y %H:%M:%S %p"), 
                                                      `LTL READY` = col_date(format = "%m/%d/%Y"), 
                                                      `OP STATUS CHANGED` = col_datetime(format = "%m/%d/%Y %H:%M:%S %p"), 
                                                      `PICKUP DATE` = col_date(format = "%m/%d/%Y"), 
                                                      SCHEDULED = col_date(format = "%m/%d/%Y"), 
                                                      SHIPPER_PHONE = col_character(), 
                                                      SHIPPER_ZIP = col_character()
                                              )
                ) %>% filter(!is.na(`NSD #`))
        })
        
        TRansportation_TT <- reactive({
                
                TT_Report <- input$button_Transportation_TT_Input
                if (is.null(TT_Report)) return(NULL)
                
                TRansportation_TT <- read_excel(TT_Report$datapath,
                                                # TRansportation_TT <- read_excel("C:/Users/hyazarloo/Desktop/NSD Track-Trace 12.18 - FINAL.xlsx",
                                                sheet = "TRACK MASTER"
                )    
        })
        
        Transportation_inbound_report <- reactive({
                TRansportation_TT <- TRansportation_TT()
                TRansportation_TM <- TRansportation_TM()
                
                if (is.null(TRansportation_TT)|is.null(TRansportation_TM)) return(NULL)
                
                Transportation_inbound_report <- 
                        TRansportation_TT %>% 
                        select(
                                `SYSTEM` = 1,	
                                `NSD#` = 2,
                                `PO#` = 3,
                                CARRIER = 6,
                                POD = 7,	
                                `UNYSON UPDATE` = 8,
                                `ORIGIN NAME` = 10,
                                `SHIP DATE` = 12,
                                ETA = 14,
                                `Carrier PRO` = 16,
                                `Origin City` = 18,	
                                `From State` = 19,
                                `Agent Code` = 21,	
                                `Destination City` = 22,	
                                `To State` = 23,
                                Pieces = 25,	
                                Weight = 26
                        ) %>% 
                        mutate(
                                `NSD#` = `NSD#` %>% as.character,
                                POD = POD %>% as.Date,
                                `SHIP DATE` = (as.numeric(`SHIP DATE`) * 60*60*24) %>% as.POSIXct(origin="1899-12-30", tz="GMT"),
                                ETA = ETA %>% as.Date
                        ) %>% 
                        filter(
                                ETA == input$inbound_report_date,
                                is.na(POD),
                                !grepl("CANCEL|OSD|DELIVERED", toupper(`UNYSON UPDATE`))
                        ) %>% 
                        left_join(
                                TRansportation_TM %>% 
                                        select(
                                                `NSD#` = `NSD #`, 
                                                `OP STATUS`,
                                                `Customer Name` = CONSIGNEE,
                                                `Customer City` = CITY,
                                                `Customer ZIP` = ZIP,
                                                `Service Type` = `SERVICE LEVEL`
                                        ), 
                                by = "NSD#"
                        ) %>% 
                        left_join(Agents %>% select(`Agent Code` = Agent, Full_Name), by = "Agent Code") %>% 
                        filter(!is.na(`OP STATUS`)) %>% 
                        select(
                                `Agent Code`,
                                AGENT = Full_Name,
                                `NSD#`, 
                                `PO#`, 
                                `Customer Name`,
                                `Customer City`,
                                `Customer ZIP`,
                                `Service Type`,
                                CARRIER, 
                                `Carrier PRO`, 
                                Pieces,
                                Weight,
                                ETA
                        ) %>% 
                        arrange(`Agent Code`)
        })
        
        observe({
                if(is.null(input$send_inbound_report_emails) || input$send_inbound_report_emails == 0) return(NULL)
                if (is.null(authorization.check(isolate(Transportation_Coordinators)))) return(NULL)
                
                Transportation_inbound_report <- isolate(Transportation_inbound_report())
                
                if (is.null(Transportation_inbound_report)) {
                        showModal(modalDialog(
                                title = "CAUTION!",
                                "Warning: No Data imported!",
                                easyClose = TRUE,
                                footer = NULL
                        ))
                        return(NULL)
                }
                foldername <- paste0(file.path(Sys.getenv("USERPROFILE"),"Desktop"), "/Inbound_Reports ", gsub(":","-",Sys.time()), "/")
                dir.create(foldername)
                
                WB_list <- 
                        Transportation_inbound_report %>% 
                        split(.$`Agent Code`) %>% 
                        lapply(. %>% select(-`Agent Code`))
                
                withProgress(message = 'Creating attachments in progress', value = 0, {
                        for (agent_name in names(WB_list)) {
                                incProgress(1/length(WB_list))
                                
                                WB_list[[agent_name]] %>% 
                                        create_inbound_report_worksheet("Return_Daily_Report") %>% 
                                        saveWorkbook(
                                                paste0(foldername, agent_name, ".xlsx"), 
                                                overwrite = TRUE
                                        )
                        }
                })
                
                all_of_the_agents <- list.files(path = foldername) %>% file_path_sans_ext
                all_of_the_agents <- all_of_the_agents[all_of_the_agents != ""]
                OutApp <- COMCreate("Outlook.Application")
                withProgress(message = 'Sending E-mails in progress', value = 0, {
                        for (agent in all_of_the_agents) {
                                incProgress(1/length(all_of_the_agents))
                                
                                outMail = OutApp$CreateItem(0)
                                outMail[["To"]] = email_address(
                                        paste0("AG_", agent,"@nonstopdelivery.com"),
                                        isolate(input$Test_Active)
                                )
                                outMail[["subject"]] = paste0("Inbound Daily Report - ", agent)
                                outMail[["HTMLBody"]] =
                                        paste0(
                                                "<p>Hello NSD Partner - </p>",
                                                "<br>",
                                                "<p>This email is to ALERT you to inbound freight coming to your facility.  The attached excel spreadsheet captures all the information you need to inbound, process, look up in the NSD system, on-dock, schedule, stage and out-for-delivery these shipments as they arrive.</p>",
                                                "<br>",
                                                "<p>Please PRINT this report daily and have it in the hands of the person receiving freight at your facility. </p>",
                                                "<p>Please ON-DOCK all accounted for freight that arrives at your facility as soon as possible. </p>",
                                                "<p>If the orders have already been recieved please update the console from Pending to On Dock and schedule for delivery. </p>",
                                                "<p>Please REPORT any issues, damages, shortages, overages to NSD as soon as possible. </p>",
                                                "<p>This report does NOT contain all shipments destined for your facility. </p>",
                                                "<p>Please be reminded that the ETAs are just that..Estimated.  The ETA provided on this report is the ETA of the carrier to your facility.</p>",
                                                "<br>",
                                                "<p>If there is an email address update please reply to this email with the email address that should be used or removed. </p>",
                                                nsd_signature(Sys.info()["user"])
                                        )
                                
                                attachment <- paste0(foldername, agent, ".xlsx")
                                if (file.exists(attachment)) {
                                        outMail[["attachments"]]$Add(attachment)
                                }
                                
                                outMail$Send()
                                
                        }
                })
                if (!is.null(all_of_the_agents))  {
                        shinyalert(
                                title = sample(good_job_words,1),
                                text = "All of the Inbound Report emails were sent to the corresponding agents",
                                closeOnEsc = TRUE,
                                closeOnClickOutside = TRUE,
                                html = FALSE,
                                type = "success",
                                showConfirmButton = TRUE,
                                showCancelButton = FALSE,
                                confirmButtonText = "OK",
                                confirmButtonCol = "#a2a8ab",
                                timer = 0,
                                imageUrl = "",
                                animation = TRUE
                        )
                } else {
                        shinyalert(
                                title = "Oops!",
                                text = "Something went wrong! There were no emails to be sent!",
                                closeOnEsc = TRUE,
                                closeOnClickOutside = TRUE,
                                html = FALSE,
                                type = "error",
                                showConfirmButton = TRUE,
                                showCancelButton = FALSE,
                                confirmButtonText = "OK",
                                confirmButtonCol = "#a2a8ab",
                                timer = 0,
                                imageUrl = "",
                                animation = TRUE
                        )
                } 
                
        })
        
        
        
        # Dashboard ---------------------------------------------------------------
        
        
        # Vol ---------------------------------------------------------------------
        
        
        
        dashboard_volumes_system <- reactive({input$dashboard_volumes_system})
        dashboard_volumes_svctyp <- reactive({input$dashboard_volumes_servicetype})
        dashboard_volumes_acc <- reactive({input$dashboard_volumes_account})
        dashboard_volumes_item <- reactive({input$dashboard_volumes_item})
        dashboard_volumes_date1 <- reactive({input$dashboard_volumes_daterange[1]})
        dashboard_volumes_date2 <- reactive({input$dashboard_volumes_daterange[2]})
        
        dashboard_volumes_category <- reactive({
                
                dashboard_volumes_svctyp <- dashboard_volumes_svctyp()
                
                dashboard_volumes_category <- 
                        dashboard_volumes_svctyp() %>% 
                        category_from_servicetype
        })
        
        dashboard_volumes_items_choices <- reactive({
                
                dashboard_volumes_category = dashboard_volumes_category()
                
                dashboard_volumes_items_choices <- 
                        switch (dashboard_volumes_category,
                                "Delivery" = Delivery.Counts.Columns ,
                                "Return" = Returns.Counts.Columns
                        )
                
        })
        # update counts_item
        observe({
                updateSelectInput(session, "dashboard_volumes_item", choices = dashboard_volumes_items_choices())
        })
        
        
        output$dashboard_volumes_download_graph_data <-
                
                downloadHandler(
                        filename = function(){
                                paste("GraphData", 
                                      dashboard_volumes_category(),
                                      dashboard_volumes_svctyp(),
                                      dashboard_volumes_acc(),
                                      dashboard_volumes_item(),
                                      "From",
                                      dashboard_volumes_date1(),
                                      "To",
                                      dashboard_volumes_date2(),
                                      "at",
                                      gsub(":","-",Sys.time()), ".csv", sep="_")
                        },
                        content = function(file) {
                                dashboard_volumes_general_graph_data() %>% write.csv(file, sep = ",", row.names = F)
                        })
        
        output$dashboard_volumes_download_raw_data <-
                
                downloadHandler(
                        filename = function(){
                                paste("RawData", 
                                      dashboard_volumes_category(),
                                      dashboard_volumes_svctyp(),
                                      dashboard_volumes_acc(),
                                      dashboard_volumes_item(),
                                      "From",
                                      dashboard_volumes_date1(),
                                      "To",
                                      dashboard_volumes_date2(),
                                      "at",
                                      gsub(":","-",Sys.time()), ".csv", sep="_")
                        },
                        content = function(file) {
                                dashboard_volumes_general_data() %>% write.csv(file, sep = ",", row.names = F)
                        })
        
        
        
        dashboard_volumes_general_data <- reactive({
                
                input$dashboard_volumes_refresh
                
                delivery_or_return = isolate(dashboard_volumes_category())
                svctyp = isolate(dashboard_volumes_svctyp())
                counts_or_metrics = "counts"
                item = isolate(dashboard_volumes_item())
                system = isolate(dashboard_volumes_system())
                acc = isolate(dashboard_volumes_acc())
                date1 = as.Date("2013-01-01")
                date2 = as.Date(Sys.Date())
                
                dashboard_volumes_general_data <-
                        filter_data(delivery_or_return, svctyp, counts_or_metrics, item, system, acc, date1, date2)
        })
        
        output$dashboard_volumes_general_7_days <- renderText({
                input$dashboard_volumes_refresh
                
                delivery_or_return = isolate(dashboard_volumes_category())
                svctyp = isolate(dashboard_volumes_svctyp())
                counts_or_metrics = "counts"
                item = isolate(dashboard_volumes_item())
                system = isolate(dashboard_volumes_system())
                acc = isolate(dashboard_volumes_acc())
                date1 = today() - 7
                date2 = today() - 1
                
                filter_data(delivery_or_return, svctyp, counts_or_metrics, item, system, acc, date1, date2) %>% 
                        count %>% 
                        as.character
        })
        
        
        output$dashboard_volumes_general_30_days <- renderText({
                
                input$dashboard_volumes_refresh
                
                delivery_or_return = isolate(dashboard_volumes_category())
                svctyp = isolate(dashboard_volumes_svctyp())
                counts_or_metrics = "counts"
                item = isolate(dashboard_volumes_item())
                system = isolate(dashboard_volumes_system())
                acc = isolate(dashboard_volumes_acc())
                date1 = today() - 30
                date2 = today() - 1
                
                filter_data(delivery_or_return, svctyp, counts_or_metrics, item, system, acc, date1, date2) %>% 
                        count %>% 
                        as.character
        })
        
        output$dashboard_volumes_general_365_days <- renderText({
                
                input$dashboard_volumes_refresh
                
                delivery_or_return = isolate(dashboard_volumes_category())
                svctyp = isolate(dashboard_volumes_svctyp())
                counts_or_metrics = "counts"
                item = isolate(dashboard_volumes_item())
                system = isolate(dashboard_volumes_system())
                acc = isolate(dashboard_volumes_acc())
                date1 = today() - 365
                date2 = today() - 1
                
                filter_data(delivery_or_return, svctyp, counts_or_metrics, item, system, acc, date1, date2) %>% 
                        count %>% 
                        as.character
        })
        
        dashboard_volumes_general_graph_data <- reactive({
                
                dashboard_volumes_general_data <- dashboard_volumes_general_data()
                
                p <- input$dashboard_volumes_general_period
                
                pr <-
                        switch (p,
                                "Daily" = "day",
                                "Weekly" = "week",
                                "Monthly" = "month",
                                "Annually" = "year"
                        )
                
                dates <-
                        seq.Date(as.Date("2013-01-01"), as.Date(Sys.Date()), "days") %>%
                        as.data.frame %>%
                        rename_("Date" = ".") %>%
                        {if (pr == "day") . else mutate(., Date = cut(Date, breaks = pr) %>% as.Date)} %>%
                        distinct
                
                dashboard_volumes_general_graph_data <- 
                        dashboard_volumes_general_data %>%
                        {if (pr == "day") . else mutate(., Date = cut(Date, breaks = pr) %>% as.Date)} %>%
                        group_by(Date) %>%
                        dplyr::summarize(Count = n()) %>%
                        right_join(dates, by = "Date") %>%
                        mutate(Count = replace_na(Count, 0)) %>% 
                        head(nrow(.)-1)
                
        })
        
        output$dashboard_volumes_general_graph <- renderGvis({
                
                dashboard_volumes_general_graph_data <- dashboard_volumes_general_graph_data()
                
                dashboard_volumes_general_graph_data %>%
                        gvisAnnotationChart(datevar="Date",
                                            numvar="Count",
                                            options=list(
                                                    width = "95%", height=550,
                                                    fill=10, displayExactValues=TRUE,
                                                    colors="['#333d4f']")
                        )
        })
        
        # output$dashboard_volumes_general_boxplot <- renderPlotly({
        #         
        #         dashboard_volumes_general_graph_data <- dashboard_volumes_general_graph_data()
        #         if (is.null(dashboard_volumes_general_graph_data)) return(NULL)
        #         
        #         dashboard_volumes_general_graph_data %>%
        #                 plot_ly(
        #                         y = ~Count, 
        #                         color = I("#8c9ac2"), 
        #                         alpha = 0.5, 
        #                         boxpoints = "suspectedoutliers",
        #                         type = "box",
        #                         name = "Boxplot"
        #                 ) %>% 
        #                 layout(
        #                         title = "Boxplot of the volume",
        #                         yaxis =  list(title = "Volume")
        #                 )
        #         
        # })
        
        
        # Mtr ---------------------------------------------------------------------
        
        
        dashboard_metrics_system <- reactive({input$dashboard_metrics_system})
        dashboard_metrics_svctyp <- reactive({input$dashboard_metrics_servicetype})
        dashboard_metrics_acc <- reactive({input$dashboard_metrics_account})
        dashboard_metrics_item <- reactive({input$dashboard_metrics_item})
        dashboard_metrics_date1 <- reactive({input$dashboard_metrics_daterange[1]})
        dashboard_metrics_date2 <- reactive({input$dashboard_metrics_daterange[2]})
        
        dashboard_metrics_category <- reactive({
                
                svctyp = dashboard_metrics_svctyp()
                
                dashboard_metrics_category <- category_from_servicetype(svctyp)
                
        })
        
        dashboard_metrics_items_choices <- reactive({
                
                # statuses are different for deliveries an returns.
                # from choice of service_type we got the category(delivery or return).
                # from category we'll get statuses.
                # this will update the drop-down menue of Status
                dashboard_metrics_category = dashboard_metrics_category()
                
                dashboard_metrics_items_choices <- 
                        Metrics %>% 
                        filter(category == dashboard_metrics_category) %>% 
                        pull(caption)
        })
        # update counts_item
        observe({
                updateSelectInput(session, "dashboard_metrics_item", choices = dashboard_metrics_items_choices())
        })
        
        
        output$dashboard_metrics_download_graph_data <-
                
                downloadHandler(
                        filename = function(){
                                paste("GraphData", dashboard_metrics_category(),
                                      dashboard_metrics_svctyp(),
                                      dashboard_metrics_acc(),
                                      dashboard_metrics_item(),
                                      "From",
                                      dashboard_metrics_date1(),
                                      "To",
                                      dashboard_metrics_date2(),
                                      "at",
                                      gsub(":","-",Sys.time()), ".csv", sep="_")
                        },
                        content = function(file) {
                                dashboard_metrics_general_graph_data() %>% write.csv(file, sep = ",", row.names = F)
                        })
        
        output$dashboard_metrics_download_raw_data <-
                
                downloadHandler(
                        filename = function(){
                                paste("RawData", 
                                      dashboard_volumes_category(),
                                      dashboard_volumes_svctyp(),
                                      dashboard_volumes_acc(),
                                      dashboard_volumes_item(),
                                      "From",
                                      dashboard_volumes_date1(),
                                      "To",
                                      dashboard_volumes_date2(),
                                      "at",
                                      gsub(":","-",Sys.time()), ".csv", sep="_")
                        },
                        content = function(file) {
                                dashboard_metrics_general_data() %>% write.csv(file, sep = ",", row.names = F)
                        })
        
        
        dashboard_metrics_general_data <- reactive({
                
                input$dashboard_metrics_refresh
                
                delivery_or_return = isolate(dashboard_metrics_category())
                svctyp = isolate(dashboard_metrics_svctyp())
                counts_or_metrics = "metrics"
                item = isolate(dashboard_metrics_item())
                system = isolate(dashboard_metrics_system())
                acc = isolate(dashboard_metrics_acc())
                date1 = as.Date("2013-01-01")
                date2 = as.Date(Sys.Date())
                
                dashboard_metrics_general_data <-
                        filter_data(delivery_or_return, svctyp, counts_or_metrics, item, system, acc, date1, date2)
        })
        
        output$dashboard_metrics_general_7_days <- renderText({
                
                input$dashboard_metrics_refresh
                
                svctyp = isolate(dashboard_metrics_svctyp())
                type <- type_from_metric(isolate(dashboard_metrics_item()), svctyp)
                delivery_or_return = isolate(dashboard_metrics_category())
                counts_or_metrics = "metrics"
                item = isolate(dashboard_metrics_item())
                system = isolate(dashboard_metrics_system())
                acc = isolate(dashboard_metrics_acc())
                date1 = today() - 7
                date2 = today() - 1
                
                filter_data(delivery_or_return, svctyp, counts_or_metrics, item, system, acc, date1, date2) %>% 
                        pull(1) %>% 
                        mean(na.rm = T) %>% 
                        round(2) %>% 
                        {
                                if (type == "number") {
                                        .
                                } else if (type == "percentage") {
                                        percent(.)
                                }
                        } %>% 
                        as.character
        })
        
        output$dashboard_metrics_general_30_days <- renderText({
                
                input$dashboard_metrics_refresh
                
                svctyp = isolate(dashboard_metrics_svctyp())
                type <- type_from_metric(isolate(dashboard_metrics_item()), svctyp)
                delivery_or_return = isolate(dashboard_metrics_category())
                counts_or_metrics = "metrics"
                item = isolate(dashboard_metrics_item())
                system = isolate(dashboard_metrics_system())
                acc = isolate(dashboard_metrics_acc())
                date1 = today() - 30
                date2 = today() - 1
                
                filter_data(delivery_or_return, svctyp, counts_or_metrics, item, system, acc, date1, date2) %>% 
                        pull(1) %>% 
                        mean(na.rm = T) %>% 
                        round(2) %>% 
                        {
                                if (type == "number") {
                                        .
                                } else if (type == "percentage") {
                                        percent(.)
                                }
                        } %>% 
                        as.character
        })
        
        output$dashboard_metrics_general_365_days <- renderText({
                
                input$dashboard_metrics_refresh
                
                svctyp = isolate(dashboard_metrics_svctyp())
                type <- type_from_metric(isolate(dashboard_metrics_item()), svctyp)
                delivery_or_return = isolate(dashboard_metrics_category())
                counts_or_metrics = "metrics"
                item = isolate(dashboard_metrics_item())
                system = isolate(dashboard_metrics_system())
                acc = isolate(dashboard_metrics_acc())
                date1 = today() - 365
                date2 = today() - 1
                
                filter_data(delivery_or_return, svctyp, counts_or_metrics, item, system, acc, date1, date2) %>% 
                        pull(1) %>% 
                        mean(na.rm = T) %>% 
                        round(2) %>% 
                        {
                                if (type == "number") {
                                        .
                                } else if (type == "percentage") {
                                        percent(.)
                                }
                        } %>% 
                        as.character
        })
        
        dashboard_metrics_general_graph_data <- reactive({
                
                dashboard_metrics_general_data <- dashboard_metrics_general_data()
                
                p <- input$dashboard_metrics_general_period
                
                pr <-
                        switch (p,
                                "Daily" = "day",
                                "Weekly" = "week",
                                "Monthly" = "month",
                                "Annually" = "year"
                        )
                dates <-
                        seq.Date(as.Date("2013-01-01"), as.Date(Sys.Date()), "days") %>%
                        as.data.frame %>%
                        rename_("Date" = ".") %>%
                        {if (pr == "day") . else mutate(., Date = cut(Date, breaks = pr) %>% as.Date)} %>%
                        distinct
                
                dashboard_metrics_general_data %>%
                        rename(metric = 1, Date = 2) %>% 
                        {if (pr == "day") . else mutate(., Date = cut(Date, breaks = pr) %>% as.Date)} %>%
                        group_by(Date) %>%
                        dplyr::summarize(
                                Mean = mean(metric, na.rm = TRUE),
                                Count = n()
                        ) %>%
                        right_join(dates, by = "Date") %>%
                        mutate(
                                Count = replace_na(Count, 0),
                                Mean = round(Mean, 2)) %>% 
                        head(nrow(.)-1)
        })
        
        output$dashboard_metrics_general_graph <- renderGvis({
                
                dashboard_metrics_general_graph_data <- dashboard_metrics_general_graph_data()
                
                dashboard_metrics_general_graph_data %>%
                        gvisAnnotationChart(datevar="Date",
                                            numvar="Mean",
                                            options=list(
                                                    width = "100%", height=500,
                                                    fill=10, displayExactValues=TRUE,
                                                    colors="['#333d4f']")
                        )
        })
        
        # output$dashboard_metrics_general_boxplot <- renderPlotly({
        #         
        #         dashboard_metrics_general_graph_data <- dashboard_metrics_general_graph_data()
        #         if (is.null(dashboard_metrics_general_graph_data)) return(NULL)
        #         
        #         dashboard_metrics_general_graph_data %>%
        #                 plot_ly(
        #                         y = ~Mean, 
        #                         color = I("#8c9ac2"), 
        #                         alpha = 0.5, 
        #                         boxpoints = "suspectedoutliers",
        #                         type = "box",
        #                         name = "Boxplot"
        #                 ) %>% 
        #                 layout(
        #                         title = "Boxplot of the metrics",
        #                         yaxis =  list(title = "metrics")
        #                 )
        #         
        # })
        
        
        
        # Volumes -----------------------------------------------------------------
        
        counts_system <- reactive({input$counts_system})
        counts_svctyp <- reactive({input$counts_servicetype})
        counts_acc <- reactive({input$counts_account})
        counts_item <- reactive({input$counts_item})
        counts_date1 <- reactive({input$counts_daterange[1]})
        counts_date2 <- reactive({input$counts_daterange[2]})
        
        counts_category <- reactive({
                
                counts_svctyp <- counts_svctyp()
                
                counts_category <- 
                        counts_svctyp() %>% 
                        category_from_servicetype
        })
        
        counts_regions_states <- reactive({
                
                counts_category <- counts_category()
                
                counts_regions_states <-
                        switch (counts_category,
                                "Delivery" = LMRegions_State %>% rename(Region = LM_Region),
                                "Return" = RRegions_State %>% rename(Region = R_Region)
                        )
        })
        
        counts_items_choices <- reactive({
                
                # statuses are different for deliveries an returns.
                # from choice of service_type we got the category(delivery or return).
                # from category we'll get statuses.
                # this will update the drop-down menue of Status
                counts_category = counts_category()
                
                counts_items_choices <- 
                        switch (counts_category,
                                "Delivery" = Delivery.Counts.Columns ,
                                "Return" = Returns.Counts.Columns
                        )
                
        })
        # update counts_item
        observe({
                updateSelectInput(session, "counts_item", choices = counts_items_choices())
        })
        # show modal if date1 > date2
        observe({
                input$counts_refresh
                
                date1 = isolate(counts_date1())
                date2 = isolate(counts_date2())
                
                if (date1 > date2) {
                        showModal(modalDialog(
                                title = "Date range is not acceptable",
                                "Beginning date should be smaller than or equal to the end date.",
                                easyClose = TRUE,
                                footer = NULL
                        ))
                }
        })
        
        counts_data <- reactive({
                
                input$counts_refresh
                
                delivery_or_return = isolate(counts_category())
                svctyp = isolate(counts_svctyp())
                counts_or_metrics = "counts"
                item = isolate(counts_item())
                system = isolate(counts_system())
                acc = isolate(counts_acc())
                date1 = isolate(counts_date1())
                date2 = isolate(counts_date2())
                
                counts_data <- filter_data(delivery_or_return, svctyp, counts_or_metrics, item, system, acc, date1, date2)
        })
        
        # counts_data_compared <- reactive({
        #         
        #         input$counts_refresh
        #         
        #         flashBack <- isolate(input$counts_monthOfFlashBack %>% as.numeric)
        #         
        #         
        #         delivery_or_return = isolate(counts_category())
        #         svctyp = isolate(counts_svctyp())
        #         counts_or_metrics = "counts"
        #         item = isolate(counts_item())
        #         system = isolate(counts_system())
        #         acc = isolate(counts_acc())
        #         date1 = isolate(counts_date1()) %m-% months(flashBack)
        #         date2 = isolate(counts_date2()) %m-% months(flashBack)
        #         
        #         counts_data_compared <- filter_data(delivery_or_return, svctyp, counts_or_metrics, item, system, acc, date1, date2)
        # })
        
        counts_results <- reactive({
                
                counts_data <- counts_data()
                
                counts_results <-
                        counts_data %>% 
                        results
        })
        
        # counts_results_compared <- reactive({
        #         
        #         counts_data_compared <- counts_data_compared()
        #         
        #         counts_results_compared <-
        #                 counts_data_compared %>% 
        #                 results
        # })
        
        
        output$counts_numberoforders <- renderText({
                
                counts_results <- counts_results()
                
                counts_results %>%
                        pull(value) %>%
                        sum(na.rm = T) %>%
                        accounting(format = "d") %>%
                        as.character
        })
        
        
        
        # Agents ------------------------------------------------------------------
        
        output$counts_download_raw_data <-
                downloadHandler(
                        filename = function(){
                                paste(
                                        "Agents_Raw_Data",
                                        counts_category(),
                                        counts_svctyp(),
                                        counts_acc(),
                                        counts_item(),
                                        "From",
                                        counts_date1(),
                                        "To",
                                        counts_date2(),
                                        "at",
                                        gsub(":","-",Sys.time()), ".csv", sep="_")
                        },
                        content = function(file) {
                                counts_data <- counts_data()
                                
                                counts_data() %>% write.csv(file, sep = ",", row.names = F)
                        })
        
        
        counts_agents <- reactive({
                
                counts_results = counts_results()
                Active_inactive <- input$counts_agent_activity
                
                counts_agents <-
                        counts_results %>%
                        group_by(Agent) %>%
                        dplyr::summarize(Volume = sum(value), Miles = sum(Miles)) %>%
                        mutate(
                                `Ave Miles` = (Miles/Volume) %>% round(2),
                                `% of Total` = (Volume / sum(Volume)) %>% round(4) %>% percent
                        ) %>% 
                        as.data.frame %>%
                        right_join(Agents %>% filter(Activity %in% Active_inactive), by = "Agent") %>%
                        filter(!is.na(Volume))
        })
        
        # counts_agents_compared <- reactive({
        #         
        #         counts_results_compared = counts_results_compared()
        #         Active_inactive <- input$counts_agent_activity
        #         
        #         counts_agents_compared <-
        #                 counts_results_compared %>%
        #                 group_by(Agent) %>%
        #                 dplyr::summarize(Volume = sum(value), Miles = sum(Miles)) %>%
        #                 mutate(
        #                         `Ave Miles` = (Miles/Volume) %>% round(2),
        #                         `% of Total` = (Volume / sum(Volume)) %>% round(4) %>% percent
        #                 ) %>% 
        #                 as.data.frame %>%
        #                 right_join(Agents %>% filter(Activity %in% Active_inactive), by = "Agent") %>%
        #                 filter(!is.na(Volume))
        # })
        
        output$counts_agents_map <- renderGvis({
                
                counts_agents <- counts_agents()
                
                Total = counts_agents %>% pull(Volume) %>% sum
                
                jscode = sprintf("var counts_agent_picked = data.getValue(chart.getSelection()[0].row,2);
                                 Shiny.onInputChange('%s', counts_agent_picked.toString())",
                                 session$ns('counts_agent_picked'))
                counts_agents %>%
                        mutate(
                                latlong = paste(Lat,Long, sep = ":"),
                                `Percentage of Total` = (Volume/Total)*100 %>% round(1)
                        ) %>%
                        gvisGeoChart("Agent", locationvar = "latlong", sizevar = "Percentage of Total", colorvar = "Volume",
                                     options =
                                             list(
                                                     region = "US",
                                                     resolution = "provinces",
                                                     displayMode = "markers",
                                                     colorAxis = "{colors:['#FFFFFF', 'green']}",
                                                     height = 347*map_scale,
                                                     width = 556*map_scale,
                                                     keepAspectRatio = F,
                                                     gvis.listener.jscode = jscode
                                             )
                        )
        })
        
        counts_agentstable_data <- reactive({
                
                counts_agents <- counts_agents()
                # counts_agents_compared <- counts_agents_compared()
                
                counts_agentstable <- 
                        counts_agents %>%
                        right_join(data.frame(state.name, State = state.abb, stringsAsFactors = FALSE), by = "State") %>%
                        as.data.frame %>%
                        select(Agent, Region = R_Region, State = state.name, Volume, `% of Total`, `Ave Miles`) %>%
                        # right_join(counts_agents_compared %>% select(Agent, `Volume Compared` = Volume, `% of Total Compared` = `% of Total`, `Ave Miles Compared` = `Ave Miles`), by = "Agent") %>%
                        filter(!is.na(Volume)) %>% 
                        arrange(-Volume)
        })
        
        output$counts_agentstable <- renderDT({
                
                counts_agentstable_data <- counts_agentstable_data()
                # monthOfFlashBack <- isolate(input$counts_monthOfFlashBack)
                
                # sketch = 
                #         htmltools::withTags(
                #                 table(
                #                         class = 'display',
                #                         thead(
                #                                 tr(
                #                                         th(rowspan = 2, 'Agent'),
                #                                         th(rowspan = 2, 'Region'),
                #                                         th(rowspan = 2, 'State'),
                #                                         th(colspan = 3, 'Interval'),
                #                                         th(colspan = 3, paste0(monthOfFlashBack, ' Months Flash Back'))
                #                                 ),
                #                                 tr(
                #                                         lapply(rep(c('Volume', '% of Total', 'Ave Miles'), 2), th)
                #                                 )
                #                         )
                #                 ))
                
                counts_agentstable <- 
                        counts_agentstable_data %>%
                        datatable(
                                options = list(
                                        dom = "tp",
                                        pageLength = 10,
                                        # initComplete = JS(
                                        #         "function(settings, json) {",
                                        #         "$(this.api().table().header()).css({'background-color': '#8895b6', 'color': '#000000'});",
                                        #         "}"),
                                        searchHighlight = TRUE
                                        # columnDefs = list(list(className = 'dt-center', targets = 3:ncol(counts_agentstable_data)))
                                ),
                                filter = 'top', 
                                # container = sketch, 
                                rownames = FALSE
                        )
                # formatStyle("Volume", backgroundColor = 'Aqua', fontWeight = 'bold') %>%
                # formatStyle("% of Total", backgroundColor = 'Azure', fontWeight = 'bold') %>%
                # formatStyle("Ave Miles", backgroundColor = 'Bisque', fontWeight = 'bold') %>% 
                # formatStyle("Volume Compared", backgroundColor = 'Aqua', fontWeight = 'bold') %>%
                # formatStyle("% of Total Compared", backgroundColor = 'Azure', fontWeight = 'bold') %>%
                # formatStyle("Ave Miles Compared", backgroundColor = 'Bisque', fontWeight = 'bold')
        })
        
        output$counts_download_agents <-
                downloadHandler(
                        filename = function(){
                                paste(
                                        "Agents",
                                        counts_category(),
                                        counts_svctyp(),
                                        counts_acc(),
                                        counts_item(),
                                        "From",
                                        counts_date1(),
                                        "To",
                                        counts_date2(),
                                        "at",
                                        gsub(":","-",Sys.time()), ".csv", sep="_")
                        },
                        content = function(file) {
                                counts_agentstable_data() %>% write.csv(file, sep = ",", row.names = F)
                        })
        
        counts_agent_modal_data <- reactive({
                
                delivery_or_return = isolate(counts_category())
                svctyp = isolate(counts_svctyp())
                counts_or_metrics = "counts"
                item = isolate(counts_item())
                system = isolate(counts_system())
                acc = isolate(counts_acc())
                date1 = as.Date("2013-01-01")
                date2 = as.Date(Sys.Date())
                
                ag = input$counts_agent_picked
                
                counts_agent_modal_data <-
                        filter_data(delivery_or_return, svctyp, counts_or_metrics, item, system, acc, date1, date2) %>% 
                        filter(Agent == ag)
        })
        output$counts_general_modal_graph <- renderGvis({
                
                counts_agent_modal_data <- counts_agent_modal_data()
                
                # p <- input$counts_general_period
                p <- "Weekly"
                
                pr <-
                        switch (p,
                                "Daily" = "day",
                                "Weekly" = "week",
                                "Monthly" = "month",
                                "Annually" = "year"
                        )
                dates <-
                        seq.Date(as.Date("2013-01-01"), as.Date(Sys.Date()), "days") %>%
                        as.data.frame %>%
                        rename_("Date" = ".") %>%
                        {if (pr == "day") . else mutate(., Date = cut(Date, breaks = pr) %>% as.Date)} %>%
                        distinct
                
                counts_general_modal_graph <-
                        counts_agent_modal_data %>%
                        {if (pr == "day") . else mutate(., Date = cut(Date, breaks = pr) %>% as.Date)} %>%
                        group_by(Date) %>%
                        dplyr::summarize(Count = n()) %>%
                        right_join(dates, by = "Date") %>%
                        mutate(Count = replace_na(Count, 0)) %>%
                        gvisAnnotationChart(datevar="Date",
                                            numvar="Count",
                                            options=list(
                                                    width=700, height=350,
                                                    fill=10, displayExactValues=TRUE,
                                                    colors="['#333d4f']"
                                            )
                        )
                
        })
        
        observeEvent(input$counts_agent_picked, {
                showModal(
                        modalDialog(
                                easyClose = T,
                                size = "l",
                                title = h2(input$counts_agent_picked),
                                paste0("Service Type: ") %>% strong, counts_svctyp(), br(),
                                paste0("Account: ") %>% strong, counts_acc(), br(),
                                paste0("Status: ") %>% strong, counts_item(), br(), br(),
                                htmlOutput("counts_general_modal_graph"),
                                footer = modalButton("Dismiss")
                        )
                )
        })
        
        
        output$counts_agents_barchart <- renderPlotly({
                
                counts_agents <- counts_agents()
                if (is.null(counts_agents)) return(NULL)
                
                counts_agents %>%
                        arrange(-Volume) %>% 
                        mutate(Agent = factor(Agent, levels = unique(Agent))) %>% 
                        plot_ly(
                                x = ~Agent,
                                y = ~Volume, 
                                color = I("#8c9ac2"), 
                                alpha = 1, 
                                type = "bar"
                        ) %>% 
                        layout(
                                xaxis = list(title = "Agent"), 
                                yaxis =  list(title = "Volume")
                        )
                
        })
        
        output$counts_agents_histogram <- renderPlotly({
                
                counts_agents <- counts_agents()
                if (is.null(counts_agents)) return(NULL)
                
                counts_agents %>%
                        plot_ly(
                                x = ~Volume, 
                                color = I("#8c9ac2"), 
                                alpha = 0.5, 
                                type = "histogram"
                        ) %>% 
                        layout(
                                title = "Histogram of the agents' volume",
                                xaxis = list(title = "Volume range"), 
                                yaxis =  list(title = "Count of agents")
                        )
                
        })
        
        output$counts_agents_boxplot <- renderPlotly({
                
                counts_agents <- counts_agents()
                if (is.null(counts_agents)) return(NULL)
                
                counts_agents %>%
                        plot_ly(
                                y = ~Volume, 
                                color = I("#8c9ac2"), 
                                alpha = 0.5, 
                                boxpoints = "suspectedoutliers",
                                type = "box",
                                name = "Boxplot"
                        ) %>% 
                        layout(
                                title = "Boxplot of the agents volume",
                                yaxis =  list(title = "Volume")
                        )
                
        })
        
        # States ------------------------------------------------------------------
        counts_states <- reactive({
                
                counts_results <- counts_results()
                counts_regions_states <- counts_regions_states()
                
                counts_states <-
                        counts_results %>%
                        arrange(Agent, -value) %>% 
                        mutate(dupl_agent = !duplicated(Agent)) %>% 
                        group_by(State) %>% 
                        dplyr::summarize(
                                Volume = sum(value, na.rm = T),
                                Miles = sum(Miles, na.rm = T),
                                `Agent Count` = sum(dupl_agent & !is.na(Agent) & Agent != "")
                        ) %>% 
                        mutate(`Ave Miles` = (Miles/Volume) %>% round(2)) %>% 
                        as.data.frame %>% 
                        right_join(counts_regions_states, by = "State") %>%
                        mutate(
                                Volume = replace_na(Volume, 0),
                                `Ave Miles` = replace_na(`Ave Miles`, 0),
                                `Agent Count` = replace_na(`Agent Count`, 0),
                                `% of Total` = (Volume / sum(Volume)) %>% round(4) %>% percent
                        )
        })
        # counts_states_compared <- reactive({
        #         
        #         counts_results_compared <- counts_results_compared()
        #         counts_regions_states <- counts_regions_states()
        #         
        #         counts_states <-
        #                 counts_results_compared %>%
        #                 arrange(Agent, -value) %>% 
        #                 mutate(dupl_agent = !duplicated(Agent)) %>% 
        #                 group_by(State) %>% 
        #                 dplyr::summarize(
        #                         Volume = sum(value, na.rm = T),
        #                         Miles = sum(Miles, na.rm = T),
        #                         `Agent Count` = sum(dupl_agent & !is.na(Agent) & Agent != "")
        #                 ) %>% 
        #                 mutate(`Ave Miles` = (Miles/Volume) %>% round(2)) %>% 
        #                 as.data.frame %>% 
        #                 right_join(counts_regions_states, by = "State") %>%
        #                 mutate(
        #                         Volume = replace_na(Volume, 0),
        #                         `Ave Miles` = replace_na(`Ave Miles`, 0),
        #                         `Agent Count` = replace_na(`Agent Count`, 0),
        #                         `% of Total` = (Volume / sum(Volume)) %>% round(4) %>% percent
        #                 )
        # })
        
        output$counts_states_map <- renderGvis({
                
                input$counts_refresh
                counts_states <- isolate(counts_states())
                
                gvisGeoChart(counts_states,
                             locationvar="State", colorvar="Volume",
                             options=list(region="US", 
                                          displayMode="regions", 
                                          showZoomOut = TRUE,
                                          resolution="provinces",
                                          colorAxis="{colors:['#FFFFFF', '#49609f']}",
                                          height = 347*map_scale,
                                          width = 556*map_scale,
                                          keepAspectRatio = F
                             )
                )
        })
        
        
        counts_states_table_data <- reactive({
                
                counts_states <- counts_states()
                # counts_states_compared <- counts_states_compared()
                
                counts_states_table <- 
                        counts_states %>%
                        select(State = State_Name, Region, `Agent Count`, Volume, `Ave Miles`, `% of Total`) %>%
                        # right_join(counts_states_compared %>% select(State = State_Name, `Agent Count Compared` = `Agent Count`, `Volume Compared` = Volume, `Ave Miles Compared` = `Ave Miles`, `% of Total Compared` = `% of Total`), by = "State") %>%
                        arrange(-Volume)
        })
        
        output$counts_states_table <- DT::renderDT({
                
                input$counts_refresh
                counts_states_table_data <- isolate(counts_states_table_data())
                
                counts_states_table <- 
                        counts_states_table_data %>%
                        datatable(
                                options = list(
                                        dom = "tp",
                                        pageLength = 10
                                        # initComplete = JS(
                                        #         "function(settings, json) {",
                                        #         "$(this.api().table().header()).css({'background-color': '#49609f', 'color': '#000000'});",
                                        #         "}")
                                ), 
                                filter = "top"
                        )
                # formatStyle("Agent Count", backgroundColor = 'CornflowerBlue', fontWeight = 'bold') %>%
                # formatStyle("Volume", backgroundColor = 'Aqua', fontWeight = 'bold') %>%
                # formatStyle("% of Total", backgroundColor = 'Azure', fontWeight = 'bold') %>%
                # formatStyle("Ave Miles", backgroundColor = 'Bisque', fontWeight = 'bold') %>%
                # formatStyle("Agent Count Compared", backgroundColor = 'CornflowerBlue', fontWeight = 'bold') %>%
                # formatStyle("Volume Compared", backgroundColor = 'Aqua', fontWeight = 'bold') %>%
                # formatStyle("% of Total Compared", backgroundColor = 'Azure', fontWeight = 'bold') %>%
                # formatStyle("Ave Miles Compared", backgroundColor = 'Bisque', fontWeight = 'bold')
        })
        
        output$counts_download_states <-
                downloadHandler(
                        filename = function(){
                                paste(
                                        "States",
                                        counts_category(),
                                        counts_svctyp(),
                                        counts_acc(),
                                        counts_item(),
                                        "From",
                                        counts_date1(),
                                        "To",
                                        counts_date2(),
                                        "at",
                                        gsub(":","-",Sys.time()), ".csv", sep="_")
                        },
                        content = function(file) {
                                counts_states_table_data() %>% write.csv(file, sep = ",", row.names = F)
                        })
        
        
        # Regions -----------------------------------------------------------------
        
        counts_regions <- reactive({
                
                counts_results <- counts_results()
                
                
                counts_regions <-
                        counts_results %>%
                        arrange(Agent, -value) %>% 
                        mutate(dupl_agent = !duplicated(Agent)) %>% 
                        group_by(Region) %>%
                        dplyr::summarize(
                                Volume = sum(value, na.rm = T),
                                Miles = sum(Miles, na.rm = T),
                                `Agent Count` = sum(dupl_agent & !is.na(Agent) & Agent != "")
                        ) %>%
                        mutate(
                                Volume = replace_na(Volume, 0),
                                `Ave Miles` = (Miles/Volume) %>% round(2) %>% replace_na(0),
                                `Agent Count` = replace_na(`Agent Count`, 0),
                                `% of Total` = (Volume / sum(Volume)) %>% round(4) %>% percent
                        ) %>% 
                        as.data.frame
        })
        # counts_regions_compared <- reactive({
        # 
        #         counts_results_compared <- counts_results_compared()
        #         
        #         
        #         counts_regions_compared <-
        #                 counts_results_compared %>%
        #                 arrange(Agent, -value) %>% 
        #                 mutate(dupl_agent = !duplicated(Agent)) %>% 
        #                 group_by(Region) %>%
        #                 dplyr::summarize(
        #                         Volume = sum(value, na.rm = T),
        #                         Miles = sum(Miles, na.rm = T),
        #                         `Agent Count` = sum(dupl_agent & !is.na(Agent) & Agent != "")
        #                 ) %>%
        #                 mutate(
        #                         Volume = replace_na(Volume, 0),
        #                         `Ave Miles` = (Miles/Volume) %>% round(2) %>% replace_na(0),
        #                         `Agent Count` = replace_na(`Agent Count`, 0),
        #                         `% of Total` = (Volume / sum(Volume)) %>% round(4) %>% percent
        #                 ) %>% 
        #                 as.data.frame
        # })
        
        output$counts_regions_map <- renderGvis({
                
                counts_regions <- counts_regions()
                counts_regions_states <- isolate(counts_regions_states())
                
                data <- 
                        counts_regions %>% 
                        left_join(counts_regions_states, by = "Region")
                
                gvisGeoChart(data %>% filter(!is.na(State), !is.na(Region), !is.na(Volume)),
                             locationvar="State", colorvar="Volume",
                             options=list(region="US", 
                                          displayMode="regions", 
                                          showZoomOut = TRUE,
                                          resolution="provinces",
                                          colorAxis="{colors:['#FFFFFF', '#49609f']}",
                                          height = 347*map_scale,
                                          width = 556*map_scale,
                                          keepAspectRatio = F
                             )
                )
        })
        
        output$counts_regions_table <- DT::renderDT({
                
                counts_regions <- counts_regions()
                # counts_regions_compared <- counts_regions_compared()
                
                counts_regions_table <- 
                        counts_regions %>%
                        select(Region, `Agent Count`, Volume, `Ave Miles`, `% of Total`) %>%
                        # right_join(counts_regions_compared %>% select(Region, `Agent Count Compared` = `Agent Count`, `Volume Compared` = Volume, `Ave Miles Compared` = `Ave Miles`, `% of Total Compared` = `% of Total`), by = "Region") %>%
                        arrange(-Volume) %>%
                        datatable(
                                options = list(
                                        dom = "Btp",
                                        buttons = c('copy', 'csv', 'excel'),
                                        pageLength = 10
                                        # initComplete = JS(
                                        #         "function(settings, json) {",
                                        #         "$(this.api().table().header()).css({'background-color': '#49609f', 'color': '#000000'});",
                                        #         "}")
                                ),
                                extensions = 'Buttons'
                        )
        })
        
        output$counts_regions_pie <- renderPlotly({
                
                counts_regions <- counts_regions()
                
                counts_regions %>%
                        arrange(Region) %>%
                        plot_ly(labels = ~Region, values = ~Volume, type = 'pie',
                                textinfo = 'label',
                                textfont = list(size = 15),
                                marker = list(line = list(color = "#000000", width = 0.5))
                        ) %>%
                        layout(font = list(
                                family = "Helvetica Neue",
                                size = 14,
                                color = 'black'),
                               legend = list(orientation = 'h')
                        )
        })
        
        
        
        
        # Metrics ----------------------------------------------------------------------
        
        metrics_system <- reactive({input$metrics_system})
        metrics_svctyp <- reactive({input$metrics_servicetype})
        metrics_acc <- reactive({input$metrics_account})
        metrics_item <- reactive({input$metrics_item})
        metrics_date1 <- reactive({input$metrics_daterange[1]})
        metrics_date2 <- reactive({input$metrics_daterange[2]})
        
        metrics_category <- reactive({
                
                svctyp = metrics_svctyp()
                
                metrics_category <- category_from_servicetype(svctyp)
                
        })
        
        metrics_regions_states <- reactive({
                
                metrics_category <- metrics_category()
                
                metrics_regions_states <-
                        switch (metrics_category,
                                "Delivery" = LMRegions_State %>% rename(Region = LM_Region),
                                "Return" = RRegions_State %>% rename(Region = R_Region)
                        )
        })
        
        metrics_items_choices <- reactive({
                
                # statuses are different for deliveries an returns.
                # from choice of service_type we got the category(delivery or return).
                # from category we'll get statuses.
                # this will update the drop-down menue of Status
                metrics_category = metrics_category()
                
                metrics_items_choices <- 
                        Metrics %>% 
                        filter(category == metrics_category) %>% 
                        pull(caption)
        })
        # update counts_item
        observe({
                updateSelectInput(session, "metrics_item", choices = metrics_items_choices())
        })
        
        # show modal if date1 > date2
        observe({
                input$metrics_refresh
                
                date1 = isolate(metrics_date1())
                date2 = isolate(metrics_date2())
                
                if (date1 > date2) {
                        showModal(modalDialog(
                                title = "Date range is not acceptable",
                                "Beginning date should be smaller than or equal to the end date.",
                                easyClose = TRUE,
                                footer = NULL
                        ))
                }
        })
        
        metrics_data <- reactive({
                
                input$metrics_refresh
                
                delivery_or_return = isolate(metrics_category())
                svctyp = isolate(metrics_svctyp())
                counts_or_metrics = "metrics"
                item = isolate(metrics_item())
                system = isolate(metrics_system())
                acc = isolate(metrics_acc())
                date1 = isolate(metrics_date1())
                date2 = isolate(metrics_date2())
                
                metrics_data <- filter_data(delivery_or_return, svctyp, counts_or_metrics, item, system, acc, date1, date2)
        })
        
        metrics_results <- reactive({
                
                metrics_data <- metrics_data()
                
                metrics_results <-
                        metrics_data %>% 
                        results
        })
        
        output$metrics_average <- renderText({
                
                metrics_data <- metrics_data()
                
                svctyp = isolate(metrics_svctyp())
                type <- type_from_metric(isolate(metrics_item()), svctyp)
                
                metrics_data %>%
                        pull(Value) %>%
                        mean(na.rm = T) %>%
                        round(2) %>% 
                        accounting %>% 
                        {
                                if (type == "number") {
                                        .
                                } else if (type == "percentage") {
                                        percent(.)
                                }
                        }
        })
        
        
        
        # WOW ---------------------------------------------------------------------
        
        period_picked <- reactive({
                switch (input$metrics_wow_periods,
                        "Week" = "weeks",
                        "Month" = "months",
                        "Year" = "years"
                )
        })
        
        wow_table_row <- function(title_item, catgr){
                
                input$metrics_refresh
                pr = period_picked()
                
                system = isolate(input$metrics_system)
                acc = isolate(input$metrics_account)
                date1 = isolate(input$metrics_daterange[1])
                date2 = isolate(input$metrics_daterange[2])
                
                svctyp = WOW %>% filter(title == title_item, category == catgr) %>% pull(service_type)
                delivery_or_return = catgr
                counts_or_metrics = "metrics"
                metric = WOW %>% filter(title == title_item, category == catgr) %>% pull(metric)
                value_format <- WOW %>% filter(title == title_item, category == catgr) %>% pull(format)
                
                wow_table_row <-
                        filter_data(
                                delivery_or_return,
                                svctyp,
                                counts_or_metrics,
                                metric,
                                system,
                                acc,
                                date1,
                                date2
                        ) %>% 
                        select(Value,Date) %>% 
                        mutate(
                                Date = Date %>% floor_date(unit = pr, week_start = week_start)
                        ) %>% 
                        group_by(Date) %>% 
                        dplyr::summarize(result = mean(Value) %>% round(2)) %>% 
                        {if (value_format == "percentage") mutate(., result = result %>% percent) else mutate(., result = result %>% as.character)}
                names(wow_table_row)[2] <- title_item
                wow_table_row
        }
        wow_create_table <- function(cat){
                lapply(
                        WOW %>% filter(category == cat) %>% pull(title), 
                        wow_table_row, catgr = cat
                ) %>% 
                        reduce(left_join, by = "Date") %>% 
                        arrange(desc(Date)) %>% 
                        mutate(Date = paste0(isolate(input$metrics_wow_periods), " starting from ", Date))
        }
        wow_delivery_table_data <- reactive({wow_create_table("Delivery")})
        wow_return_table_data <- reactive({wow_create_table("Return")})
        
        output$metrics_wow_delivery_table <- DT::renderDT({
                
                wow_delivery_table_data() %>%
                        datatable(
                                options = list(
                                        dom = "Btp",
                                        buttons = c('copy', 'csv', 'excel'),
                                        pageLength = 4,
                                        initComplete = JS(
                                                "function(settings, json) {",
                                                "$(this.api().table().header()).css({'background-color': '#ed470b', 'color': '#000000'});",
                                                "}"),
                                        columnDefs = list(list(className = 'dt-center', targets = 1:(WOW %>% filter(category == "Delivery") %>% nrow) + 1))
                                ), 
                                extensions = 'Buttons'
                        )
        })
        output$metrics_wow_return_table <- DT::renderDT({
                
                wow_return_table_data() %>%
                        datatable(
                                options = list(
                                        dom = "Btp",
                                        buttons = c('copy', 'csv', 'excel'),
                                        pageLength = 4,
                                        initComplete = JS(
                                                "function(settings, json) {",
                                                "$(this.api().table().header()).css({'background-color': '#2180d3', 'color': '#000000'});",
                                                "}"),
                                        columnDefs = list(list(className = 'dt-center', targets = 1:(WOW %>% filter(category == "Return") %>% nrow) + 1))
                                ), 
                                extensions = 'Buttons'
                        )
        })
        
        
        
        # Region ------------------------------------------------------------------
        
        metrics_regions <- reactive({
                
                metrics_data <- metrics_data()
                metrics_regions_states <- isolate(metrics_regions_states())
                
                metrics_regions <-
                        metrics_data %>%
                        group_by(Region) %>%
                        dplyr::summarize(Volume = n(),
                                         value = mean(Value, na.rm = T)) %>%
                        left_join(metrics_regions_states, by = "Region") %>%
                        as.data.frame
                
        })
        
        output$metrics_regions_map <- renderGvis({
                
                metrics_regions <- metrics_regions()
                svctyp = isolate(metrics_svctyp())
                
                gvisGeoChart(metrics_regions %>% filter(!is.na(State), !is.na(Region), !is.na(value), !is.na(Volume)),
                             locationvar="State", colorvar="value",
                             options=list(region="US", 
                                          displayMode="regions", 
                                          showZoomOut = TRUE,
                                          resolution="provinces",
                                          colorAxis = paste0("{colors:['", lowest_color(isolate(metrics_item()), svctyp), "', '", highest_color(isolate(metrics_item()), svctyp), "']}"),
                                          height = 347*map_scale,
                                          width = 556*map_scale,
                                          keepAspectRatio = F
                             )
                )
        })
        
        output$metrics_regions_table <- DT::renderDT({
                metrics_regions <- metrics_regions()
                metrics_results <- metrics_results()
                
                metrics_regions %>%
                        select(Region, Volume, value) %>% 
                        mutate(
                                value = value %>% round(2), 
                                Volume = replace_na(Volume, 0)
                        ) %>%
                        as.data.frame %>%
                        distinct %>% 
                        arrange(-value) %>%
                        datatable(
                                options = list(
                                        dom = "tp",
                                        pageLength = 10,
                                        initComplete = JS(
                                                "function(settings, json) {",
                                                "$(this.api().table().header()).css({'background-color': '#8895b6', 'color': '#000000'});",
                                                "}")
                                )
                        )
        })
        
        
        # States ------------------------------------------------------------------
        metrics_states <- reactive({
                
                metrics_data <- metrics_data()
                metrics_regions_states <- isolate(metrics_regions_states())
                
                metrics_states <-
                        metrics_data %>%
                        group_by(State) %>% 
                        dplyr::summarize(Volume = n(),
                                         value = mean(Value, na.rm = T)) %>% 
                        as.data.frame %>% 
                        right_join(metrics_regions_states, by = "State") %>% 
                        mutate(value = value %>% round(2))
        })
        
        output$metrics_states_map <- renderGvis({
                
                metrics_states <- metrics_states()
                svctyp = isolate(metrics_svctyp())
                
                gvisGeoChart(metrics_states,
                             locationvar = "State", 
                             colorvar = "value",
                             options = list(region = "US", 
                                            displayMode = "regions", 
                                            showZoomOut = TRUE,
                                            resolution = "provinces",
                                            colorAxis = paste0("{colors:[\'", lowest_color(isolate(metrics_item()), svctyp), "', \'", highest_color(isolate(metrics_item()), svctyp), "']}"),
                                            height = 347 * map_scale,
                                            width = 556 * map_scale,
                                            keepAspectRatio = F
                             )
                )
        })
        
        output$metrics_states_table <- DT::renderDT({
                
                metrics_states <- metrics_states()
                
                svctyp = isolate(metrics_svctyp())
                type <- type_from_metric(isolate(metrics_item()), svctyp)
                
                metrics_states %>%
                        filter(!is.na(value), !is.na(Volume)) %>%
                        arrange(-value) %>%
                        mutate(value = value %>%
                                       round(2) %>% 
                                       accounting %>% 
                                       {
                                               if (type == "number") {
                                                       .
                                               } else if (type == "percentage") {
                                                       percent(.)
                                               }
                                       }
                        ) %>% 
                        select(State = State_Name, Region, Value = value) %>%
                        datatable(
                                options = list(
                                        dom = "ftp",
                                        pageLength = 10,
                                        initComplete = JS(
                                                "function(settings, json) {",
                                                "$(this.api().table().header()).css({'background-color': '#8895b6', 'color': '#000000'});",
                                                "}")
                                )
                        )
                
        })
        
        output$metrics_download_states <-
                downloadHandler(
                        filename = function(){
                                paste(
                                        "States",
                                        metrics_category(),
                                        metrics_svctyp(),
                                        metrics_acc(),
                                        metrics_item(),
                                        "From",
                                        metrics_date1(),
                                        "To",
                                        metrics_date2(),
                                        "at",
                                        gsub(":","-",Sys.time()), ".csv", sep="_")
                        },
                        content = function(file) {
                                metrics_states() %>%
                                        select(State, State_Name, Region, Volume, Value = value) %>%
                                        arrange(-Value) %>%
                                        write.csv(file, sep = ",", row.names = F)
                        })
        
        
        # Agents ------------------------------------------------------------------
        
        metrics_agents <- reactive({
                
                metrics_data <- metrics_data()
                Active_inactive <- input$metrics_agent_activity
                
                metrics_agents <-
                        metrics_data %>%
                        group_by(Agent) %>% 
                        dplyr::summarize(Volume = n(),
                                         value = mean(Value, na.rm = T)) %>% 
                        as.data.frame %>% 
                        inner_join(Agents %>% filter(Activity %in% Active_inactive), by = "Agent") %>%
                        filter(!is.na(value), !is.na(Volume)) %>% 
                        mutate(value = value %>% round(2), 
                               latlong = paste(Lat,Long, sep = ":")
                        ) %>% 
                        arrange(-Volume)
                
        })
        output$metrics_agents_map <- renderGvis({
                
                metrics_agents <- metrics_agents()
                svctyp = isolate(metrics_svctyp())
                
                jscode = sprintf("var metrics_agent_picked = data.getValue(chart.getSelection()[0].row,2);
                                 Shiny.onInputChange('%s', metrics_agent_picked.toString())",
                                 session$ns('metrics_agent_picked'))
                metrics_agents %>%
                        mutate(`Data Available` = Volume) %>% 
                        gvisGeoChart("Agent", locationvar = "latlong", colorvar = "value", sizevar = "Data Available",
                                     options =
                                             list(
                                                     region = "US",
                                                     resolution = "provinces",
                                                     displayMode = "markers",
                                                     colorAxis = paste0("{colors:['", lowest_color(isolate(metrics_item()), svctyp), "', '", highest_color(isolate(metrics_item()), svctyp), "']}"),
                                                     height = 347*map_scale,
                                                     width = 556*map_scale,
                                                     keepAspectRatio = F,
                                                     gvis.listener.jscode = jscode
                                             )
                        )
        })
        
        output$metrics_agentstable <- renderDT({
                
                metrics_agents <- metrics_agents()
                
                svctyp = isolate(metrics_svctyp())
                type <- type_from_metric(isolate(metrics_item()), svctyp)
                
                metrics_agents %>%
                        filter(!is.na(Volume), !is.na(value)) %>%
                        arrange(-value) %>%
                        mutate(
                                `Data Available` = Volume,
                                value = value %>%
                                        round(2) %>% 
                                        accounting %>% 
                                        {
                                                if (type == "number") {
                                                        .
                                                } else if (type == "percentage") {
                                                        percent(.)
                                                }
                                        }
                        ) %>% 
                        select(Agent, State = State_Name, `Data Available`, Value = value) %>%
                        datatable(
                                options = list(
                                        dom = "ftp",
                                        pageLength = 10,
                                        initComplete = JS(
                                                "function(settings, json) {",
                                                "$(this.api().table().header()).css({'background-color': '#8895b6', 'color': '#000000'});",
                                                "}")
                                )
                        )
        })
        
        output$metrics_download_agents <-
                downloadHandler(
                        filename = function(){
                                paste(
                                        "Agents",
                                        counts_category(),
                                        counts_svctyp(),
                                        counts_acc(),
                                        counts_item(),
                                        "From",
                                        counts_date1(),
                                        "To",
                                        counts_date2(),
                                        "at",
                                        gsub(":","-",Sys.time()), ".csv", sep="_")
                        },
                        content = function(file) {
                                metrics_agents() %>%
                                        arrange(-Volume) %>%
                                        select(Agent, City, State, `Available Data` = Volume, value) %>%
                                        write.csv(file, sep = ",", row.names = F)
                        })
        
        metrics_agent_modal_data <- reactive({
                
                delivery_or_return = isolate(metrics_category())
                svctyp = isolate(metrics_svctyp())
                counts_or_metrics = "metrics"
                item = isolate(metrics_item())
                system = isolate(metrics_system())
                acc = isolate(metrics_acc())
                date1 = as.Date("2013-01-01")
                date2 = as.Date(Sys.Date())
                
                ag = input$metrics_agent_picked
                
                metrics_agent_modal_data <-
                        filter_data(delivery_or_return, svctyp, counts_or_metrics, item, system, acc, date1, date2) %>%
                        filter(Agent == ag)
        })
        output$metrics_general_modal_graph <- renderGvis({
                
                metrics_agent_modal_data <- metrics_agent_modal_data()
                
                # p <- input$counts_general_period
                p <- "Weekly"
                
                pr <-
                        switch (p,
                                "Daily" = "day",
                                "Weekly" = "week",
                                "Monthly" = "month",
                                "Annually" = "year"
                        )
                dates <-
                        seq.Date(as.Date("2013-01-01"), as.Date(Sys.Date()), "days") %>%
                        as.data.frame %>%
                        rename_("Date" = ".") %>%
                        {if (pr == "day") . else mutate(., Date = cut(Date, breaks = pr) %>% as.Date)} %>%
                        distinct
                
                metrics_general_modal_graph <-
                        metrics_agent_modal_data %>%
                        {if (pr == "day") . else mutate(., Date = cut(Date, breaks = pr) %>% as.Date)} %>%
                        group_by(Date) %>%
                        dplyr::summarize(Volume = n(),
                                         value = mean(Value, na.rm = T)) %>% 
                        right_join(dates, by = "Date") %>%
                        mutate(Volume = replace_na(Volume, 0)) %>%
                        gvisAnnotationChart(datevar="Date",
                                            numvar="value",
                                            options=list(
                                                    width=700, height=350,
                                                    fill=10, displayExactValues=TRUE,
                                                    colors="['#ed6d5a']")
                        )
                
        })
        
        observeEvent(input$metrics_agent_picked, {
                showModal(
                        modalDialog(
                                easyClose = T,
                                size = "l",
                                title = h2(input$metrics_agent_picked),
                                paste0("Service Type: ") %>% strong, metrics_svctyp(), br(),
                                paste0("Account: ") %>% strong, metrics_acc(), br(),
                                paste0("Metrics: ") %>% strong, metrics_item(), br(), br(),
                                htmlOutput("metrics_general_modal_graph"),
                                footer = modalButton("Dismiss")
                        )
                )
        })
        
        # observeEvent(input$metrics_scorecards_agent, {
        #         
        #         number_of_weeks <- 4
        #         
        #         delivery_metrics_needed <-
        #                 c("Receiving LeadTime",
        #                   "LM Transit",
        #                   "Window Compliance",
        #                   "Date Compliance",
        #                   "First Note Compliance"
        #                 )
        #         delivery_metrics_format <- c("#", "#", "%", "%", "%")
        #         return_metrics_needed <-
        #                 c("Job Date to Pick UP",
        #                   "Pick UP to LTL POD"
        #                 )
        #         return_metrics_format <- c("#", "#")
        #         
        #         
        #         
        #         average_4_weeks <- function(agent, metric, return_designation, number_of_weeks) {
        #                 data <-
        #                         switch (return_designation,
        #                                 "Delivery" = Deliveries,
        #                                 "Return" = Returns
        #                         ) %>%
        #                         filter(Agent == agent)
        #                 column <-
        #                         switch (return_designation,
        #                                 "Delivery" = Metrics %>% filter(category == "Delivery", caption == metric) %>% pull(column),
        #                                 "Return" = Metrics %>% filter(category == "Return", caption == metric) %>% pull(column)
        #                         )
        #                 date_column <-
        #                         switch (return_designation,
        #                                 "Delivery" = Metrics %>% filter(category == "Delivery", caption == metric) %>% pull(date_column),
        #                                 "Return" = Metrics %>% filter(category == "Return", caption == metric) %>% pull(date_column)
        #                         )
        #                 days_needed <- (number_of_weeks + 1) * 7
        #                 data_needed <-
        #                         data %>%
        #                         select(column = column, date_column = date_column) %>%
        #                         filter(!is.na(column), !is.na(date_column), date_column > isolate(today()) - days_needed)
        #                 
        #                 all_NAs <- data.frame(Week = as.Date(cut(c(today()-7*(1:4)), breaks = "week")), average = rep(NA, 4))
        #                 
        #                 if (!nrow(data_needed)) {
        #                         final_results <- all_NAs
        #                 } else {
        #                         results <-
        #                                 data_needed %>%
        #                                 mutate(Week = as.Date(cut(date_column, breaks = "week"))) %>%
        #                                 group_by(Week) %>%
        #                                 dplyr::summarize(average = mean(column, na.rm = TRUE)) %>%
        #                                 filter(!Week %in% c(cut(isolate(today()), breaks = "week") %>% as.Date, cut(today() - days_needed, breaks = "week") %>% as.Date))
        #                         if (!nrow(results)) {
        #                                 final_results <- all_NAs
        #                         } else {
        #                                 final_results <-
        #                                         results %>%
        #                                         mutate(average = round(average, 2)) %>%
        #                                         as.data.frame %>%
        #                                         merge(all_NAs %>% select(Week), by = "Week", all.y = T)
        #                         }
        #                 }
        #                 
        #                 names(final_results) <- c("Week", metric)
        #                 final_results
        #         }
        #         
        #         agent_category_metrics_table <- function(agent, return_designation, number_of_weeks) {
        #                 result <-
        #                         switch (return_designation,
        #                                 "Delivery" = delivery_metrics_needed,
        #                                 "Return" = return_metrics_needed
        #                         ) %>%
        #                         lapply(function(metric) average_4_weeks(agent, metric, return_designation, number_of_weeks)) %>%
        #                         reduce(left_join, by = "Week") %>%
        #                         `rownames<-`(.$Week) %>%
        #                         select(-Week) %>%
        #                         t
        #                 result
        #         }
        #         
        #         agent_scorecard_dir <- "//nsdstor1/SHARED/Agent Scorecards/"
        #         agent_scorecard_subdir <- paste0("//nsdstor1/SHARED/Agent Scorecards/", isolate(gsub(":","-",Sys.time())))
        #         agent_scorecard_subdir_deliveryregions <- paste0(agent_scorecard_subdir, "/", LMRegions_LMC %>% pull(LM_Region))
        #         dir.create(agent_scorecard_dir, showWarnings = FALSE)
        #         dir.create(agent_scorecard_subdir)
        #         sapply(agent_scorecard_subdir_deliveryregions, dir.create)
        #         wd <- isolate(getwd())
        #         isolate(setwd(agent_scorecard_subdir))
        #         
        #         
        #         active_agents <-
        #                 Agents %>%
        #                 filter(Activity == "Active") %>%
        #                 pull(Agent)
        #         
        #         count_of_active_agents <- active_agents %>% length
        #         
        #         withProgress(message = 'Creating agent scorecards', value = 0, {
        #                 
        #                 for (i in 1:count_of_active_agents) {
        #                         
        #                         agent_name <- active_agents[i]
        #                         agent_region <- Agents %>% filter(Agent == agent_name) %>% pull(LM_Region)
        #                         incProgress(
        #                                 1/count_of_active_agents,
        #                                 message = 'Creating agent scorecards',
        #                                 detail = paste(i, "/", count_of_active_agents," ", agent_name)
        #                         )
        #                         Delivery_table <- agent_category_metrics_table(agent_name, "Delivery", number_of_weeks)
        #                         Delivery_table[delivery_metrics_format == "%",] <- ifelse(is.na(Delivery_table[delivery_metrics_format == "%",]), NA, sprintf("%1.2f%%", 100*as.numeric(Delivery_table[delivery_metrics_format == "%",])))
        #                         Return_table <- agent_category_metrics_table(agent_name, "Return", number_of_weeks)
        #                         save(agent_name, Delivery_table, Return_table, file='//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/Hamed/tables.Rda')
        #                         output_dir <- paste0(agent_scorecard_subdir, "/", agent_region)
        #                         rmarkdown::render('//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/Hamed/Agent_Scorecard_Template.Rmd', "pdf_document", 
        #                                           clean = FALSE, 
        #                                           output_dir = output_dir
        #                         )
        #                         file.rename(paste0(output_dir, "/Agent_Scorecard_Template.pdf"), paste0(output_dir, "/", agent_name, ".pdf"))
        #                 }
        #         })
        #         isolate(setwd(wd))
        # })        
        # 
        
        output$metrics_agents_barchart <- renderPlotly({
                
                metrics_agents <- metrics_agents()
                if (is.null(metrics_agents)) return(NULL)
                
                metrics_agents %>%
                        arrange(-value) %>% 
                        mutate(Agent = factor(Agent, levels = unique(Agent))) %>% 
                        plot_ly(
                                x = ~Agent,
                                y = ~value, 
                                color = I("#8c9ac2"), 
                                alpha = 1, 
                                type = "bar"
                        ) %>% 
                        layout(
                                title = "Bar chart of the agents' metrics",
                                xaxis = list(title = "Agent"), 
                                yaxis =  list(title = "Metrics")
                        )
                
        })
        
        output$metrics_agents_histogram <- renderPlotly({
                
                metrics_agents <- metrics_agents()
                if (is.null(metrics_agents)) return(NULL)
                
                metrics_agents %>%
                        plot_ly(
                                x = ~value, 
                                color = I("#8c9ac2"), 
                                alpha = 0.5, 
                                type = "histogram"
                        ) %>% 
                        layout(
                                title = "Histogram of the agents' metrics",
                                xaxis = list(title = "Metric range"), 
                                yaxis =  list(title = "Count of agents")
                        )
                
        })
        
        output$metrics_agents_boxplot <- renderPlotly({
                
                metrics_agents <- metrics_agents()
                if (is.null(metrics_agents)) return(NULL)
                
                metrics_agents %>%
                        plot_ly(
                                y = ~value, 
                                color = I("#8c9ac2"), 
                                alpha = 0.5, 
                                boxpoints = "suspectedoutliers",
                                type = "box",
                                name = "Boxplot"
                        ) %>% 
                        layout(
                                title = "Boxplot of the agents metrics",
                                yaxis =  list(title = "Metrics")
                        )
                
        })
        
        
        # JCP ---------------------------------------------------------------------
        create.retail <- function(retail.file.name) {
                # read data from cda text file
                # retail.file.name <- "Amelia 06.14.txt"
                retail_table <- read.table(retail.file.name, sep="|", header=FALSE, fill=TRUE, comment.char="", quote = "")
                # return NA if there is no items in the file
                if (retail_table %>% filter(V1 == 1) %>% nrow == 0) return(NULL)
                # specify columns need from each of the code tables and their names
                order_detail <- c(PO_Number = 3,
                                  Ref_2 = 6,
                                  Customer_Last_Name = 16,
                                  Customer_First_Name = 17,
                                  Customer_Street_Address_1 = 19,
                                  Customer_Street_Address_2 = 20,
                                  Customer_Street_Address_3 = 21,
                                  Customer_City = 22,
                                  Customer_State = 23,
                                  Customer_Zip = 24,
                                  Customer_Phone = 25,
                                  Customer_Cell_Phone = 27,
                                  Customer_Email = 33)
                order_item <- c(PO_Number = 3,
                                Piece_Count = 7,
                                Sub_Lot = 10,
                                Product_Description = 11)
                spec_instruction <- c(PO_Number = 3, Special_Instruction = 6)
                
                # specify code tables to create the retail table
                table_codes <- c(order_detail = 1, order_item = 2, spec_instruction = 4)
                #create code tables
                for (item in 1:length(table_codes)){
                        table.name <- names(table_codes)[item]
                        code <- table_codes[item]
                        table <- 
                                retail_table %>% 
                                filter(V1 == table_codes[item]) %>% 
                                select(
                                        switch (table.name,
                                                "order_detail" = order_detail,
                                                "order_item" = order_item,
                                                "spec_instruction" = spec_instruction
                                        )
                                ) %>% 
                                lapply(as.character) %>% 
                                as.data.frame(stringsAsFactors=FALSE) %>% 
                                # trim whitespaces in both sides
                                lapply(trimws) %>% 
                                as.data.frame(stringsAsFactors=FALSE)
                        
                        # name the table
                        assign(paste0("retail_0",code,"_",table.name),table)
                        # delete temporary variable
                        rm(table)
                        
                }
                
                
                # add Sub and Lot columns to order items table
                if (exists("retail_02_order_item")) {
                        retail_02_order_item <-
                                retail_02_order_item %>% 
                                mutate(Sub = substr(Sub_Lot, 1, 3), Lot = substr(Sub_Lot, 5, 8))
                }
                
                # aggregate retail_04_spec_instruction rows
                if (exists("retail_04_spec_instruction") & nrow(retail_04_spec_instruction) > 1) {
                        retail_04_spec_instruction <- 
                                aggregate(
                                        retail_04_spec_instruction$Special_Instruction, 
                                        by = list(PO_Number=retail_04_spec_instruction$PO_Number), 
                                        paste, 
                                        collapse = ""
                                ) %>% 
                                rename_("Special_Instruction" = "x")
                }
                
                ## create the final dataframe of order items
                
                create.retail <- 
                        retail_02_order_item %>% 
                        merge(retail_01_order_detail,by="PO_Number") %>% 
                        merge(retail_04_spec_instruction,by="PO_Number",all.x = TRUE) %>% 
                        merge(EFF(), by = c("Sub","Lot"), all.x = TRUE) %>% 
                        mutate(Customer_Name = paste(Customer_First_Name, Customer_Last_Name), Identifier = "Retail") %>% 
                        select(PO_Number, 
                               Sub_Lot, 
                               Product_Description, 
                               Piece_Count, 
                               Ref_2, 
                               Length2, 
                               Width, 
                               Height, 
                               Weight,
                               `Delivery Experience`, 
                               Customer_Name, 
                               Customer_Street_Address_1, 
                               Customer_Street_Address_2,
                               Customer_Street_Address_3, 
                               Customer_City, 
                               Customer_State, 
                               Customer_Zip, 
                               Customer_Phone,
                               Customer_Cell_Phone, 
                               Customer_Email, 
                               Special_Instruction, 
                               Identifier
                        )
        }
        
        create.online.input <- function(online.input.filename) {
                
                online <- 
                        read_csv(online.input.filename) %>% 
                        {if (nrow(.) < 2) return(NULL) else .} %>% 
                        as.data.frame %>% 
                        # delete empty rows
                        filter(rowSums(is.na(.)) != ncol(.)) %>% 
                        filter(!duplicated(.))
                
                # make the column names uniform
                colnames(online) <- LTL.colnames()
                rownames(online) <- seq(length = nrow(online))
                return(online)
        }
        
        
        EFF <- reactive({
                
                EFF_button <- input$JCP_input_EFF
                
                EFF_path <- 
                        if (is.null(EFF_button)) {
                                "//nsdstor1/SHARED/JCP_R/EFF/EFF.csv"
                        } else {
                                EFF_button$datapath
                        }
                
                EFF <- 
                        EFF_path %>%
                        # "//nsdstor1/SHARED/JCP_R/EFF/EFF.csv" %>%
                        read.csv(stringsAsFactors = FALSE, check.names = F, fileEncoding="UTF-8-BOM") %>%
                        as.data.frame %>%                                               
                        filter(rowSums(is.na(.)) != ncol(.)) %>%                        
                        filter(!duplicated(.)) %>%                                      
                        filter(!is.na(Sub), !is.na(Lot)) %>%                            
                        mutate(Sub = as.character(Sub), Lot = as.character(Lot)) %>%    
                        mutate(Lot = sprintf("%04d", as.numeric(Lot))) %>%              
                        select(Sub, Lot, Length2, Width, Height, Weight, `Delivery Experience`)
                
                
        })
        
        retail <- reactive({
                
                if (is.null(input$JCP_input_retail)) return(NULL)
                
                retail <- 
                        input$JCP_input_retail %>%
                        pull(datapath) %>%
                        lapply(create.retail) %>% 
                        list.clean %>% 
                        list.rbind %>% 
                        mutate(Weight = Weight*as.numeric(Piece_Count)) %>% 
                        distinct
                
        })        
        
        LTL <- reactive({
                
                if (is.null(input$JCP_input_LTL)) return(NULL)
                
                LTL <- 
                        input$JCP_input_LTL %>%
                        pull(datapath) %>%
                        read_csv %>% 
                        as.data.frame %>% 
                        filter(rowSums(is.na(.)) != ncol(.)) %>% 
                        distinct
                
        })
        
        LTL.colnames <- reactive({
                LTL <- LTL()
                
                
                LTL.colnames <- 
                        if (!is.null(LTL)) {
                                colnames(LTL)
                        } else {
                                c(
                                        "Action(Transaction)",
                                        "Vendor(Order)",
                                        "PO_Number(Order)",
                                        "SubDivision",
                                        "Invoice_Number",
                                        "Merchant_SKU(Order Line)",
                                        "Sub_Lot",
                                        "Order_Date(Order)",
                                        "Merchant_Line Number(Order Line)",
                                        "Description(Order Line)",
                                        "Quantity_Ordered(Order Line)",
                                        "Status(Order Line)",
                                        "Applied_Date(Transaction)",
                                        "ShipTo_City(Order)",
                                        "ShipTo_State(Order)",
                                        "Tracking_Number(Transaction)",
                                        "Tracking_Delivery_Status(Transaction)",
                                        "ShipTo_Last_Name(Order)",
                                        "ShipTo_First_Name(Order)",
                                        "ShipTo_Address1(Order)",
                                        "ShipTo_Address2(Order)",
                                        "ShipTo_Address3(Order)",
                                        "ShipTo_Postal_Code(Order)",
                                        "ShipTo_Day_Phone(Order)",
                                        "Customer_Email(Order)",
                                        "SMS_Phonenumber",
                                        "Supplier_number"
                                )
                        }
        })
        
        online <- reactive({
                
                if (is.null(input$JCP_input_online)) return(NULL)
                
                online <- 
                        input$JCP_input_online %>%
                        pull(datapath) %>%
                        lapply(create.online.input) %>% 
                        list.clean %>% 
                        list.rbind
        })
        
        zip <- reactive({
                
                zip_button <- input$JCP_input_zip
                
                zip_path <- 
                        if (is.null(zip_button)) {
                                "//nsdstor1/SHARED/JCP_R/zip/zip.csv"
                        } else {
                                zip_button$datapath
                        }
                
                zip <- zip_path %>% read_csv
                
        })
        
        last_list_of_orders <- reactive({
                
                last_list_button <- input$JCP_input_last
                
                last_list_path <- 
                        if (is.null(last_list_button)) {
                                "//nsdstor1/SHARED/JCP_R/last_list_of_orders/last.csv"
                        } else {
                                last_list_button$datapath
                        }
                
                last_list_of_orders <- last_list_path %>% read_csv
                
        })
        
        
        flat.file <- reactive({
                
                EFF <- EFF()
                LTL <- LTL()
                online <- online()
                retail <- retail()
                zip <- zip()
                last_list_of_orders <- last_list_of_orders()
                
                flat.file <- 
                        online %>% 
                        {if (!is.null(LTL)) filter(., !`PO_Number(Order)` %in% LTL$`PO_Number(Order)`) else .} %>% 
                        list(LTL) %>% 
                        list.clean %>% 
                        list.rbind %>% 
                        mutate(
                                Sub = substr(Sub_Lot, 1, 3),
                                Lot = substr(Sub_Lot, 5, 8)
                        ) %>% 
                        left_join(EFF, by = c("Sub","Lot")) %>% 
                        mutate(
                                Identifier = "Online",
                                Customer_Cell_Phone = `SMS_Phonenumber`, 
                                Customer_Email = `Customer_Email(Order)`, 
                                Special_Instruction = NA,
                                Weight = Weight * as.numeric(`Quantity_Ordered(Order Line)`)
                        ) %>%  
                        select(
                                PO_Number = `PO_Number(Order)`,
                                Sub_Lot,
                                Product_Description = `Description(Order Line)`,
                                Piece_Count = `Quantity_Ordered(Order Line)`,
                                Ref_2 = `Tracking_Number(Transaction)`,
                                Length2,
                                Width,
                                Height,
                                Weight,
                                `Delivery Experience`,
                                Customer_Name = `ShipTo_First_Name(Order)`,
                                Customer_Street_Address_1 = `ShipTo_Address1(Order)`,
                                Customer_Street_Address_2 = `ShipTo_Address2(Order)`,
                                Customer_Street_Address_3 = `ShipTo_Address3(Order)`,
                                Customer_City = `ShipTo_City(Order)`,
                                Customer_State = `ShipTo_State(Order)`,
                                Customer_Zip = `ShipTo_Postal_Code(Order)`,
                                Customer_Phone = `ShipTo_Day_Phone(Order)`,
                                Customer_Cell_Phone,
                                Customer_Email,
                                Special_Instruction,
                                Identifier
                        ) %>% 
                        list(retail) %>% 
                        list.clean %>% 
                        list.rbind %>% 
                        select(PO_Number,
                               Sub_Lot,
                               Product_Description,
                               Ref_2,
                               Piece_Count,
                               Length2,
                               Width,
                               Height,
                               Weight,
                               `Delivery Experience`,
                               `Customer_Name`,
                               Customer_Street_Address_1,
                               Customer_Street_Address_2,
                               Customer_Street_Address_3,
                               Customer_City,
                               Customer_State,
                               Customer_Zip,
                               Customer_Phone,
                               Customer_Cell_Phone,
                               Customer_Email,
                               Special_Instruction,
                               Identifier
                        ) %>% 
                        filter(
                                # remove the orders not in areas covered by us
                                Customer_Zip %in% zip$`5 ZIP`,
                                # remove the orders existing in the last list of orders
                                !PO_Number %in% last_list_of_orders[,3]
                        ) %>% 
                        # sort by po number
                        arrange(PO_Number) %>% 
                        # make the flatfile
                        dplyr::rename(Weight2 = Weight, Width2 = Width, Height2 = Height) %>% 
                        mutate(
                                `Shipper ID` = "11649",
                                `Shipper Name` = "J.C. Penney Corporation, Inc.",
                                `Shipper Address` = "6501 Legacy Drive",
                                `Shipper Address 2` = NA,
                                `Shipper City` = "Plano",
                                `Shipper State` = "TX",
                                `Shipper Zip` = "75024",
                                `Shipper Phone` = "8003221189",
                                `Requested P/U Begin` = NA,
                                `Requested P/U End` = NA,
                                `Consginee ID` = NA,
                                `Cons Name` = Customer_Name,
                                `Cons Address` = Customer_Street_Address_1,
                                `Cons Address 2` = Customer_Street_Address_2,
                                `Consginee City` = Customer_City,
                                `Cons State` = Customer_State,
                                `Cons Zip` = Customer_Zip,
                                `Cons Phone` = Customer_Phone,
                                `Cons Alt. Phone` = Customer_Cell_Phone,
                                `Cons Email` = Customer_Email,
                                `Requested Deliver-By Begin` = NA,
                                `Requested Deliver-By End` = NA,
                                `NSD Bill#` = NA,
                                `Service Level` = `Delivery Experience`,
                                `Op Code` = "LM",
                                `Movement` = "DELIVERY",
                                `Primary Ref` = PO_Number,
                                `Secondary Ref` = "",
                                Commodity = `Delivery Experience`,
                                RMA = NA,
                                PLT = NA,
                                `PLT Type/Unit` = NA,
                                PCS = Piece_Count,
                                `PC Type` = "CTN",
                                `Product Description` = Product_Description,
                                Weight = Weight2,
                                `Cubic Weight` = NA,
                                `Declared Value` = NA,
                                Length = Length2,
                                Width = Width2,
                                Height = Height2,
                                Disposal =
                                        case_when(
                                                `Product Description` %in% c("MATTRESS PICK UP", "MATTRESS REMOVAL FEE", "HAUL AWAY SERVICE") ~ "Yes",
                                                grepl("^CA RECYCLE FEE",`Product Description`) ~ "Yes",
                                                T ~ NA_character_
                                        ),
                                Assembly = NA,
                                `Mattress Removal` =
                                        case_when(
                                                `Product Description` %in% c("MATTRESS PICK UP", "MATTRESS REMOVAL FEE") ~ "Yes",
                                                grepl("^CA RECYCLE FEE|HAUL AWAY SERVICE",`Product Description`) ~ "Yes",
                                                T ~ NA_character_
                                        ),
                                `Extra-Man` = NA,
                                Notes = Special_Instruction,
                                `3rd Reference` = Identifier
                        ) %>% 
                        filter(!grepl("\\Protector$|\\Protectr$|\\Encase$|^UPH|^ACCT PLAN", `Product Description`)) %>% 
                        mutate(
                                `Product Description` = paste(Sub_Lot, `Product Description`),
                                `Cons Alt. Phone` = ifelse(`Cons Alt. Phone` == "N/A", NA, `Cons Alt. Phone`),
                                `Cons Address 2` = ifelse(`Cons Address 2` == "N/A", NA, `Cons Address 2`)
                        ) %>% 
                        select(-(1:22))
                
        })
        
        output$JCP_input_download <- downloadHandler(
                filename = function() {
                        paste0("Z:/JCP_R/output/JCP_Flatfile_", Sys.Date(),".csv")
                },
                content = function(file) {
                        
                        if (is.null(authorization.check(JCP_team))) return(NULL)
                        flat.file <- flat.file()
                        
                        flat.file %>% write.csv(file, sep = ",", row.names = F, na = "")
                        
                }
        )
        
        observe({
                updateSelectInput(session, "JCP_zip_agent", choices = c("ALL", unique(zip()$`NSD Terminal`)))
        })
        
        output$JCP_retail_table <- renderDT({
                
                retail <- retail()
                if (is.null(retail)) return(NULL)
                
                retail %>% 
                        mutate(
                                `Dimension(LxWxH)` = paste0(Length2, "x", Width, "x", Height)
                        ) %>% 
                        select(
                                `PO #` = PO_Number,
                                `SubLot` = Sub_Lot,
                                Description = Product_Description,
                                Items = Piece_Count,
                                `Ref2` = Ref_2,
                                `Dimension LxWxH` = `Dimension(LxWxH)`,
                                Weight,
                                `ServiceType` = `Delivery Experience`,
                                Instruction = Special_Instruction
                        ) %>% 
                        datatable(
                                options = list(
                                        dom = "Brtip",
                                        pageLength = 8
                                ),
                                rownames = F,
                                selection = list(mode = 'single', selected = 1)
                        ) 
        })
        output$JCP_retail_info  <-  renderPrint({
                
                if (length(input$JCP_retail_table_rows_selected) > 0) {
                        name <- retail()[input$JCP_retail_table_rows_selected, "Customer_Name"] %>% str_trim
                        address1 <- retail()[input$JCP_retail_table_rows_selected, "Customer_Street_Address_1"] %>% str_trim
                        address2 <- retail()[input$JCP_retail_table_rows_selected, "Customer_Street_Address_2"] %>% str_trim
                        address3 <- retail()[input$JCP_retail_table_rows_selected, "Customer_Street_Address_3"] %>% str_trim
                        city <- retail()[input$JCP_retail_table_rows_selected, "Customer_City"] %>% str_trim
                        state <- retail()[input$JCP_retail_table_rows_selected, "Customer_State"] %>% str_trim
                        zip <- retail()[input$JCP_retail_table_rows_selected, "Customer_Zip"] %>% str_trim
                        phone <- retail()[input$JCP_retail_table_rows_selected, "Customer_Phone"] %>% str_trim
                        cell <- retail()[input$JCP_retail_table_rows_selected, "Customer_Cell_Phone"] %>% str_trim
                        email <- retail()[input$JCP_retail_table_rows_selected, "Customer_Email"] %>% str_trim
                        
                        
                        cat("**Customer Contacts**\n\n")
                        
                        if (!is.na(name) & name != "") cat(paste0(name, "\n"))
                        if (!is.na(address1) & address1 != "") cat(paste0(address1, "\n"))
                        if (!is.na(address2) & address2 != "") cat(paste0(address2, "\n"))
                        if (!is.na(address3) & address3 != "") cat(paste0(address3, "\n"))
                        if (!is.na(city) & city != "") cat(paste0(city, ", "))
                        if (!is.na(state) & state != "") cat(paste0(state, " "))
                        if (!is.na(zip) & zip != "") cat(paste0(zip, "\n"))
                        if (!is.na(phone) & phone != "") cat(paste0("Phone: ", phone, "\n"))
                        if (!is.na(cell) & cell != "") cat(paste0("Cell: ", cell, "\n"))
                        if (!is.na(email) & email != "") cat(paste0("Email: ", email, "\n"))
                }
        })
        
        output$JCP_online_table <- renderDT({
                
                online <- online()
                if (is.null(online)) return(NULL)
                
                online %>% 
                        select(
                                `PO #` = `PO_Number(Order)`,
                                `SubLot` = Sub_Lot,
                                Description = `Description(Order Line)`,
                                SKU = `Merchant_SKU(Order Line)`,
                                Items = `Quantity_Ordered(Order Line)`,
                                `Order Date` = `Order_Date(Order)`,
                                `Applied Date` = `Applied_Date(Transaction)`,
                                Vendor = `Vendor(Order)`,
                                Status = `Status(Order Line)`,
                                `Supplier #` = Supplier_number
                        ) %>%
                        datatable(
                                options = list(
                                        dom = "Brtip",
                                        pageLength = 8
                                ),
                                rownames = FALSE,
                                selection = list(mode = 'single', selected = 1)
                        )
        })
        output$JCP_online_info = renderPrint({
                
                if (length(input$JCP_online_table_rows_selected) > 0) {
                        name <- online()[input$JCP_online_table_rows_selected, "ShipTo_First_Name(Order)"] %>% str_trim
                        address1 <- online()[input$JCP_online_table_rows_selected, "ShipTo_Address1(Order)"] %>% str_trim
                        address2 <- online()[input$JCP_online_table_rows_selected, "ShipTo_Address2(Order)"] %>% str_trim
                        address3 <- online()[input$JCP_online_table_rows_selected, "ShipTo_Address3(Order)"] %>% str_trim
                        city <- online()[input$JCP_online_table_rows_selected, "ShipTo_City(Order)"] %>% str_trim
                        state <- online()[input$JCP_online_table_rows_selected, "ShipTo_State(Order)"] %>% str_trim
                        zip <- online()[input$JCP_online_table_rows_selected, "ShipTo_Postal_Code(Order)"] %>% str_trim
                        phone <- online()[input$JCP_online_table_rows_selected, "ShipTo_Day_Phone(Order)"] %>% str_trim
                        sms <- online()[input$JCP_online_table_rows_selected, "SMS_Phonenumber"] %>% str_trim
                        email <- online()[input$JCP_online_table_rows_selected, "Customer_Email(Order)"] %>% str_trim
                        
                        
                        cat("**Customer Contacts**\n\n")
                        
                        if (!is.na(name) & !name %in% c("", "N/A")) cat(paste0(name, "\n"))
                        if (!is.na(address1) & !address1 %in% c("", "N/A")) cat(paste0(address1, "\n"))
                        if (!is.na(address2) & !address2 %in% c("", "N/A")) cat(paste0(address2, "\n"))
                        if (!is.na(address3) & !address3 %in% c("", "N/A")) cat(paste0(address3, "\n"))
                        if (!is.na(city) & !city %in% c("", "N/A")) cat(paste0(city, ", "))
                        if (!is.na(state) & !state %in% c("", "N/A")) cat(paste0(state, " "))
                        if (!is.na(zip) & !zip %in% c("", "N/A")) cat(paste0(zip, "\n"))
                        if (!is.na(phone) & !phone %in% c("", "N/A")) cat(paste0("Phone: ", phone, "\n"))
                        if (!is.na(sms) & !sms %in% c("", "N/A")) cat(paste0("SMS: ", sms, "\n"))
                        if (!is.na(email) & !email %in% c("", "N/A")) cat(paste0("Email: ", email, "\n"))
                }
        })
        
        output$JCP_LTL_table <- renderDT({
                
                LTL <- LTL()
                if (is.null(LTL)) return(NULL)
                
                LTL %>% 
                        select(
                                `PO #` = `PO_Number(Order)`,
                                `SubLot` = Sub_Lot,
                                Description = `Description(Order Line)`,
                                SKU = `Merchant_SKU(Order Line)`,
                                Items = `Quantity_Ordered(Order Line)`,
                                `Order Date` = `Order_Date(Order)`,
                                `Applied Date` = `Applied_Date(Transaction)`,
                                Vendor = `Vendor(Order)`,
                                Status = `Status(Order Line)`,
                                `Supplier #` = Supplier_number
                        ) %>%
                        datatable(
                                options = list(
                                        dom = "Brtip",
                                        pageLength = 8
                                ),
                                rownames = FALSE,
                                selection = list(mode = 'single', selected = 1)
                        )
        })
        output$JCP_LTL_info = renderPrint({
                
                if (length(input$JCP_LTL_table_rows_selected) > 0) {
                        name <- LTL()[input$JCP_LTL_table_rows_selected, "ShipTo_First_Name(Order)"] %>% str_trim
                        address1 <- LTL()[input$JCP_LTL_table_rows_selected, "ShipTo_Address1(Order)"] %>% str_trim
                        address2 <- LTL()[input$JCP_LTL_table_rows_selected, "ShipTo_Address2(Order)"] %>% str_trim
                        address3 <- LTL()[input$JCP_LTL_table_rows_selected, "ShipTo_Address3(Order)"] %>% str_trim
                        city <- LTL()[input$JCP_LTL_table_rows_selected, "ShipTo_City(Order)"] %>% str_trim
                        state <- LTL()[input$JCP_LTL_table_rows_selected, "ShipTo_State(Order)"] %>% str_trim
                        zip <- LTL()[input$JCP_LTL_table_rows_selected, "ShipTo_Postal_Code(Order)"] %>% str_trim
                        phone <- LTL()[input$JCP_LTL_table_rows_selected, "ShipTo_Day_Phone(Order)"] %>% str_trim
                        sms <- LTL()[input$JCP_LTL_table_rows_selected, "SMS_Phonenumber"] %>% str_trim
                        email <- LTL()[input$JCP_LTL_table_rows_selected, "Customer_Email(Order)"] %>% str_trim
                        
                        
                        cat("**Customer Contacts**\n\n")
                        
                        if (!is.na(name) & !name %in% c("", "N/A")) cat(paste0(name, "\n"))
                        if (!is.na(address1) & !address1 %in% c("", "N/A")) cat(paste0(address1, "\n"))
                        if (!is.na(address2) & !address2 %in% c("", "N/A")) cat(paste0(address2, "\n"))
                        if (!is.na(address3) & !address3 %in% c("", "N/A")) cat(paste0(address3, "\n"))
                        if (!is.na(city) & !city %in% c("", "N/A")) cat(paste0(city, ", "))
                        if (!is.na(state) & !state %in% c("", "N/A")) cat(paste0(state, " "))
                        if (!is.na(zip) & !zip %in% c("", "N/A")) cat(paste0(zip, "\n"))
                        if (!is.na(phone) & !phone %in% c("", "N/A")) cat(paste0("Phone: ", phone, "\n"))
                        if (!is.na(sms) & !sms %in% c("", "N/A")) cat(paste0("SMS: ", sms, "\n"))
                        if (!is.na(email) & !email %in% c("", "N/A")) cat(paste0("Email: ", email, "\n"))
                }
        })
        
        
        JCP_zip_map_data <- reactive({
                zip <- zip()
                data(zipcode)
                
                JCP_zip_map_data <- 
                        zip %>% 
                        {if (input$JCP_zip_agent != "ALL") filter(., `NSD Terminal` == input$JCP_zip_agent) else .} %>% 
                        rename(zip = `5 ZIP`) %>% 
                        mutate(zip = zip %>% as.character) %>% 
                        left_join(zipcode %>% select(zip, latitude, longitude), by = "zip") %>% 
                        filter(!is.na(latitude), !is.na(longitude)) %>% 
                        mutate(
                                latlong = paste0(latitude, ":", longitude),
                                size = 0.05
                        )
        })
        
        output$JCP_zip_count <- renderText({
                JCP_zip_map_data <- JCP_zip_map_data()
                
                JCP_zip_map_data %>% 
                        nrow %>% 
                        {if (. == 0) return(NULL) else .} %>% 
                        as.character %>% 
                        paste0("Number of zipcodes covered: ", .)
        })
        
        output$JCP_zip_map <- renderGvis({
                
                JCP_zip_map_data <- JCP_zip_map_data()
                
                JCP_zip_map_data %>%
                        gvisGeoChart("size", locationvar = "latlong", sizevar = "size", hovervar = "zip",
                                     options =
                                             list(
                                                     region = "US",
                                                     resolution = "provinces",
                                                     displayMode = "markers",
                                                     height = 347*map_scale,
                                                     width = 556*map_scale,
                                                     keepAspectRatio = F,
                                                     colorAxis = paste0("{colors:['", picked_color, "']}"),
                                                     legend = "none",
                                                     sizeAxis = "{maxSize: 3}"
                                             )
                        )
        })
        
        # Accounts -----------------------------------------------------------
        
        
        Accounts_TM <- reactive({
                
                TM_report <- input$accounts_TM_update_data
                
                if (is.null(TM_report)) {
                        return(NULL)
                } else {
                        Accounts_TM_path <- TM_report$datapath
                } 
                
                Accounts_TM <- 
                        Accounts_TM_path %>% 
                        read_csv(
                                col_types = 
                                        cols(
                                                `ADMIN STATUS CHANGED` = col_datetime(format = "%m/%d/%Y %H:%M:%S %p"), 
                                                `AGENT ETA` = col_date(format = "%m/%d/%Y"), 
                                                `CLIENT ID` = col_character(), 
                                                `CONTACT ATTEMPTS` = col_integer(), 
                                                CREATED = col_date(format = "%m/%d/%Y"), 
                                                CUSTOMER_CONTACT_COUNT = col_integer(), 
                                                `DAYS ON DOCK` = col_integer(), 
                                                `DOCKED DATE` = col_date(format = "%m/%d/%Y"), 
                                                `LAST COMMENT DATE` = col_datetime(format = "%m/%d/%Y %H:%M:%S %p"), 
                                                `LAST STATUS UPDATE` = col_datetime(format = "%m/%d/%Y %H:%M:%S %p"), 
                                                `LTL ETA DATE` = col_datetime(format = "%m/%d/%Y %H:%M:%S %p"), 
                                                `LTL P/U DATE` = col_datetime(format = "%m/%d/%Y %H:%M:%S %p"), 
                                                LTL_LANE_DAYS = col_integer(), 
                                                LTL_POD_DATE_TIME = col_datetime(format = "%m/%d/%Y %H:%M:%S %p"), 
                                                `OP STATUS CHANGED` = col_datetime(format = "%m/%d/%Y %H:%M:%S %p"), 
                                                `PHYSICAL_CARRIER_PRO#` = col_character(), 
                                                PIECES = col_integer(), 
                                                `REFERENCE#2` = col_character(), 
                                                `RMA/RGA_NUMBER` = col_character(), 
                                                SCHEDULED = col_date(format = "%m/%d/%Y")
                                        )
                        ) %>% 
                        arrange(AGENT, CREATED)
        })

# Home Depot --------------------------------------------------------------


# HD Past OFD -------------------------------------------------------------

        
        Accounts_HD_past_OFD <- reactive({
                
                Accounts_TM <- Accounts_TM()
                if (is.null(Accounts_TM)) return(NULL)
                
                Accounts_HD_past_OFD <- 
                        Accounts_TM %>% 
                        mutate(
                                date = as.Date(`OP STATUS CHANGED`)  
                        ) %>% 
                        filter(
                                CLIENT == "THE HOME DEPOT",
                                `ORDER TYPE` == "DELIVERY",
                                `OP STATUS` == "OFD",
                                `ADMIN STATUS` %in% c(NA, "", "PCF"), 
                                date < today()
                        ) %>% 
                        select(
                                `NSD #`,
                                `PO#`,
                                CLIENT,
                                `SERVICE LEVEL`,
                                `ORDER TYPE`,
                                AGENT,
                                CONSIGNEE,
                                CITY,
                                ZIP,
                                PIECES,
                                MILEAGE,
                                `DOCKED DATE`,
                                `DAYS ON DOCK`,
                                SCHEDULED,
                                `OP STATUS`,
                                `OP STATUS CHANGED`
                        ) %>% 
                        lapply(as.character) %>% 
                        as.data.frame
        })
        
        output$HD_OFD_last_sent <- renderText({
                input$HD_past_OFD_send_emails
                "//nsdstor1/SHARED/NSDComp/HomeDepot/HD_OFD_last_sent.txt" %>% 
                        readChar(file.info(.)$size) %>% 
                        paste0("Last Sent: ", .)
        })
        
        observeEvent(input$HD_past_OFD_subject_save, {
                
                if (is.null(authorization.check(HD_team))) return(NULL)
                
                writeChar(
                        isolate(input$HD_past_OFD_subject), 
                        "//nsdstor1/SHARED/NSDComp/HomeDepot/HD_OFD_template_subject.txt"
                )
                
                shinyalert(
                        title = "Saved",
                        text = "Subject",
                        closeOnEsc = TRUE,
                        closeOnClickOutside = TRUE,
                        html = FALSE,
                        type = "success",
                        showConfirmButton = TRUE,
                        showCancelButton = FALSE,
                        confirmButtonText = "OK",
                        confirmButtonCol = "#a2a8ab",
                        timer = 0,
                        imageUrl = "",
                        animation = TRUE
                )
        })
        
        observeEvent(input$HD_past_OFD_body_save, {
                
                if (is.null(authorization.check(HD_team))) return(NULL)
                
                
                writeChar(
                        isolate(input$HD_past_OFD_body), 
                        "//nsdstor1/SHARED/NSDComp/HomeDepot/HD_OFD_template_body.txt"
                )
                
                shinyalert(
                        title = "Saved",
                        text = "Body",
                        closeOnEsc = TRUE,
                        closeOnClickOutside = TRUE,
                        html = FALSE,
                        type = "success",
                        showConfirmButton = TRUE,
                        showCancelButton = FALSE,
                        confirmButtonText = "OK",
                        confirmButtonCol = "#a2a8ab",
                        timer = 0,
                        imageUrl = "",
                        animation = TRUE
                )
                
        })
        
        observeEvent(input$HD_past_OFD_send_emails, {
                
                if (is.null(authorization.check(HD_team))) return(NULL)
                
                Accounts_HD_past_OFD <- isolate(Accounts_HD_past_OFD())
                if (is.null(Accounts_HD_past_OFD)) {
                        showModal(modalDialog(
                                title = "CAUTION!",
                                "No data imported!",
                                easyClose = TRUE,
                                footer = NULL
                        ))
                        return(NULL)
                }
                
                agents_names <- Accounts_HD_past_OFD$AGENT %>% unique
                
                HD_OFD_body <- paste(paste0('<p>', strsplit(isolate(input$HD_past_OFD_body),"\n")[[1]], "</p>"), collapse="")
                
                if (length(agents_names) > 0) {
                        shinyalert(
                                paste0("Past OFD emails are going to be sent to ", length(agents_names), " agents"),
                                type = "info",
                                showCancelButton = T,
                                showConfirmButton = T,
                                callbackR =
                                        function(x) {
                                                if(x != FALSE) {
                                                        OutApp <- COMCreate("Outlook.Application")
                                                        withProgress(message = 'Sending E-mails in progress', value = 0, {
                                                                for (agent in agents_names) {
                                                                        incProgress(1/length(agents_names))
                                                                        data <- Accounts_HD_past_OFD %>% filter(AGENT == agent)
                                                                        HD_OFD_subject <- isolate(input$HD_past_OFD_subject) %>% paste0(agent)
                                                                        
                                                                        outMail = OutApp$CreateItem(0)
                                                                        outMail[["To"]] = email_address(
                                                                                if (agent == "CLAUDIO") {"chantillyoperations@nonstopdelivery.com;"} else paste0("AG_", agent,"@nonstopdelivery.com;"),
                                                                                isolate(input$Test_Active)
                                                                        )
                                                                        outMail[["subject"]] = HD_OFD_subject
                                                                        
                                                                        outMail[["HTMLBody"]] = 
                                                                                paste0(
                                                                                        HD_OFD_body,
                                                                                        data %>% 
                                                                                                kable(format = "html") %>%
                                                                                                kable_styling(bootstrap_options = c("bordered"), full_width = T, position = "left", font_size = 12) %>% 
                                                                                                row_spec(1:nrow(data), color = "white", background = "#D7261E") %>% 
                                                                                                HTML, 
                                                                                        nsd_signature(Sys.info()["user"][[1]])
                                                                                )
                                                                        outMail$Send() 
                                                                }
                                                        })
                                                        
                                                        writeChar(isolate(now() %>% as.character), "//nsdstor1/SHARED/NSDComp/HomeDepot/HD_OFD_last_sent.txt")
                                                        
                                                }
                                        }
                        )
                }
                                                        
                
                
                
        })

        # HD Past Scheduled -------------------------------------------------------
        
        Accounts_HD_past_scheduled <- reactive({
                
                Accounts_TM <- Accounts_TM()
                if (is.null(Accounts_TM)) return(NULL)
                
                Accounts_HD_past_scheduled <- 
                        Accounts_TM %>% 
                        mutate(
                                date = as.Date(SCHEDULED)  
                        ) %>% 
                        filter(
                                CLIENT == "THE HOME DEPOT",
                                `ORDER TYPE` == "DELIVERY",
                                `OP STATUS` == "SCHEDULED",
                                `ADMIN STATUS` %in% c(NA, "", "PCF"), 
                                date < today()
                        ) %>% 
                        select(
                                `NSD #`,
                                `PO#`,
                                CLIENT,
                                `SERVICE LEVEL`,
                                `ORDER TYPE`,
                                AGENT,
                                CONSIGNEE,
                                CITY,
                                ZIP,
                                PIECES,
                                MILEAGE,
                                `DOCKED DATE`,
                                `DAYS ON DOCK`,
                                SCHEDULED,
                                `OP STATUS`,
                                `OP STATUS CHANGED`
                        ) %>% 
                        lapply(as.character) %>% 
                        as.data.frame
        })
        
        output$HD_scheduled_last_sent <- renderText({
                input$HD_past_scheduled_send_emails
                "//nsdstor1/SHARED/NSDComp/HomeDepot/HD_scheduled_last_sent.txt" %>% 
                        readChar(file.info(.)$size) %>% 
                        paste0("Last Sent: ", .)
        })
        
        observeEvent(input$HD_past_scheduled_subject_save, {
                
                if (is.null(authorization.check(HD_team))) return(NULL)
                
                writeChar(
                        isolate(input$HD_past_scheduled_subject), 
                        "//nsdstor1/SHARED/NSDComp/HomeDepot/HD_scheduled_template_subject.txt"
                )
                
                shinyalert(
                        title = "Saved",
                        text = "Subject",
                        closeOnEsc = TRUE,
                        closeOnClickOutside = TRUE,
                        html = FALSE,
                        type = "success",
                        showConfirmButton = TRUE,
                        showCancelButton = FALSE,
                        confirmButtonText = "OK",
                        confirmButtonCol = "#a2a8ab",
                        timer = 0,
                        imageUrl = "",
                        animation = TRUE
                )
        })
        
        observeEvent(input$HD_past_scheduled_body_save, {
                
                if (is.null(authorization.check(HD_team))) return(NULL)
                
                
                writeChar(
                        isolate(input$HD_past_scheduled_body), 
                        "//nsdstor1/SHARED/NSDComp/HomeDepot/HD_scheduled_template_body.txt"
                )
                
                shinyalert(
                        title = "Saved",
                        text = "Body",
                        closeOnEsc = TRUE,
                        closeOnClickOutside = TRUE,
                        html = FALSE,
                        type = "success",
                        showConfirmButton = TRUE,
                        showCancelButton = FALSE,
                        confirmButtonText = "OK",
                        confirmButtonCol = "#a2a8ab",
                        timer = 0,
                        imageUrl = "",
                        animation = TRUE
                )
                
        })
        
        observeEvent(input$HD_past_scheduled_send_emails, {
                
                if (is.null(authorization.check(HD_team))) return(NULL)
                
                Accounts_HD_past_scheduled <- isolate(Accounts_HD_past_scheduled())
                if (is.null(Accounts_HD_past_scheduled)) {
                        showModal(modalDialog(
                                title = "CAUTION!",
                                "No data imported!",
                                easyClose = TRUE,
                                footer = NULL
                        ))
                        return(NULL)
                }
                
                agents_names <- Accounts_HD_past_scheduled$AGENT %>% unique
                
                
                
                HD_scheduled_body <- paste(paste0('<p>', strsplit(isolate(input$HD_past_scheduled_body),"\n")[[1]], "</p>"), collapse="")
                
                
                if (length(agents_names) > 0) {
                        shinyalert(
                                paste0("Past Scheduled emails are going to be sent to ", length(agents_names), " agents"),
                                type = "info",
                                showCancelButton = T,
                                showConfirmButton = T,
                                callbackR =
                                        function(x) {
                                                if(x != FALSE) {
                                                        OutApp <- COMCreate("Outlook.Application")
                                                        withProgress(message = 'Sending E-mails in progress', value = 0, {
                                                                for (agent in agents_names) {
                                                                        incProgress(1/length(agents_names))
                                                                        data <- Accounts_HD_past_scheduled %>% filter(AGENT == agent)
                                                                        HD_scheduled_subject <- isolate(input$HD_past_scheduled_subject) %>% paste0(agent)
                                                                        
                                                                        outMail = OutApp$CreateItem(0)
                                                                        outMail[["To"]] = email_address(
                                                                                if (agent == "CLAUDIO") {"chantillyoperations@nonstopdelivery.com;"} else paste0("AG_", agent,"@nonstopdelivery.com;"),
                                                                                isolate(input$Test_Active)
                                                                        )
                                                                        outMail[["subject"]] = HD_scheduled_subject
                                                                        
                                                                        outMail[["HTMLBody"]] = 
                                                                                paste0(
                                                                                        HD_scheduled_body,
                                                                                        data %>% 
                                                                                                kable(format = "html") %>%
                                                                                                kable_styling(bootstrap_options = c("bordered"), full_width = T, position = "left", font_size = 12) %>% 
                                                                                                row_spec(1:nrow(data), color = "black", background = "#ddc84d") %>% 
                                                                                                HTML, 
                                                                                        nsd_signature(Sys.info()["user"][[1]])
                                                                                )
                                                                        outMail$Send() 
                                                                }
                                                        })
                                                        
                                                        writeChar(isolate(now() %>% as.character), "//nsdstor1/SHARED/NSDComp/HomeDepot/HD_scheduled_last_sent.txt")
                                                        
                                                }
                                        }
                        )
                }
                
                
        })
        # ETA Report  -------------------------------------------------------
        
        Accounts_HD_ETA_report <- reactive({
                
                Accounts_TM <- Accounts_TM()
                if (is.null(Accounts_TM)) return(NULL)
                
                Accounts_HD_ETA_report <- 
                        Accounts_TM %>% 
                        mutate(
                                date = as.Date(`LTL ETA DATE`)  
                        ) %>% 
                        filter(
                                CLIENT == "THE HOME DEPOT",
                                `ORDER TYPE` == "DELIVERY",
                                `ADMIN STATUS` %in% c(NA, "", "PCF"), 
                                `OP STATUS` %in% c("AVAIL", "SHPTACK", "SHPTACK-T"),
                                date < today()
                        ) %>% 
                        select(
                                `NSD #`,
                                `PO#`,
                                CLIENT,
                                `SERVICE LEVEL`,
                                `ORDER TYPE`,
                                AGENT,
                                CONSIGNEE,
                                CITY,
                                ZIP,
                                PIECES,
                                MILEAGE,
                                `DOCKED DATE`,
                                `DAYS ON DOCK`,
                                SCHEDULED,
                                `OP STATUS`,
                                `OP STATUS CHANGED`
                        ) %>% 
                        lapply(as.character) %>% 
                        as.data.frame
        })
        
        output$HD_ETA_last_sent <- renderText({
                input$HD_ETA_report_send_emails
                "//nsdstor1/SHARED/NSDComp/HomeDepot/HD_ETA_last_sent.txt" %>% 
                        readChar(file.info(.)$size) %>% 
                        paste0("Last Sent: ", .)
        })
        
        
        observeEvent(input$HD_ETA_report_subject_save, {
                
                if (is.null(authorization.check(HD_team))) return(NULL)
                
                writeChar(
                        isolate(input$HD_ETA_report_subject), 
                        "//nsdstor1/SHARED/NSDComp/HomeDepot/HD_ETA_template_subject.txt"
                )
                
                shinyalert(
                        title = "Saved",
                        text = "Subject",
                        closeOnEsc = TRUE,
                        closeOnClickOutside = TRUE,
                        html = FALSE,
                        type = "success",
                        showConfirmButton = TRUE,
                        showCancelButton = FALSE,
                        confirmButtonText = "OK",
                        confirmButtonCol = "#a2a8ab",
                        timer = 0,
                        imageUrl = "",
                        animation = TRUE
                )
        })
        
        observeEvent(input$HD_ETA_report_body_save, {
                
                if (is.null(authorization.check(HD_team))) return(NULL)
                
                
                writeChar(
                        isolate(input$HD_ETA_report_body), 
                        "//nsdstor1/SHARED/NSDComp/HomeDepot/HD_ETA_template_body.txt"
                )
                
                shinyalert(
                        title = "Saved",
                        text = "Body",
                        closeOnEsc = TRUE,
                        closeOnClickOutside = TRUE,
                        html = FALSE,
                        type = "success",
                        showConfirmButton = TRUE,
                        showCancelButton = FALSE,
                        confirmButtonText = "OK",
                        confirmButtonCol = "#a2a8ab",
                        timer = 0,
                        imageUrl = "",
                        animation = TRUE
                )
                
        })
        
        observeEvent(input$HD_ETA_report_send_emails, {
                
                if (is.null(authorization.check(HD_team))) return(NULL)
                
                Accounts_HD_ETA_report <- isolate(Accounts_HD_ETA_report())
                if (is.null(Accounts_HD_ETA_report)) {
                        showModal(modalDialog(
                                title = "CAUTION!",
                                "No data imported!",
                                easyClose = TRUE,
                                footer = NULL
                        ))
                        return(NULL)
                }
                
                agents_names <- Accounts_HD_ETA_report$AGENT %>% unique
                
                
                
                HD_ETA_body <- paste(paste0('<p>', strsplit(isolate(input$HD_ETA_report_body),"\n")[[1]], "</p>"), collapse="")
                
                
                if (length(agents_names) > 0) {
                        shinyalert(
                                paste0("ETA Report emails are going to be sent to ", length(agents_names), " agents"),
                                type = "info",
                                showCancelButton = T,
                                showConfirmButton = T,
                                callbackR =
                                        function(x) {
                                                if(x != FALSE) {
                                                        OutApp <- COMCreate("Outlook.Application")
                                                        withProgress(message = 'Sending E-mails in progress', value = 0, {
                                                                for (agent in agents_names) {
                                                                        incProgress(1/length(agents_names))
                                                                        data <- Accounts_HD_ETA_report %>% filter(AGENT == agent)
                                                                        HD_ETA_subject <- isolate(input$HD_ETA_report_subject) %>% paste0(agent)
                                                                        
                                                                        outMail = OutApp$CreateItem(0)
                                                                        outMail[["To"]] = email_address(
                                                                                if (agent == "CLAUDIO") {"chantillyoperations@nonstopdelivery.com;"} else paste0("AG_", agent,"@nonstopdelivery.com;"),
                                                                                isolate(input$Test_Active)
                                                                        )
                                                                        outMail[["subject"]] = HD_ETA_subject
                                                                        
                                                                        outMail[["HTMLBody"]] = 
                                                                                paste0(
                                                                                        HD_ETA_body,
                                                                                        data %>% 
                                                                                                kable(format = "html") %>%
                                                                                                kable_styling(bootstrap_options = c("bordered"), full_width = T, position = "left", font_size = 12) %>% 
                                                                                                row_spec(1:nrow(data), color = "black", background = "#f2902e") %>% 
                                                                                                HTML, 
                                                                                        nsd_signature(Sys.info()["user"][[1]])
                                                                                )
                                                                        outMail$Send() 
                                                                }
                                                        })
                                                        
                                                        writeChar(isolate(now() %>% as.character), "//nsdstor1/SHARED/NSDComp/HomeDepot/HD_ETA_last_sent.txt")
                                                        
                                                }
                                        }
                        )
                }
                
                
                
        })
        
# Orders ------------------------------------------------------------------

        
        observe({updateSelectInput(session, "accounts_orders_selectinput", choices = c(Accounts_TM()$`NSD #`) %>% sort)})
        
        accounts_order_info <- reactive({

                jobno <- input$accounts_orders_selectinput
                Accounts_TM <- Accounts_TM()
                if (jobno == "-" | is.null(Accounts_TM)) return(NULL)

                accounts_order_info <- Accounts_TM %>% filter(`NSD #` == jobno)

        })
        
        # returns_orders_shipper_data <- reactive({
        #         
        #         order_info <- order_info()
        #         if (is.null(order_info)) return(NULL)
        #         
        #         zipa <- order_info %>% pull(SHIPPER_ZIP)
        #         
        #         returns_orders_shipper_data <- 
        #                 list(
        #                         name = order_info %>% pull(SHIPPER),
        #                         jobdate = order_info %>% pull(CREATED),
        #                         phone = order_info %>% pull(SHIPPER_PHONE),
        #                         email = order_info %>% pull(SHIPPER_EMAIL),
        #                         city = order_info %>% pull(SHIPPER_CITY),
        #                         zip = zipa,
        #                         state = zipcode %>% filter(zip == zipa) %>% pull(state)
        #                 )
        # })
        # output$returns_orders_shipper_info <- renderText({
        #         
        #         shipper <- returns_orders_shipper_data()
        #         if (is.null(shipper)) return(NULL)
        #         
        #         paste0(
        #                 "<b>Name: </b>", shipper[["name"]], "<br/>",
        #                 "<b>Created Date: </b>", shipper[["jobdate"]], "<br/>",
        #                 "<b>Address: </b>", shipper[["city"]], ", ", shipper[["state"]], " ", shipper[["zip"]], "<br/>",
        #                 "<b>Phone: </b>", shipper[["phone"]], "<br/>",
        #                 "<b>Email: </b>", shipper[["email"]], "<br/>"
        #                 
        #         ) %>% HTML
        # })
        # returns_orders_agent_data <- reactive({
        #         
        #         order_info <- order_info()
        #         if (is.null(order_info)) return(NULL)
        #         
        #         agent_name <- order_info %>% pull(AGENT)
        #         
        #         returns_orders_agent_data <- 
        #                 list(
        #                         name = agent_name,
        #                         full_name = order_info %>% pull(AGENT_NAME),
        #                         phone = order_info %>% pull(`AGENT PHONE`),
        #                         email = order_info %>% pull(`AGENT EMAIL`),
        #                         city = order_info %>% pull(AGENT_CITY),
        #                         zip = if (agent_name %in% Agents$Agent) Agents %>% filter(Agent == agent_name) %>% pull(zip),
        #                         state = order_info %>% pull(`AGENT STATE`)
        #                 )
        # })
        # output$returns_orders_agent_info <- renderText({
        #         
        #         agent <- returns_orders_agent_data()
        #         if (is.null(agent)) return(NULL)
        #         
        #         paste0(
        #                 "<b>Code: </b>", agent[["name"]], "<br/>",
        #                 "<b>Name: </b>", agent[["full_name"]], "<br/>",
        #                 "<b>Address: </b>", agent[["city"]], ", ", agent[["state"]], " ", agent[["zip"]], "<br/>",
        #                 "<b>Phone: </b>", agent[["phone"]], "<br/>",
        #                 "<b>Email: </b>", agent[["email"]]
        #         ) %>% HTML
        #         
        # })
        # returns_orders_carrier_data <- reactive({
        #         
        #         order_info <- order_info()
        #         if (is.null(order_info)) return(NULL)
        #         
        #         returns_orders_carrier_data <- 
        #                 list(
        #                         name = order_info %>% pull(`CARRIER NAME`),
        #                         carrier_pro = order_info %>% pull(`CARRIER PRO`),
        #                         unyson_pro = order_info %>% pull(`UNYSON PRO`),
        #                         ltl_ready = order_info %>% pull(`LTL READY`),
        #                         ltl_Pickup = order_info %>% pull(`LTL P/U DATE`)
        #                         
        #                 )
        # })
        # output$returns_orders_carrier_info <- renderText({
        #         
        #         carrier <- returns_orders_carrier_data()
        #         if (is.null(carrier)) return(NULL)
        #         
        #         paste0(
        #                 "<b>Carrier: </b>", carrier[["name"]], "<br/>",
        #                 "<b>Carrier Pro: </b>", carrier[["carrier_pro"]], "<br/>",
        #                 "<b>Unyson Pro: </b>", carrier[["unyson_pro"]], "<br/>",
        #                 "<b>LTL Ready: </b>", carrier[["ltl_ready"]], "<br/>",
        #                 "<b>LTL Pick-Up: </b>", carrier[["ltl_Pickup"]]
        #         ) %>% HTML
        #         
        # })
        # returns_orders_consignee_data <- reactive({
        #         
        #         order_info <- order_info()
        #         if (is.null(order_info)) return(NULL)
        #         
        #         returns_orders_consignee_data <- 
        #                 list(
        #                         client = order_info %>% pull(CLIENT),
        #                         client_id = order_info %>% pull(`CLIENT ID`),
        #                         name = order_info %>% pull(CONSIGNEE),
        #                         po = order_info %>% pull(`PO#`),
        #                         city = order_info %>% pull(CONSIGNEE_CITY),
        #                         state = order_info %>% pull(CONSIGNEE_STATE),
        #                         zip = order_info %>% pull(CONSIGNEE_ZIP)
        #                 )
        # })
        # output$returns_orders_consignee_info <- renderText({
        #         
        #         consignee <- returns_orders_consignee_data()
        #         if (is.null(consignee)) return(NULL)
        #         
        #         paste0(
        #                 "<b>Client: </b>", consignee[["client"]], " (id: ", consignee[["client_id"]], ")", "<br/>",
        #                 "<b>Name: </b>", consignee[["name"]], "<br/>",
        #                 "<b>PO number: </b>", consignee[["po"]], "<br/>",
        #                 "<b>Address: </b>", consignee[["city"]], ", ", consignee[["state"]], " ", consignee[["zip"]], "<br/>",
        #                 " <br/>"
        #         ) %>% HTML
        #         
        # })
        # output$return_orders_map <- renderLeaflet({
        #         
        #         order_info <- order_info()
        #         if (is.null(order_info)) return(leaflet() %>% addTiles() %>% setView(-96, 37.8, 4))
        #         
        #         shipper <- returns_orders_shipper_data()
        #         agent <- returns_orders_agent_data()
        #         consignee <- returns_orders_consignee_data()
        #         
        #         shipper_zip <- shipper[["zip"]]
        #         agent_zip <- agent[["zip"]]
        #         consignee_zip <- consignee[["zip"]]
        #         
        #         
        #         if(!agent[["name"]] %in% Agents$Agent | is.null(agent_zip) | !shipper_zip %in% zipcode$zip) {
        #                 shinyalert(
        #                         title = "Oops!",
        #                         text = "Shipper zipcode does not exist in the zipcodes source",
        #                         closeOnEsc = TRUE,
        #                         closeOnClickOutside = TRUE,
        #                         html = FALSE,
        #                         type = "error",
        #                         showConfirmButton = TRUE,
        #                         showCancelButton = FALSE,
        #                         confirmButtonText = "OK",
        #                         confirmButtonCol = "#a2a8ab",
        #                         timer = 0,
        #                         imageUrl = "",
        #                         animation = TRUE
        #                 )
        #                 return(NULL)
        #         }
        #         if(is.null(agent_zip) | !agent_zip %in% zipcode$zip) {
        #                 shinyalert(
        #                         title = "Oops!",
        #                         text = "Agent zipcode does not exist in the zipcodes source",
        #                         closeOnEsc = TRUE,
        #                         closeOnClickOutside = TRUE,
        #                         html = FALSE,
        #                         type = "error",
        #                         showConfirmButton = TRUE,
        #                         showCancelButton = FALSE,
        #                         confirmButtonText = "OK",
        #                         confirmButtonCol = "#a2a8ab",
        #                         timer = 0,
        #                         imageUrl = "",
        #                         animation = TRUE
        #                 )
        #                 return(NULL)
        #         }
        #         if(is.null(consignee_zip) | !consignee_zip %in% zipcode$zip) {
        #                 shinyalert(
        #                         title = "Oops!",
        #                         text = "Consignee zipcode does not exist in the zipcodes source",
        #                         closeOnEsc = TRUE,
        #                         closeOnClickOutside = TRUE,
        #                         html = FALSE,
        #                         type = "error",
        #                         showConfirmButton = TRUE,
        #                         showCancelButton = FALSE,
        #                         confirmButtonText = "OK",
        #                         confirmButtonCol = "#a2a8ab",
        #                         timer = 0,
        #                         imageUrl = "",
        #                         animation = TRUE
        #                 )
        #                 return(NULL)
        #         }
        #         
        #         zip_lat <- function(x) {if (x %in% zipcode$zip) zipcode %>% filter(zip == x) %>% pull(latitude)}
        #         zip_lon <- function(x) {if (x %in% zipcode$zip) zipcode %>% filter(zip == x) %>% pull(longitude)}
        #         
        #         shipper_lat <- zip_lat(shipper_zip)
        #         agent_lat <- zip_lat(agent_zip)
        #         consignee_lat <- zip_lat(consignee_zip)
        #         
        #         shipper_long <- zip_lon(shipper_zip)
        #         agent_long <- zip_lon(agent_zip)
        #         consignee_long <- zip_lon(consignee_zip)
        #         
        #         shipper_label <- 
        #                 paste0(
        #                         "<b>Shipper</b><br/>",
        #                         shipper[["name"]], "<br/>",
        #                         shipper[["city"]], ", ", shipper[["state"]], " ", shipper[["zip"]]
        #                 )
        #         
        #         agent_label <- 
        #                 paste0(
        #                         "<b>Agent</b><br/>",
        #                         agent[["name"]], "<br/>",
        #                         agent[["full_name"]], "<br/>",
        #                         agent[["city"]], ", ", agent[["state"]], " ", agent[["zip"]]
        #                 )
        #         
        #         consignee_label <- 
        #                 paste0(
        #                         "<b>Consignee</b><br/>",
        #                         consignee[["name"]], "<br/>",
        #                         consignee[["city"]], ", ", consignee[["state"]], " ", consignee[["zip"]]
        #                 )
        #         
        #         
        #         map_data <- 
        #                 data.frame(
        #                         stop_title = c("Shipper", "Agent", "Consignee"),
        #                         stop_name = c(shipper[["name"]], agent[["name"]], consignee[["name"]]),
        #                         stop_num = c(1, 2, 3),
        #                         latitude = c(shipper_lat, agent_lat, consignee_lat),
        #                         longitude = c(shipper_long, agent_long, consignee_long)
        #                         
        #                 )
        #         map <- 
        #                 leaflet() %>% 
        #                 addTiles() %>%
        #                 addMarkers(
        #                         data = map_data %>% filter(stop_title == "Shipper"), 
        #                         lng = ~longitude, 
        #                         lat = ~latitude, 
        #                         group = "Shipper", 
        #                         popup = shipper_label,
        #                         popupOptions = popupOptions(closeOnClick = FALSE)
        #                 ) %>% 
        #                 addMarkers(
        #                         data = map_data %>% filter(stop_title == "Agent"), 
        #                         lng = ~longitude, 
        #                         lat = ~latitude, 
        #                         group = "Agent", 
        #                         popup = agent_label,
        #                         popupOptions = popupOptions(closeOnClick = FALSE)
        #                 ) %>% 
        #                 addMarkers(
        #                         data = map_data %>% filter(stop_title == "Consignee"), 
        #                         lng = ~longitude, 
        #                         lat = ~latitude, 
        #                         group = "Consignee", 
        #                         popup = consignee_label,
        #                         popupOptions = popupOptions(closeOnClick = FALSE)
        #                 )
        # })
        output$accounts_orders_men <- renderImage({
                order_info <- accounts_order_info()
                if (is.null(order_info)) return(list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50))

                men <- order_info %>% pull(MAN)

                if (is.na(men)) {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                } else if (men == "2-MAN") {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/twoman.png", height = 50, width = 50)
                } else if (men == "1-MAN") {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/oneman.png", height = 50, width = 50)
                } else {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                }

        }, deleteFile = FALSE)
        # output$return_orders_servicetype <- renderImage({
        #         order_info <- order_info()
        #         if (is.null(order_info)) return(list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50))
        #         
        #         service <- order_info %>% pull(`SERVICE LEVEL`)
        #         mode <- order_info %>% pull(MODE)
        #         
        #         if (is.na(service)) {
        #                 list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
        #         } else if (mode == "LHR") {
        #                 list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
        #         } else if (service == "WHITEGLOVE") {
        #                 list(src = "//nsdstor1/SHARED/NSDComp/www/whiteglove.png", height = 50, width = 50)
        #         } else if (service == "THRESHOLD") {
        #                 list(src = "//nsdstor1/SHARED/NSDComp/www/threshold.png", height = 50, width = 50)
        #         } else {
        #                 list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
        #         }
        #         
        #         
        # }, deleteFile = FALSE)
        # output$return_orders_contact_attempts <- renderImage({
        #         order_info <- order_info()
        #         if (is.null(order_info)) return(list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50))
        #         
        #         attempts <- 
        #                 order_info %>% 
        #                 pull(`CONTACT ATTEMPTS`) %>% 
        #                 as.character
        #         mode <- order_info %>% pull(MODE)
        #         
        #         outfile <- 
        #                 "//nsdstor1/SHARED/NSDComp/www/phone.png" %>% 
        #                 image_read %>% 
        #                 image_annotate(attempts, size = 200, gravity = "southwest", location = "+350-25") %>%
        #                 image_write(tempfile(fileext='png'), format = 'png')
        #         
        #         if (is.na(attempts)) {
        #                 list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
        #         } else if (mode == "LHR") {
        #                 list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
        #         } else {
        #                 list(src = outfile, height = 50, width = 50)
        #         }
        #         
        #         
        #         
        # }, deleteFile = FALSE)
        output$accounts_orders_items <- renderImage({
                order_info <- accounts_order_info()
                if (is.null(order_info)) return(list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50))

                items <-
                        order_info %>%
                        pull(`PIECES`) %>%
                        as.character

                sizes = list("1" = "140", "2" = "110", "3" = "80", "4" = "40")

                outfile <-
                        "//nsdstor1/SHARED/NSDComp/www/item.png" %>%
                        image_read %>%
                        image_annotate(items, size = 100, gravity = "southwest", location = paste0("+", sizes[[nchar(items)]] %>% as.character, "+100")) %>%
                        image_write(tempfile(fileext='png'), format = 'png')

                if (is.na(items)) {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                } else {
                        list(src = outfile, height = 50, width = 50)
                }



        }, deleteFile = FALSE)
        output$accounts_orders_weight <- renderImage({
                order_info <- accounts_order_info()
                if (is.null(order_info)) return(list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50))

                weight <-
                        order_info %>%
                        pull(WEIGHT) %>%
                        {if (as.numeric(.) > 10) round(.) else round(., 2)} %>%
                        as.character

                sizes = list("1" = "190", "2" = "150", "3" = "100", "4" = "60")


                outfile <-
                        "//nsdstor1/SHARED/NSDComp/www/weight.png" %>%
                        image_read %>%
                        image_annotate(weight, size = 150, gravity = "southwest", location = paste0("+", sizes[[nchar(weight)]] %>% as.character, "+90")) %>%
                        image_write(tempfile(fileext='png'), format = 'png')

                if (is.na(weight)) {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                } else {
                        list(src = outfile, height = 50, width = 50)
                }



        }, deleteFile = FALSE)
        output$accounts_orders_miles <- renderImage({
                order_info <- accounts_order_info()
                if (is.null(order_info)) return(list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50))

                miles <- order_info %>% pull(MILEAGE) %>% as.character

                outfile <-
                        "//nsdstor1/SHARED/NSDComp/www/mileage.png" %>%
                        image_read %>%
                        image_annotate(miles, size = 100, gravity = "southwest", location = paste0("+150+20")) %>%
                        image_write(tempfile(fileext='png'), format = 'png')

                if (is.na(miles)) {
                        list(src = "//nsdstor1/SHARED/NSDComp/www/blank.png", height = 50, width = 50)
                } else {
                        list(src = outfile, height = 50, width = 50)
                }



        }, deleteFile = FALSE)
        # output$returns_orders_op <- renderText({
        #         
        #         order_info <- order_info()
        #         if (is.null(order_info)) return(NULL)
        #         
        #         status <- order_info %>% pull(`OP STATUS`)
        #         reason_code <- order_info %>% pull(`OP REASON CODE`)
        #         date <- order_info %>% pull(`OP STATUS CHANGED`)
        #         comment <- order_info %>% pull(`OP STATUS COMMENT`)
        #         
        #         paste0(
        #                 "<b>", if (!is.na(status)) status, "</b><br/>",
        #                 if (!is.na(date)) date, "<br/>",
        #                 if (!is.na(reason_code)) reason_code, "<br/>",
        #                 "<br/>",
        #                 if (!is.na(comment)) comment
        #                 
        #         ) %>% HTML
        # })
        # output$returns_orders_admin <- renderText({
        #         
        #         order_info <- order_info()
        #         if (is.null(order_info)) return(NULL)
        #         
        #         status <- order_info %>% pull(`ADMIN STATUS`)
        #         reason_code <- order_info %>% pull(`ADMIN REASON CODE`)
        #         date <- order_info %>% pull(`ADMIN STATUS CHANGED`)
        #         comment <- order_info %>% pull(`ADMIN STATUS COMMENT`)
        #         
        #         paste0(
        #                 "<b>", if (!is.na(status)) status, "</b><br/>",
        #                 if (!is.na(date)) date, "<br/>",
        #                 if (!is.na(reason_code)) reason_code, "<br/>",
        #                 "<br/>",
        #                 if (!is.na(comment)) comment
        #                 
        #         ) %>% HTML
        # })
        # output$returns_orders_last <- renderText({
        #         
        #         order_info <- order_info()
        #         if (is.null(order_info)) return(NULL)
        #         
        #         date <- order_info %>% pull(`LAST COMMENT DATE`)
        #         comment <- order_info %>% pull(`LAST STATUS COMMENT`)
        #         
        #         paste0(
        #                 if (!is.na(date)) date, "<br/>",
        #                 "<br/>",
        #                 if (!is.na(comment)) comment
        #                 
        #         ) %>% HTML
        # })
        
        
        
        # observe({
        #         if (input$close > 0) stopApp()                             # stop shiny
        # })
        
}
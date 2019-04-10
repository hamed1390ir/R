# General -----------------------------------------------------------------
info.address <- "//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/Hamed/info.xlsx"
setwd("//nsdstor1/SHARED/NSDComp")


# Package Checking -----------------------------------------------------------------
check.packages <- function(pkg){
        
        new.pkg <- pkg[!(pkg %in% installed.packages()[, "Package"])]
        if (length(new.pkg)) 
                install.packages(new.pkg, dependencies = TRUE)
        sapply(pkg, require, character.only = TRUE)
}

packages <- 
        c(
                "shinycssloaders", "markdown", "shiny", "readxl", 
                "markdown", "readr", "googleVis", "dplyr", 
                "ggplot2", "bizdays", "timeDate", "openxlsx", 
                "stringr", "magick"
        )
check.packages(packages)


# Authorized People -------------------------------------------------------

ICON_emails_authorized_people <- 
        c(
                "hyazarloo",
                "nbowley"
        )
Return_Coordinators <- 
        c(
                "hyazarloo", 
                "mreyes", "swilliams", "skamrani", "mreyes", 
                "ljacobscobb", "canderson", "SRaza", "hsylla", 
                "shassanzada", "eperfilova", "amilstead"
        )
Transportation_Coordinators <- 
        c(
                "hyazarloo", 
                "rdadhich"
        )
Amazon_team <- 
        c(
                "hyazarloo", 
                "ushah", "jyoo2", "nkomarova", "jhooshangi",
                "dramos", "sghazal", "kpiczon", "wunderwood", "jhooshangi"
        )
JCP_team <- 
        c(
                "hyazarloo",
                "mhannah"
        )
Agent_Scorecards_emails_authorized_people <- 
        c(
                "hyazarloo",
                "brichardson", "dlowery", "tosborne", "cwashington", 
                "jsomarriba", "ssmith", "cjones", "dpaccione", 
                "jparis", "jdixon"
        )

HD_team <- 
        c(
                "hyazarloo",
                "wdavey"
        )

# calander ----------------------------------------------------------------

Beg.Year <- 2010
End.Year <- getRmetricsOptions("currentYear") + 1

# find NSD holidays between Beg.Year and current year
holidays <- 
        c(
                USNewYearsDay(Beg.Year:End.Year),
                USMemorialDay(Beg.Year:End.Year),
                USIndependenceDay(Beg.Year:End.Year),
                USLaborDay(Beg.Year:End.Year),
                USThanksgivingDay(Beg.Year:End.Year),
                USChristmasDay(Beg.Year:End.Year)
        ) %>% 
        as.Date

# create the NSD calander
create.calendar(
        name="NSD_CalandeOr", 
        holidays=holidays, 
        weekdays=c("sunday", "saturday"),
        adjust.from=adjust.next, 
        adjust.to=adjust.previous
)

rm("holidays", "Beg.Year", "End.Year")


# Import ------------------------------------------------------------------

LMRegions_State <- read_excel(info.address, sheet = "LMRegions_State")
RRegions_State <- read_excel(info.address, sheet = "RRegions_State")
LMRegions_LMC <- read_excel(info.address, sheet = "LMRegions_LMC")
RRegions_RC <- read_excel(info.address, sheet = "RRegions_RC")
Agents <- 
        read_excel(info.address, sheet = "Agents") %>% 
                merge(LMRegions_State, by = "State") %>% 
                merge(RRegions_State, by = "State") %>% 
                merge(LMRegions_LMC, by = "LM_Region") %>% 
                merge(RRegions_RC, by = "R_Region") %>% 
                select(-State_Name.y) %>% 
                dplyr::rename(State_Name = State_Name.x)
People <- read_excel(info.address, sheet = "People")
Metrics <- read_excel(info.address, sheet = "Metrics")
WOW <- read_excel(info.address, sheet = "WOW")

Accounts <- readRDS('//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/Accounts.rds')
Delivery.Counts.Columns <- readRDS('//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/Delivery.Counts.Columns.rds')
Delivery.Metrics.Columns <- readRDS('//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/Delivery.Metrics.Columns.rds')
Returns.Counts.Columns <- readRDS('//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/Returns.Counts.Columns.rds')
Return.Metrics.Columns <- readRDS('//nsdstor1/SHARED/Hamed/Customer_360_Report/R_360/Return.Metrics.Columns.rds')


# Functions ---------------------------------------------------------------

email_address <- function(address, mode) {
        if (mode) {
                agents_email_address <- paste0(Sys.info()["user"][[1]], "@nonstopdelivery.com")
        } else {
                agents_email_address <- address
        }
}
nsd_signature <- function(nsd_id) {
        first_name <- People %>% filter(id == nsd_id) %>% pull(first_name)
        last_name <- People %>% filter(id == nsd_id) %>% pull(last_name)
        ext <- People %>% filter(id == nsd_id) %>% pull(ext)
        title <- People %>% filter(id == nsd_id) %>% pull(title)
        
        paste0(
                '<br>', '<br>', '<br>', 'Thank you,', '<br>',
                '<br><p style="margin: 0;">', first_name," ", last_name, "</p>",
                '<p style="margin: 0;">', title, "</p>",
                '<p style="margin: 0;">P: (703) 964-9500 Ext:', ext, "</p>",
                '<p style="margin: 0;">', nsd_id, "@shipnsd.com</p>",
                '<p style="margin: 0;">www.shipnsd.com</p>'
        )
}
RH_return_service_types <- c("RT1", "T1R", "T2R", "WH1R", "WH2R")
RC_from_Region <- function(region) {
        RRegions_RC %>% 
                filter(R_Region %in% region) %>% 
                pull(RC_Name)
}
type_from_metric <- function(metric, svctyp) {
        
        return_designation <- category_from_servicetype(svctyp)
                
        Metrics %>% 
                filter(caption == metric, category == return_designation) %>% 
                pull(type)
}
authorization.check <- function(authorized_people){
        if (!Sys.info()["user"] %in% authorized_people) {
                showModal(modalDialog(
                        title = "CAUTION!",
                        "It looks like you do not have permission to take this action!",
                        easyClose = TRUE,
                        footer = NULL
                ))
                return(NULL)
        } else {return(1)}
}
difdat<-function(first.date,second.date){
        # This is a function to calculte the difference between first.date and second.date.
        # If second.date < first.date or (second.date - first.date) > 120, returns NA
        
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

# report_name <- "Return_Daily_Report"
# data <- Return_TM_Region_Categorised

UTF8_compatible_dataframe <- function(x) {
        
        if (is.null(x)) return(NULL)
        les_noms <- names(x)
        A <- x %>%
                lapply(function(x) {if (is.character(x)) {x %>% iconv(to = "UTF-8")} else x}) %>%
                as.data.frame(stringsAsFactors = FALSE)
        names(A) <- les_noms
        A
        
        
}
create_return_daily_worksheet <- function(data, report_name) {
        wb <- openxlsx::createWorkbook()
        openxlsx::addWorksheet(wb, report_name, zoom = 85)
        # data <-
        #         data %>%
        #         lapply(function(x) {if (is.character(x)) {x %>% iconv(to = "UTF-8")} else x}) %>%
        #         as.data.frame(stringsAsFactors = FALSE)
        openxlsx::writeData(wb, 1, data, startRow = 1, startCol = 1)
        
        hidden_cols <- rep(FALSE, 23)
        hidden_cols[3] <- TRUE
        hidden_cols[20] <- TRUE
        
        widths <- c(6, 6, 12, 9, 10, 10, 17, 9, 6, 15, 16, 5, 5, 7, 11, 6, 17, 27, 6, 6, 17, 43, 17)
        setColWidths(wb, sheet = report_name, 1:23, widths = widths, hidden = hidden_cols)
        
        setRowHeights(wb, sheet = report_name, 1:nrow(data), heights = 39)
        
        options("openxlsx.dateFormat" = "m/dd")
        
        generalStyle <-
                createStyle(
                        fontName = "Arial" ,
                        fontSize = 10,
                        fontColour = "#000000",
                        halign = "center",
                        valign = "bottom",
                        fgFill = "#FFFF99",
                        border="TopBottomLeftRight",
                        borderColour = "#000000",
                        wrapText = TRUE
                )
        
        jobdateStyle <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            fgFill = "#FFFF99",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = TRUE, 
                            numFmt = "Date")
        conditionalFormatting(wb, sheet = report_name, cols = 1, rows = 1:nrow(data)+1, rule = NULL, style = c("red", "yellow", "green"), type = "colourScale")
        acctStyle <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            fgFill = "#FFFF99",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = FALSE, 
                            numFmt = "TEXT")
        rcStyle <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            fgFill = "#FFCC99",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = FALSE)
        amStyle <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = FALSE, 
                            fgFill = "#CCFFCC")
        nsdnumstyle <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            fgFill = "#FFFF99",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = FALSE, 
                            textDecoration = c("bold"))
        agentstyle <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            fgFill = "#FFFF99",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = FALSE, 
                            textDecoration = c("bold"))
        svcstyle <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            fgFill = "#FFFF99",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = FALSE)
        citystyle <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            fgFill = "#FFFF99",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = TRUE)
        ststyle <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            fgFill = "#FFFF99",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = FALSE)
        wgtstyle <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            fgFill = "#FFFF99",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = FALSE)
        milestyle <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            fgFill = "#FFFF99",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = FALSE)
        statstyle <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            fgFill = "#FFFF99",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = FALSE, 
                            textDecoration = c("bold"))
        
        options("openxlsx.dateFormat" = "m/dd")
        notedatestyle <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            fgFill = "#FFFF99",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = FALSE, 
                            numFmt = "Date")
        instructionsstyle <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = TRUE, 
                            fgFill = "#FFFFFF")
        emailstyle <-
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = TRUE, 
                            fgFill = "#FFCC99")
        telnumstyle <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = FALSE, 
                            fgFill = "#FFCC99")
        
        
        headerStyle <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = TRUE, 
                            valign = "top", 
                            fgFill = "#FFC000", 
                            textDecoration = c("bold", "underline"))
        
        addStyle(wb, sheet = 1, generalStyle, rows = 1:nrow(data)+1, cols = 1:23, gridExpand = TRUE)
        addStyle(wb, sheet = 1, jobdateStyle, rows = 1:nrow(data)+1, cols = 1, gridExpand = TRUE)
        addStyle(wb, sheet = 1, acctStyle, rows = 1:nrow(data)+1, cols = 2, gridExpand = TRUE)
        addStyle(wb, sheet = 1, rcStyle, rows = 1:nrow(data)+1, cols = 4, gridExpand = TRUE)
        addStyle(wb, sheet = 1, amStyle, rows = 1:nrow(data)+1, cols = 5, gridExpand = TRUE)
        addStyle(wb, sheet = 1, nsdnumstyle, rows = 1:nrow(data)+1, cols = 6, gridExpand = TRUE)
        addStyle(wb, sheet = 1, agentstyle, rows = 1:nrow(data)+1, cols = 8, gridExpand = TRUE)
        addStyle(wb, sheet = 1, svcstyle, rows = 1:nrow(data)+1, cols = 9, gridExpand = TRUE)
        addStyle(wb, sheet = 1, citystyle, rows = 1:nrow(data)+1, cols = 11, gridExpand = TRUE)
        addStyle(wb, sheet = 1, ststyle, rows = 1:nrow(data)+1, cols = 12, gridExpand = TRUE)
        addStyle(wb, sheet = 1, wgtstyle, rows = 1:nrow(data)+1, cols = 13, gridExpand = TRUE)
        addStyle(wb, sheet = 1, milestyle, rows = 1:nrow(data)+1, cols = 14, gridExpand = TRUE)
        addStyle(wb, sheet = 1, statstyle, rows = 1:nrow(data)+1, cols = 15, gridExpand = TRUE)
        addStyle(wb, sheet = 1, notedatestyle, rows = 1:nrow(data)+1, cols = 19, gridExpand = TRUE)
        addStyle(wb, sheet = 1, instructionsstyle, rows = 1:nrow(data)+1, cols = 21, gridExpand = TRUE)
        addStyle(wb, sheet = 1, emailstyle, rows = 1:nrow(data)+1, cols = 22, gridExpand = TRUE)
        addStyle(wb, sheet = 1, telnumstyle, rows = 1:nrow(data)+1, cols = 23, gridExpand = TRUE)
        
        addStyle(wb, sheet = 1, headerStyle, rows = 1, cols = 1:23, gridExpand = TRUE)
        
        addFilter(wb, sheet = 1, row = 1, cols = 1:ncol(data))
        freezePane(wb, sheet = 1, firstRow = TRUE) 
        
        wb
        # saveWorkbook(wb, "addStyleExample.xlsx", overwrite = TRUE)
        # openxlsx::openXL(wb)
}
create_inbound_report_worksheet <- function(data, report_name) {
        wb <- openxlsx::createWorkbook()
        openxlsx::addWorksheet(wb, report_name, zoom = 85)
        openxlsx::writeData(wb, 1, data, startRow = 1, startCol = 1)
        
        widths <- c(11.46, 8.36, 8.91, 21.36, 21.55, 15.55, 14.91, 8.36, 13.82, 7.91, 8.18, 5.18)
        setColWidths(wb, sheet = report_name, 1:12, widths = widths)
        
        setRowHeights(wb, sheet = report_name, 1:nrow(data), heights = 14.5)
        
        
        
        general_style <-
                createStyle(
                        fontName = "Arial" ,
                        fontSize = 10,
                        fontColour = "#000000",
                        halign = "center",
                        valign = "bottom",
                        fgFill = "#ffffff",
                        border="TopBottomLeftRight",
                        borderColour = "#000000",
                        wrapText = TRUE,
                        numFmt = "TEXT"
                )
        
        NSD_number_style <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            fgFill = "#f1f442",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = TRUE, 
                            numFmt = "TEXT")
        
        options("openxlsx.dateFormat" = "m/dd")
        ETA_style <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            fgFill = "#61f95c",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = FALSE, 
                            numFmt = "Date")
        headers_style <- 
                createStyle(fontName = "Arial" ,
                            fontSize = 10,
                            fontColour = "#000000",
                            halign = "center",
                            valign = "bottom",
                            fgFill = "#f9c75c",
                            border="TopBottomLeftRight",
                            borderColour = "#000000",
                            wrapText = FALSE, 
                            numFmt = "TEXT")
        
        
        addStyle(wb, sheet = 1, general_style, rows = 1:nrow(data)+1, cols = 1:12, gridExpand = TRUE)
        addStyle(wb, sheet = 1, NSD_number_style, rows = 1:nrow(data)+1, cols = 2, gridExpand = TRUE)
        addStyle(wb, sheet = 1, ETA_style, rows = 1:nrow(data)+1, cols = 12, gridExpand = TRUE)
        addStyle(wb, sheet = 1, headers_style, rows = 1, cols = 1:12, gridExpand = TRUE)
        freezePane(wb, sheet = 1, firstRow = TRUE) 
        
        wb
        # saveWorkbook(wb, "addStyleExample.xlsx", overwrite = TRUE)
        # openxlsx::openXL(wb)
}
simpleCap <- function(x) {
        s <- strsplit(x, " ")[[1]]
        paste(toupper(substring(s, 1, 1)), substring(s, 2),
              sep = "", collapse = " ")
}
amazon_email_item_path <- function(template, type) {
        paste0(
                "//nsdstor1/SHARED/NSDComp/Amazon/",
                "amazon_",
                switch (template,
                        "Template 1" = "template_1_",
                        "Template 2" = "template_2_",
                        "Template 3" = "template_3_",
                        "Template 4" = "template_4_",
                        "Template 5" = "template_5_",
                        "Template 6" = "template_6_",
                        "Estes" = "template_estes_"
                ),
                type,
                ".txt"
        )
}
amazon_email_item <- function(template, type) {
        amazon_email_item_path(template, type) %>% 
                readChar(file.info(.)$size)
}
results <- function(data){
        # by can take "Agent", "State", "Region"
        data %>%
                group_by(Region, State, Agent) %>%
                {
                        if (names(data)[1] == "Date") {
                                dplyr::summarize(., 
                                                 value = n(),
                                                 Miles = sum(Mileage, na.rm = T)
                                                 )
                        } else {
                                dplyr::summarize(., 
                                                 value = mean(Value),
                                                 Count = n(),
                                                 Miles = sum(Mileage, na.rm = T)
                                )
                        }
                } %>% 
                mutate(Agent = Agent %>% str_replace_all("UNYSON", "") %>% str_trim) %>% 
                as.data.frame
}
category_from_servicetype <- function(svctyp){
        # input: Service Type
        # Output: Delivery or Return
        if (svctyp %in% service.types[["Delivery"]]) "Delivery" else "Return"
}
percent <- function(x, digits = 0, format = "f", ...) {
        paste0(formatC(100 * x, format = format, digits = digits, ...), "%")
}
highest_color <- function(metric, svctyp) {
        
        return_designation <- category_from_servicetype(svctyp)
        
        Metrics %>% 
                filter(caption == metric, category == return_designation) %>% 
                pull(higher_better) %>% 
                {if (.) "#4286f4" else "#f24b18"}
        
}
lowest_color <- function(metric, svctyp) {
        
        
        return_designation <- category_from_servicetype(svctyp)
        
        Metrics %>% 
                filter(caption == metric, category == return_designation) %>% 
                pull(higher_better) %>% 
                {if (.) "#f24b18" else "#4286f4"}
        
}
Amazon_Additiona_Email_from_Agent <- function(agent_name){
  if (!agent_name %in% (Agents$Agent %>% unique)) return(NULL)
        additional_email <- 
                Agents %>% 
                filter(Agent == agent_name) %>% 
                pull(`Amazon Additional Emails`)
        if (is.na(additional_email)) {return(NULL)} else {additional_email}
}
Assistant_Process <- function(data) {
        rbind(
                data %>% 
                        select(Date, `Tracking ID`, Driver, `Driver Accurate Conf#`) %>% 
                        mutate(`DA Job` = "Driver") %>% 
                        dplyr::rename(`Delivery Associate` = Driver, `Accurate Conf#` = `Driver Accurate Conf#`),
                data %>% 
                        select(Date, `Tracking ID`, Helper, `Helper Accurate Conf#`) %>% 
                        mutate(`DA Job` = "Helper") %>% 
                        dplyr::rename(`Delivery Associate` = Helper, `Accurate Conf#` = `Helper Accurate Conf#`)
        )
}
Finalize <- function(data, location) {
        data %>% 
                mutate(Location = location) %>% 
                filter(!is.na(`Accurate Conf#`)) %>% 
                select(Date, Location, `Tracking ID`, `Delivery Associate`, `Accurate Conf#`, `DA Job`) %>% 
                arrange(Date, `Tracking ID`)
}



# JCP ---------------------------------------------------------------------




# Constant ----------------------------------------------------------------

amazon_east_coast <- 
        c(
                "MIAMI_FL",
                "CHANTILLY_VA",
                "PHILADELPHIA_PA",
                "CHICAGO_IL",
                "SALT LAKE CITY_UT",
                "BELTSVILLE_MD",
                "LORIS_SC",
                "ORLANDO_FL",
                "TAMPA_FL",
                "JACKSONVILLE_FL",
                "CLINTON_PA",
                "DURHAM_NC",
                "BIRMINGHAM_AL",
                "OWENSBORO_KY",
                "WILLISTON_VT"
        )
amazon_west_coast <- 
        c(
                "DES MOINES_IA",
                "PORTLAND_OR",
                "COON RAPIDS_MN",
                "HILLSIDE_IL",
                "WEST VALLEY CITY_UT",
                "FARMERS BRANCH_TX",
                "LAS VEGAS_NV",
                "SAINT LOUIS_MO",
                "SAN ANTONIO_TX"
        )

service.types <- 
        list(`Delivery` = c("All of the Deliveries", "Basic Delivery", "Threshold Delivery", "White Glove Delivery"),
             `Return` = c("All of the Returns", "Threshold Return", "White Glove Return", "Linehaul Return")
        )
map_scale = 1.8
week_start <- 1 # Monday

good_job_words <- c(
        "You've got it made!", "Sensational!", "You're doing fine.", "Super!", 
        "You've got your brain in gear today.", "Good thinking.", "That's right!", 
        "That's better.", "Good going.", "That's good!", "Excellent!", "Wonderful!",
        "You are very good at that.", "That was first class work.", "That's a real work of art.",
        "Good work!", "That's the best ever.", "Superb!",
        "Exactly right!", "You did that very well.", "Good remembering!", 
        "You've just about got it.", "Perfect!", "You've got that down pat.", 
        "You are doing a good job!", "That's better than ever.", "You certainly did well today.",
        "That's it!", "Much better!", "Keep it up!", "Now you've figured it out.", "Fine!", 
        "Outstanding!", "Great!", "Nice going.", "You're really improving.",
        "I knew you could do it.", "Fantastic!", "You are learning a lot.",
        "Congratulations!", "Tremendous!", "Good going.",
        "Not bad.", "That's great.", "I'm impressed.", "Keep working on it; you're improving.",
        "Congratulations, you got it right!", "You must have been practicing.",
        "Now you have it.", "You did a lot of work today.", "That's it.",
        "You are learning fast.", "Marvelous!", "I like that.",
        "Good for you!", "Cool!", "Way to go.",
        "Couldn't have done it better myself.",
        "Now that's what I call a fine job.",
        "You've just about mastered that.",
        "Beautiful!", "You've got the hang of it!", "That's an interesting way of looking at it.",
        "One more time and you'll have it.", "I've never seen anyone do it better.",
        "That's the right way to do it.", "It's a classic.", "Super-Duper!",
        "You did it that time!", "Right on!", 'Out of sight.',
        "You're getting better and better.",
        "Congratulations, you only missed ..",
        "It looks like you've put a lot of work into this.",
        "You're on the right track now.", "Keep on trying!", "Good for you!",
        "Nice going.", "Good job!", "You remembered!",
        "You haven't missed a thing.", "That's really nice.", "Thanks!",
        "Wow!", "What neat work!", "That's A work.",
        "That's the way.", "That's clever.", "Very interesting.",
        "Keep up the good work.", "You make it look easy.", "Good thinking",
        'Terrific!', "Muy Bien!", "That's a good point.",
        "Nothing can stop you now.", "Superior work.", 'Nice going.',
        "That's the way to do it.", "I knew you could do it.", "That's coming along nicely."
)







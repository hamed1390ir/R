# General -----------------------------------------------------------------
check.packages <- function(pkg){
        new.pkg <- pkg[!(pkg %in% installed.packages()[, "Package"])]
        if (length(new.pkg)) 
                install.packages(new.pkg, dependencies = TRUE)
        sapply(pkg, require, character.only = TRUE)
}

packages <- c("DT", "shinycustomloader", "leaflet", "shinythemes", "shiny", "markdown", "lubridate", "readr", "dplyr", "readxl", "googleVis", "ggplot2", "bizdays", "timeDate", "openxlsx", "plotly", "shinydashboard", "shinyWidgets", "shinyalert")
check.packages(packages)

source("//nsdstor1/SHARED/NSDComp/init - 3.R")

picked_color <- "#e95420"
picked_color_2 <- "#ffc3a0"

navbarPage("NSDcomp v.3",
           position = "fixed-top",
           tabPanel("Home", align = "center",
                    fluidRow(
                            p("NSDcomp"),
                            tags$head(
                                    tags$style("p {color: #e95420;font-size: 100px;font-style: bold;}"
                                    )
                            )
                    ),
                    fluidRow(
                            img(src='wordcloud_picture.png', align = "center")
                    ),
                    fluidRow(style = "height:50px;"),
                    fluidRow(
                            column(2, 
                                   img(src='amazon.png', align = "center", width = paste0(dev.size("px")[1]/6, "px"))
                            ),
                            column(2,
                                   img(src='homedepot.png', align = "center", width = paste0(dev.size("px")[1]/6, "px"))
                            ),
                            column(2,
                                   img(src='jcpenny.png', align = "center", width = paste0(dev.size("px")[1]/6, "px"))
                            ),
                            column(2,
                                   img(src='purchasingpower.png', align = "center", width = paste0(dev.size("px")[1]/6, "px"))
                            ),
                            column(2,
                                   img(src='wayfair.png', align = "center", width = paste0(dev.size("px")[1]/6, "px"))
                            ),
                            column(2,
                                   img(src='icon.png', align = "center", width = paste0(dev.size("px")[1]/6, "px"))
                            )
                    ),
                    fluidRow(style = "height:50px;")
           ),
           tabPanel("Dashboard",
                    
                    # Dashboard ---------------------------------------------------------------
                    
                    absolutePanel(width = "16%", fixed = F, top = "300px", right = "50px",
                                  wellPanel(
                                          selectInput("dashboard_volumes_account", "Account", c("ALL", Accounts), width = 240),
                                          selectInput("dashboard_volumes_servicetype", "Service Type", service.types, width = 240),
                                          selectInput("dashboard_volumes_item", "Status", Delivery.Counts.Columns, selected = "Created", width = 240),
                                          checkboxGroupInput("dashboard_volumes_system", "System",
                                                             choiceNames =
                                                                     list("RockHopper", "TruckMate"),
                                                             choiceValues =
                                                                     list("RockHopper", "TruckMate"),
                                                             selected = c("RockHopper", "TruckMate")
                                          ),
                                          actionButton("dashboard_volumes_refresh", "Refresh", icon = NULL)
                                  )
                    ),
                    fluidRow(
                            column(8, offset = 1,
                                   h1("Volumes"),
                                   fluidRow(
                                           column(6, 
                                                  radioButtons("dashboard_volumes_general_period", "", c("Daily", "Weekly", "Monthly", "Annually"), selected = "Weekly", inline = T, width = NULL)
                                           ),
                                           column(2, offset = 2,
                                                  downloadButton("dashboard_volumes_download_graph_data", "Graph Data")
                                           ),
                                           column(2,
                                                  downloadButton("dashboard_volumes_download_raw_data", "Raw Data")
                                           )
                                   ),
                                   fluidRow(
                                           tags$div(style = "height:550px;", htmlOutput("dashboard_volumes_general_graph")  %>% withSpinner(color=picked_color))
                                   ),
                                   br(),
                                   fluidRow(
                                           column(4,
                                                  wellPanel(align = "center",
                                                            h4("Last 7 Days"),
                                                            h4(textOutput("dashboard_volumes_general_7_days"))
                                                  )
                                           ),
                                           column(4,
                                                  wellPanel(align = "center",
                                                            h4("Last 30 Days"),
                                                            h4(textOutput("dashboard_volumes_general_30_days"))
                                                  )
                                           ),
                                           column(4,
                                                  wellPanel(align = "center",
                                                            h4("Last 365 Days"),
                                                            h4(textOutput("dashboard_volumes_general_365_days"))
                                                  ) 
                                           )
                                   )
                            )
                            
                    ),
                    fluidRow(style = "height:50px;"),
                    fluidRow(style = "height:25px;background-color:#e95420;"),
                    fluidRow(style = "height:50px;"),
                    absolutePanel(width = "16%", fixed = F, top = "1250px", right = "50px",
                                  wellPanel(
                                          selectInput("dashboard_metrics_account", "Account", c("ALL", Accounts), width = 240),
                                          selectInput("dashboard_metrics_servicetype", "Service Type", service.types, width = 240),
                                          selectInput("dashboard_metrics_item", "Metrics", Metrics %>% filter(category == "Delivery") %>% pull(caption), selected = "Same Day Receiving %", width = 240),
                                          checkboxGroupInput("dashboard_metrics_system", "System",
                                                             choiceNames =
                                                                     list("RockHopper", "TruckMate"),
                                                             choiceValues =
                                                                     list("RockHopper", "TruckMate"),
                                                             selected = c("RockHopper", "TruckMate")
                                          ),
                                          actionButton("dashboard_metrics_refresh", "Refresh", icon = NULL)
                                  )
                    ),
                    fluidRow(
                            column(8, offset = 1,
                                   h1("Metrics"),
                                   
                                   fluidRow(
                                           column(6, 
                                                  radioButtons("dashboard_metrics_general_period", "", c("Daily", "Weekly", "Monthly", "Annually"), selected = "Weekly", inline = T, width = NULL)
                                           ),
                                           column(2, offset = 2,
                                                  downloadButton("dashboard_metrics_download_graph_data", "Graph Data")
                                           ),
                                           column(2,
                                                  downloadButton("dashboard_metrics_download_raw_data", "Raw Data")
                                           )
                                   ),
                                   fluidRow(
                                           tags$div(style = "height:550px;", htmlOutput("dashboard_metrics_general_graph")  %>% withSpinner(color=picked_color))
                                   ),
                                   br(),
                                   fluidRow(
                                           column(4,
                                                  wellPanel(align = "center",
                                                            h4("Last 7 Days"),
                                                            h4(textOutput("dashboard_metrics_general_7_days"))
                                                  )
                                           ),
                                           column(4,
                                                  wellPanel(align = "center",
                                                            h4("Last 30 Days"),
                                                            h4(textOutput("dashboard_metrics_general_30_days"))
                                                  )
                                           ),
                                           column(4,
                                                  wellPanel(align = "center",
                                                            h4("Last 365 Days"),
                                                            h4(textOutput("dashboard_metrics_general_365_days"))
                                                  ) 
                                           )
                                   )
                            )
                    )
           ),
           # Volumes -----------------------------------------------------------------
           tabPanel("Volumes",
                    
                    absolutePanel(width = "16%", fixed = TRUE, top = "100px", right = "10px",
                                  wellPanel(
                                          selectInput("counts_account", "Account", c("ALL", Accounts), width = 240),
                                          selectInput("counts_servicetype", "Service Type", service.types, width = 240),
                                          selectInput("counts_item", "Status", Delivery.Counts.Columns, selected = "Created", width = 240),
                                          checkboxGroupInput("counts_system", "System",
                                                             choiceNames =
                                                                     list("RockHopper", "TruckMate"),
                                                             choiceValues =
                                                                     list("RockHopper", "TruckMate"),
                                                             selected = c("RockHopper", "TruckMate")
                                          ),
                                          dateRangeInput(
                                                  "counts_daterange", 
                                                  "", 
                                                  start = today() - 90, 
                                                  end = today(), 
                                                  min = "2013-01-01", 
                                                  max = Sys.Date(), 
                                                  separator = " to "
                                          ),
                                          actionButton("counts_refresh", "Refresh", icon = NULL)
                                  ),
                                  wellPanel(align = "center",
                                            h2(textOutput("counts_numberoforders"))
                                  ),
                                  downloadButton("counts_download_raw_data", "Download Raw Data")
                    ),
                    fluidRow(
                            column(8, offset = 1,
                                   tabsetPanel(type = "tabs",
                                               tabPanel("Agents",
                                                        fluidRow(
                                                                column(2, 
                                                                       fluidRow(style = "height:280px;"),
                                                                       checkboxGroupInput("counts_agent_activity", "",
                                                                                          choiceNames =
                                                                                                  list("Active", "Inactive"),
                                                                                          choiceValues =
                                                                                                  list("Active", "Inactive"),
                                                                                          selected = c("Active", "Inactive")
                                                                       )
                                                                ),
                                                                column(10, align = "center",
                                                                       tags$div(style = "height:700px;", htmlOutput("counts_agents_map")  %>% withSpinner(color=picked_color))
                                                                )
                                                        ),
                                                        fluidRow(
                                                                
                                                        ),
                                                        fluidRow(style = "height:100px;"),
                                                        fluidRow(align = "center",
                                                                 column(10, offset = 1,
                                                                        tags$div(style = "height:600px;", DTOutput("counts_agentstable")  %>% withSpinner(color=picked_color))
                                                                 )
                                                        ),
                                                        fluidRow(
                                                                column(2,
                                                                       downloadButton("counts_download_agents", "Download Table", width = 100)
                                                                )
                                                        ),
                                                        fluidRow(style = "height:100px;"),
                                                        fluidRow(align = "center",
                                                                 column(11,
                                                                        tags$div(style = "height:400px;", plotlyOutput("counts_agents_barchart")  %>% withSpinner(color=picked_color))
                                                                 )   
                                                        ),
                                                        fluidRow(style = "height:100px;"),
                                                        fluidRow(
                                                                column(4, offset = 1,
                                                                       tags$div(style = "height:400px;", plotlyOutput("counts_agents_histogram")  %>% withSpinner(color=picked_color))
                                                                ),
                                                                column(4,  offset = 1,
                                                                       tags$div(style = "height:400px;", plotlyOutput("counts_agents_boxplot")  %>% withSpinner(color=picked_color))
                                                                )
                                                        ),
                                                        fluidRow(style = "height:100px;")
                                                        
                                               ),
                                               tabPanel("States", 
                                                        fluidRow(
                                                                column(10, offset = 2,
                                                                       tags$div(style = "height:700px;", htmlOutput("counts_states_map")  %>% withSpinner(color=picked_color))
                                                                )
                                                        ),
                                                        fluidRow(style = "height:100px;"),
                                                        fluidRow(
                                                                column(11,
                                                                       tags$div(style = "height:600px;", DTOutput("counts_states_table")  %>% withSpinner(color=picked_color))
                                                                )
                                                        ),
                                                        fluidRow(
                                                                column(2,
                                                                       br(),
                                                                       downloadButton("counts_download_states", "States Table", width = 200)
                                                                )
                                                        ),
                                                        fluidRow(style = "height:100px;")
                                               ),
                                               tabPanel("Regions", 
                                                        fluidRow(
                                                                column(10, offset = 2, 
                                                                       tags$div(style = "height:700px;", htmlOutput("counts_regions_map")  %>% withSpinner(color=picked_color))
                                                                )
                                                        ),
                                                        fluidRow(style = "height:100px;"),
                                                        fluidRow(
                                                                column(8, offset = 2, 
                                                                       tags$div(style = "height:400px;", DTOutput("counts_regions_table")  %>% withSpinner(color=picked_color))
                                                                )
                                                        ),
                                                        fluidRow(style = "height:100px;"),
                                                        fluidRow(
                                                                column(8, offset = 2,
                                                                       tags$div(style = "height:400px;", plotlyOutput("counts_regions_pie")  %>% withSpinner(color=picked_color))
                                                                )
                                                        ),
                                                        fluidRow(style = "height:100px;")
                                               )
                                   )
                                   
                            )
                    )
           ),
           # Metrics -----------------------------------------------------------------
           tabPanel("Metrics",
                    absolutePanel(width = "16%", fixed = TRUE, top = "100px", left = "10px",
                                  wellPanel(
                                          selectInput("metrics_account", "Account", c("ALL", Accounts), width = 240),
                                          selectInput("metrics_servicetype", "Service Type", service.types, width = 240),
                                          selectInput("metrics_item", "Metrics", Metrics %>% filter(category == "Delivery") %>% pull(caption), selected = "Same Day Receiving %", width = 240),
                                          checkboxGroupInput("metrics_system", "System",
                                                             choiceNames =
                                                                     list("RockHopper", "TruckMate"),
                                                             choiceValues =
                                                                     list("RockHopper", "TruckMate"),
                                                             selected = c("RockHopper", "TruckMate")
                                          ),
                                          dateRangeInput(
                                                  "metrics_daterange", 
                                                  "", 
                                                  start = today() - 90, 
                                                  end = today(), 
                                                  min = "2013-01-01", 
                                                  max = Sys.Date(), 
                                                  separator = " to "
                                          ),
                                          actionButton("metrics_refresh", "Refresh", icon = NULL)
                                  ),
                                  wellPanel(align = "center",
                                            h2(textOutput("metrics_average"))
                                  )
                                  # actionButton("delivery_scorecards_emails_button", "Send Agent Scorecards", width = 200)
                    ),
                    fluidRow(
                            column(8, offset = 3,
                                   tabsetPanel(type = "tabs",
                                               tabPanel("Agents",
                                                        fluidRow(
                                                                column(2, 
                                                                       fluidRow(style = "height:280px;"),
                                                                       checkboxGroupInput("metrics_agent_activity", "",
                                                                                          choiceNames =
                                                                                                  list("Active", "Inactive"),
                                                                                          choiceValues =
                                                                                                  list("Active", "Inactive"),
                                                                                          selected = c("Active", "Inactive")
                                                                       )
                                                                ),
                                                                column(10,
                                                                       tags$div(style = "height:700px;", htmlOutput("metrics_agents_map")  %>% withSpinner(color=picked_color))
                                                                )
                                                        ),
                                                        fluidRow(style = "height:100px;"),
                                                        fluidRow(
                                                                column(10, offset = 1,
                                                                       tags$div(style = "height:600px;", DTOutput("metrics_agentstable")  %>% withSpinner(color=picked_color))
                                                                )
                                                        ),
                                                        fluidRow(
                                                                column(2,
                                                                       downloadButton("metrics_download_agents", "Download Table", width = 100)
                                                                )
                                                        ),
                                                        fluidRow(style = "height:100px;"),
                                                        fluidRow(
                                                                column(11,
                                                                       tags$div(style = "height:400px;", plotlyOutput("metrics_agents_barchart")  %>% withSpinner(color=picked_color))
                                                                )   
                                                        ),
                                                        fluidRow(style = "height:100px;"),
                                                        fluidRow(
                                                                column(4, offset = 1,
                                                                       tags$div(style = "height:400px;", plotlyOutput("metrics_agents_histogram")  %>% withSpinner(color=picked_color))
                                                                ),
                                                                column(4,  offset = 1,
                                                                       tags$div(style = "height:400px;", plotlyOutput("metrics_agents_boxplot")  %>% withSpinner(color=picked_color))
                                                                )
                                                        ),
                                                        fluidRow(style = "height:100px;")
                                                        
                                               ),
                                               tabPanel("States", 
                                                        fluidRow(
                                                                column(10, offset = 2,
                                                                       tags$div(style = "height:700px;", htmlOutput("metrics_states_map")  %>% withSpinner(color=picked_color))
                                                                )
                                                        ),
                                                        fluidRow(style = "height:100px;"),
                                                        fluidRow(
                                                                column(11,
                                                                       tags$div(style = "height:600px;", DTOutput("metrics_states_table")  %>% withSpinner(color=picked_color))
                                                                )
                                                        ),
                                                        fluidRow(
                                                                column(2,
                                                                       br(),
                                                                       downloadButton("metrics_download_states", "States Table", width = 200)
                                                                )
                                                        ),
                                                        fluidRow(style = "height:100px;")
                                               ),
                                               tabPanel("Regions", 
                                                        fluidRow(
                                                                column(10, offset = 2, 
                                                                       tags$div(style = "height:700px;", htmlOutput("metrics_regions_map")  %>% withSpinner(color=picked_color))
                                                                )
                                                        ),
                                                        fluidRow(style = "height:100px;"),
                                                        fluidRow(
                                                                column(8, offset = 2, 
                                                                       tags$div(style = "height:400px;", DTOutput("metrics_regions_table")  %>% withSpinner(color=picked_color))
                                                                )
                                                        ),
                                                        fluidRow(style = "height:100px;")
                                               ),
                                               tabPanel("Wow",
                                                        column(1, 
                                                               radioButtons("metrics_wow_periods", "", c("Week", "Month", "Year"), selected = "Week", inline = FALSE, width = NULL)
                                                        ),
                                                        column(11, 
                                                               
                                                               fluidRow(
                                                                       h2("Delivery"),
                                                                       tags$div(style = "height:300px;", DTOutput("metrics_wow_delivery_table")  %>% withSpinner(color=picked_color))
                                                               ),
                                                               br(),
                                                               fluidRow(
                                                                       h2("Return"),
                                                                       tags$div(style = "height:300px;", DTOutput("metrics_wow_return_table")  %>% withSpinner(color=picked_color))
                                                               )
                                                        )
                                               )
                                   )
                                   
                            )
                    )
           ),
           
           # Returns -----------------------------------------------------------------
           
           
           tabPanel(
                  "Returns",
                   fluidPage(
                           useShinyalert(),
                           # Valid themes are: cerulean, cosmo, cyborg, darkly, flatly, journal, lumen, paper, readable, sandstone, simplex, slate, spacelab, superhero, united, yeti
                           
                           theme = shinytheme("united"),
                           tags$style(type="text/css", "#RH_returns_returns_calendarplot.recalculating { opacity: 1.0; }"),
                           tags$style(type="text/css", "#TM_returns_returns_calendarplot.recalculating { opacity: 1.0; }"),
                           tags$style(type="text/css", "body {padding-top: 100px;}"),
                           shinyjs::useShinyjs(),
                           column(1,
                                  fluidRow(
                                          dropdownButton(
                                                  tags$h3("Settings"),
                                                  selectInput("region", 
                                                              "Region",
                                                              c("ALL", RRegions_RC$R_Region), 
                                                              selected = "ALL",
                                                              multiple = T
                                                  ),
                                                  fileInput("button_TM_update_data", "TruckMate Report", accept = ".csv", buttonLabel = "Browse..."),
                                                  circle = TRUE, status = "danger", icon = icon("gear"), width = "300px",
                                                  tooltip = tooltipOptions(title = "Click to change your preferences !")
                                          )  
                                  ),
                                  fluidRow(style = "height:10px;"),
                                  fluidRow(
                                          circleButton(
                                                  inputId = "return_agent_database_check",
                                                  size = "default",
                                                  status = "warning",
                                                  icon = icon("thumbs-up"),
                                                  tooltip = tooltipOptions(title = "Click to change your preferences !")
                                          )
                                  )
                           ),
                           column(10,
                                  tabsetPanel(type = "tabs",
                                              tabPanel("Orders", align = "left", 
                                                       br(),
                                                       fluidRow(align = "left",
                                                                column(6,
                                                                       column(4,
                                                                              selectInput("return_orders_selectinput", "Order Number", choices = "-")
                                                                       ),
                                                                       column(1,
                                                                              imageOutput("returns_orders_items", height = 75, width = 75)
                                                                       ),
                                                                       column(1,
                                                                              imageOutput("returns_orders_weight", height = 75, width = 75)
                                                                       ),
                                                                       column(1,
                                                                              imageOutput("returns_orders_miles", height = 75, width = 75)
                                                                       ),
                                                                       column(1,
                                                                              imageOutput("return_orders_men", height = 75, width = 75)
                                                                       ),
                                                                       column(1,
                                                                              imageOutput("return_orders_servicetype", height = 75, width = 75)
                                                                       ),
                                                                       column(1,
                                                                              imageOutput("return_orders_contact_attempts", height = 75, width = 75)
                                                                       ),
                                                                       column(1,
                                                                              imageOutput("return_orders_disposal", height = 75, width = 75)
                                                                       ),
                                                                       column(1,
                                                                              imageOutput("return_orders_replacement", height = 75, width = 75)
                                                                       )
                                                                )
                                                       ),
                                                       fluidRow(
                                                               column(6,
                                                                      fluidRow(style = "height:250px;",
                                                                               column(12,
                                                                                      DTOutput("returns_orders_description")
                                                                               )
                                                                      ),
                                                                      fluidRow(
                                                                              column(4, align = "center",
                                                                                     wellPanel(id = "tPanel",style = "background:white;overflow-y:scroll;height: 150px",
                                                                                               tags$b(h4("OP")),
                                                                                               br(),
                                                                                               htmlOutput(align = "left","returns_orders_op")
                                                                                               
                                                                                     )
                                                                              ),
                                                                              column(4, align = "center",
                                                                                     wellPanel(id = "tPanel",style = "background:white;overflow-y:scroll;height: 150px",
                                                                                               tags$b(h4("Admin")),
                                                                                               br(),
                                                                                               htmlOutput(align = "left","returns_orders_admin")
                                                                                     )
                                                                              ),
                                                                              column(4, align = "center",
                                                                                     wellPanel(id = "tPanel",style = "background:white;overflow-y:scroll;height: 150px",
                                                                                               tags$b(h4("Last Comment")),
                                                                                               br(),
                                                                                               htmlOutput(align = "left","returns_orders_last")
                                                                                     )
                                                                              )
                                                                      )
                                                                      
                                                               ),
                                                               column(6, 
                                                                      tags$div(leafletOutput("return_orders_map", height = 400) %>% withSpinner(color=picked_color))
                                                                      
                                                               )
                                                       ),
                                                       fluidRow(style = "height:20px;"),
                                                       fluidRow(
                                                               column(3,align = "center",
                                                                      wellPanel(style = paste0("background: ", picked_color_2),
                                                                                tags$b(h4("Shipper")),
                                                                                htmlOutput(align = "left", "returns_orders_shipper_info")
                                                                      )
                                                               ),
                                                               column(3,align = "center",
                                                                      wellPanel(style = paste0("background: ", picked_color_2),
                                                                                tags$b(h4("Agent")),
                                                                                htmlOutput(align = "left","returns_orders_agent_info")
                                                                      )
                                                               ),
                                                               column(3,align = "center",
                                                                      wellPanel(style = paste0("background: ", picked_color_2),
                                                                                tags$b(h4("Carrier")),
                                                                                htmlOutput(align = "left","returns_orders_carrier_info")
                                                                      )
                                                               ),
                                                               column(3,align = "center",
                                                                      wellPanel(style = paste0("background: ", picked_color_2),
                                                                                tags$b(h4("Consignee")),
                                                                                htmlOutput(align = "left","returns_orders_consignee_info")
                                                                      )
                                                               )
                                                       )
                                              ),
                                              tabPanel("Operations",
                                                       column(10,
                                                              column(3,
                                                                     fluidRow(style = "height:30px;"),
                                                                     fluidRow(
                                                                             column(9,
                                                                                    selectInput(
                                                                                            inputId = "return_agent_code",
                                                                                            label = "Select Agent",
                                                                                            choices = c("-")
                                                                                    )
                                                                             )
                                                                     ),
                                                                     fluidRow(
                                                                             align = "center", 
                                                                             h2("Count of Orders"),
                                                                             h3(align = "center", textOutput("returns_explore_TM_count"))
                                                                     ) %>% wellPanel(style = paste0("background: ", picked_color_2)),
                                                                     verbatimTextOutput('return_agent_info')
                                                              ),
                                                              column(9,
                                                                     fluidRow(style = "height:30px;"),
                                                                     tags$div(
                                                                             style = "height:300px;", 
                                                                             htmlOutput("TM_returns_returns_calendarplot")  %>% 
                                                                                     withSpinner(color=picked_color)
                                                                     ),
                                                                     tags$div(
                                                                             style = "height:500px;", 
                                                                             plotlyOutput("TM_return_category_pie", height = 500)  %>% 
                                                                                     withSpinner(color=picked_color)
                                                                     ),
                                                                     fluidRow(style = "height:160px;")
                                                              )
                                                              
                                                       ),
                                                       column(2,
                                                              fluidRow(style = "height:30px;"),
                                                              downloadButton(outputId = "TM_Returns_Daily_Report_download", icon = icon("download"), width = "100%"),
                                                              fluidRow(style = "height:30px;"),
                                                              fileInput("TM_Returns_Daily_Report_upload", "Upload Modified TM", accept = ".xlsx", width = "100%", buttonLabel = "Browse..."),
                                                              fluidRow(style = "height:30px;"),
                                                              actionButton("Returns_send_escalation_emails", "Escalations", width = "100%", height = "200px", style = paste0("background: ", picked_color)),
                                                              fluidRow(style = "height:30px;"),
                                                              actionButton("Returns_send_2_week_emails", "2 Weeks", width = "100%", height = "200px", style = paste0("background: ", picked_color)),
                                                              fluidRow(style = "height:30px;"),
                                                              actionButton("Returns_send_update_emails", "Updates", width = "100%", height = "200px", style = paste0("background: ", picked_color)),
                                                              fluidRow(style = "height:30px;"),
                                                              actionButton("Returns_send_WPU_emails", "WPUs", width = "100%", height = "200px", style = paste0("background: ", picked_color)),
                                                              fluidRow(style = "height:30px;"),
                                                              actionButton("Returns_send_active_emails", "Actives", width = "100%", height = "200px", style = paste0("background: ", picked_color)),
                                                              fluidRow(style = "height:30px;"),
                                                              actionButton("Returns_send_HD_emails", "HD Customers", width = "100%", height = "200px", style = paste0("background: ", picked_color)),
                                                              fluidRow(style = "height:30px;"),
                                                              actionButton("Returns_send_wayfair_emails", "Wayfair Customers", width = "100%", height = "200px", style = paste0("background: ", picked_color)),
                                                              fluidRow(style = "height:30px;"),
                                                              downloadButton("Returns_Create_Unyson_worksheet", "Unyson Worksheet", width = "100%", height = "200px", style = paste0("background: ", picked_color)),
                                                              fluidRow(style = "height:30px;")
                                                       )
                                              ),
                                              tabPanel("Consolidation", align = "center", 
                                                       fluidRow(style = "height:600px;",
                                                                column(6, 
                                                                       h2("Destinations"),
                                                                       tags$div(htmlOutput("return_consolidation_map_destinations_circle", height = 300) %>% withSpinner(color=picked_color)),
                                                                       wellPanel(style = paste0("background: ", picked_color_2),
                                                                                 htmlOutput("returns_consolidation_destination_info")
                                                                       )
                                                                ),
                                                                column(6, 
                                                                       h2("Origins"),
                                                                       tags$div(htmlOutput("return_consolidation_map_origins_circle", height = 300) %>% withSpinner(color=picked_color)),
                                                                       wellPanel(style = paste0("background: ", picked_color_2),
                                                                                 htmlOutput("returns_consolidation_origin_info")
                                                                       )
                                                                )
                                                       ),
                                                       fluidRow(
                                                               column(8, offset = 2,
                                                                      tags$div(DTOutput("return_consolidation_map_table")  %>% withSpinner(color=picked_color))
                                                               )
                                                               
                                                       )
                                                       
                                                       
                                              )
                                  )
                           )
                   )
           ),
           
           # # Accounts  --------------------------------------------------------------
           tabPanel(
                   "Accounts",
                   fluidPage(
                           column(1,
                                  fluidRow(
                                          dropdownButton(
                                                  tags$h3("Settings"),
                                                  fileInput("accounts_TM_update_data", "TruckMate Report", accept = ".csv", buttonLabel = "Browse..."),
                                                  circle = TRUE, status = "danger", icon = icon("gear"), width = "300px",
                                                  tooltip = tooltipOptions(title = "Click to change your preferences!")
                                          )
                                  ),
                                  fluidRow(style = "height:10px;"),
                                  fluidRow(
                                          circleButton(
                                                  inputId = "accounts_agent_database_check",
                                                  size = "default",
                                                  status = "warning",
                                                  icon = icon("thumbs-up"),
                                                  tooltip = tooltipOptions(title = "Check if all of the agents in the TM report exist in the database")
                                          )
                                  )
                           ),
                           column(10,
                                  tabsetPanel(type = "pills",
                                              tabPanel("Home Depot",

                                                       # Home Depot --------------------------------------------------------------
                                                       br(),
                                                       wellPanel(style = "height:500px;",
                                                                 br(),
                                                                 fluidRow(
                                                                         column(8, offset =1,
                                                                                h3("Past OFD Date"),
                                                                                textOutput("HD_OFD_last_sent")
                                                                         ),
                                                                         column(3, style = "margin-top: 25px;",
                                                                                actionButton("HD_past_OFD_send_emails", "Send Past OFD emails", width = 200)
                                                                         )
                                                                 ),
                                                                 br(),
                                                                 fluidRow(
                                                                         column(8, offset =1,
                                                                                textInput("HD_past_OFD_subject", "Subject", value = "//nsdstor1/SHARED/NSDComp/HomeDepot/HD_OFD_template_subject.txt" %>% readChar(file.info(.)$size))
                                                                         ),
                                                                         column(3, style = "margin-top: 25px;",
                                                                                actionButton("HD_past_OFD_subject_save", "Save")
                                                                         )
                                                                 ),
                                                                 br(),
                                                                 fluidRow(
                                                                         column(8, offset =1,
                                                                                textAreaInput("HD_past_OFD_body", "Body", value = "//nsdstor1/SHARED/NSDComp/HomeDepot/HD_OFD_template_body.txt" %>% readChar(file.info(.)$size), resize = "vertical", width = "100%", height = 150)
                                                                         ),
                                                                         column(3, style = "margin-top: 25px;",
                                                                                actionButton("HD_past_OFD_body_save", "Save")
                                                                         )
                                                                 )
                                                                 
                                                       ),
                                                       br(),
                                                       wellPanel(style = "height:500px;",
                                                                 br(),
                                                                 fluidRow(
                                                                         column(8, offset =1,
                                                                                h3("Past Scheduled Date"),
                                                                                textOutput("HD_scheduled_last_sent")
                                                                         ),
                                                                         column(3, style = "margin-top: 25px;",
                                                                                actionButton("HD_past_scheduled_send_emails", "Send Past Scheduled emails", width = 200)
                                                                         )
                                                                 ),
                                                                 br(),
                                                                 fluidRow(
                                                                         column(8, offset =1,
                                                                                textInput("HD_past_scheduled_subject", "Subject", value = "//nsdstor1/SHARED/NSDComp/HomeDepot/HD_scheduled_template_subject.txt" %>% readChar(file.info(.)$size))
                                                                         ),
                                                                         column(3, style = "margin-top: 25px;",
                                                                                actionButton("HD_past_scheduled_subject_save", "Save")
                                                                         )
                                                                 ),
                                                                 br(),
                                                                 fluidRow(
                                                                         column(8, offset =1,
                                                                                textAreaInput("HD_past_scheduled_body", "Body", value = "//nsdstor1/SHARED/NSDComp/HomeDepot/HD_scheduled_template_body.txt" %>% readChar(file.info(.)$size), resize = "vertical", width = "100%", height = 150)
                                                                         ),
                                                                         column(3, style = "margin-top: 25px;",
                                                                                actionButton("HD_past_scheduled_body_save", "Save")
                                                                         )
                                                                 )
                                                                 
                                                       ),
                                                       br(),
                                                       wellPanel(style = "height:500px;",
                                                                 br(),
                                                                 fluidRow(
                                                                         column(8, offset =1,
                                                                                h3("ETA Report"),
                                                                                textOutput("HD_ETA_last_sent")
                                                                         ),
                                                                         column(3, style = "margin-top: 25px;",
                                                                                actionButton("HD_ETA_report_send_emails", "Send ETA report emails", width = 200)
                                                                         )
                                                                 ),
                                                                 br(),
                                                                 fluidRow(
                                                                         column(8, offset =1,
                                                                                textInput("HD_ETA_report_subject", "Subject", value = "//nsdstor1/SHARED/NSDComp/HomeDepot/HD_ETA_template_subject.txt" %>% readChar(file.info(.)$size))
                                                                         ),
                                                                         column(3, style = "margin-top: 25px;",
                                                                                actionButton("HD_ETA_report_subject_save", "Save")
                                                                         )
                                                                 ),
                                                                 br(),
                                                                 fluidRow(
                                                                         column(8, offset =1,
                                                                                textAreaInput("HD_ETA_report_body", "Body", value = "//nsdstor1/SHARED/NSDComp/HomeDepot/HD_ETA_template_body.txt" %>% readChar(file.info(.)$size), resize = "vertical", width = "100%", height = 150)
                                                                         ),
                                                                         column(3, style = "margin-top: 25px;",
                                                                                actionButton("HD_ETA_report_body_save", "Save")
                                                                         )
                                                                 )
                                                                 
                                                       )
                                              ),
                                              tabPanel("J. C. Penney",
                                                       
                                                       # JC Penny ----------------------------------------------------------------
                                                       br(),
                                                       fluidRow(style = "height:800px;",
                                                                fluidRow(
                                                                        column(2, align = "center",
                                                                               wellPanel(
                                                                                       fileInput("JCP_input_retail", "Upload retail files (.txt)", accept = ".txt", buttonLabel = "Browse...", multiple = T),
                                                                                       fileInput("JCP_input_online", "Upload online files (.csv)", accept = ".csv", buttonLabel = "Browse...", multiple = T),
                                                                                       fileInput("JCP_input_LTL", "Upload LTL file (.csv)", accept = ".csv", buttonLabel = "Browse...", multiple = FALSE),
                                                                                       fileInput("JCP_input_last", "Upload last list of orders (.csv)", accept = ".csv", buttonLabel = "Browse...", multiple = F)
                                                                               ),
                                                                               wellPanel(
                                                                                       fileInput("JCP_input_EFF", "Upload EFF file (.csv)", accept = ".csv", buttonLabel = "Browse...", multiple = F),
                                                                                       fileInput("JCP_input_zip", "Upload zip file (.csv)", accept = ".csv", buttonLabel = "Browse...", multiple = F)
                                                                               ),
                                                                               downloadButton('JCP_input_download', 'Flat File')
                                                                        ),
                                                                        column(10, align = "center",
                                                                               tabsetPanel(type = "tabs",
                                                                                           tabPanel("ZIPs",
                                                                                                    selectInput("JCP_zip_agent", "Agent", c("ALL"), width = 200),
                                                                                                    textOutput("JCP_zip_count"),
                                                                                                    tags$div(htmlOutput("JCP_zip_map")  %>% withSpinner(color=picked_color))
                                                                                           ),
                                                                                           tabPanel("RETAIL Orders",
                                                                                                    
                                                                                                    fluidRow(align = "center",
                                                                                                             column(11, 
                                                                                                                    DTOutput("JCP_retail_table")  %>% withSpinner(color=picked_color)
                                                                                                             )
                                                                                                    )
                                                                                                    ,
                                                                                                    fluidRow(
                                                                                                            column(3, 
                                                                                                                   verbatimTextOutput('JCP_retail_info')
                                                                                                            )
                                                                                                    )
                                                                                                    
                                                                                                    
                                                                                           ),
                                                                                           tabPanel("Online Orders",
                                                                                                    
                                                                                                    fluidRow(align = "center",
                                                                                                             column(11, 
                                                                                                                    DTOutput("JCP_online_table")  %>% withSpinner(color=picked_color)
                                                                                                             )
                                                                                                    )
                                                                                                    ,
                                                                                                    fluidRow(
                                                                                                            column(3, 
                                                                                                                   verbatimTextOutput('JCP_online_info')
                                                                                                            )
                                                                                                    )
                                                                                                    
                                                                                           ),
                                                                                           tabPanel("LTL Orders",
                                                                                                    fluidRow(align = "center",
                                                                                                             column(11, 
                                                                                                                    DTOutput("JCP_LTL_table")  %>% withSpinner(color=picked_color)
                                                                                                             )
                                                                                                    ),
                                                                                                    fluidRow(
                                                                                                            column(3, 
                                                                                                                   verbatimTextOutput('JCP_LTL_info')
                                                                                                            )
                                                                                                    )
                                                                                                    
                                                                                           )
                                                                                           
                                                                               )
                                                                        )
                                                                )
                                                       )
                                              ),
                                              
                                              # Amazon ------------------------------------------------------------------
                                              
                                              
                                              
                                              tabPanel("Amazon", align = "left", 
                                                       fluidRow(style = "height:25px;"),
                                                       fluidRow(
                                                               column(4,
                                                                      column(9, offset = 2,
                                                                             
                                                                             wellPanel(style = "height:675px;",
                                                                                       h3("OFD, PDD Terminal mass email"),
                                                                                       br(),
                                                                                       fluidRow(
                                                                                               column(12,
                                                                                                      fileInput("delivery_amazon_pdd_input", "Input Excel Workbook", accept = ".xlsx")
                                                                                               )
                                                                                       ),
                                                                                       fluidRow(
                                                                                               column(4,
                                                                                                      selectInput("delivery_amazon_select_template", "Select Template", choices = c("Template 1", "Template 2", "Template 3", "Template 4", "Template 5", "Template 6"))
                                                                                               )
                                                                                       ),
                                                                                       fluidRow(
                                                                                               column(9,
                                                                                                      textInput("delivery_amazon_pdd_subject", "Subject", value = amazon_email_item("Template 1", "subject"))
                                                                                               ),
                                                                                               column(3, style = "margin-top: 25px;",
                                                                                                      actionButton("delivery_amazon_pdd_subject_save", "Save")
                                                                                               )
                                                                                       ),
                                                                                       br(),
                                                                                       fluidRow(
                                                                                               column(9,
                                                                                                      textAreaInput("delivery_amazon_pdd_body", "Body", value = amazon_email_item("Template 1", "body"), resize = "vertical", width = "100%", height = 150)
                                                                                               ),
                                                                                               column(3, style = "margin-top: 25px;",
                                                                                                      actionButton("delivery_amazon_pdd_body_save", "Save")
                                                                                               )
                                                                                       ),
                                                                                       br(),
                                                                                       actionButton("delivery_amazon_send_button", "Send Emails", width = 200)
                                                                             )
                                                                      )
                                                               ),
                                                               column(4,
                                                                      column(9, offset = 1,
                                                                             wellPanel(style = "height:675px;",
                                                                                       h3("Daily Delivery Report "),
                                                                                       br(),
                                                                                       fileInput("delivery_amazon_daily_TM_input", "TM", accept = ".csv"),
                                                                                       fileInput("delivery_amazon_daily_raw_data_input", "Inputs", accept = ".csv", multiple = T),
                                                                                       fluidRow(
                                                                                               column(6, style = "margin-top: 25px;",
                                                                                                      downloadButton("delivery_amazon_create_east_coast_button", "Create East Coast")
                                                                                               ),
                                                                                               column(6, style = "margin-top: 25px;",
                                                                                                      downloadButton("delivery_amazon_create_west_coast_button", "Create West Coast")
                                                                                               )
                                                                                       )
                                                                                       
                                                                             )
                                                                      )
                                                               ),
                                                               column(4,
                                                                      column(9, offset = 1,
                                                                             wellPanel(style = "height:675px;",
                                                                                       h3("Estes Inbound List mass email"),
                                                                                       br(),
                                                                                       fileInput("delivery_amazon_estes_raw_data_input", "Inputs", accept = ".xlsx"),
                                                                                       fluidRow(
                                                                                               column(9,
                                                                                                      textInput("delivery_amazon_estes_subject", "Subject", value = amazon_email_item("Estes", "subject"))
                                                                                               ),
                                                                                               column(3, style = "margin-top: 25px;",
                                                                                                      actionButton("delivery_amazon_estes_subject_save", "Save")
                                                                                               )
                                                                                       ),
                                                                                       br(),
                                                                                       fluidRow(
                                                                                               column(9,
                                                                                                      textAreaInput("delivery_amazon_estes_body", "Body", value = amazon_email_item("Estes", "body"), resize = "vertical", width = "100%", height = 150)
                                                                                               ),
                                                                                               column(3, style = "margin-top: 25px;",
                                                                                                      actionButton("delivery_amazon_estes_body_save", "Save")
                                                                                               )
                                                                                       ),
                                                                                       br(),
                                                                                       actionButton("delivery_amazon_estes_send_button", "Send Emails", width = 200)
                                                                             )
                                                                      )
                                                               )
                                                               
                                                       )
                                              )     
                                          
                                  )
                           )
                   )
           ),
           
           
           # Transportations ---------------------------------------------------------
           
           
           tabPanel("Transportations",
                    fluidRow(
                            column(2, 
                                   wellPanel(style = "height:450px;", align = "center",
                                             h2("Step 1"),
                                             helpText("Upload the input files"),
                                             br(),
                                             fileInput("button_Transportation_TM_Input", "TruckMate Report", accept = ".csv", buttonLabel = "Browse..."),
                                             fileInput("button_Transportation_TT_Input", "Track & Trace Report", accept = ".xlsx", buttonLabel = "Browse..."),
                                             dateInput("inbound_report_date", "Date", today() + 1, width = "66%")
                                   ),
                                   wellPanel(style = "height:250px;", align = "center",
                                             h2("Step 2"),
                                             helpText("Send Emails"),
                                             br(),
                                             actionButton("send_inbound_report_emails", "Send Inbound Report E-mails", width = "100%", height = "200px")
                                   )
                                   
                            )
                    )
           ),
           
           # Test_Active -------------------------------------------------------------
           
           
           tabPanel(
                   switchInput(
                           inputId = "Test_Active",
                           onStatus = "success", 
                           offStatus = "danger",
                           onLabel = "Test",
                           offLabel = "Active",
                           size = "mini",
                           value = TRUE
                   )
           ),
           footer = 
                   div(height = "700px",
                       style = "background-color:#e95420;",
                       align = "center",
                       h3(style = "color:white;", "NSDcomp 3")
                   )
           
           
           # tabPanel(
           #         tags$button(style = "background-color:red;",
           #                     id = 'close',
           #                     type = "button",
           #                     class = "btn action-button",
           #                     onclick = "setTimeout(function(){window.close();},500);",  # close browser
           #                     "Exit App"
           #         )
           # )
           
           
           
           # tabPanel("test",
           #          textOutput("test_output")
           #          )
)
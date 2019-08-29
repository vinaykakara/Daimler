library(shiny)

library(shinydashboard)

library(ggplot2)

library(plotly)

library(lubridate)

library(shinythemes)

library(reshape)

library(dplyr)

library(tidyr)

library(xts)

library(xlsx)

library(plyr)


library(dplyr)

library(DT)

library(rhandsontable)
library(shinyalert)
library(shinyBS)


#User Interface Start from here


shinyUI(
  
  dashboardPage(skin="blue",
                dashboardHeader(title="BLUE SKY KPI",
                                dropdownMenu(type="notifications", badgeStatus = "warning",
                                             
                                             notificationItem(icon = icon("warning"), status = "info",
                                                              
                                                              "Batch Transfermation Happened")
                                             
                                ),
                                
                                #Static Messages
                                
                                dropdownMenu(type = "messages", badgeStatus = "info",
                                             
                                             messageItem(from = "Measuring Department", message = "Master detected", icon=icon("bookmark"))
                                             
                                ),
                                
                                #Static Tasks
                                
                                dropdownMenu(type="tasks", badgeStatus = "info",
                                             
                                             taskItem(value = 30, color = "red", "Amount of R" )
                                             
                                )),
                dashboardSidebar(
                  sidebarMenu(
                    menuItem("KPI",tabName = "kpi",icon=icon("dashboard")),
                    menuItem("Safety",tabName = "safety",icon=icon("dashboard"),
                             menuSubItem("Major Accidents",tabName = "major"),
                             menuSubItem("Minor Accidents",tabName = "minor"),
                             menuSubItem("First Aid",tabName = "firstaid"),
                             menuSubItem("Accidents Counter Measure",tabName = "counter_measure")),
                    menuItem("Data",tabName = "data",icon=icon("dashboard"),
                             menuSubItem("Safety",tabName = "safety_data"),
                             menuSubItem("Progress",tabName = "data_completed"))
                  )
                ),
                dashboardBody(
                  tabItems(
                    tabItem(tabName = "kpi",
                            fluidRow(  
                      titlePanel("Safety"),
                      infoBoxOutput("ibox1"),
                      infoBoxOutput("ibox2"),
                      infoBoxOutput("ibox3"),
                      infoBoxOutput("ibox4")
                     
                      
                    )
                            ),
                    tabItem(tabName = "major",
                            tabsetPanel(
                              tabPanel("Major Accidents",
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=8,(plotlyOutput("plot_major",height = "600px", width = "800px"))),
                                         column(width=4,selectInput("choose_major","Select Month",selected = 'Jul',choices = c('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')),dataTableOutput("table_major")),
                                         column(width=10,selectInput("comment_choose_major","Select Month",selected = 'Jul',choices = c('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')),selectInput("choose_dept_major","Select Department",selected = 'Jul',choices = c('Chasis','Engine','PTI','FM')),textInput("description_major","Description"),actionButton("save_comment_major","Save Comment"))
                                       )
                                       )
                            )
                            ),
                    tabItem(tabName = "minor",
                            tabsetPanel(
                              tabPanel("Minor Accidents",
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=8,(plotlyOutput("plot_minor",height = "600px", width = "800px"))),
                                         column(width=4,selectInput("choose_minor","Select Month",selected = 'Jul',choices = c('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')),dataTableOutput("table_minor")),
                                         column(width=10,selectInput("comment_choose_minor","Select Month",selected = 'Jul',choices = c('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')),selectInput("choose_dept_minor","Select Department",selected = 'Jul',choices = c('Chasis','Engine','PTI','FM')),textInput("description_minor","Description"),actionButton("save_comment_minor","Save Comment"))
                                       )
                              )
                            )
                    ),
                    
                    tabItem(tabName = "firstaid",
                            tabsetPanel(
                              tabPanel("First Aid",
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=8,(plotlyOutput("plot_firstaid",height = "600px", width = "800px"))),
                                         column(width=4,selectInput("choose_firstaid","Select Month",selected = 'Jul',choices = c('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')),dataTableOutput("table_firstaid")),
                                         column(width=10,selectInput("comment_choose_firstaid","Select Month",selected = 'Jul',choices = c('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')),selectInput("choose_dept_firstaid","Select Department",selected = 'Jul',choices = c('Chasis','Engine','PTI','FM')),textInput("description_firstaid","Description"),actionButton("save_comment_firstaid","Save Comment"))
                                       )
                              ),
                              tabPanel("Deviation trend",
                                       selectInput("choose_deviation_firstaid","Select Department",selected = "All",choices = c("All","Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","IPL","Frame","TOS/others")),
                                       plotlyOutput("deviation_firstaid"),
                                       dataTableOutput("table_deviation_firstaid")
                                       
                              ),
                              tabPanel("Comparision",
                                       column(width=12,plotlyOutput("comp_firstaid")),
                                       selectInput("choose_comp_firstaid","Select Month",selected = 'Jul',choices = c('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')),
                                       plotlyOutput("comp_pie_firstaid")
                                       
                              ),
                              tabPanel("Department",
                                       selectInput("choose_indiv_firstaid","Select Department",selected = "Chassis",choices = c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","IPL","Frame","TOS/others")),
                                       plotlyOutput("dept_firstaid"),
                                       dataTableOutput("table_dept_firstaid")
                              )
                            )
                    ),
                    
                    tabItem(tabName = "counter_measure",
                            tabsetPanel(
                              tabPanel("Accidents Counter Measure",
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=8,(plotlyOutput("plot_counter",height = "600px", width = "800px"))),
                                         column(width=4,selectInput("choose_counter","Select Month",selected = 'Jul',choices = c('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')),dataTableOutput("table_counter")),
                                         column(width=10,selectInput("comment_choose_counter","Select Month",selected = 'Jul',choices = c('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')),selectInput("choose_dept_counter","Select Department",selected = 'Jul',choices = c('Chasis','Engine','PTI','FM')),textInput("description_counter","Description"),actionButton("save_comment_countor","Save Comment"))
                                       )
                              ),
                              tabPanel("Deviation trend",
                                       selectInput("choose_deviation_counter","Select Department",selected = "All",choices = c("All","Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","IPL","Frame","TOS/others")),
                                       plotlyOutput("deviation_counter"),
                                       dataTableOutput("table_deviation_counter")
                                
                              ),
                              tabPanel("Comparision",
                                       column(width=12,plotlyOutput("comp_counter")),
                                       selectInput("choose_comp_counter","Select Month",selected = 'Jul',choices = c('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')),
                                       plotlyOutput("comp_pie_counter")
                                
                              ),
                              tabPanel("Department",
                                       selectInput("choose_indiv_counter","Select Department",selected = "Chassis",choices = c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","IPL","Frame","TOS/others")),
                                       plotlyOutput("dept_counter"),
                                       dataTableOutput("table_dept_counter")
                                       )
                            )
                    ),
                    tabItem(tabName = "safety_data",
                            
                            tabsetPanel(
                              tabPanel("Major Accidents",
                                       useShinyalert(),
                                       actionButton("save_safety1","save"),
                                       rHandsontableOutput("hotable1")),
                              tabPanel("Minor Accidents",
                                       useShinyalert(),
                                       actionButton("save_safety2","save"),
                                       rHandsontableOutput("hotable2")),
                              tabPanel("First Aid",
                                       useShinyalert(),
                                       actionButton("save_safety3","save"),
                                       rHandsontableOutput("hotable3")),
                              tabPanel("Accident Counter Measure",
                                       useShinyalert(),
                                       actionButton("save_safety4","save"),
                                       rHandsontableOutput("hotable4"))
                            
                            
                            )),
                    tabItem(tabName = "data_completed",
                            titlePanel("Data"),
                            infoBoxOutput("ibox20"),
                            infoBoxOutput("ibox21"),
                            infoBoxOutput("ibox22"),
                            infoBoxOutput("ibox23"),
                            textOutput("text_safety")
                            )
                  
                )
    
  )
  
))

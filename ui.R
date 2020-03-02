library(shiny)
library(shinydashboard)
library(ggplot2)
library(plotly)
library(lubridate)
library(shinythemes)
library(reshape)
library(tidyr)
library(xts)
library(xlsx)
library(plyr)
library(dplyr)
library(DT)
library(rhandsontable)
library(shinyalert)
library(shinyBS)
library(slickR)
library(formattable)
library(shinymaterial)
#User Interface Start from here


shinyUI(
  
  dashboardPage(skin="blue",
                
                dashboardHeader(title="BLUE SKY KPI",
                                dropdownMenuOutput("menu")
                                ),
                dashboardSidebar(width = 200, div(style="overflow-y: scroll"),
                  sidebarMenu(
                    menuItem("KPI",tabName = "kpi",icon=icon("dashboard")),
                    menuItem("Safety",tabName = "safety",icon=icon("walking"),
                             menuSubItem("Major Accidents",tabName = "major"),
                             menuSubItem("Minor Accidents",tabName = "minor"),
                             menuSubItem("First Aid",tabName = "firstaid"),
                             menuSubItem("Accidents Counter Measure",tabName = "counter_measure"),
                             menuSubItem("Unsafe Acts",tabName = "unsafe_acts")),
                    menuItem("Quality",tabName = "quality",icon = icon("check-circle"),
                             menuSubItem("DPU @ QFL4(Overall)",tabName = "dpuqfl4"),
                             menuSubItem("HDT DPU @ QFL4(Ops)",tabName = "dpuqfl4_ops_hdt"),
                             menuSubItem("MDT DPU @ QFL4(Ops)",tabName = "dpuqfl4_ops_mdt"),
                             menuSubItem("Teardown Audit",tabName = "tear"),
                             menuSubItem("DPU @ QFL2",tabName = "dpuqfl2"),
                             menuSubItem("FTT",tabName="ftt"),
                             menuSubItem("SPR",tabName="spr")
                             ),
                    menuItem("Delivery",tabName = "delivery",icon=icon("bus"),
                             menuSubItem("QC OK",tabName = "qc_ok"),
                             menuSubItem("Roll Out",tabName = "roll_out"),
                             menuSubItem("Capacity Utilization",tabName = "cap_uti"),
                             menuSubItem("Non Forecasted Shortages",tabName = "non_for"),
                             menuSubItem("Vehicle loss",tabName = "veh_loss")
                             ),
                    menuItem("Cost",tabName = "cost",icon=icon("rupee-sign"),
                             menuSubItem("HPU per capacity",tabName = "hpu_capacity"),
                             menuSubItem("Indirect Consumables",tabName = "indirect_cons"),
                             menuSubItem("Rejection cost/truck",tabName = "rej_cost"),
                             menuSubItem("Electricity/Propane",tabName = "ele_pro")
                             ),
                    menuItem("Morale",tabName = "morale",icon = icon("award"),
                             menuSubItem("White Collar",tabName = "white_collar"),
                             menuSubItem("Kaizen per BCA",tabName = "bca_participation"),
                             menuSubItem("CA/BA participation",tabName = "caba_participation"),
                             menuSubItem("Attrition rate of Managers + Engineers",tabName = "man_attrition"),
                             menuSubItem("Attrition rate of BCA/BCAT/CA",tabName = "bca_attrition"),
                             menuSubItem("Attrition rate of Contractors",tabName = "con_attrition")
                             ),
                    menuItem("Data",tabName = "data",icon=icon("database"),
                             menuSubItem("Safety",tabName = "safety_data"),
                             menuSubItem("QM",tabName = "qm_data"),
                             menuSubItem("Chassis",tabName = "chassis_data"),
                             menuSubItem("CabTrim",tabName = "cabtrim_data"),
                             menuSubItem("EOL",tabName = "eol_data"),
                             menuSubItem("FBV",tabName = "fbv_data"),
                             menuSubItem("CIW",tabName = "ciw_data"),
                             menuSubItem("Paint",tabName = "paint_data"),
                             menuSubItem("Engine",tabName = "engine_data"),
                             menuSubItem("Transmission",tabName = "transmission_data"),
                             menuSubItem("Frame",tabName = "frame_data"),
                             menuSubItem("IPL",tabName = "ipl_data"),
                             menuSubItem("FM",tabName = "fm_data"),
                             menuSubItem("HPU per capacity",tabName = "hpu_data"),
                             menuSubItem("White Collar Attrition",tabName = "wc_data"),
                             menuSubItem("HPU",tabName = "hours_data"),
                             menuSubItem("FTT/SPR",tabName = "fttspr_data"),
                             menuSubItem("Progress",tabName = "data_completed")
                             )
                  )
                ),
                dashboardBody(
                  tags$head( tags$script(type="text/javascript",'$(document).ready(function(){
                             $(".main-sidebar").css("height","100%");
                             $(".main-sidebar .sidebar").css({"position":"relative","max-height": "100%","overflow": "auto"})
                             })')),
                  tabItems(
                    tabItem(tabName = "kpi",
                            fluidRow( 
                              actionButton("preview", "Generate Report"),
                              #bsModal("modalExample", "Blue SKY KPI report", "preview", size = "small",downloadButton("downloadReport", "Download Report")),
                      titlePanel("Safety"),
                      infoBoxOutput("ibox1"),
                      infoBoxOutput("ibox2"),
                      infoBoxOutput("ibox3"),
                      infoBoxOutput("ibox4"),
                      infoBoxOutput("ibox5")
                     
                      
                    )
                            ),
                    tabItem(tabName = "major",
                            tabsetPanel(
                              tabPanel("Major Accidents",
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=4,selectInput("choose_plot_major","Select Department",selected = "Plant level",choices = c("Plant level","Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS/others"))),
                                         column(width=4,selectInput("choose_plot_year_major","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=8,plotlyOutput("plot_major",height = "400px", width = "800px"),tableOutput("table_plot_major")),
                                         column(width=4,textAreaInput("enab_major",label = "Enablers",width=400,height=150)),
                                         column(width=4,textAreaInput("task_major",label = "Key Tasks",width=400,height=150),actionButton("save_comm_major","Save"))
                                       )
                                       ),
                              
                              tabPanel("Shop-wise",
                                       selectInput("choose_comp_major","Select Month",selected =months(Sys.Date()-30) ,choices = c('January','February','March','April','May','June','July','August','September','October','November','December')),
                                       column(width=8,plotlyOutput("comp_major"),tableOutput("table_comp_major")),
                                       column(width=4,textAreaInput("enab_major_shop",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_major_shop",label = "Key Tasks",width=400,height=150),actionButton("save_comm_major_shop","Save"))
                                       
                                       
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_major", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_major","Delete One Pagers"),
                                         slickROutput("slickr_major", width="1100px",height = "500px")
                                       
                                       )
                            )
                            ),
                    tabItem(tabName = "minor",
                            tabsetPanel(
                              tabPanel("Minor Accidents",
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=4,selectInput("choose_plot_minor","Select Department",selected = "Plant level",choices = c("Plant level","Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS/others"))),
                                         column(width=4,selectInput("choose_plot_year_minor","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=8,plotlyOutput("plot_minor",height = "400px", width = "800px"),tableOutput("table_plot_minor")),
                                         column(width=4,textAreaInput("enab_minor",label = "Enablers",width=400,height=150)),
                                         column(width=4,textAreaInput("task_minor",label = "Key Tasks",width=400,height=150),actionButton("save_comm_minor","Save"))
                                         
                                       )
                              ),
                              tabPanel("Shop-wise",
                                       selectInput("choose_comp_minor","Select Month",selected =months(Sys.Date()-30) ,choices = c('January','February','March','April','May','June','July','August','September','October','November','December')),
                                       column(width=8,plotlyOutput("comp_minor"),tableOutput("table_comp_minor")),
                                       column(width=4,textAreaInput("enab_minor_shop",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_minor_shop",label = "Key Tasks",width=400,height=150),actionButton("save_comm_minor_shop","Save"))
                                       
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_minor", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_minor","Delete One Pagers"),
                                       slickROutput("slickr_minor", width="1100px",height = "500px")
                              )
                            )
                    ),
                    
                    tabItem(tabName = "firstaid",
                            tabsetPanel(
                              tabPanel("First Aid",
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=4,selectInput("choose_plot_firstaid","Select Department",selected = "Plant level",choices = c("Plant level","Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS/others"))),
                                         column(width=4,selectInput("choose_plot_year_firstaid","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=8,plotlyOutput("plot_firstaid",height = "400px", width = "800px"),tableOutput("table_plot_firstaid")),
                                         column(width=4,textAreaInput("enab_firstaid",label = "Enablers",width=400,height=150)),
                                         column(width=4,textAreaInput("task_firstaid",label = "Key Tasks",width=400,height=150),actionButton("save_comm_firstaid","Save"))
                                         
                                         
                                         )
                              ),
                              tabPanel("Shop-wise",
                                       selectInput("choose_comp_firstaid","Select Month",selected =months(Sys.Date()-30) ,choices = c('January','February','March','April','May','June','July','August','September','October','November','December')),
                                       column(width=8,plotlyOutput("comp_firstaid"),tableOutput("table_comp_firstaid")),
                                       column(width=4,textAreaInput("enab_firstaid_shop",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_firstaid_shop",label = "Key Tasks",width=400,height=150),actionButton("save_comm_firstaid_shop","Save"))
                                       
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_firstaid", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_firstaid","Delete One Pagers"),
                                       slickROutput("slickr_firstaid", width="1100px",height = "500px")
                              )
                            )
                    ),
                    
                    tabItem(tabName = "counter_measure",
                            tabsetPanel(
                              tabPanel("Accidents Counter Measure",
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=4,selectInput("choose_plot_counter","Select Department",selected = "Plant level",choices = c("Plant level","Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS/others"))),
                                         column(width=4,selectInput("choose_plot_year_counter","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=8,plotlyOutput("plot_counter",height = "400px", width = "800px"),tableOutput("table_plot_counter")),
                                         column(width=4,textAreaInput("enab_counter",label = "Enablers",width=400,height=150)),
                                         column(width=4,textAreaInput("task_counter",label = "Key Tasks",width=400,height=150),actionButton("save_comm_counter","Save"))
                                         
                                         
                                          )
                              ),
                              
                              tabPanel("Shop-wise",
                                       selectInput("choose_comp_counter","Select Month",selected =months(Sys.Date()-30) ,choices = c('January','February','March','April','May','June','July','August','September','October','November','December')),
                                       column(width=8,plotlyOutput("comp_counter"),tableOutput("table_comp_counter")),
                                       column(width=4,textAreaInput("enab_counter_shop",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_counter_shop",label = "Key Tasks",width=400,height=150),actionButton("save_comm_counter_shop","Save"))
                                       
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_counter", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_counter","Delete One Pagers"),
                                       slickROutput("slickr_counter", width="1100px",height = "500px")
                              )
                            )
                    ),
                    tabItem(tabName = "unsafe_acts",
                            tabsetPanel(
                              tabPanel("Unsafe Acts",
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=4,selectInput("choose_plot_unsafe","Select Department",selected = "Plant level",choices = c("Plant level","Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS/others"))),
                                         column(width=4,selectInput("choose_plot_year_unsafe","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=8,plotlyOutput("plot_unsafe",height = "400px", width = "800px"),tableOutput("table_plot_unsafe")),
                                         column(width=4,textAreaInput("enab_unsafe",label = "Enablers",width=400,height=150)),
                                         column(width=4,textAreaInput("task_unsafe",label = "Key Tasks",width=400,height=150),actionButton("save_comm_unsafe","Save"))
                                         
                                         )
                              ),
                              
                              tabPanel("Shop-wise",
                                       selectInput("choose_comp_unsafe","Select Month",selected =months(Sys.Date()-30) ,choices = c('January','February','March','April','May','June','July','August','September','October','November','December')),
                                       column(width=8,plotlyOutput("comp_unsafe"),tableOutput("table_comp_unsafe")),
                                       column(width=4,textAreaInput("enab_unsafe_shop",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_unsafe_shop",label = "Key Tasks",width=400,height=150),actionButton("save_comm_unsafe_shop","Save"))
                                       
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_unsafe", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_unsafe","Delete One Pagers"),
                                       slickROutput("slickr_unsafe", width="1100px",height = "500px")
                              )
                            )
                    ),
                    tabItem(tabName = "dpuqfl4",
                            tabsetPanel(
                              tabPanel("HDT",
                                       useShinyalert(),
                                       fluidRow(
                                           useShinyalert(),
                                         column(width=9,selectInput("choose_plot_year_qm_hdt_dpu_qfl4","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=6,plotlyOutput("plot_qm_hdt_dpu_qfl4",height = "400px", width = "625px")),
                                         column(width=6,plotlyOutput("plot_qm_hdt_dpu_qfl4_ab",height = "400px", width = "625px")),
                                         column(width=12,tableOutput("table_plot_qm_hdt_dpu_qfl4"))
                                         
                                       )
                                        ),
                              tabPanel("MDT",
                                       useShinyalert(),
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=9,selectInput("choose_plot_year_qm_mdt_dpu_qfl4","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=6,plotlyOutput("plot_qm_mdt_dpu_qfl4",height = "400px", width = "625px")),
                                         column(width=6,plotlyOutput("plot_qm_mdt_dpu_qfl4_ab",height = "400px", width = "625px")),
                                         column(width=12,tableOutput("table_plot_qm_mdt_dpu_qfl4"))
                                       )
                                       ),
                              tabPanel("One pager",
                                       fileInput("myFile_qua_qfl4", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_qua_qfl4","Delete One Pagers"),
                                       slickROutput("slickr_qua_qfl4", width="1100px",height = "500px")
                              )
                              )),
                    tabItem(tabName = "tear",
                            tabsetPanel(
                                       tabPanel("Engine",
                                                useShinyalert(),
                                                fluidRow(
                                                  useShinyalert(),
                                                  column(width=9,selectInput("choose_plot_year_qm_qfl4_eng","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                                  column(width=6,plotlyOutput("plot_qm_qfl4_eng",height = "400px", width = "625px")),
                                                  column(width=6,plotlyOutput("plot_qm_qfl4_eng_ab",height = "400px", width = "625px")),
                                                  column(width=12,tableOutput("table_plot_qm_qfl4_eng"))
                                                )
                                             ),
                              tabPanel("Transmission",
                                       useShinyalert(),
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=9,selectInput("choose_plot_year_qm_qfl4_tra","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=6,plotlyOutput("plot_qm_qfl4_tra",height = "400px", width = "625px")),
                                         column(width=6,plotlyOutput("plot_qm_qfl4_tra_ab",height = "400px", width = "625px")),
                                         column(width=12,tableOutput("table_plot_qm_qfl4_tra"))
                                       )
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_qua_tear", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_qua_tear","Delete One Pagers"),
                                       slickROutput("slickr_qua_tear", width="1100px",height = "500px")
                              )
                              )),
                    tabItem(tabName = "dpuqfl4_ops_hdt",
                            tabsetPanel(
                              tabPanel("HDT (Overall)-Ops",
                                       useShinyalert(),
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=8,selectInput("choose_plot_year_qm_qfl4_hdt_ops","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=8,plotlyOutput("plot_qm_qfl4_hdt_ops"),tableOutput("table_plot_qm_qfl4_hdt_ops")),
                                         column(width=4,textAreaInput("enab_dpu_hdt_ops",label = "Enablers",width=400,height=150)),
                                         column(width=4,textAreaInput("task_dpu_hdt_ops",label = "Key Tasks",width=400,height=150),actionButton("save_comm_dpu_hdt_ops","Save"))
                                         
                                       )
                              ),
                              tabPanel("HDT (A+B)-Ops",
                                       useShinyalert(),
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=8,selectInput("choose_plot_year_qm_qfl4_hdt_ops_ab","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=8,plotlyOutput("plot_qm_qfl4_hdt_ops_ab"),tableOutput("table_plot_qm_qfl4_hdt_ops_ab")),
                                         column(width=4,textAreaInput("enab_dpu_hdt_ops_ab",label = "Enablers",width=400,height=150)),
                                         column(width=4,textAreaInput("task_dpu_hdt_ops_ab",label = "Key Tasks",width=400,height=150),actionButton("save_comm_dpu_hdt_ops_ab","Save"))
                                         
                                       )
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_qua_ops_hdt", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_qua_ops_hdt","Delete One Pagers"),
                                       slickROutput("slickr_qua_ops_hdt", width="1100px",height = "500px")
                              )
                              )),
                    tabItem(tabName = "dpuqfl4_ops_mdt",
                            tabsetPanel(
                              tabPanel("MDT (Overall)-Ops",
                                       useShinyalert(),
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=8,selectInput("choose_plot_year_qm_qfl4_mdt_ops","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=8,plotlyOutput("plot_qm_qfl4_mdt_ops"),tableOutput("table_plot_qm_qfl4_mdt_ops")),
                                         column(width=4,textAreaInput("enab_dpu_mdt_ops",label = "Enablers",width=400,height=150)),
                                         column(width=4,textAreaInput("task_dpu_mdt_ops",label = "Key Tasks",width=400,height=150),actionButton("save_comm_dpu_mdt_ops","Save"))
                                         
                                       )
                              ),
                              
                              tabPanel("MDT (A+B)-Ops",
                                       useShinyalert(),
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=8,selectInput("choose_plot_year_qm_qfl4_mdt_ops_ab","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=8,plotlyOutput("plot_qm_qfl4_mdt_ops_ab"),tableOutput("table_plot_qm_qfl4_mdt_ops_ab")),
                                         column(width=4,textAreaInput("enab_dpu_mdt_ops_ab",label = "Enablers",width=400,height=150)),
                                         column(width=4,textAreaInput("task_dpu_mdt_ops_ab",label = "Key Tasks",width=400,height=150),actionButton("save_comm_dpu_mdt_ops_ab","Save"))
                                         
                                       )
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_qua_ops_mdt", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_qua_ops_mdt","Delete One Pagers"),
                                       slickROutput("slickr_qua_ops_mdt", width="1100px",height = "500px")
                              )
                            )
                            ),
                    tabItem(tabName = "dpuqfl2",
                            tabsetPanel(
                    
                              tabPanel("Vehicle DPU @QFL2 HDT",
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=8,selectInput("choose_plot_year_qm_qfl2_hdt","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=8,plotlyOutput("plot_qm_qfl2_hdt"),tableOutput("table_plot_qm_qfl2_hdt")),
                                         column(width=4,textAreaInput("enab_dpu_hdt_qfl2",label = "Enablers",width=400,height=150)),
                                         column(width=4,textAreaInput("task_dpu_hdt_qfl2",label = "Key Tasks",width=400,height=150),actionButton("save_comm_dpu_hdt_qfl2","Save"))
                                         
                                       )
                                       ),
                              tabPanel("Vehicle DPU @QFL2 MDT",
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=8,selectInput("choose_plot_year_qm_qfl2_mdt","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=8,plotlyOutput("plot_qm_qfl2_mdt"),tableOutput("table_plot_qm_qfl2_mdt")),
                                         column(width=4,textAreaInput("enab_dpu_mdt_qfl2",label = "Enablers",width=400,height=150)),
                                         column(width=4,textAreaInput("task_dpu_mdt_qfl2",label = "Key Tasks",width=400,height=150),actionButton("save_comm_dpu_mdt_qfl2","Save"))
                                         
                                       )
                              ),
                              tabPanel("Engine",
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=8,selectInput("choose_plot_year_qm_qfl2_eng","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=8,plotlyOutput("plot_qm_qfl2_eng"),tableOutput("table_plot_qm_qfl2_eng")),
                                         column(width=4,textAreaInput("enab_dpu_eng",label = "Enablers",width=400,height=150)),
                                         column(width=4,textAreaInput("task_dpu_eng",label = "Key Tasks",width=400,height=150),actionButton("save_comm_dpu_eng","Save"))
                                         
                                          )
                              ),
                              tabPanel("Transmission",
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=8,selectInput("choose_plot_year_qm_qfl2_tra","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=8,plotlyOutput("plot_qm_qfl2_tra"),tableOutput("table_plot_qm_qfl2_tra")),
                                         column(width=4,textAreaInput("enab_dpu_tra",label = "Enablers",width=400,height=150)),
                                         column(width=4,textAreaInput("task_dpu_tra",label = "Key Tasks",width=400,height=150),actionButton("save_comm_dpu_tra","Save"))
                                         
                                          )
                              ),
                              tabPanel("CiW",
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=8,selectInput("choose_plot_year_qm_qfl2_ciw","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=8,plotlyOutput("plot_qm_qfl2_ciw"),tableOutput("table_plot_qm_qfl2_ciw")),
                                         column(width=4,textAreaInput("enab_dpu_ciw",label = "Enablers",width=400,height=150)),
                                         column(width=4,textAreaInput("task_dpu_ciw",label = "Key Tasks",width=400,height=150),actionButton("save_comm_dpu_ciw","Save"))
                                         
                                          )
                              ),
                              tabPanel("Paint",
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=8,selectInput("choose_plot_year_qm_qfl2_pai","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=8,plotlyOutput("plot_qm_qfl2_pai"),tableOutput("table_plot_qm_qfl2_pai")),
                                         column(width=4,textAreaInput("enab_dpu_pai",label = "Enablers",width=400,height=150)),
                                         column(width=4,textAreaInput("task_dpu_pai",label = "Key Tasks",width=400,height=150),actionButton("save_comm_dpu_pai","Save"))
                                         
                                         )
                              ),
                              tabPanel("Frame",
                                       fluidRow(
                                         useShinyalert(),
                                         column(width=8,selectInput("choose_plot_year_qm_qfl2_fra","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                         column(width=8,plotlyOutput("plot_qm_qfl2_fra"),tableOutput("table_plot_qm_qfl2_fra")),
                                         column(width=4,textAreaInput("enab_dpu_fra",label = "Enablers",width=400,height=150)),
                                         column(width=4,textAreaInput("task_dpu_fra",label = "Key Tasks",width=400,height=150),actionButton("save_comm_dpu_fra","Save"))
                                         
                                         )
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_qua_qfl2", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_qua_qfl2","Delete One Pagers"),
                                       slickROutput("slickr_qua_qfl2", width="1100px",height = "500px")
                              )
                            )
                            ),
                    tabItem(tabName = "ftt",
                            tabsetPanel(
                              tabPanel("FTT HDT",
                                       useShinyalert(),
                                       column(width=8,selectInput("choose_plot_year_qm_ftt_hdt","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                       column(width=8,plotlyOutput("plot_qm_ftt_hdt"),tableOutput("table_plot_qm_ftt_hdt")),
                                       column(width=4,textAreaInput("enab_ftt_hdt",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_ftt_hdt",label = "Key Tasks",width=400,height=150),actionButton("save_comm_ftt_hdt","Save"))
                                       
                                       ),
                              tabPanel("FTT MDT",
                                       useShinyalert(),
                                       column(width=8,selectInput("choose_plot_year_qm_ftt_mdt","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                       column(width=8,plotlyOutput("plot_qm_ftt_mdt"),tableOutput("table_plot_qm_ftt_mdt")),
                                       column(width=4,textAreaInput("enab_ftt_mdt",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_ftt_mdt",label = "Key Tasks",width=400,height=150),actionButton("save_comm_ftt_mdt","Save"))
                                       
                              ),
                              tabPanel("FTT LDT",
                                       useShinyalert(),
                                       column(width=8,selectInput("choose_plot_year_qm_ftt_ldt","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                       column(width=8,plotlyOutput("plot_qm_ftt_ldt"),tableOutput("table_plot_qm_ftt_ldt")),
                                       column(width=4,textAreaInput("enab_ftt_ldt",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_ftt_ldt",label = "Key Tasks",width=400,height=150),actionButton("save_comm_ftt_ldt","Save"))
                                       
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_qua_ftt", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_qua_ftt","Delete One Pagers"),
                                       slickROutput("slickr_qua_ftt", width="1100px",height = "500px")
                              )
                              )),
                    tabItem(tabName = "spr",
                            tabsetPanel(
                              tabPanel("SPR HDT",
                                       useShinyalert(),
                                       column(width=8,selectInput("choose_plot_year_qm_spr_hdt","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                       column(width=8,plotlyOutput("plot_qm_spr_hdt"),tableOutput("table_plot_qm_spr_hdt")),
                                       column(width=4,textAreaInput("enab_spr_hdt",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_spr_hdt",label = "Key Tasks",width=400,height=150),actionButton("save_comm_spr_hdt","Save"))
                                       
                              ),
                              tabPanel("SPR MDT",
                                       useShinyalert(),
                                       column(width=8,selectInput("choose_plot_year_qm_spr_mdt","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                       column(width=8,plotlyOutput("plot_qm_spr_mdt"),tableOutput("table_plot_qm_spr_mdt")),
                                       column(width=4,textAreaInput("enab_spr_mdt",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_spr_mdt",label = "Key Tasks",width=400,height=150),actionButton("save_comm_spr_mdt","Save"))
                                       
                              ),
                              tabPanel("SPR LDT",
                                       useShinyalert(),
                                       column(width=8,selectInput("choose_plot_year_qm_spr_ldt","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                       column(width=8,plotlyOutput("plot_qm_spr_ldt"),tableOutput("table_plot_qm_spr_ldt")),
                                       column(width=4,textAreaInput("enab_spr_ldt",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_spr_ldt",label = "Key Tasks",width=400,height=150),actionButton("save_comm_spr_ldt","Save"))
                                       
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_qua_spr", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_qua_spr","Delete One Pagers"),
                                       slickROutput("slickr_qua_spr", width="1100px",height = "500px")
                              )
                              
                            )
                            ),
                    tabItem(tabName = "qc_ok",
                      tabsetPanel(
                        tabPanel("QC Ok",
                                 useShinyalert(),
                                 column(width=9,selectInput("choose_plot_del_qcok_veh","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                 column(width=6,plotlyOutput("plot_del_qcok_veh")),
                                 column(width=6,plotlyOutput("plot_del_qcok_bus")),
                                 column(width=12,tableOutput("table_plot_del_qcok_veh"))
                        ),
                        tabPanel("CKD QC Ok",
                                 useShinyalert(),
                                 column(width=9,selectInput("choose_plot_del_qcok_ckd","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                 column(width=12,plotlyOutput("plot_del_qcok_ckd")),
                                 column(width=12,tableOutput("table_plot_del_qcok_ckd"))
                        ),
                        tabPanel("One pager",
                                 fileInput("myFile_del_qcok", "Choose a file", accept = c('image/png')),
                                 actionButton("delete_slickr_del_qcok","Delete One Pagers"),
                                 slickROutput("slickr_del_qcok", width="1100px",height = "500px")
                        )
                        )),
                    tabItem(tabName = "roll_out",
                            tabsetPanel(
                        tabPanel("HDT Roll Out",
                                 useShinyalert(),
                                 column(width=8,selectInput("choose_plot_del_rollout","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                 column(width=8,plotlyOutput("plot_del_rollout_hdt"),tableOutput("table_plot_del_rollout")),
                                 column(width=4,textAreaInput("enab_roll_hdt",label = "Enablers",width=400,height=150)),
                                 column(width=4,textAreaInput("task_roll_hdt",label = "Key Tasks",width=400,height=150),actionButton("save_comm_roll_hdt","Save"))
                                 
                        ),
                        tabPanel("MDT Roll Out",
                                 useShinyalert(),
                                 column(width=9,selectInput("choose_plot_del_rollout_mdt","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                 column(width=8,plotlyOutput("plot_del_rollout_mdt"),tableOutput("table_plot_del_rollout_mdt")),
                                 column(width=4,textAreaInput("enab_roll_mdt",label = "Enablers",width=400,height=150)),
                                 column(width=4,textAreaInput("task_roll_mdt",label = "Key Tasks",width=400,height=150),actionButton("save_comm_roll_mdt","Save"))
                                 
                        ),
                        tabPanel("One pager",
                                 fileInput("myFile_del_roll", "Choose a file", accept = c('image/png')),
                                 actionButton("delete_slickr_del_roll","Delete One Pagers"),
                                 slickROutput("slickr_del_roll", width="1100px",height = "500px")
                        )
                        )),
                    tabItem(tabName = "cap_uti",
                            tabsetPanel(
                        tabPanel("HDT Capacity Utilization",
                                 useShinyalert(),
                                 column(width=8,selectInput("choose_plot_del_capacity","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                 column(width=8,plotlyOutput("plot_del_capacity_hdt"),tableOutput("table_plot_del_capacity")),
                                 column(width=4,textAreaInput("enab_cap_hdt",label = "Enablers",width=400,height=150)),
                                 column(width=4,textAreaInput("task_cap_hdt",label = "Key Tasks",width=400,height=150),actionButton("save_comm_cap_hdt","Save"))
                                 
                        ),
                        tabPanel("MDT Capacity Utilization",
                                 useShinyalert(),
                                 column(width=8,selectInput("choose_plot_del_capacity_mdt","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                 column(width=8,plotlyOutput("plot_del_capacity_mdt"),tableOutput("table_plot_del_capacity_mdt")),
                                 column(width=4,textAreaInput("enab_cap_mdt",label = "Enablers",width=400,height=150)),
                                 column(width=4,textAreaInput("task_cap_mdt",label = "Key Tasks",width=400,height=150),actionButton("save_comm_cap_mdt","Save"))
                                 
                        ),
                        tabPanel("One pager",
                                 fileInput("myFile_del_cap", "Choose a file", accept = c('image/png')),
                                 actionButton("delete_slickr_del_cap","Delete One Pagers"),
                                 slickROutput("slickr_del_cap", width="1100px",height = "500px")
                        )
                        )),
                    tabItem(tabName = "non_for",
                            tabsetPanel(
                        tabPanel("Non Forecasted Shortages",
                                 useShinyalert(),
                                 column(width=8,selectInput("choose_plot_del_nonforecast","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                 column(width=8,plotlyOutput("plot_del_nonforecast"),tableOutput("table_plot_del_nonforecast")),
                                 column(width=4,textAreaInput("enab_forecasted",label = "Enablers",width=400,height=150)),
                                 column(width=4,textAreaInput("task_forecasted",label = "Key Tasks",width=400,height=150),actionButton("save_comm_forecasted","Save"))
                                 
                        ),
                        tabPanel("One pager",
                                 fileInput("myFile_del_fore", "Choose a file", accept = c('image/png')),
                                 actionButton("delete_slickr_del_fore","Delete One Pagers"),
                                 slickROutput("slickr_del_fore", width="1100px",height = "500px")
                        ))),
                    tabItem(tabName = "veh_loss",
                            tabsetPanel(
                        tabPanel("Vehicle loss(Operations)",
                                 useShinyalert(),
                                 column(width=8,selectInput("choose_plot_del_opp_loss","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                 column(width=8,plotlyOutput("plot_del_opp_loss"),tableOutput("table_plot_del_opp_loss")),
                                 column(width=4,textAreaInput("enab_loss_ope",label = "Enablers",width=400,height=150)),
                                 column(width=4,textAreaInput("task_loss_ope",label = "Key Tasks",width=400,height=150),actionButton("save_comm_loss_ope","Save"))
                                 
                        ),
                        tabPanel("Vehicle loss(Aggregates)",
                                 useShinyalert(),
                                 column(width=8,selectInput("choose_plot_del_agg_loss","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                 column(width=8,plotlyOutput("plot_del_agg_loss"),tableOutput("table_plot_del_agg_loss")),
                                 column(width=4,textAreaInput("enab_loss_agg",label = "Enablers",width=400,height=150)),
                                 column(width=4,textAreaInput("task_loss_agg",label = "Key Tasks",width=400,height=150),actionButton("save_comm_loss_agg","Save"))
                                 
                        ),
                        tabPanel("One pager",
                                 fileInput("myFile_del_loss", "Choose a file", accept = c('image/png')),
                                 actionButton("delete_slickr_del_loss","Delete One Pagers"),
                                 slickROutput("slickr_del_loss", width="1100px",height = "500px")
                        )
                        )),
                    tabItem(tabName = "hpu_capacity",
                            tabsetPanel(
                              tabPanel("Plant Level- HPU per capacity",
                                       useShinyalert(),
                                       column(width=8,selectInput("choose_plot_cos_hpu_capacity","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                       column(width=8,plotlyOutput("plot_cos_hpu_capacity"),tableOutput("table_plot_cos_hpu_capacity")),
                                       column(width=4,textAreaInput("enab_hpu_cap_plant",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_hpu_cap_plant",label = "Key Tasks",width=400,height=150),actionButton("save_comm_hpu_cap_plant","Save"))
                                       
                                       ),
                              tabPanel("Shop-wise",
                                       selectInput("choose_comp_hpu_capacity_shop","Select Month",selected =months(Sys.Date()-30) ,choices = c('January','February','March','April','May','June','July','August','September','October','November','December')),
                                       column(width=8,plotlyOutput("comp_hpu_capacity_shop"),tableOutput("table_comp_hpu_capacity_shop")),
                                       column(width=4,textAreaInput("enab_hpu_cap_shop",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_hpu_cap_shop",label = "Key Tasks",width=400,height=150),actionButton("save_comm_hpu_cap_shop","Save"))
                                       
                                       
                                       
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_cost_hpucapacity", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_cost_hpucapacity","Delete One Pagers"),
                                       slickROutput("slickr_cost_hpucapacity", width="1100px",height = "500px")
                              )
                            )
                            ),
                    tabItem(tabName = "indirect_cons",
                            tabsetPanel(
                              tabPanel("Plant Level- Indirect Consumables",
                                       useShinyalert(),
                                       column(width=8,selectInput("choose_plot_cos_indirect_cons","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                       column(width=8,plotlyOutput("plot_cos_indirect_cons"),tableOutput("table_plot_cos_indirect_cons")),
                                       column(width=4,textAreaInput("enab_cons_plant",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_cons_plant",label = "Key Tasks",width=400,height=150),actionButton("save_comm_cons_plant","Save"))
                                       
                              ),
                              tabPanel("Shop-wise",
                                       selectInput("choose_cos_indirect_cons_shop","Select Month",selected =months(Sys.Date()-30) ,choices = c('January','February','March','April','May','June','July','August','September','October','November','December')),
                                       column(width=8,plotlyOutput("comp_cos_indirect_cons_shop"),tableOutput("table_comp_cos_indirect_cons_shop")),
                                       column(width=4,textAreaInput("enab_cons_shop",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_cons_shop",label = "Key Tasks",width=400,height=150),actionButton("save_comm_cons_shop","Save"))
                                       
                                       
                                       
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_cost_indirect", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_cost_indirect","Delete One Pagers"),
                                       slickROutput("slickr_cost_indirect", width="1100px",height = "500px")
                              )
                            )
                    ),
                    tabItem(tabName = "rej_cost",
                            tabsetPanel(
                              tabPanel("Plant Level- rejection cost/truck",
                                       useShinyalert(),
                                       column(width=8,selectInput("choose_plot_cos_rej_cost","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                       column(width=8,plotlyOutput("plot_cos_rej_cost"),tableOutput("table_plot_cos_rej_cost")),
                                       column(width=4,textAreaInput("enab_rej_plant",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_rej_plant",label = "Key Tasks",width=400,height=150),actionButton("save_comm_rej_plant","Save"))
                                       
                              ),
                              tabPanel("Shop-wise",
                                       selectInput("choose_comp_rej_cost_shop","Select Month",selected =months(Sys.Date()-30) ,choices = c('January','February','March','April','May','June','July','August','September','October','November','December')),
                                       column(width=8,plotlyOutput("comp_rej_cost_shop"),tableOutput("table_comp_rej_cost_shop")),
                                       column(width=4,textAreaInput("enab_rej_shop",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_rej_shop",label = "Key Tasks",width=400,height=150),actionButton("save_comm_rej_shop","Save"))
                                       
                                       
                                       
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_cost_rej", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_cost_rej","Delete One Pagers"),
                                       slickROutput("slickr_cost_rej", width="1100px",height = "500px")
                              )
                            )
                    ),
                    tabItem(tabName = "ele_pro",
                            tabsetPanel(
                              tabPanel("Electricity Consumption",
                                       useShinyalert(),
                                       column(width=8,selectInput("choose_plot_cos_ele","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                       column(width=8,plotlyOutput("plot_cos_ele"),tableOutput("table_plot_cos_ele")),
                                       column(width=4,textAreaInput("enab_electricity",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_electricity",label = "Key Tasks",width=400,height=150),actionButton("save_comm_electricity","Save"))
                                       
                              ),
                              tabPanel("Propane consumption",
                                       useShinyalert(),
                                       column(width=8,selectInput("choose_plot_cos_pro","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                       column(width=8,plotlyOutput("plot_cos_pro"),tableOutput("table_plot_cos_pro")),
                                       column(width=4,textAreaInput("enab_propane",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_propane",label = "Key Tasks",width=400,height=150),actionButton("save_comm_propane","Save"))
                                       
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_cost_ele", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_cost_ele","Delete One Pagers"),
                                       slickROutput("slickr_cost_ele", width="1100px",height = "500px")
                              )
                            )
                    ),
                    tabItem(tabName = "white_collar",
                            tabsetPanel(
                              tabPanel("Plant Level- White collars attrition rate",
                                       useShinyalert(),
                                       column(width=9,selectInput("choose_plot_mor_white_collar","Select year",selected = "2020",choices = c("2015","2016","2017","2018","2019","2020"))),
                                       column(width=8,plotlyOutput("plot_mor_white_collar"),tableOutput("table_plot_mor_white_collar")),
                                       column(width=4,textAreaInput("enab_att_white_plant",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_att_white_plant",label = "Key Tasks",width=400,height=150),actionButton("save_comm_att_white_plant","Save"))
                                       
                              ),
                              tabPanel("Shop-wise",
                                       selectInput("choose_comp_white_collar_shop","Select Month",selected =months(Sys.Date()-30) ,choices = c('January','February','March','April','May','June','July','August','September','October','November','December')),
                                       column(width=8,plotlyOutput("comp_white_collar_shop"),tableOutput("table_comp_white_collar_shop")),
                                       column(width=4,textAreaInput("enab_att_white_shop",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_att_white_shop",label = "Key Tasks",width=400,height=150),actionButton("save_comm_att_white_shop","Save"))
                                       
                                       
                                       
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_morale_white", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_morale_white","Delete One Pagers"),
                                       slickROutput("slickr_morale_white", width="1100px",height = "500px")
                              )
                            )
                    ),
                    tabItem(tabName = "bca_participation",
                            tabsetPanel(
                              tabPanel("Kaizen per BCA per Year",
                                       useShinyalert(),
                                       column(width=8,plotlyOutput("plot_mor_bca_participation"),tableOutput("table_plot_mor_bca_participation")),
                                       column(width=4,textAreaInput("enab_att_bca_plant",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_att_bca_plant",label = "Key Tasks",width=400,height=150),actionButton("save_comm_att_bca_plant","Save"))
                                       
                              ),
                              tabPanel("Shop-wise",
                                       selectInput("choose_comp_bca_participation_shop","Select Month",selected =months(Sys.Date()-30) ,choices = c('January','February','March','April','May','June','July','August','September','October','November','December')),
                                       column(width=8,plotlyOutput("comp_bca_participation_shop"),tableOutput("table_comp_bca_participation_shop")),
                                       column(width=4,textAreaInput("enab_att_bca_shop",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_att_bca_shop",label = "Key Tasks",width=400,height=150),actionButton("save_comm_att_bca_shop","Save"))
                                       
                                       
                                       
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_morale_kai_bca", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_morale_kai_bca","Delete One Pagers"),
                                       slickROutput("slickr_morale_kai_bca", width="1100px",height = "500px")
                              )
                            )
                    ),
                    tabItem(tabName = "caba_participation",
                            tabsetPanel(
                              tabPanel("Participation in AOM - CA/BA",
                                       useShinyalert(),
                                       column(width=8,plotlyOutput("plot_mor_caba_participation"),tableOutput("table_plot_mor_caba_participation")),
                              column(width=4,textAreaInput("enab_att_baca_plant",label = "Enablers",width=400,height=150)),
                              column(width=4,textAreaInput("task_att_baca_plant",label = "Key Tasks",width=400,height=150),actionButton("save_comm_att_baca_plant","Save"))
                              
                              ),
                              tabPanel("Shop-wise",
                                       selectInput("choose_comp_caba_participation_shop","Select Month",selected =months(Sys.Date()-30) ,choices = c('January','February','March','April','May','June','July','August','September','October','November','December')),
                                       column(width=8,plotlyOutput("comp_caba_participation_shop"),tableOutput("table_comp_caba_participation_shop")),
                                       column(width=4,textAreaInput("enab_att_baca_shop",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_att_baca_shop",label = "Key Tasks",width=400,height=150),actionButton("save_comm_att_baca_shop","Save"))
                                       
                                       
                                       
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_morale_kai_ca", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_morale_kai_ca","Delete One Pagers"),
                                       slickROutput("slickr_morale_kai_ca", width="1100px",height = "500px")
                              )
                            )
                    ),
                    tabItem(tabName = "man_attrition",
                            tabsetPanel(
                              tabPanel("Attrition rate of Managers + Engineers",
                                       useShinyalert(),
                                       column(width=8,plotlyOutput("plot_mor_man_attrition"),tableOutput("table_plot_mor_man_attrition"))
                                       
                              ),
                              tabPanel("Shop-wise",
                                       selectInput("choose_comp_man_attrition_shop","Select Month",selected =months(Sys.Date()-30) ,choices = c('January','February','March','April','May','June','July','August','September','October','November','December')),
                                       column(width=8,plotlyOutput("comp_man_attrition_shop"),tableOutput("table_comp_man_attrition_shop"))
                                       
                                       
                                       
                              )
                            )
                    ),
                    tabItem(tabName = "bca_attrition",
                            tabsetPanel(
                              tabPanel("Attrition rate of BCA/BCAT/CA",
                                       useShinyalert(),
                                       column(width=8,plotlyOutput("plot_mor_bca_attrition"),tableOutput("table_plot_mor_bca_attrition"))
                                        
                              ),
                              tabPanel("Shop-wise",
                                       selectInput("choose_comp_bca_attrition_shop","Select Month",selected =months(Sys.Date()-30) ,choices = c('January','February','March','April','May','June','July','August','September','October','November','December')),
                                       column(width=8,plotlyOutput("comp_bca_attrition_shop"),tableOutput("table_comp_bca_attrition_shop"))
                                       
                                       
                                       
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_morale_att_bca", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_morale_att_bca","Delete One Pagers"),
                                       slickROutput("slickr_morale_att_bca", width="1100px",height = "500px")
                              )
                            )
                    ),
                    tabItem(tabName = "con_attrition",
                            tabsetPanel(
                              tabPanel("Attrition rate of Contractors",
                                       useShinyalert(),
                                       column(width=8,plotlyOutput("plot_mor_con_attrition"),tableOutput("table_plot_mor_con_attrition")),
                                       column(width=4,textAreaInput("enab_att_con_plant",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_att_con_plant",label = "Key Tasks",width=400,height=150),actionButton("save_comm_att_con_plant","Save"))
                                       
                              ),
                              tabPanel("Shop-wise",
                                       selectInput("choose_comp_con_attrition_shop","Select Month",selected =months(Sys.Date()-30) ,choices = c('January','February','March','April','May','June','July','August','September','October','November','December')),
                                       column(width=8,plotlyOutput("comp_con_attrition_shop"),tableOutput("table_comp_con_attrition_shop")),
                                       column(width=4,textAreaInput("enab_att_con_shop",label = "Enablers",width=400,height=150)),
                                       column(width=4,textAreaInput("task_att_con_shop",label = "Key Tasks",width=400,height=150),actionButton("save_comm_att_con_shop","Save"))
                                       
                                       
                                       
                              ),
                              tabPanel("One pager",
                                       fileInput("myFile_morale_att_con", "Choose a file", accept = c('image/png')),
                                       actionButton("delete_slickr_morale_att_con","Delete One Pagers"),
                                       slickROutput("slickr_morale_att_con", width="1100px",height = "500px")
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
                                       rHandsontableOutput("hotable4")),
                              tabPanel("Unsafe Acts",
                                       useShinyalert(),
                                       actionButton("save_safety5","save"),
                                       rHandsontableOutput("hotable5"))
                            
                            
                            )),
                    tabItem(tabName = "qm_data",
                            
                            tabsetPanel(
                              tabPanel("KPI",
                                       useShinyalert(),
                                       actionButton("save_qm_kpi","save"),
                                       rHandsontableOutput("hotable_qm_kpi")),
                              tabPanel("DPU @ QFL4",
                                       useShinyalert(),
                                       actionButton("save_qm_dpu_qfl4","save"),
                                       rHandsontableOutput("hotable_qm_dpu_qfl4")),
                              tabPanel("DPU @ QFL4 Ops related",
                                       useShinyalert(),
                                       actionButton("save_qm_dpu_qfl4_ops","save"),
                                       rHandsontableOutput("hotable_qm_dpu_qfl4_ops"))
                              
                              
                            )),
                      tabItem(tabName = "chassis_data",
                              tabsetPanel(
                                tabPanel("KPI",
                                         useShinyalert(),
                                         actionButton("save_chassis_kpi","save"),
                                         rHandsontableOutput("hotable_chassis_kpi")
                                ),
                                tabPanel("Roll Out",
                                         useShinyalert(),
                                         actionButton("save_chassis_rollout","save"),
                                         rHandsontableOutput("hotable_chassis_rollout")
                                ),
                                tabPanel("QC OK",
                                         useShinyalert(),
                                         actionButton("save_chassis_qcok","save"),
                                         rHandsontableOutput("hotable_chassis_qcok")
                                ),
                                tabPanel("Capacity Utilization",
                                         useShinyalert(),
                                         actionButton("save_chassis_capacity","save"),
                                         rHandsontableOutput("hotable_chassis_capacity")
                                )
                                
                              )
                        
                      ),
                    tabItem(tabName = "cabtrim_data",
                            
                                       useShinyalert(),
                                       actionButton("save_cabtrim_kpi","save"),
                                       rHandsontableOutput("hotable_cabtrim_kpi")
                              
                            
                    ),
                    tabItem(tabName = "eol_data",
                            
                                       useShinyalert(),
                                       actionButton("save_eol_kpi","save"),
                                       rHandsontableOutput("hotable_eol_kpi")
                              
                    ),
                    tabItem(tabName = "fbv_data",
                            
                                       useShinyalert(),
                                       actionButton("save_fbv_kpi","save"),
                                       rHandsontableOutput("hotable_fbv_kpi")
                             
                            
                    ),
                    tabItem(tabName = "ciw_data",
                            
                                       useShinyalert(),
                                       actionButton("save_ciw_kpi","save"),
                                       rHandsontableOutput("hotable_ciw_kpi")
                              
                            
                    ),
                    tabItem(tabName = "paint_data",
                            
                                       useShinyalert(),
                                       actionButton("save_paint_kpi","save"),
                                       rHandsontableOutput("hotable_paint_kpi")
                              
                    ),
                    tabItem(tabName = "engine_data",
                            
                                       useShinyalert(),
                                       actionButton("save_engine_kpi","save"),
                                       rHandsontableOutput("hotable_engine_kpi")
                             
                            
                    ),
                    tabItem(tabName = "transmission_data",
                            
                                       useShinyalert(),
                                       actionButton("save_transmission_kpi","save"),
                                       rHandsontableOutput("hotable_transmission_kpi")
                             
                            
                    ),
                    tabItem(tabName = "frame_data",
                            
                                       useShinyalert(),
                                       actionButton("save_frame_kpi","save"),
                                       rHandsontableOutput("hotable_frame_kpi")
                              
                            
                    ),
                    
                    tabItem(tabName = "ipl_data",
                                       useShinyalert(),
                                       actionButton("save_ipl_kpi","save"),
                                       rHandsontableOutput("hotable_ipl_kpi")
                              
                            
                    ),
                    tabItem(tabName = "fm_data",
                                       useShinyalert(),
                                       actionButton("save_fm_kpi","save"),
                                       rHandsontableOutput("hotable_fm_kpi")
                              
                            
                    ),
                    tabItem(tabName = "wc_data",
                            useShinyalert(),
                            actionButton("save_wc_kpi","save"),
                            rHandsontableOutput("hotable_wc_kpi")
                            
                            
                    ),
                    tabItem(tabName = "hpu_data",
                            tabsetPanel(
                              tabPanel("HPU per capacity",
                                       useShinyalert(),
                                       actionButton("save_hpu_capacity","save"),
                                       rHandsontableOutput("hotable_hpu_capacity")
                                       )
                            )
                            ),
                    
                    tabItem(tabName = "hours_data",
                            tabsetPanel(
                              tabPanel("CabTrim",
                                       useShinyalert(),
                                       actionButton("save_hpu_cabtrim","save"),
                                       rHandsontableOutput("hotable_hpu_cabtrim")
                              )
                            )
                    ),
                    tabItem(tabName = "fttspr_data",
                            tabsetPanel(
                              tabPanel("FTt/SPR",
                                       useShinyalert(),
                                       actionButton("save_fttspr","save"),
                                       rHandsontableOutput("hotable_fttspr")
                              )
                            )
                    ),
                    
                    tabItem(tabName = "data_completed",
                            titlePanel("Data"),
                            infoBoxOutput("ibox20"),
                            infoBoxOutput("ibox21"),
                            infoBoxOutput("ibox22"),
                            infoBoxOutput("ibox23"),
                            infoBoxOutput("ibox24"),
                            column(width=10,textOutput("text_major")),
                            column(width=10,textOutput("text_minor")),
                            column(width=10,textOutput("text_firstaid")),
                            column(width=10,textOutput("text_counter")),
                            column(width=10,textOutput("text_unsafe"))
                            
                            )
                  
                )
    
  )
  
))
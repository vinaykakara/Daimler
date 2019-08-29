library(readxl)
library(rhandsontable)
library(shiny)
library(plotly)
library(dplyr)

server <- function(input, output) {
  
  
 
  output$ibox1 <- renderInfoBox({    
    te<-input$save_safety1
    d<-read.xlsx("safety/major_accidents.xlsx",sheetIndex = 1)
    if(d$Jul[28]>d$Jul[27]){
      ic<-"thumbs-down"
      co="red"
    }
    else{
      ic<-"thumbs-up"
      co="green"
    }
    infoBox(
      value=va<-paste("T:",d$Jul[27]," A:",d$Jul[28],sep=""),
      title="Major Accidents",
      icon=icon(ic),
      color=co
    )
  })
  output$ibox2 <- renderInfoBox({    
    te<-input$save_safety2
    d<-read.xlsx("safety/minor_accidents.xlsx",sheetIndex = 1)
    if(d$Jul[28]>d$Jul[27]){
      ic<-"thumbs-down"
      co="red"
    }
    else{
      ic<-"thumbs-up"
      co="green"
    }
    infoBox(
      value=va<-paste("T:",d$Jul[27]," A:",d$Jul[28],sep=""),
      title="Minor Accidents",
      icon=icon(ic),
      color=co
    )
  })

  output$ibox3 <- renderInfoBox({    
    te<-input$save_safety3
    d<-read.xlsx("safety/first_aid.xlsx",sheetIndex = 1)
    if(d$Jul[28]>d$Jul[27]){
      ic<-"thumbs-down"
      co="red"
    }
    else{
      ic<-"thumbs-up"
      co="green"
    }
    infoBox(
      value=va<-paste("T:",d$Jul[27]," A:",d$Jul[28],sep=""),
      title="First Aid",
      icon=icon(ic),
      color=co
    )
  })
  output$ibox4 <- renderInfoBox({    
    te<-input$save_safety4
    d<-read.xlsx("safety/accidents_countermeasure.xlsx",sheetIndex = 1)
    if(d$Jul[28]<d$Jul[27]){
      ic<-"thumbs-down"
      co="red"
    }
    else{
      ic<-"thumbs-up"
      co="green"
    }
    infoBox(
      value=va<-paste("T:",d$Jul[27]," A:",d$Jul[28],sep=""),
      title="Accidents Counter Measure",
      icon=icon(ic),
      color=co
    )
  })
  
  
  output$table_counter<-renderDataTable({
    te<-input$save_comment_countor
    d<-read.xlsx("comments/counter.xlsx",sheetIndex = 1)
    d <-d[d$Month==input$choose_counter,]
    d
  })
  
  observeEvent(input$save_comment_countor,{
    d<-read.xlsx("comments/counter.xlsx",sheetIndex = 1)
    da<-data.frame(Month=input$comment_choose_counter,Department=input$choose_dept_counter,Description=input$description_counter)
    d<-rbind(d,da)
    write.xlsx(d,"comments/counter.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  output$table_firstaid<-renderDataTable({
    te<-input$save_comment_firstaid
    d<-read.xlsx("comments/firstaid.xlsx",sheetIndex = 1)
    d <-d[d$Month==input$choose_firstaid,]
    d
  })
  
  observeEvent(input$save_comment_firstaid,{
    d<-read.xlsx("comments/firstaid.xlsx",sheetIndex = 1)
    da<-data.frame(Month=input$comment_choose_firstaid,Department=input$choose_dept_firstaid,Description=input$description_firstaid)
    d<-rbind(d,da)
    write.xlsx(d,"comments/firstaid.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  output$table_minor<-renderDataTable({
    te<-input$save_comment_minor
    d<-read.xlsx("comments/minor.xlsx",sheetIndex = 1)
    d <-d[d$Month==input$choose_minor,]
    d
  })
  
  observeEvent(input$save_comment_minor,{
    d<-read.xlsx("comments/minor.xlsx",sheetIndex = 1)
    da<-data.frame(Month=input$comment_choose_minor,Department=input$choose_dept_minor,Description=input$description_minor)
    d<-rbind(d,da)
    write.xlsx(d,"comments/minor.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  output$table_major<-renderDataTable({
    te<-input$save_comment_major
    d<-read.xlsx("comments/major.xlsx",sheetIndex = 1)
    d <-d[d$Month==input$choose_major,]
    d
  })
  
  observeEvent(input$save_comment_major,{
    d<-read.xlsx("comments/major.xlsx",sheetIndex = 1)
    da<-data.frame(Month=input$comment_choose_major,Department=input$choose_dept_major,Description=input$description_major)
    d<-rbind(d,da)
    write.xlsx(d,"comments/major.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  values1 <- reactiveValues()
  
  
   previous1 <- reactive({
     read.xlsx("safety/major_accidents.xlsx",sheetIndex = 1)
   })
  
  MyChanges1 <- reactive({
    if(is.null(input$hotable1)){return(previous1())}
    else if(!identical(previous1(),input$hotable1)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable1 <- as.data.frame(hot_to_r(input$hotable1))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable1 <- mytable1[1:nrow(previous1()),]
      
      for(i in 3:14)
      mytable1[27,i]<-sum(mytable1[seq(1,26,2),i],na.rm=TRUE)
      
      for(i in 3:14)
        mytable1[28,i]<-sum(mytable1[seq(2,27,2),i],na.rm=TRUE)
      mytable1
    }
  })
  
  output$hotable1<-renderRHandsontable({
    col_highlight = 0:7
    row_highlight = c(27,26)
    
    rhandsontable(MyChanges1(), col_highlight = col_highlight, row_highlight = row_highlight)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:8,readOnly = TRUE)%>%
      hot_row(c(27,28), readOnly = TRUE) %>%
      hot_cols(renderer = "
            function(instance, td, row, col, prop, value, cellProperties) {
               Handsontable.renderers.NumericRenderer.apply(this, arguments);
                if (instance.params) {
                    hcols = instance.params.col_highlight
                    hcols = hcols instanceof Array ? hcols : [hcols]
                    hrows = instance.params.row_highlight
                    hrows = hrows instanceof Array ? hrows : [hrows]
                }
                if (instance.params && hcols.includes(col)) td.style.background = 'lightgrey';
                if (instance.params && hrows.includes(row)) td.style.background = 'lightgreen';
            }")
            
  })
  observeEvent(input$save_safety1,{
    #
    write.xlsx(hot_to_r(input$hotable1),"safety/major_accidents.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  
  values2 <- reactiveValues()
  
  
  previous2 <- reactive({
    read.xlsx("safety/minor_accidents.xlsx",sheetIndex = 1)
  })
  
  MyChanges2 <- reactive({
    if(is.null(input$hotable2)){return(previous2())}
    else if(!identical(previous2(),input$hotable2)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable2 <- as.data.frame(hot_to_r(input$hotable2))
      # here 2he second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable2 <- mytable2[1:nrow(previous2()),]
      
      for(i in 3:14)
        mytable2[27,i]<-sum(mytable2[seq(1,26,2),i],na.rm=TRUE)
      
      for(i in 3:14)
        mytable2[28,i]<-sum(mytable2[seq(2,27,2),i],na.rm=TRUE)
      mytable2
    }
  })
  
  output$hotable2<-renderRHandsontable({
    col_highlight = 0:7
    row_highlight = c(27,26)
    
    rhandsontable(MyChanges2(), col_highlight = col_highlight, row_highlight = row_highlight)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:8,readOnly = TRUE)%>%
      hot_row(c(27,28), readOnly = TRUE) %>%
      hot_cols(renderer = "
            function(instance, td, row, col, prop, value, cellProperties) {
               Handsontable.renderers.NumericRenderer.apply(this, arguments);
                if (instance.params) {
                    hcols = instance.params.col_highlight
                    hcols = hcols instanceof Array ? hcols : [hcols]
                    hrows = instance.params.row_highlight
                    hrows = hrows instanceof Array ? hrows : [hrows]
                }
                if (instance.params && hcols.includes(col)) td.style.background = 'lightgrey';
                if (instance.params && hrows.includes(row)) td.style.background = 'lightgreen';
            }")
    
  })
  observeEvent(input$save_safety2,{
    #
    write.xlsx(hot_to_r(input$hotable2),"safety/minor_accidents.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  
  values3 <- reactiveValues()
  
  
  previous3 <- reactive({
    read.xlsx("safety/first_aid.xlsx",sheetIndex = 1)
  })
  
  MyChanges3 <- reactive({
    if(is.null(input$hotable3)){return(previous3())}
    else if(!identical(previous3(),input$hotable3)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable3 <- as.data.frame(hot_to_r(input$hotable3))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable3 <- mytable3[1:nrow(previous3()),]
      
      for(i in 3:14)
        mytable3[27,i]<-sum(mytable3[seq(1,26,2),i],na.rm=TRUE)
      
      for(i in 3:14)
        mytable3[28,i]<-sum(mytable3[seq(2,27,2),i],na.rm=TRUE)
      mytable3
    }
  })
  
  output$hotable3<-renderRHandsontable({
    col_highlight = 0:7
    row_highlight = c(27,26)
    
    rhandsontable(MyChanges3(), col_highlight = col_highlight, row_highlight = row_highlight)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:8,readOnly = TRUE)%>%
      hot_row(c(27,28), readOnly = TRUE) %>%
      hot_cols(renderer = "
            function(instance, td, row, col, prop, value, cellProperties) {
               Handsontable.renderers.NumericRenderer.apply(this, arguments);
                if (instance.params) {
                    hcols = instance.params.col_highlight
                    hcols = hcols instanceof Array ? hcols : [hcols]
                    hrows = instance.params.row_highlight
                    hrows = hrows instanceof Array ? hrows : [hrows]
                }
                if (instance.params && hcols.includes(col)) td.style.background = 'lightgrey';
                if (instance.params && hrows.includes(row)) td.style.background = 'lightgreen';
            }")
    
  })
  observeEvent(input$save_safety3,{
    
    write.xlsx(hot_to_r(input$hotable3),"safety/first_aid.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  
  values4 <- reactiveValues()
  
  
  previous4<- reactive({
    read.xlsx("safety/accidents_countermeasure.xlsx",sheetIndex = 1)
  })
  
  MyChanges4 <- reactive({
    if(is.null(input$hotable4)){return(previous4())}
    else if(!identical(previous4(),input$hotable4)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable4 <- as.data.frame(hot_to_r(input$hotable4))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable4 <- mytable4[1:nrow(previous3()),]
      
      for(i in 3:14)
        mytable4[27,i]<-sum(mytable4[seq(1,26,2),i],na.rm=TRUE)
      
      for(i in 3:14)
        mytable4[28,i]<-sum(mytable4[seq(2,27,2),i],na.rm=TRUE)
      mytable4
    }
  })
  
  output$hotable4<-renderRHandsontable({
    col_highlight = 0:7
    row_highlight = c(27,26)
    
    rhandsontable(MyChanges4(), col_highlight = col_highlight, row_highlight = row_highlight)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:8,readOnly = TRUE)%>%
      hot_row(c(27,28), readOnly = TRUE) %>%
      hot_cols(renderer = "
            function(instance, td, row, col, prop, value, cellProperties) {
               Handsontable.renderers.NumericRenderer.apply(this, arguments);
                if (instance.params) {
                    hcols = instance.params.col_highlight
                    hcols = hcols instanceof Array ? hcols : [hcols]
                    hrows = instance.params.row_highlight
                    hrows = hrows instanceof Array ? hrows : [hrows]
                }
                if (instance.params && hcols.includes(col)) td.style.background = 'lightgrey';
                if (instance.params && hrows.includes(row)) td.style.background = 'lightgreen';
            }")
    
  })
  observeEvent(input$save_safety4,{
    
    write.xlsx(hot_to_r(input$hotable4),"safety/accidents_countermeasure.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  output$major_plot<-renderPlot({
    
    df2<-read.xlsx("safety/major_accidents.xlsx",sheetIndex = 1)
    
    ggplot(data=df2, aes(x=dose, y=len, fill=supp)) +
      geom_bar(stat="identity")
  })
  
  
  
  
  
  
  
  
  
  output$majorPlot <- renderPlotly({
    
    Products <- c("A", "B", "C","D","E","F")
    Trucks <- c(20, 14, 23,45,30,15)
    Busses <- c(12, 18, 29,12,2,67)
    data <- data.frame(Products, Trucks, Busses)
    
    p <- plot_ly(data, x = ~Products, y = ~Trucks, type = 'bar', name = 'Trucks') %>%
      add_trace(y = ~Busses, name = 'Busses') %>%
      layout(yaxis = list(title = 'Count'), barmode = 'stack')
    p
  })
  
  output$ibox20 <- renderInfoBox({
    
    d1<-read.xlsx("safety/major_accidents.xlsx",sheetIndex = 1)
    x1<-input$save_safety1
    if(sum(is.na(d1$Jul))==0){
      col="green"
      ic="thumbs-up"
      val="Filled"
    }
    else{
      col="red"
      ic="thumbs-down"
      val="Yet to fill"
    }
    
    infoBox(
      value=val,
      title="Major Accidents",
      icon=icon(ic),
      color=col
    )
  })

  
  output$ibox21 <- renderInfoBox({
    
    d1<-read.xlsx("safety/minor_accidents.xlsx",sheetIndex = 1)
    x1<-input$save_safety2
    if(sum(is.na(d1$Jul))==0){
      col="green"
      ic="thumbs-up"
      val="Filled"
    }
    else{
      col="red"
      ic="thumbs-down"
      val="Yet to fill"
    }
    
    infoBox(
      value=val,
      title="Minor Accidents",
      icon=icon(ic),
      color=col
    )
  })
  
  
  output$ibox22 <- renderInfoBox({
    
    d1<-read.xlsx("safety/first_aid.xlsx",sheetIndex = 1)
    x1<-input$save_safety3
    if(sum(is.na(d1$Jul))==0){
      col="green"
      ic="thumbs-up"
      val="Filled"
    }
    else{
      col="red"
      ic="thumbs-down"
      val="Yet to fill"
    }
    
    infoBox(
      value=val,
      title="First Aid",
      icon=icon(ic),
      color=col
    )
  })
  
  
  output$ibox23 <- renderInfoBox({
    
    d1<-read.xlsx("safety/accidents_countermeasure.xlsx",sheetIndex = 1)
    x1<-input$save_safety4
    if(sum(is.na(d1$Jul))==0){
      col="green"
      ic="thumbs-up"
      val="Filled"
    }
    else{
      col="red"
      ic="thumbs-down"
      val="Yet to fill"
    }
    
    infoBox(
      value=val,
      title="Accident Closure",
      icon=icon(ic),
      color=col
    )
  })
  
  output$text_safety<-renderText({
    d1<-read.xlsx("safety/accidents_countermeasure.xlsx",sheetIndex = 1)
    text<-''
  })
  
  
  
  
  output$plot_major<-renderPlotly({
    
    te<-input$save_safety1
    d<-read.xlsx("safety/major_accidents.xlsx",sheetIndex = 1)
    da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[27:27,3:14]))))
    da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[28:28,3:14]))))
    da1$co<-NA
    for(i in 1:nrow(da1)){
      if(da$yval[i]>=da1$yval[i])
        da1$co[i]<-"green"
      else
        da1$co[i]<-"red"
    }
    
    
    da1$xval<-factor(da1$xval,levels=c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"))
    da$xval<-factor(da$xval,levels=c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"))
    da1$yval<-as.numeric(da1$yval)
    p<-plot_ly(data = da1, x = ~xval, y = ~yval, type = "bar",color=~co, showlegend=FALSE) %>%
      add_lines(y = da$yval, showlegend=FALSE, color = 'black') %>%
      layout(showlegend=FALSE, xaxis = list(side="right", showgrid=FALSE),
             yaxis=list(showgrid=TRUE))
    
    
    p
  })
  
  output$plot_minor<-renderPlotly({
    
    te<-input$save_safety2
    d<-read.xlsx("safety/minor_accidents.xlsx",sheetIndex = 1)
    da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[27:27,3:14]))))
    da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[28:28,3:14]))))
    da1$co<-NA
    for(i in 1:nrow(da1)){
      if(da$yval[i]<=da1$yval[i])
        da1$co[i]<-"green"
      else
        da1$co[i]<-"red"
    }
    
    
    da1$xval<-factor(da1$xval,levels=c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"))
    da$xval<-factor(da$xval,levels=c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"))
    da1$yval<-as.numeric(da1$yval)
    p<-plot_ly(data = da1, x = ~xval, y = ~yval, type = "bar",color=~co, showlegend=FALSE) %>%
      add_lines(y = da$yval, showlegend=FALSE, color = 'black') %>%
      layout(showlegend=FALSE, xaxis = list(side="right", showgrid=FALSE),
             yaxis=list(showgrid=TRUE))
    
    
    p
  })
  
  output$plot_firstaid<-renderPlotly({
    
    te<-input$save_safety3
    d<-read.xlsx("safety/first_aid.xlsx",sheetIndex = 1)
    da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[27:27,3:14]))))
    da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[28:28,3:14]))))
    
    da1$co<-NA
    for(i in 1:nrow(da1)){
      if(da$yval[i]<=da1$yval[i])
        da1$co[i]<-"green"
      else
        da1$co[i]<-"red"
    }
    
    da1$xval<-factor(da1$xval,levels=c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"))
    da$xval<-factor(da$xval,levels=c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"))
    da1$yval<-as.numeric(da1$yval)
    p<-plot_ly(data = da1, x = ~xval, y = ~yval,color= ~co, type = "bar", showlegend=FALSE) %>%
      add_lines(y = da$yval, showlegend=FALSE, color = 'black') %>%
      layout(showlegend=FALSE, xaxis = list(side="right", showgrid=FALSE),
             yaxis=list(showgrid=TRUE))
    
    
    p
  })
  
  output$plot_counter<-renderPlotly({
    
    te<-input$save_safety4
    d<-read.xlsx("safety/accidents_countermeasure.xlsx",sheetIndex = 1)
    da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[27:27,3:14]))))
    da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[28:28,3:14]))))
    da1$co<-NA
    for(i in 1:nrow(da1)){
      if(da$yval[i]>=da1$yval[i])
        da1$co[i]<-"green"
      else
        da1$co[i]<-"red"
    }
      
    da1$xval<-factor(da1$xval,levels=c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"))
    da$xval<-factor(da$xval,levels=c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"))
    da1$yval<-as.numeric(da1$yval)
    p<-plot_ly(data = da1, x = ~xval, y = ~yval, color = ~co,type = "bar", showlegend=TRUE) %>%
      add_lines(y = da$yval, showlegend=FALSE, color = 'black') %>%
      layout(showlegend=FALSE, xaxis = list(side="right", showgrid=FALSE),
             yaxis=list(showgrid=TRUE))
    
    
    p
  })
  
 data_counter<-reactive({
   te<-input$save_safety4
   d<-read.xlsx("safety/accidents_countermeasure.xlsx",sheetIndex = 1)
   d<-d[d$Department!="Plant level",]
   da<-d[d$Category=='Actual',]
   dt<-d[d$Category=='Target',]
   
   da<-da%>%gather(month,value,Jan:Dec)
   da$Category<-NULL
   da$month<-factor(da$month,levels=c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"))
   da
 })
  
 data_dept_counter<-reactive({
   te<-input$save_safety4
   d<-read.xlsx("safety/accidents_countermeasure.xlsx",sheetIndex = 1)
   d<-d[d$Department!="Plant level",]
   
   d<-d%>%gather(month,value,Jan:Dec)
   
   d$month<-factor(d$month,levels=c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"))
   d
 })
  
  output$comp_counter<-renderPlotly({
    te<-input$save_safety4
    da<-data_counter()
    p <- plot_ly(da, x = ~month, y = ~value, type = 'bar',color = ~Department) %>%
      
      layout(yaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE), barmode = 'stack')
    p
  })
  
  output$comp_pie_counter<-renderPlotly({
    te<-input$save_safety4
    da<-data_counter()
    da<-da[da$month==input$choose_comp_counter,]
    
    p <- plot_ly(da, labels = ~Department, values = ~value, type = 'pie') %>%
      layout(title = 'Pie chart ',
             xaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE),
             yaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE))
    
  })
  output$dept_counter<-renderPlotly({
    te<-input$save_safety4
    da<-data_dept_counter()
    da<-da[da$Department==input$choose_indiv_counter,]
    da$Category<-factor(da$Category,levels=c("Target","Actual"))
    p <- plot_ly(da, x = ~month, y = ~value, type = 'bar',color = ~Category) %>%
      
      layout(yaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE), barmode = 'group')
    p
  })
  output$table_dept_counter<-DT::renderDataTable({
    te<-input$save_safety4
    da<-data_dept_counter()
    da<-da[da$Department==input$choose_indiv_counter,]
    da$Category<-factor(da$Category,levels=c("Target","Actual"))
    d<-spread(da,key = "month",value = value)
    d
  })
  output$deviation_counter<-renderPlotly({
    te<-input$save_safety4
    d<-read.xlsx("safety/accidents_countermeasure.xlsx",sheetIndex = 1)
    d<-d[d$Department!="Plant level",]
    if(input$choose_deviation_counter!='All')
      d<-d[d$Department==input$choose_deviation_counter,]
    da<-d[d$Category=='Actual',]
    dt<-d[d$Category=='Target',]
    
    da$Category<-NULL
    dt$Category<-NULL
    
    for (i in 1:nrow(da)){
      for(j in 2:ncol(dt)){
        da[i,j]<-dt[i,j]-da[i,j]
      }
    }
    
    da<-da%>%gather(month,value,Jan:Dec)
    da$month<-factor(da$month,levels=c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"))
    
    plot_ly(da,x=~month,y=~value,color = ~Department,type = 'scatter',mode='lines')
    
  })
  
  output$table_deviation_counter<-DT::renderDataTable({
    te<-input$save_safety4
    d<-read.xlsx("safety/accidents_countermeasure.xlsx",sheetIndex = 1)
    d<-d[d$Department!="Plant level",]
    if(input$choose_deviation_counter!='All')
      d<-d[d$Department==input$choose_deviation_counter,]
    d
  })
  
  data_firstaid<-reactive({
    te<-input$save_safety3
    d<-read.xlsx("safety/first_aid.xlsx",sheetIndex = 1)
    d<-d[d$Department!="Plant level",]
    da<-d[d$Category=='Actual',]
    dt<-d[d$Category=='Target',]
    
    da<-da%>%gather(month,value,Jan:Dec)
    da$Category<-NULL
    da$month<-factor(da$month,levels=c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"))
    da
  })
  
  data_dept_firstaid<-reactive({
    te<-input$save_safety3
    d<-read.xlsx("safety/first_aid.xlsx",sheetIndex = 1)
    d<-d[d$Department!="Plant level",]
    
    d<-d%>%gather(month,value,Jan:Dec)
    
    d$month<-factor(d$month,levels=c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"))
    d
  })
  
  output$comp_firstaid<-renderPlotly({
    te<-input$save_safety3
    da<-data_firstaid()
    p <- plot_ly(da, x = ~month, y = ~value, type = 'bar',color = ~Department) %>%
      
      layout(yaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE), barmode = 'stack')
    p
  })
  
  output$comp_pie_firstaid<-renderPlotly({
    te<-input$save_safety3
    da<-data_firstaid()
    da<-da[da$month==input$choose_comp_firstaid,]
    
    p <- plot_ly(da, labels = ~Department, values = ~value, type = 'pie') %>%
      layout(title = 'Pie chart ',
             xaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE),
             yaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE))
    
  })
  output$dept_firstaid<-renderPlotly({
    te<-input$save_safety3
    da<-data_dept_firstaid()
    da<-da[da$Department==input$choose_indiv_firstaid,]
    da$Category<-factor(da$Category,levels=c("Target","Actual"))
    p <- plot_ly(da, x = ~month, y = ~value, type = 'bar',color = ~Category) %>%
      
      layout(yaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE), barmode = 'group')
    p
  })
  output$table_dept_firstaid<-DT::renderDataTable({
    te<-input$save_safety3
    da<-data_dept_firstaid()
    da<-da[da$Department==input$choose_indiv_firstaid,]
    da$Category<-factor(da$Category,levels=c("Target","Actual"))
    d<-spread(da,key = "month",value = value)
    d
  })
  output$deviation_firstaid<-renderPlotly({
    te<-input$save_safety3
    d<-read.xlsx("safety/first_aid.xlsx",sheetIndex = 1)
    d<-d[d$Department!="Plant level",]
    if(input$choose_deviation_firstaid!='All')
      d<-d[d$Department==input$choose_deviation_firstaid,]
    da<-d[d$Category=='Actual',]
    dt<-d[d$Category=='Target',]
    
    da$Category<-NULL
    dt$Category<-NULL
    
    for (i in 1:nrow(da)){
      for(j in 2:ncol(dt)){
        da[i,j]<-dt[i,j]-da[i,j]
      }
    }
    
    da<-da%>%gather(month,value,Jan:Dec)
    da$month<-factor(da$month,levels=c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"))
    
    plot_ly(da,x=~month,y=~value,color = ~Department,type = 'scatter',mode='lines')
    
  })
  
  output$table_deviation_firstaid<-DT::renderDataTable({
    te<-input$save_safety3
    d<-read.xlsx("safety/first_aid.xlsx",sheetIndex = 1)
    d<-d[d$Department!="Plant level",]
    if(input$choose_deviation_firstaid!='All')
      d<-d[d$Department==input$choose_deviation_firstaid,]
    d
  })
}
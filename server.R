library(readxl)
library(rhandsontable)
library(shiny)
library(plotly)
library(dplyr)
library(RColorBrewer)
library(zoo)
library(lubridate)
library(tidyr)
library(xlsx)
library(plotly)
library(tibble)
library(ggplot2)
library(flextable)
library(reshape)
library(plyr)
library(xts)
library(magrittr)
library(officer)
library(flextable)

mon<-c("cum_2017","cum_2018","cum_2019","cum_2020","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
server <- function(session,input, output) {
  
  

  
  observeEvent(input$preview,{
    
    withProgress(message = 'Generating Report', value = 0, {
      incProgress(1/43, detail = paste("Under Progress"))
      doc<-read_pptx(path ="template.pptx")
      
      com<-read.xlsx("comments/comments_2020_1.xlsx",sheetIndex = 1)
      
      doc<-on_slide(x=doc,index=2)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("safety/major_accidents/major_accidents_","2020",".xlsx",sep="")
      d <- read_excel(f)
      d<-d[d$Department=="Plant level",]
      da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[1:1,3:14]))))
      da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[2:2,3:14]))))
      
      l<-list.files(path="safety/major_accidents/")
      
      for(z in l){
        na<-paste("safety/major_accidents/",z,sep='')
        dt<-read_excel(na)
        dt<-dt[dt$Department=="Plant level",]
        dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[1:1,3:14]))))
        dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[2:2,3:14]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=sum(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=sum(dta$yval,na.rm = TRUE))
      }
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=sum(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=as.integer(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=1.1) )
      
      
      
      d<-read_excel("safety/major_accidents/major_accidents_2020.xlsx")
      d<-d[d$Department!="Plant level",]
      da<-d[d$Category=='Actual',]
      dt<-d[d$Category=='Target',]
      
      da<-da%>%gather(month,value,3:14)
      da$Category<-NULL
      
      d<-read_excel("safety/major_accidents/major_accidents_2020.xlsx")
      d<-d[d$Department!="Plant level",]
      dt<-d[d$Category=='Target',]
      
      dt<-dt%>%gather(month,value,3:14)
      dt$Category<-NULL
      
      
      d<-read_excel("safety/major_accidents/major_accidents_2020.xlsx")
      d<-d[d$Department!="Plant level",]
      d<-d%>%gather(month,value,3:14)
      d$month<-as.yearmon(d$month,"%b %Y")
      d<-d[months(d$month)==months(Sys.Date()-30),]
      d$Department<-factor(d$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
      
      d<-spread(d,Department,value)
      
      d$Category<-factor(d$Category,levels=c("Target","Actual"))
      d$month<-months(d$month)
      
      
      
      
      
      da$month<-as.yearmon(da$month,"%b %Y")
      dt$month<-as.yearmon(dt$month,"%b %Y")
      
      da1<-da[months(da$month)==months(Sys.Date()-30),]
      da<-dt[months(dt$month)==months(Sys.Date()-30),]
      
      da$Department<-factor(da$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
      da1$Department<-factor(da1$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
      
      da$yval<-as.numeric(da$value)
      da1$yval<-as.numeric(da1$value)
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      p<-ggplot(da,aes(x=Department,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=as.integer(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=4.3) )
      
      
      co1<-com[com$KPI=="major",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=1.4)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=3.1)
      
      co1<-com[com$KPI=="major_shop",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=4.8)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=6.5)
      
      
      
      #Minor Accidents
      
      
      doc<-on_slide(x=doc,index=3)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("safety/minor_accidents/minor_accidents_","2020",".xlsx",sep="")
      d <- read_excel(f)
      d<-d[d$Department=="Plant level",]
      da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[1:1,3:14]))))
      da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[2:2,3:14]))))
      
      l<-list.files(path="safety/minor_accidents/")
      
      for(z in l){
        na<-paste("safety/minor_accidents/",z,sep='')
        dt<-read_excel(na)
        dt<-dt[dt$Department=="Plant level",]
        dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[1:1,3:14]))))
        dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[2:2,3:14]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=max(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=max(dta$yval,na.rm = TRUE))
      }
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=max(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=as.integer(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=1.1) )
      
      
      
      d<-read_excel("safety/minor_accidents/minor_accidents_2020.xlsx")
      d<-d[d$Department!="Plant level",]
      da<-d[d$Category=='Actual',]
      dt<-d[d$Category=='Target',]
      
      da<-da%>%gather(month,value,3:14)
      da$Category<-NULL
      
      d<-read_excel("safety/minor_accidents/minor_accidents_2020.xlsx")
      d<-d[d$Department!="Plant level",]
      dt<-d[d$Category=='Target',]
      
      dt<-dt%>%gather(month,value,3:14)
      dt$Category<-NULL
      
      
      d<-read_excel("safety/minor_accidents/minor_accidents_2020.xlsx")
      d<-d[d$Department!="Plant level",]
      d<-d%>%gather(month,value,3:14)
      d$month<-as.yearmon(d$month,"%b %Y")
      d<-d[months(d$month)==months(Sys.Date()-30),]
      d$Department<-factor(d$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
      
      d<-spread(d,Department,value )
      
      d$Category<-factor(d$Category,levels=c("Target","Actual"))
      d$month<-months(d$month)
      
      
      
      
      
      da$month<-as.yearmon(da$month,"%b %Y")
      dt$month<-as.yearmon(dt$month,"%b %Y")
      
      da1<-da[months(da$month)==months(Sys.Date()-30),]
      da<-dt[months(dt$month)==months(Sys.Date()-30),]
      
      da$Department<-factor(da$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
      da1$Department<-factor(da1$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
      
      da$yval<-as.numeric(da$value)
      da1$yval<-as.numeric(da1$value)
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      p<-ggplot(da,aes(x=Department,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=as.integer(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=4.3) )
      
      co1<-com[com$KPI=="minor",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=1.4)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=3.1)
      
      co1<-com[com$KPI=="minor_shop",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=4.8)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=6.5)
      
      
      # First Aid
      
      
      doc<-on_slide(x=doc,index=4)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("safety/first_aid/first_aid_","2020",".xlsx",sep="")
      d <- read_excel(f)
      d<-d[d$Department=="Plant level",]
      da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[1:1,3:14]))))
      da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[2:2,3:14]))))
      
      l<-list.files(path="safety/first_aid/")
      
      for(z in l){
        na<-paste("safety/first_aid/",z,sep='')
        dt<-read_excel(na)
        dt<-dt[dt$Department=="Plant level",]
        dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[1:1,3:14]))))
        dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[2:2,3:14]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=max(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=max(dta$yval,na.rm = TRUE))
      }
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=max(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=as.integer(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=1.1) )
      
      
      
      d<-read_excel("safety/first_aid/first_aid_2020.xlsx")
      d<-d[d$Department!="Plant level",]
      da<-d[d$Category=='Actual',]
      dt<-d[d$Category=='Target',]
      
      da<-da%>%gather(month,value,3:14)
      da$Category<-NULL
      
      d<-read_excel("safety/first_aid/first_aid_2020.xlsx")
      d<-d[d$Department!="Plant level",]
      dt<-d[d$Category=='Target',]
      
      dt<-dt%>%gather(month,value,3:14)
      dt$Category<-NULL
      
      
      d<-read_excel("safety/first_aid/first_aid_2020.xlsx")
      d<-d[d$Department!="Plant level",]
      d<-d%>%gather(month,value,3:14)
      d$month<-as.yearmon(d$month,"%b %Y")
      d<-d[months(d$month)==months(Sys.Date()-30),]
      d$Department<-factor(d$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
      
      d<-spread(d,Department,value )
      
      d$Category<-factor(d$Category,levels=c("Target","Actual"))
      d$month<-months(d$month)
      
      
      
      
      
      da$month<-as.yearmon(da$month,"%b %Y")
      dt$month<-as.yearmon(dt$month,"%b %Y")
      
      da1<-da[months(da$month)==months(Sys.Date()-30),]
      da<-dt[months(dt$month)==months(Sys.Date()-30),]
      
      da$Department<-factor(da$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
      da1$Department<-factor(da1$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
      
      da$yval<-as.numeric(da$value)
      da1$yval<-as.numeric(da1$value)
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      p<-ggplot(da,aes(x=Department,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=as.integer(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=4.3) )
      
      co1<-com[com$KPI=="firstaid",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=1.4)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=3.1)
      
      co1<-com[com$KPI=="firstaid_shop",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=4.8)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=6.5)
      
      
      
      
      #Unsafe Acts
      
      
      doc<-on_slide(x=doc,index=5)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("safety/unsafe_acts/unsafe_acts_","2020",".xlsx",sep="")
      d <- read_excel(f)
      d<-d[d$Department=="Plant level",]
      da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[1:1,3:14]))))
      da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[2:2,3:14]))))
      
      l<-list.files(path="safety/unsafe_acts/")
      
      for(z in l){
        na<-paste("safety/unsafe_acts/",z,sep='')
        dt<-read_excel(na)
        dt<-dt[dt$Department=="Plant level",]
        dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[1:1,3:14]))))
        dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[2:2,3:14]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=max(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=max(dta$yval,na.rm = TRUE))
      }
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=max(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=as.integer(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=1.1) )
      
      
      
      d<-read_excel("safety/unsafe_acts/unsafe_acts_2020.xlsx")
      d<-d[d$Department!="Plant level",]
      da<-d[d$Category=='Actual',]
      dt<-d[d$Category=='Target',]
      
      da<-da%>%gather(month,value,3:14)
      da$Category<-NULL
      
      d<-read_excel("safety/unsafe_acts/unsafe_acts_2020.xlsx")
      d<-d[d$Department!="Plant level",]
      dt<-d[d$Category=='Target',]
      
      dt<-dt%>%gather(month,value,3:14)
      dt$Category<-NULL
      
      
      d<-read_excel("safety/unsafe_acts/unsafe_acts_2020.xlsx")
      d<-d[d$Department!="Plant level",]
      d<-d%>%gather(month,value,3:14)
      d$month<-as.yearmon(d$month,"%b %Y")
      d<-d[months(d$month)==months(Sys.Date()-30),]
      d$Department<-factor(d$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
      
      d<-spread(d,Department,value )
      
      d$Category<-factor(d$Category,levels=c("Target","Actual"))
      d$month<-months(d$month)
      
      
      
      
      
      da$month<-as.yearmon(da$month,"%b %Y")
      dt$month<-as.yearmon(dt$month,"%b %Y")
      
      da1<-da[months(da$month)==months(Sys.Date()-30),]
      da<-dt[months(dt$month)==months(Sys.Date()-30),]
      
      da$Department<-factor(da$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
      da1$Department<-factor(da1$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
      
      da$yval<-as.numeric(da$value)
      da1$yval<-as.numeric(da1$value)
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      p<-ggplot(da,aes(x=Department,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=as.integer(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=4.3) )
      
      co1<-com[com$KPI=="unsafe",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=1.4)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=3.1)
      
      co1<-com[com$KPI=="unsafe_shop",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=4.8)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=6.5)
      
      
      
      
      
      #Accidets counter measure
      
      
      doc<-on_slide(x=doc,index=6)
      incProgress(1/43, detail = paste("Under Progress"))
      
      
      f<-paste("safety/accidents_countermeasure/accidents_countermeasure_","2020",".xlsx",sep="")
      d <- read_excel(f)
      d<-d[d$Department=="Plant level",]
      da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[1:1,3:14]))))
      da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[2:2,3:14]))))
      
      l<-list.files(path="safety/accidents_countermeasure/")
      
      for(z in l){
        na<-paste("safety/accidents_countermeasure/",z,sep='')
        dt<-read_excel(na)
        dt<-dt[dt$Department=="Plant level",]
        dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[1:1,3:14]))))
        dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[2:2,3:14]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(da$yval[i]!=0)
          da$yval2[i]<-as.integer(100*da$yval2[i]/da$yval[i])
        else
          da$yval2[i]<-0
      }
      da$yval<-100
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=as.integer(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8,height = 2.7,left=0,top=1.1) )
      
      
      
      d<-read_excel("safety/accidents_countermeasure/accidents_countermeasure_2020.xlsx")
      d<-d[d$Department!="Plant level",]
      da<-d[d$Category=='Actual',]
      dt<-d[d$Category=='Target',]
      
      da<-da%>%gather(month,value,3:14)
      da$Category<-NULL
      
      d<-read_excel("safety/accidents_countermeasure/accidents_countermeasure_2020.xlsx")
      d<-d[d$Department!="Plant level",]
      dt<-d[d$Category=='Target',]
      
      dt<-dt%>%gather(month,value,3:14)
      dt$Category<-NULL
      
      
      d<-read_excel("safety/accidents_countermeasure/accidents_countermeasure_2020.xlsx")
      d<-d[d$Department!="Plant level",]
      d<-d%>%gather(month,value,3:14)
      d$month<-as.yearmon(d$month,"%b %Y")
      d<-d[months(d$month)==months(Sys.Date()-30),]
      d$Department<-factor(d$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
      
      d<-spread(d,Department,value )
      
      d$Category<-factor(d$Category,levels=c("Target","Actual"))
      d$month<-months(d$month)
      
      
      
      
      
      da$month<-as.yearmon(da$month,"%b %Y")
      dt$month<-as.yearmon(dt$month,"%b %Y")
      
      da1<-da[months(da$month)==months(Sys.Date()-30),]
      da<-dt[months(dt$month)==months(Sys.Date()-30),]
      
      da$Department<-factor(da$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
      da1$Department<-factor(da1$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
      
      da$yval<-as.numeric(da$value)
      da1$yval<-as.numeric(da1$value)
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      p<-ggplot(da,aes(x=Department,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=as.integer(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.5,height = 2.7,left=0,top=4.3) )
      
      co1<-com[com$KPI=="counter",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=1.4)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=3.1)
      
      co1<-com[com$KPI=="counter_shop",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=4.8)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=6.5)
      
      
      
      
      
      
      #DPU 
      
      
      doc<-on_slide(x=doc,index=8)
      incProgress(1/43, detail = paste("Under Progress"))
      
      
      f<-paste("QM/QM_DPU_QFL4/QM_DPU_QFL4_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[1:1,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))))
      
      l<-list.files(path="QM/QM_DPU_QFL4/")
      
      for(z in l){
        na<-paste("QM/QM_DPU_QFL4/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[1:1,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[2:2,4:15]))))
        
        na<-strsplit(z,"_")[[1]][4]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      
      da<-da[order(da$xval),]
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=6.6,height = 2.75,left=0,top=1.1) )
      
      f<-paste("QM/QM_DPU_QFL4/QM_DPU_QFL4_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[3:3,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[4:4,4:15]))))
      
      l<-list.files(path="QM/QM_DPU_QFL4/")
      
      for(z in l){
        na<-paste("QM/QM_DPU_QFL4/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[3:3,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[4:4,4:15]))))
        
        na<-strsplit(z,"_")[[1]][4]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      #da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      
      da<-da[order(da$xval),]
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=6.6,height = 2.75,left=0,top=4.3) )
      
      
      f<-paste("QM/QM_DPU_QFL4/QM_DPU_QFL4_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[5:5,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[6:6,4:15]))))
      
      l<-list.files(path="QM/QM_DPU_QFL4/")
      
      for(z in l){
        na<-paste("QM/QM_DPU_QFL4/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[5:5,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[6:6,4:15]))))
        
        na<-strsplit(z,"_")[[1]][4]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      
      da<-da[order(da$xval),]
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=6.6,height = 2.75,left=6.4,top=1.1) )
      
      f<-paste("QM/QM_DPU_QFL4/QM_DPU_QFL4_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[7:7,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[8:8,4:15]))))
      
      l<-list.files(path="QM/QM_DPU_QFL4/")
      
      for(z in l){
        na<-paste("QM/QM_DPU_QFL4/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[7:7,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[8:8,4:15]))))
        
        na<-strsplit(z,"_")[[1]][4]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      
      da<-da[order(da$xval),]
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=6.6,height = 2.75,left=6.4,top=4.3) )
      
      
      
      
      
      #DPU Teardown audit
      
      
      doc<-on_slide(x=doc,index=9)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("QM/QM/QM_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[15:15,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[16:16,4:15]))))
      
      l<-list.files(path="QM/QM/")
      
      for(z in l){
        na<-paste("QM/QM/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[15:15,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[16:16,4:15]))))
        
        na<-strsplit(z,"_")[[1]][2]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      
      da<-da[order(da$xval),]
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=2.5)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=6.6,height = 2.75,left=0,top=1.1) )
      
      f<-paste("QM/QM/QM_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[17:17,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[18:18,4:15]))))
      
      l<-list.files(path="QM/QM/")
      
      for(z in l){
        na<-paste("QM/QM/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[17:17,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[18:18,4:15]))))
        
        na<-strsplit(z,"_")[[1]][2]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      
      da<-da[order(da$xval),]
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=2.5)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=6.6,height = 2.75,left=0,top=4.3) )
      
      
      f<-paste("QM/QM/QM_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[19:19,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[20:20,4:15]))))
      
      l<-list.files(path="QM/QM/")
      
      for(z in l){
        na<-paste("QM/QM/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[19:19,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[20:20,4:15]))))
        
        na<-strsplit(z,"_")[[1]][2]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      
      da<-da[order(da$xval),]
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=2.5)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=6.6,height = 2.75,left=6.4,top=1.1) )
      
      f<-paste("QM/QM/QM_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[21:21,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[22:22,4:15]))))
      
      l<-list.files(path="QM/QM/")
      
      for(z in l){
        na<-paste("QM/QM/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[21:21,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[22:22,4:15]))))
        
        na<-strsplit(z,"_")[[1]][2]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      
      da<-da[order(da$xval),]
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=2.5)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      #theme(panel.background = element_blank())
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=6.6,height = 2.75,left=6.4,top=4.3) )
      
      
      
      #DPU Ops related
      
      doc<-on_slide(x=doc,index=10)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("QM/QM_DPU_QFL4_OPS/QM_DPU_QFL4_OPS_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[1:1,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))))
      
      l<-list.files(path="QM/QM_DPU_QFL4_OPS/")
      
      for(z in l){
        na<-paste("QM/QM_DPU_QFL4_OPS/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[1:1,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[2:2,4:15]))))
        
        na<-strsplit(z,"_")[[1]][5]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=1.1) )
      
      f<-paste("QM/QM_DPU_QFL4_OPS/QM_DPU_QFL4_OPS_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[3:3,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[4:4,4:15]))))
      
      l<-list.files(path="QM/QM_DPU_QFL4_OPS/")
      
      for(z in l){
        na<-paste("QM/QM_DPU_QFL4_OPS/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[3:3,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[4:4,4:15]))))
        
        na<-strsplit(z,"_")[[1]][5]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=2.5)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=4.3) )
      
      co1<-com[com$KPI=="dpu_hdt_ops",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=1.4)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=3.1)
      
      co1<-com[com$KPI=="dpu_hdt_ops_ab",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=4.8)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=6.5)
      
      
      
      #DPU MDT Ops related
      
      doc<-on_slide(x=doc,index=11)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("QM/QM_DPU_QFL4_OPS/QM_DPU_QFL4_OPS_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[5:5,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[6:6,4:15]))))
      
      l<-list.files(path="QM/QM_DPU_QFL4_OPS/")
      
      for(z in l){
        na<-paste("QM/QM_DPU_QFL4_OPS/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[5:5,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[6:6,4:15]))))
        
        na<-strsplit(z,"_")[[1]][5]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=1.1) )
      
      f<-paste("QM/QM_DPU_QFL4_OPS/QM_DPU_QFL4_OPS_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[7:7,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[8:8,4:15]))))
      
      l<-list.files(path="QM/QM_DPU_QFL4_OPS/")
      
      for(z in l){
        na<-paste("QM/QM_DPU_QFL4_OPS/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[7:7,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[8:8,4:15]))))
        
        na<-strsplit(z,"_")[[1]][5]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=2.5)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=4.3) )
      
      
      co1<-com[com$KPI=="dpu_mdt_ops",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=1.4)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=3.1)
      
      co1<-com[com$KPI=="dpu_mdt_ops_ab",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=4.8)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=6.5)
      
      
      
      #DPU QFL2
      
      doc<-on_slide(x=doc,index=12)
      incProgress(1/43, detail = paste("Under Progress"))
      
      
      f<-paste("Chassis/KPI/chassis_kpi_",2020,".xlsx",sep="")
      d <- read_excel(f)
      
      d<-d[d$Description=="DPU @ QFL2 - HDT",]
      
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))),tar=c(t(array(d[1:1,4:15]))),co=c(t(array(d[2:2,4:15]+d[3:3,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[3:3,4:15]))),tar=c(t(array(d[1:1,4:15]))),co=c(t(array(d[2:2,4:15]+d[3:3,4:15]))))
      da$te<-"Veh_ass"
      da1$te<-"Scratches"
      da<-rbind(da,da1)
      
      da<-da%>%add_row(xval="2018",yval=2.87,tar=18,co=18,te="Veh_ass")
      da<-da%>%add_row(xval="2018",yval=15.13,tar=18,co=18,te="Scratches")
      
      da<-da%>%add_row(xval="2019",yval=1.61,tar=12,co=10.64,te="Veh_ass")
      da<-da%>%add_row(xval="2019",yval=9.03,tar=12,co=10.64,te="Scratches")
      
      
      da<-da%>%add_row(xval="2020",yval=rowMeans(d[2:2,4:15],na.rm=TRUE),tar=rowMeans(d[1:1,4:15],na.rm=TRUE),co=rowMeans(d[2:2,4:15],na.rm=TRUE)+rowMeans(d[3:3,4:15],na.rm=TRUE),te="Veh_ass")
      da<-da%>%add_row(xval="2020",yval=rowMeans(d[3:3,4:15],na.rm=TRUE),tar=rowMeans(d[1:1,4:15],na.rm=TRUE),co=rowMeans(d[2:2,4:15],na.rm=TRUE)+rowMeans(d[3:3,4:15],na.rm=TRUE),te="Scratches")
      
      
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=0,tar=12,co=0,te="Veh_ass")
      da<-da%>%add_row(xval=na,yval=0,tar=12,co=0,te="Scratches")
      
      
      
      for(i in 1:nrow(da))
        if(is.na(da$yval[i]))
          da$yval[i]<-0
      
      for(i in 1:nrow(da))
        if(is.na(da$co[i]))
          da$co[i]<-0
      
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$yval<-as.numeric(da$yval)
      
      da<-da[order(da$xval),]
      da$zval<-NA
      for(i in 1:nrow(da)){
        if(da$te[i]=="Scratches")
          da$zval[i]<-"grey"
        else if(da$co[i]>da$tar[i])
          da$zval[i]<-"red"
        else
          da$zval[i]<-"green4"
      }
      
      p<-ggplot(data=da, aes(x=xval, y=yval, fill=te)) +
        geom_bar(stat="identity",fill=da$zval)+
        geom_line(aes(y=tar,group = 1),color='black')+
        geom_text(aes(label=round(yval,2)), vjust=1.5, color="white", size=2.5)+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=1.1) )
      
      
      
      f<-paste("Chassis/KPI/chassis_kpi_",2020,".xlsx",sep="")
      d <- read_excel(f)
      
      d<-d[d$Description=="DPU @ QFL2 - MDT",]
      
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))),tar=c(t(array(d[1:1,4:15]))),co=c(t(array(d[2:2,4:15]+d[3:3,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[3:3,4:15]))),tar=c(t(array(d[1:1,4:15]))),co=c(t(array(d[2:2,4:15]+d[3:3,4:15]))))
      da$te<-"Veh_ass"
      da1$te<-"Scratches"
      da<-rbind(da,da1)
      
      da<-da%>%add_row(xval="2018",yval=1.74,tar=18,co=12.87,te="Veh_ass")
      da<-da%>%add_row(xval="2018",yval=11.13,tar=18,co=12.87,te="Scratches")
      
      
      da<-da%>%add_row(xval="2019",yval=2.38,tar=12,co=11.43,te="Veh_ass")
      da<-da%>%add_row(xval="2019",yval=9.05,tar=12,co=11.43,te="Scratches")
      
      da<-da%>%add_row(xval="2020",yval=rowMeans(d[2:2,4:15],na.rm=TRUE),tar=rowMeans(d[1:1,4:15],na.rm=TRUE),co=rowMeans(d[2:2,4:15],na.rm=TRUE)+rowMeans(d[3:3,4:15],na.rm=TRUE),te="Veh_ass")
      da<-da%>%add_row(xval="2020",yval=rowMeans(d[3:3,4:15],na.rm=TRUE),tar=rowMeans(d[1:1,4:15],na.rm=TRUE),co=rowMeans(d[2:2,4:15],na.rm=TRUE)+rowMeans(d[3:3,4:15],na.rm=TRUE),te="Scratches")
      
      
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=0,tar=12,co=0,te="Veh_ass")
      da<-da%>%add_row(xval=na,yval=0,tar=12,co=0,te="Scratches")
      
      
      
      for(i in 1:nrow(da))
        if(is.na(da$yval[i]))
          da$yval[i]<-0
      for(i in 1:nrow(da))
        if(is.na(da$co[i]))
          da$co[i]<-0
      
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$yval<-as.numeric(da$yval)
      
      da<-da[order(da$xval),]
      da$zval<-NA
      for(i in 1:nrow(da)){
        if(da$te[i]=="Scratches")
          da$zval[i]<-"grey"
        else if(da$co[i]>da$tar[i])
          da$zval[i]<-"red"
        else
          da$zval[i]<-"green4"
      }
      
      p<-ggplot(data=da, aes(x=xval, y=yval, fill=te)) +
        geom_bar(stat="identity",fill=da$zval)+
        geom_line(aes(y=tar,group = 1),color='black')+
        geom_text(aes(label=round(yval,2)), vjust=1.5, color="white", size=2.5)+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=4.3) )
      
      
      co1<-com[com$KPI=="dpu_hdt_qfl2",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=1.4)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=3.1)
      
      co1<-com[com$KPI=="dpu_mdt_qfl2",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=4.8)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=6.5)
      
      
      
      #DPU Engine
      
      doc<-on_slide(x=doc,index=13)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("QM/QM/QM_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[35:35,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[36:36,4:15]))))
      
      l<-list.files(path="QM/QM/")
      
      for(z in l){
        na<-paste("QM/QM/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[35:35,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[36:36,4:15]))))
        
        na<-strsplit(z,"_")[[1]][2]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      
      co1<-com[com$KPI=="dpu_eng",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      
      #DPU Transmission
      
      doc<-on_slide(x=doc,index=14)
      incProgress(1/43, detail = paste("Under Progress"))
      f<-paste("QM/QM/QM_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[37:37,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[38:38,4:15]))))
      
      l<-list.files(path="QM/QM/")
      
      for(z in l){
        na<-paste("QM/QM/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[37:37,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[38:38,4:15]))))
        
        na<-strsplit(z,"_")[[1]][2]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      co1<-com[com$KPI=="dpu_tra",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      #DPU CiW
      
      doc<-on_slide(x=doc,index=15)
      incProgress(1/43, detail = paste("Under Progress"))
      f<-paste("QM/QM/QM_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[41:41,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[42:42,4:15]))))
      
      l<-list.files(path="QM/QM/")
      
      for(z in l){
        na<-paste("QM/QM/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[41:41,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[42:42,4:15]))))
        
        na<-strsplit(z,"_")[[1]][2]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      co1<-com[com$KPI=="dpu_ciw",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      #DPU Paint
      
      doc<-on_slide(x=doc,index=16)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("QM/QM/QM_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[39:39,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[40:40,4:15]))))
      
      l<-list.files(path="QM/QM/")
      
      for(z in l){
        na<-paste("QM/QM/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[39:39,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[40:40,4:15]))))
        
        na<-strsplit(z,"_")[[1]][2]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      co1<-com[com$KPI=="dpu_pai",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      #DPU Frame
      
      doc<-on_slide(x=doc,index=17)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("QM/QM/QM_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[43:43,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[44:44,4:15]))))
      
      l<-list.files(path="QM/QM/")
      
      for(z in l){
        na<-paste("QM/QM/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[43:43,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[44:44,4:15]))))
        
        na<-strsplit(z,"_")[[1]][2]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      
      co1<-com[com$KPI=="dpu_fra",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      #FTT HDT
      
      doc<-on_slide(x=doc,index=18)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("spr_ftt/spr_ftt_",2020,".xlsx",sep="")
      d <- read_excel(f)
      
      d<-d[d$KPI=="FTT",]
      d<-d[d$Description=="HDT",]
      
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[1:1,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))))
      
      
      l<-list.files(path="spr_ftt/")
      
      for(z in l){
        na<-paste("spr_ftt/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[1:1,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[2:2,4:15]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      co1<-com[com$KPI=="ftt_hdt",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      #FTT MDT
      
      doc<-on_slide(x=doc,index=19)
      incProgress(1/43, detail = paste("Under Progress"))
      f<-paste("spr_ftt/spr_ftt_",2020,".xlsx",sep="")
      d <- read_excel(f)
      
      d<-d[d$KPI=="FTT",]
      d<-d[d$Description=="MDT",]
      
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[1:1,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))))
      
      
      l<-list.files(path="spr_ftt/")
      
      for(z in l){
        na<-paste("spr_ftt/",z,sep='')
        dt<-read_excel(na)
        
        dt<-dt[dt$KPI=="FTT",]
        dt<-dt[dt$Description=="MDT",]
        
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[1:1,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[2:2,4:15]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date()-30)+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date()-30)+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date()-30)+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      co1<-com[com$KPI=="ftt_mdt",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      #FTT LDT
      
      doc<-on_slide(x=doc,index=20)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("spr_ftt/spr_ftt_",2020,".xlsx",sep="")
      d <- read_excel(f)
      
      d<-d[d$KPI=="FTT",]
      d<-d[d$Description=="LDT",]
      
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[1:1,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))))
      
      
      l<-list.files(path="spr_ftt/")
      
      for(z in l){
        na<-paste("spr_ftt/",z,sep='')
        dt<-read_excel(na)
        dt<-dt[dt$KPI=="FTT",]
        dt<-dt[dt$Description=="LDT",]
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[1:1,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[2:2,4:15]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=dtt$yval[month(Sys.Date()-30)])
        da1<-da1%>%add_row(xval=na,yval=dta$yval[month(Sys.Date()-30)])
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      co1<-com[com$KPI=="ftt_ldt",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      #SPR HDT
      
      doc<-on_slide(x=doc,index=21)
      incProgress(1/43, detail = paste("Under Progress"))
      f<-paste("spr_ftt/spr_ftt_",2020,".xlsx",sep="")
      d <- read_excel(f)
      
      d<-d[d$KPI=="SPR",]
      d<-d[d$Description=="HDT",]
      
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[1:1,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))))
      
      
      l<-list.files(path="spr_ftt/")
      
      for(z in l){
        na<-paste("spr_ftt/",z,sep='')
        dt<-read_excel(na)
        dt<-dt[dt$KPI=="SPR",]
        dt<-dt[dt$Description=="HDT",]
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[1:1,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[2:2,4:15]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=dtt$yval[month(Sys.Date()-30)])
        da1<-da1%>%add_row(xval=na,yval=dta$yval[month(Sys.Date()-30)])
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      co1<-com[com$KPI=="spr_hdt",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      #SPR MDT
      
      doc<-on_slide(x=doc,index=22)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("spr_ftt/spr_ftt_",2020,".xlsx",sep="")
      d <- read_excel(f)
      
      d<-d[d$KPI=="SPR",]
      d<-d[d$Description=="MDT",]
      
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[1:1,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))))
      
      
      l<-list.files(path="spr_ftt/")
      
      for(z in l){
        na<-paste("spr_ftt/",z,sep='')
        dt<-read_excel(na)
        dt<-dt[dt$KPI=="SPR",]
        dt<-dt[dt$Description=="MDT",]
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[1:1,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[2:2,4:15]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=dtt$yval[month(Sys.Date()-30)])
        da1<-da1%>%add_row(xval=na,yval=dta$yval[month(Sys.Date()-30)])
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      
      co1<-com[com$KPI=="spr_mdt",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      
      #SPR LDT
      
      doc<-on_slide(x=doc,index=23)
      incProgress(1/43, detail = paste("Under Progress"))
      f<-paste("spr_ftt/spr_ftt_",2020,".xlsx",sep="")
      d <- read_excel(f)
      
      d<-d[d$KPI=="SPR",]
      d<-d[d$Description=="LDT",]
      
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[1:1,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))))
      
      
      l<-list.files(path="spr_ftt/")
      
      for(z in l){
        na<-paste("spr_ftt/",z,sep='')
        dt<-read_excel(na)
        dt<-dt[dt$KPI=="SPR",]
        dt<-dt[dt$Description=="LDT",]
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[1:1,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[2:2,4:15]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=dtt$yval[month(Sys.Date()-30)])
        da1<-da1%>%add_row(xval=na,yval=dta$yval[month(Sys.Date()-30)])
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      
      co1<-com[com$KPI=="spr_ldt",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      
      #Delivery QC Ok
      
      
      doc<-on_slide(x=doc,index=25)
      incProgress(1/43, detail = paste("Under Progress"))
      
      
      f<-paste("Chassis/Chassis_qcok/chassis_qcok_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[5:5,3:14]))))
      da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[6:6,3:14]))))
      
      l<-list.files(path="Chassis/Chassis_qcok/")
      
      for(z in l){
        na<-paste("Chassis/Chassis_qcok/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[5:5,3:14]))))
        dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[6:6,3:14]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2)), vjust=1.5, color="white", size=2.5)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=6.6,height = 2.75,left=0,top=1.1) )
      
      f<-paste("Chassis/Chassis_qcok/chassis_qcok_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[7:7,3:14]))))
      da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[8:8,3:14]))))
      
      l<-list.files(path="Chassis/Chassis_qcok/")
      
      for(z in l){
        na<-paste("Chassis/Chassis_qcok/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[7:7,3:14]))))
        dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[8:8,3:14]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=6.6,height = 2.75,left=6.4,top=1.1) )
      
      f<-paste("Chassis/Chassis_capacity/chassis_capacity_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[5:5,3:14]))))
      da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[6:6,3:14]))))
      
      l<-list.files(path="Chassis/Chassis_capacity/")
      
      for(z in l){
        na<-paste("Chassis/Chassis_capacity/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[5:5,3:14]))))
        dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[6:6,3:14]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2)), vjust=1.5, color="white", size=2.5)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=12,height = 2.65,left=0,top=4.4) )
      
      #Roll Out
      
      doc<-on_slide(x=doc,index=26)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("Chassis/Chassis_rollout/chassis_rollout_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[1:1,3:14]))))
      da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[2:2,3:14]))))
      
      l<-list.files(path="Chassis/Chassis_rollout/")
      
      for(z in l){
        na<-paste("Chassis/Chassis_rollout/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[1:1,3:14]))))
        dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[2:2,3:14]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=1.1) )
      
      f<-paste("Chassis/Chassis_rollout/chassis_rollout_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[3:3,3:14]))))
      da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[4:4,3:14]))))
      
      l<-list.files(path="Chassis/Chassis_rollout/")
      
      for(z in l){
        na<-paste("Chassis/Chassis_rollout/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[3:3,3:14]))))
        dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[4:4,3:14]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=2.5)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=4.3) )
      co1<-com[com$KPI=="roll_hdt",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=1.4)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=3.1)
      
      co1<-com[com$KPI=="roll_mdt",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=4.8)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=6.5)
      
      
      #capacity utilization
      
      doc<-on_slide(x=doc,index=27)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("Chassis/Chassis_capacity/chassis_capacity_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[1:1,3:14]))))
      da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[2:2,3:14]))))
      
      l<-list.files(path="Chassis/Chassis_capacity/")
      
      for(z in l){
        na<-paste("Chassis/Chassis_capacity/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[1:1,3:14]))))
        dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[2:2,3:14]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=1.1) )
      
      f<-paste("Chassis/Chassis_capacity/chassis_capacity_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[3:3,3:14]))))
      da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[4:4,3:14]))))
      
      l<-list.files(path="Chassis/Chassis_capacity/")
      
      for(z in l){
        na<-paste("Chassis/Chassis_capacity/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[3:3,3:14]))))
        dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[4:4,3:14]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2)), vjust=1.5, color="white", size=2.5)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=4.3) )
      
      co1<-com[com$KPI=="cap_hdt",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=1.4)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=3.1)
      
      co1<-com[com$KPI=="cap_mdt",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=4.8)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=6.5)
      
      
      #non forecasted shortages
      
      doc<-on_slide(x=doc,index=28)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("Chassis/KPI/chassis_kpi_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[21:21,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[22:22,4:15]))))
      
      l<-list.files(path="Chassis/KPI/")
      
      for(z in l){
        na<-paste("Chassis/KPI/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[21:21,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[22:22,4:15]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        # if(da$yval[i]>=da$yval2[i])
        #   da$zval[i]<-"green4"
        # else
        da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        # geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      
      co1<-com[com$KPI=="forecasted",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      
      #Vehicle losses due to Operations
      
      
      doc<-on_slide(x=doc,index=29)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("Chassis/KPI/chassis_kpi_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[29:29,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[30:30,4:15]))))
      
      l<-list.files(path="Chassis/KPI/")
      
      for(z in l){
        na<-paste("Chassis/KPI/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[29:29,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[30:30,4:15]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      
      co1<-com[com$KPI=="loss_ope",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      
      #Vehicle losses due to aggregates
      
      
      doc<-on_slide(x=doc,index=30)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("Chassis/KPI/chassis_kpi_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[31:31,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[32:32,4:15]))))
      
      l<-list.files(path="Chassis/KPI/")
      
      for(z in l){
        na<-paste("Chassis/KPI/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[31:31,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[32:32,4:15]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      
      co1<-com[com$KPI=="loss_agg",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      
      #indirect consumables
      
      doc<-on_slide(x=doc,index=33)
      incProgress(1/43, detail = paste("Under Progress"))
      
      cost_data<-function(){
        d1<-read_excel("Chassis/KPI/chassis_kpi_2020.xlsx")
        d1<-d1[d1$KPI=="Cost",]
        d1$KPI<-"Chassis"
        
        d<-d1
        
        d1<-read_excel("Cabtrim/KPI/cabtrim_kpi_2020.xlsx")
        d1<-d1[d1$KPI=="Cost",]
        d1$KPI<-"Cabtrim"
        
        d<-rbind(d,d1)
        
        d1<-read_excel("EOL/KPI/eol_kpi_2020.xlsx")
        d1<-d1[d1$KPI=="Cost",]
        d1$KPI<-"EOL"
        
        d<-rbind(d,d1)
        
        d1<-read_excel("FBV/KPI/fbv_kpi_2020.xlsx")
        d1<-d1[d1$KPI=="Cost",]
        d1$KPI<-"FBV"
        
        d<-rbind(d,d1)
        
        d1<-read_excel("CIW/KPI/ciw_kpi_2020.xlsx")
        d1<-d1[d1$KPI=="Cost",]
        d1$KPI<-"CiW"
        
        d<-rbind(d,d1)
        
        d1<-read_excel("PAINT/KPI/paint_kpi_2020.xlsx")
        d1<-d1[d1$KPI=="Cost",]
        d1$KPI<-"Paint"
        
        d<-rbind(d,d1)
        
        d1<-read_excel("engine/KPI/engine_kpi_2020.xlsx")
        d1<-d1[d1$KPI=="Cost",]
        d1$KPI<-"Engine"
        
        d<-rbind(d,d1)
        
        d1<-read_excel("transmission/KPI/transmission_kpi_2020.xlsx")
        d1<-d1[d1$KPI=="Cost",]
        d1$KPI<-"Transmission"
        
        d<-rbind(d,d1)
        
        d1<-read_excel("ipl/KPI/ipl_kpi_2020.xlsx")
        d1<-d1[d1$KPI=="Cost",]
        d1$KPI<-"IPL"
        
        d<-rbind(d,d1)
        
        
        d1<-read_excel("QM/QM/QM_2020.xlsx")
        
        d1<-d1[d1$KPI=="Cost",]
        d1$KPI<-"QM"
        
        d<-rbind(d,d1)
        
        d1<-read_excel("fm/KPI/fm_kpi_2020.xlsx")
        
        d1<-d1[d1$KPI=="Cost",]
        d1$KPI<-"FM"
        
        d<-rbind(d,d1)
        
        
        d1<-read_excel("frame/KPI/frame_kpi_2020.xlsx")
        
        d1<-d1[d1$KPI=="Cost",]
        d1$KPI<-"Frame"
        
        d<-rbind(d,d1)
        d}
      
      d<-cost_data()
      d<-d[d$Description=="Indirect Consumables",]
      
      dd<-d[d$Category=="Target",]
      dd2<-d[d$Category=="Actual",]
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd[,4:15],sum,na.rm=TRUE)))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd2[,4:15],sum,na.rm=TRUE)))))
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=1.1) )
      
      data_cos_indirect_cons_shop_act<-function(){
        d <- cost_data()
        
        d<-d[d$Description=="Indirect Consumables",]
        
        dd<-d[d$Category=="Actual",]
        dd2<-d[d$Category=="Target",]
        
        da<-dd%>%gather(month,value,4:15)
        da$Category<-NULL
        da
      }
      data_cos_indirect_cons_shop_tar<-function(){
        d <- cost_data()
        
        d<-d[d$Description=="Indirect Consumables",]
        
        dd<-d[d$Category=="Actual",]
        dd2<-d[d$Category=="Target",]
        
        dt<-dd2%>%gather(month,value,4:15)
        
        dt$Category<-NULL
        dt
      }
      
      da<-data_cos_indirect_cons_shop_act()
      dt<-data_cos_indirect_cons_shop_tar()
      
      da$month<-as.yearmon(da$month,"%b %Y")
      dt$month<-as.yearmon(dt$month,"%b %Y")
      
      da1<-da[months(da$month)==months(Sys.Date()-30),]
      da<-dt[months(dt$month)==months(Sys.Date()-30),]
      
      da$KPI<-factor(da$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","IPL","QM","FM","Frame"))
      da1$KPI<-factor(da1$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","IPL","QM","FM","Frame"))
      
      da$yval<-as.numeric(da$value)
      da1$yval<-as.numeric(da1$value)
      
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      p<-ggplot(da,aes(x=KPI,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=4.3) )
      
      co1<-com[com$KPI=="cons_plant",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=1.4)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=3.1)
      
      co1<-com[com$KPI=="cons_shop",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=4.8)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=6.5)
      
      
      doc<-on_slide(x=doc,index=34)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("HPU_Capacity/HPU_Capacity/hpu_capacity_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[1:1,3:14]))))
      da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[2:2,3:14]))))
      
      l<-list.files(path="HPU_Capacity/HPU_Capacity/")
      
      for(z in l){
        na<-paste("HPU_Capacity/HPU_Capacity/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[1:1,3:14]))))
        dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[2:2,3:14]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=1.1) )
      
      
      d<-read_excel("HPU_Capacity/HPU_Capacity/hpu_capacity_2020.xlsx")
      d<-d[d$Department!="Plant level",]
      da<-d[d$Category=='Actual',]
      
      da<-da%>%gather(month,value,3:14)
      da$Category<-NULL
      
      d<-read_excel("HPU_Capacity/HPU_Capacity/hpu_capacity_2020.xlsx")
      d<-d[d$Department!="Plant level",]
      dt<-d[d$Category=='Target',]
      
      dt<-dt%>%gather(month,value,3:14)
      dt$Category<-NULL
      
      da$month<-as.yearmon(da$month,"%b %Y")
      dt$month<-as.yearmon(dt$month,"%b %Y")
      
      da1<-da[months(da$month)==months(Sys.Date()-30),]
      da<-dt[months(dt$month)==months(Sys.Date()-30),]
      
      da$Department<-factor(da$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","Frame","TOS"))
      da1$Department<-factor(da1$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","Frame","TOS"))
      
      da$yval<-as.numeric(da$value)
      da1$yval<-as.numeric(da1$value)
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      p<-ggplot(da,aes(x=Department,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=4.3) )
      co1<-com[com$KPI=="hpu_cap_plant",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=1.4)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=3.1)
      
      co1<-com[com$KPI=="hpu_cap_shop",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=4.8)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=6.5)
      
      
      # electricity and propane
      
      doc<-on_slide(x=doc,index=35)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("fm/KPI/fm_kpi_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[13:13,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[14:14,4:15]))))
      
      l<-list.files(path="fm/KPI/")
      
      for(z in l){
        na<-paste("fm/KPI/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[13:13,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[14:14,4:15]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,0)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=1.1) )
      
      f<-paste("fm/KPI/fm_kpi_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[15:15,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[16:16,4:15]))))
      
      l<-list.files(path="fm/KPI/")
      
      for(z in l){
        na<-paste("fm/KPI/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[15:15,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[16:16,4:15]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,0)), vjust=1.5, color="white", size=2.5)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=8.6,height = 2.7,left=0,top=4.3) )
      
      co1<-com[com$KPI=="electricity",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=1.4)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=3.1)
      
      co1<-com[com$KPI=="propane",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=4.8)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=8.75,top=6.5)
      
      
      #Rejection cost
      
      
      doc<-on_slide(x=doc,index=36)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("QM/QM/QM_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[57:57,4:15]))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[58:58,4:15]))))
      da$yval<-150
      
      l<-list.files(path="QM/QM/")
      
      for(z in l){
        na<-paste("QM/QM/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[57:57,4:15]))))
        dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[58:58,4:15]))))
        dtt$yval<-150
        
        na<-strsplit(z,"_")[[1]][2]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      co1<-com[com$KPI=="rej_plant",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      
      #Rejection cost
      
      
      doc<-on_slide(x=doc,index=37)
      incProgress(1/43, detail = paste("Under Progress"))
      
      data_rej_cost_shop_act<-function(){
        
        d<-read_excel("QM/QM/QM_2020.xlsx")
        d<-d[d$Description=="Rejection cost",]
        da<-d[d$Category!='Cost/truck',]
        da<-da[da$Category!='Volume',]
        da<-da[da$Category!='Plant level',]
        da<-da[da$Category!='Veh prod',]
        
        da<-da%>%gather(month,value,4:15)
        
        da
      }
      data_rej_cost_shop_tar<-function(){
        
        d<-read_excel("QM/QM/QM_2020.xlsx")
        d<-d[d$Description=="Rejection cost",]
        dt<-d[d$Category!='Cost/truck',]
        dt<-dt[dt$Category!='Volume',]
        dt<-dt[dt$Category!='Plant level',]
        dt<-dt[dt$Category!='Veh prod',]
        
        
        dt<-dt%>%gather(month,value,4:15)
        
        dt
      }
      da<-data_rej_cost_shop_act()
      dt<-data_rej_cost_shop_tar()
      
      da$month<-as.yearmon(da$month,"%b %Y")
      dt$month<-as.yearmon(dt$month,"%b %Y")
      
      da1<-da[months(da$month)==months(Sys.Date()-30),]
      da<-dt[months(dt$month)==months(Sys.Date()-30),]
      
      da$value<-c(38000,19000,7000,0,1000,1000,32000,4000,1000,47000)
      
      da$Category<-factor(da$Category,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QA-Vehicle","QA-PTI"))
      da1$Category<-factor(da1$Category,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QA-Vehicle","QA-PTI"))
      
      da$yval<-as.numeric(da$value)
      da1$yval<-as.numeric(da1$value)
      
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      p<-ggplot(da,aes(x=Category,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      co1<-com[com$KPI=="rej_shop",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      
      #white collar attrition
      
      
      doc<-on_slide(x=doc,index=39)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("white_collar/white_collar_",2020,".xlsx",sep="")
      d <- read_excel(f)
      da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[45:45,3:14]))))
      da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[46:46,3:14]))))
      
      l<-list.files(path="white_collar/")
      
      for(z in l){
        na<-paste("white_collar/",z,sep='')
        dt<-read_excel(na)
        dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[45:45,3:14]))))
        dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[46:46,3:14]))))
        
        na<-strsplit(z,"_")[[1]][3]
        na<-strsplit(na,"[.]")[[1]][1]
        da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
        da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      }
      
      na=toString(year(Sys.Date())+1)
      
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      co1<-com[com$KPI=="att_white_plant",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      
      #white collar
      
      
      doc<-on_slide(x=doc,index=40)
      incProgress(1/43, detail = paste("Under Progress"))
      
      
      data_white_collar_shop_act<-function(){
        
        d<-read_excel("white_collar/white_collar_2020.xlsx")
        d<-d[d$KPI!="Plant level",]
        da<-d[d$Description=='Ratio',]
        dt<-d[d$Description=='Target',]
        
        da<-da%>%gather(month,value,3:14)
        da$Description<-NULL
        da
      }
      data_white_collar_shop_tar<-function(){
        
        d<-read_excel("white_collar/white_collar_2020.xlsx")
        d<-d[d$KPI!="Plant level",]
        da<-d[d$Description=='Ratio',]
        dt<-da
        
        dt<-dt%>%gather(month,value,3:14)
        dt$value<-3.3
        dt$Description<-NULL
        dt
      }
      
      
      
      da<-data_white_collar_shop_act()
      dt<-data_white_collar_shop_tar()
      
      da$month<-as.yearmon(da$month,"%b %Y")
      dt$month<-as.yearmon(dt$month,"%b %Y")
      
      da1<-da[months(da$month)==months(Sys.Date()-30),]
      da<-dt[months(dt$month)==months(Sys.Date()-30),]
      
      da$KPI<-factor(da$KPI,levels=c("Chassis","Cabtrim","EOL/ FBV","CiW","Paint","Engine","Transmn.","QM","FM","IPL","ME","Frame","VP office","TOS & OMCD"))
      da1$KPI<-factor(da1$KPI,levels=c("Chassis","Cabtrim","EOL/ FBV","CiW","Paint","Engine","Transmn.","QM","FM","IPL","ME","Frame","VP office","TOS & OMCD"))
      
      da$yval<-as.numeric(da$value)
      da1$yval<-as.numeric(da1$value)
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      p<-ggplot(da,aes(x=KPI,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      co1<-com[com$KPI=="att_white_shop",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      #Plant level: Kaizen (No. of kaizen/BCA & Engineer/month)
      
      
      
      doc<-on_slide(x=doc,index=41)
      incProgress(1/43, detail = paste("Under Progress"))
      
      
      morale_data<-function(){
        
        d1<-read_excel("Chassis/KPI/chassis_kpi_2020.xlsx")
        d1<-d1[d1$KPI=="Morale",]
        d1$KPI<-"Chassis"
        
        d<-d1
        
        d1<-read_excel("Cabtrim/KPI/cabtrim_kpi_2020.xlsx")
        d1<-d1[d1$KPI=="Morale",]
        d1$KPI<-"Cabtrim"
        
        d<-rbind(d,d1)
        
        d1<-read_excel("EOL/KPI/eol_kpi_2020.xlsx")
        d1<-d1[d1$KPI=="Morale",]
        d1$KPI<-"EOL"
        
        d<-rbind(d,d1)
        
        d1<-read_excel("FBV/KPI/fbv_kpi_2020.xlsx")
        d1<-d1[d1$KPI=="Morale",]
        d1$KPI<-"FBV"
        
        d<-rbind(d,d1)
        
        d1<-read_excel("CIW/KPI/ciw_kpi_2020.xlsx")
        d1<-d1[d1$KPI=="Morale",]
        d1$KPI<-"CiW"
        
        d<-rbind(d,d1)
        
        d1<-read_excel("PAINT/KPI/paint_kpi_2020.xlsx")
        d1<-d1[d1$KPI=="Morale",]
        d1$KPI<-"Paint"
        
        d<-rbind(d,d1)
        
        d1<-read_excel("engine/KPI/engine_kpi_2020.xlsx")
        d1<-d1[d1$KPI=="Morale",]
        d1$KPI<-"Engine"
        
        d<-rbind(d,d1)
        
        d1<-read_excel("transmission/KPI/transmission_kpi_2020.xlsx")
        d1<-d1[d1$KPI=="Morale",]
        d1$KPI<-"Transmission"
        
        d<-rbind(d,d1)
        
        d1<-read_excel("ipl/KPI/ipl_kpi_2020.xlsx")
        d1<-d1[d1$KPI=="Morale",]
        d1$KPI<-"IPL"
        
        d<-rbind(d,d1)
        
        
        d1<-read_excel("QM/QM/QM_2020.xlsx")
        
        d1<-d1[d1$KPI=="Morale",]
        d1$KPI<-"QM"
        
        d<-rbind(d,d1)
        
        d1<-read_excel("fm/KPI/fm_kpi_2020.xlsx")
        
        d1<-d1[d1$KPI=="Morale",]
        d1$KPI<-"FM"
        
        d<-rbind(d,d1)
        
        
        d1<-read_excel("frame/KPI/frame_kpi_2020.xlsx")
        
        d1<-d1[d1$KPI=="Morale",]
        d1$KPI<-"Frame"
        
        d<-rbind(d,d1)
        
        d
      }
      d <- morale_data()
      
      d<-d[d$Description=="Participation in AOM - BCA/T & Engineers",]
      
      dd<-d[d$Category=="Actual",]
      dd2<-d[d$Category=="Target",]
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd[,4:15],mean,na.rm=TRUE)))))
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd2[,4:15],mean,na.rm=TRUE)))))
      
      da<-da%>%add_row(xval="2020",yval=mean(da$yval))
      da1<-da1%>%add_row(xval="2020",yval=mean(da1$yval))
      
      da<-da%>%add_row(xval="2017",yval=0.38)
      da1<-da1%>%add_row(xval="2017",yval=0.39)
      
      da<-da%>%add_row(xval="2018",yval=0.42)
      da1<-da1%>%add_row(xval="2018",yval=0.42)
      
      da<-da%>%add_row(xval="2019",yval=0.42)
      da1<-da1%>%add_row(xval="2019",yval=0.43)
      
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      co1<-com[com$KPI=="att_bca_plant",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      
      #Plant level: Kaizen (No. of kaizen/BCA & Engineer/month)
      
      
      doc<-on_slide(x=doc,index=42)
      incProgress(1/43, detail = paste("Under Progress"))
      
      data_bca_participation_shop_act<-function(){
        d <- morale_data()
        
        d<-d[d$Description=="Participation in AOM - BCA/T & Engineers",]
        
        dd<-d[d$Category=="Actual",]
        dd2<-d[d$Category=="Target",]
        
        da<-dd%>%gather(month,value,4:15)
        da$Category<-NULL
        da
      }
      data_bca_participation_shop_tar<-function(){
        d <- morale_data()
        
        d<-d[d$Description=="Participation in AOM - BCA/T & Engineers",]
        
        dd<-d[d$Category=="Actual",]
        dd2<-d[d$Category=="Target",]
        
        dt<-dd2%>%gather(month,value,4:15)
        
        dt$Category<-NULL
        dt
      }
      
      da<-data_bca_participation_shop_act()
      dt<-data_bca_participation_shop_tar()
      
      da$month<-as.yearmon(da$month,"%b %Y")
      dt$month<-as.yearmon(dt$month,"%b %Y")
      
      da1<-da[months(da$month)==months(Sys.Date()-30),]
      da<-dt[months(dt$month)==months(Sys.Date()-30),]
      
      da$KPI<-factor(da$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL"))
      da1$KPI<-factor(da1$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL"))
      
      da$yval<-as.numeric(da$value)
      da1$yval<-as.numeric(da1$value)
      
      da<-da[order(da$KPI),]
      da1<-da1[order(da1$KPI),]
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      p<-ggplot(da,aes(x=KPI,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      co1<-com[com$KPI=="att_bca_shop",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      
      #Plant level: Kaizen (No. of kaizen/CA/month)
      
      
      doc<-on_slide(x=doc,index=43)
      incProgress(1/43, detail = paste("Under Progress"))
      
      
      d <- morale_data()
      d<-d[d$KPI!="IPL",]
      d<-d[d$Description=="Participation in AOM - CA/BA",]
      
      dd<-d[d$Category=="Actual",]
      dd2<-d[d$Category=="Target",]
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd[,4:15],mean,na.rm=TRUE)))))
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd2[,4:15],mean,na.rm=TRUE)))))
      
      
      da<-da%>%add_row(xval="2020",yval=mean(da$yval))
      da1<-da1%>%add_row(xval="2020",yval=mean(da1$yval))
      
      
      da<-da%>%add_row(xval="2019",yval=0.42)
      da1<-da1%>%add_row(xval="2019",yval=0.39)
      
      
      # l<-list.files(path="caba_participation/")
      # 
      # for(z in l){
      #   na<-paste("caba_participation/",z,sep='')
      #   dt<-read_excel(na)
      #   dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[45:45,3:14]))))
      #   dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[46:46,3:14]))))
      # 
      #   na<-strsplit(z,"_")[[1]][3]
      #   na<-strsplit(na,"[.]")[[1]][1]
      #   da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      #   da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
      # }
      # 
      # na=toString(year(Sys.Date())+1)
      # 
      # da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      # da1<-da1%>%add_row(xval=na,yval=0)
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      co1<-com[com$KPI=="att_baca_plant",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      
      #Shop level: Kaizen (No. of kaizen/CA/month)
      
      
      doc<-on_slide(x=doc,index=44)
      incProgress(1/43, detail = paste("Under Progress"))
      
      
      data_caba_participation_shop_act<-function(){
        
        d<-d[d$KPI!="IPL",]
        d<-d[d$Description=="Participation in AOM - CA/BA",]
        
        dd<-d[d$Category=="Actual",]
        dd2<-d[d$Category=="Target",]
        
        da<-dd%>%gather(month,value,4:15)
        da$Category<-NULL
        da
      }
      data_caba_participation_shop_tar<-function(){
        
        d<-d[d$KPI!="IPL",]
        d<-d[d$Description=="Participation in AOM - CA/BA",]
        
        dd<-d[d$Category=="Actual",]
        dd2<-d[d$Category=="Target",]
        
        dt<-dd2%>%gather(month,value,4:15)
        
        dt$Category<-NULL
        dt
      }
      
      da<-data_caba_participation_shop_act()
      dt<-data_caba_participation_shop_tar()
      
      da$month<-as.yearmon(da$month,"%b %Y")
      dt$month<-as.yearmon(dt$month,"%b %Y")
      
      da1<-da[months(da$month)==months(Sys.Date()-30),]
      da<-dt[months(dt$month)==months(Sys.Date()-30),]
      
      da$KPI<-factor(da$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","QM","FM"))
      da1$KPI<-factor(da1$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","QM","FM"))
      
      da$yval<-as.numeric(da$value)
      da1$yval<-as.numeric(da1$value)
      
      da<-da[order(da$KPI),]
      da1<-da1[order(da1$KPI),]
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]<=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      p<-ggplot(da,aes(x=KPI,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      co1<-com[com$KPI=="att_baca_shop",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      
      #Plant level: Kaizen (No. of kaizen/CA)
      
      
      
      doc<-on_slide(x=doc,index=45)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("bca_attrition/bca_attrition_",2020,".xlsx",sep="")
      d <- read_excel(f)
      d<-d[d$Department=="Plant level",]
      
      d<-d[d$Description=="Others_ratio",]
      da<-gather(d,"xval","yval2",4:15)
      da$yval<-2
      
      
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$yval<-as.numeric(da$yval)
      
      
      da<-da[order(da$xval),]
      
      da$zval<-NA
      
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      
      
      
      
      doc<-on_slide(x=doc,index=46)
      incProgress(1/43, detail = paste("Under Progress"))
      
      f<-paste("bca_attrition/bca_attrition_",2020,".xlsx",sep="")
      d <- read_excel(f)
      
      d<-d[d$Situation=="Attrition",]
      
      da<-gather(d,"month","yval2",4:15)
      
      da$month<-as.yearmon(da$month,"%b %Y")
      
      da<-da[months(da$month)==months(Sys.Date()-30),]
      da$yval<-2
      
      
      
      
      
      
      da$zval<-NA
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      p<-ggplot(da,aes(x=Department,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      
      
      #Plant level: contractors attrition rate
      
      
      
      
      doc<-on_slide(x=doc,index=47)
      incProgress(1/43, detail = paste("Under Progress"))
      
      d <- morale_data()
      
      d<-d[d$Description=="Attrition rate of Contractors",]
      
      dd<-d[d$Category=="Required",]
      dd2<-d[d$Category=="Left",]
      da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd[,4:15],sum,na.rm=TRUE)))))
      da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd2[,4:15],sum,na.rm=TRUE)))))
      
      
      da<-da%>%add_row(xval="2020",yval=mean(da$yval))
      da1<-da1%>%add_row(xval="2020",yval=mean(da1$yval))
      
      
      da1$yval<-100*da1$yval/(da$yval)
      da$yval<-2
      
      da<-da%>%add_row(xval="2019",yval=2)
      da1<-da1%>%add_row(xval="2019",yval=7.1)
      
      da<-da%>%add_row(xval="2018",yval=0.8)
      da1<-da1%>%add_row(xval="2018",yval=3.6)
      
      da<-da%>%add_row(xval="2017",yval=0.8)
      da1<-da1%>%add_row(xval="2017",yval=17.3)
      
      da<-da%>%add_row(xval="2021",yval=2)
      da1<-da1%>%add_row(xval="2021",yval=0)
      
      
      da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
      da1$yval<-as.numeric(da1$yval)
      da$yval<-as.numeric(da$yval)
      
      
      da<-da[order(da$xval),]
      da1<-da1[order(da1$xval),]
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      
      
      p<-ggplot(da,aes(x=xval,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      
      co1<-com[com$KPI=="att_con_plant",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      
      
      
      
      doc<-on_slide(x=doc,index=48)
      incProgress(1/43, detail = paste("Under Progress"))
      
      data_con_attrition_shop_act<-function(){
        d <- morale_data()
        
        d<-d[d$Description=="Attrition rate of Contractors",]
        
        dd<-d[d$Category=="Rate",]
        dd2<-d[d$Category=="Target",]
        
        da<-dd%>%gather(month,value,4:15)
        da$Category<-NULL
        da
      }
      data_con_attrition_shop_tar<-function(){
        d <- morale_data()
        
        d<-d[d$Description=="Attrition rate of Contractors",]
        
        dd<-d[d$Category=="Rate",]
        dd2<-d[d$Category=="Target",]
        
        dt<-dd2%>%gather(month,value,4:15)
        
        dt$Category<-NULL
        dt
      }
      da<-data_con_attrition_shop_act()
      dt<-data_con_attrition_shop_tar()
      
      da$month<-as.yearmon(da$month,"%b %Y")
      dt$month<-as.yearmon(dt$month,"%b %Y")
      
      da1<-da[months(da$month)==months(Sys.Date()-30),]
      da<-dt[months(dt$month)==months(Sys.Date()-30),]
      
      da$KPI<-factor(da$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","IPL","QM","FM","Frame"))
      da1$KPI<-factor(da1$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","IPL","QM","FM","Frame"))
      
      da$yval<-as.numeric(da$value)
      da1$yval<-as.numeric(da1$value)
      
      da<-da[order(da$KPI),]
      da1<-da1[order(da1$KPI),]
      
      
      da$zval<-NA
      da$yval2<-da1$yval
      
      for(i in 1:nrow(da)){
        if(is.na(da$yval[i]))
          da$yval[i]<-0
        if(is.na(da$yval2[i]))
          da$yval2[i]<-0
      }
      
      for (i in 1:nrow(da)){
        if(da$yval[i]>=da$yval2[i])
          da$zval[i]<-"green4"
        else
          da$zval[i]<-"red"
      }
      p<-ggplot(da,aes(x=KPI,y=yval2))+
        geom_bar(stat="identity", fill=da$zval, width = 0.75)+
        geom_text(aes(label=round(yval2,2)), vjust=1.5, color="white", size=3)+
        geom_line(aes(y=yval,group = 1),color='black')+
        # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        labs(x="",y="")
      
      doc <- ph_with(x = doc, value = p, 
                     location = ph_location(width=13,height = 3.75,left=0,top=1.1) )
      co1<-com[com$KPI=="att_con_shop",]
      tb<-flextable(co1[3])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=0.8,top=5.45)
      
      tb<-flextable(co1[4])
      tb<-delete_part(x=tb, part = "header")
      tb<-border_remove(x=tb)
      tb<-align(x=tb,  align = "left", part = "body")
      tb <- width(tb, width = 4)
      doc <-ph_with_flextable_at(x=doc,tb,left=7.15,top=5.45)
      
      
      count<-48
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/att_con/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/att_con/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=49)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/att_bca/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/att_bca/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=47)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/kai_ca/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/kai_ca/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=45)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/kai_bca/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/kai_bca/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=43)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/white/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/white/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=41)
        }
      
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/rej/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/rej/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=38)
        }
      
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/ele/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/ele/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=36)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/hpucapacity/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/hpucapacity/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=35)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/indirect/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/indirect/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=34)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/loss/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/loss/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=31)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/fore/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/fore/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=29)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/cap/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/cap/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=28)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/roll/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/roll/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=27)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/qcok/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/qcok/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=26)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/spr/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/spr/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=24)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/ftt/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/ftt/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=21)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/qfl2/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/qfl2/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=18)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/ops_mdt/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/ops_mdt/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=12)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/ops_hdt/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/ops_hdt/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=11)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/tear/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/tear/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=10)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/qfl4/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/qfl4/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=9)
        }
      
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/counter/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/counter/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=7)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/unsafe/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/unsafe/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=6)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/firstaid/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/firstaid/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=5)
        }
      
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/minor/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/minor/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=4)
        }
      li<-list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/major/",sep=""))
      
      if(length(li)!=0)
        for (i in length(li):1){
          na<-paste(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/major/",sep=""),i,".png",sep="")
          count<-count+1
          doc<-add_slide(doc,layout = "Blank", master = "Blank")
          doc<-on_slide(x=doc,index=count)
          doc<-ph_with_img(x=doc, src=na,location = ph_location(width=13.33,height = 7.5,left=0,top=0))
          
          doc<-move_slide(doc,index = count,to=3)
        }
      
      print(doc,target=paste(choose.dir(default = "", caption = "Select folder location to save report"),"Report.pptx",sep = "\\")) 
      
      
    })
  })
  
 
  # suppress warnings  
  storeWarn<- getOption("warn")
  options(warn = -1)
  
  output$ibox1 <- renderInfoBox({    
    te<-input$save_safety1
    d<-read_excel("safety/major_accidents/major_accidents_2020.xlsx")
    d<-d[d$Department=='Plant level',]
    d1<-d%>%gather(date,value,3:ncol(d))
    
    d1$date<-gsub('-',' ',d1$date)
    d1$date<-as.yearmon(d1$date,"%b %Y")
    d1<-d1[(month(d1$date)==month(Sys.Date()-30)),]
    a<-d1[d1$Category=='Actual',]$value
    t<-d1[d1$Category=='Target',]$value
    if(a>t){
      ic<-"thumbs-down"
      co="red"
    }
    else{
      ic<-"thumbs-up"
      co="green"
    }
    infoBox(
      value=va<-paste("T:",as.integer(t)," A:",as.integer(a),sep=""),
      title="Major Accidents",
      icon=icon(ic),
      color=co
    )
  })
  output$ibox2 <- renderInfoBox({    
    te<-input$save_safety2
    d<-read_excel("safety/minor_accidents/minor_accidents_2020.xlsx")
    d<-d[d$Department=='Plant level',]
    d1<-d%>%gather(date,value,3:ncol(d))
    
    d1$date<-gsub('-',' ',d1$date)
    d1$date<-as.yearmon(d1$date,"%b %Y")
    d1<-d1[(month(d1$date)==month(Sys.Date()-30)),]
    a<-d1[d1$Category=='Actual',]$value
    t<-d1[d1$Category=='Target',]$value
    if(a>t){
      ic<-"thumbs-down"
      co="red"
    }
    else{
      ic<-"thumbs-up"
      co="green"
    }
    infoBox(
      value=va<-paste("T:",as.integer(t)," A:",as.integer(a),sep=""),
      title="Minor Accidents",
      icon=icon(ic),
      color=co
    )
  })

  output$ibox3 <- renderInfoBox({    
    te<-input$save_safety3
    d<-read_excel("safety/first_aid/first_aid_2020.xlsx")
    d<-d[d$Department=='Plant level',]
    d1<-d%>%gather(date,value,3:ncol(d))
    
    d1$date<-gsub('-',' ',d1$date)
    d1$date<-as.yearmon(d1$date,"%b %Y")
    d1<-d1[(month(d1$date)==month(Sys.Date()-30)),]
    a<-d1[d1$Category=='Actual',]$value
    t<-d1[d1$Category=='Target',]$value
    if(a>t){
      ic<-"thumbs-down"
      co="red"
    }
    else{
      ic<-"thumbs-up"
      co="green"
    }
    infoBox(
      value=va<-paste("T:",as.integer(t)," A:",a,sep=""),
      title="First Aid",
      icon=icon(ic),
      color=co
    )
  })
  output$ibox4 <- renderInfoBox({    
    te<-input$save_safety4
    d<-read_excel("safety/accidents_countermeasure/accidents_countermeasure_2020.xlsx")
    d<-d[d$Department=='Plant level',]
    d1<-d%>%gather(date,value,3:ncol(d))
    
    d1$date<-gsub('-',' ',d1$date)
    d1$date<-as.yearmon(d1$date,"%b %Y")
    d1<-d1[(month(d1$date)==month(Sys.Date()-30)),]
    a<-d1[d1$Category=='Actual',]$value
    t<-d1[d1$Category=='Target',]$value
    if(a<t){
      ic<-"thumbs-down"
      co="red"
    }
    else{
      ic<-"thumbs-up"
      co="green"
    }
    infoBox(
      value=va<-paste("T:",t," A:",a,sep=""),
      title="Accidents Counter Measure",
      icon=icon(ic),
      color=co
    )
  })
  
  output$ibox5 <- renderInfoBox({    
    te<-input$save_safety5
    d <- read_excel("safety/unsafe_acts/unsafe_acts_2020.xlsx")
    d<-d[d$Department=='Plant level',]
    d1<-d%>%gather(date,value,3:ncol(d))
    
    d1$date<-gsub('-',' ',d1$date)
    d1$date<-as.yearmon(d1$date,"%b %Y")
    d1<-d1[(month(d1$date)==month(Sys.Date()-30)),]
    a<-d1[d1$Category=='Actual',]$value
    t<-d1[d1$Category=='Target',]$value
    if(a<t){
      ic<-"thumbs-down"
      co="red"
    }
    else{
      ic<-"thumbs-up"
      co="green"
    }
    infoBox(
      value=va<-paste("T:",t," A:",a,sep=""),
      title="Unsafe Acts",
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
  
  
  
  output$comments_table_major<-renderDataTable({
    te<-input$save_comment_major
    d<-read_excel("safety/comments/major_accidents_comments.xlsx")
    d <-d[d$Month==input$choose_major_display_month,]
    d <-d[d$Year==input$choose_major_display_year,]
    d$Year<-NULL
    d$Month<-NULL
    d
  })

    

  observeEvent(input$save_comment_major,{
    d<-read_excel("safety/comments/major_accidents_comments.xlsx")
    da<-data.frame(Year=input$comment_choose_major_year,Month=input$comment_choose_major_month,Gap_analysis=input$comment_desc_enablers,Key_tasks=input$comment_desc_tasks)
    d<-rbind(d,da)
    write.xlsx(as.data.frame(d),"safety/comments/major_accidents_comments.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  output$table_unsafe<-renderDataTable({
    te<-input$save_comment_unsafe
    d<-read.xlsx("comments/unsafe.xlsx",sheetIndex = 1)
    d <-d[d$Month==input$choose_unsafe,]
    d
  })
  
  
  observeEvent(input$save_comment_unsafe,{
    d<-read.xlsx("comments/unsafe.xlsx",sheetIndex = 1)
    da<-data.frame(Month=input$comment_choose_unsafe,Department=input$choose_dept_unsafe,Description=input$description_unsafe)
    d<-rbind(d,da)
    write.xlsx(d,"comments/unsafe.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #Major accidents
  values1 <- reactiveValues()
  
  
   previous1 <- reactive({
     read_excel("safety/major_accidents/major_accidents_2020.xlsx")
   })
  
  MyChanges1 <- reactive({
    if(is.null(input$hotable1)){return(previous1())}
    else if(!identical(previous1(),input$hotable1)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable1 <- as.data.frame(hot_to_r(input$hotable1))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable1 <- mytable1[1:nrow(previous1()),]
      
      for(i in 3:ncol(mytable1))
      mytable1[27,i]<-sum(mytable1[seq(1,26,2),i],na.rm=TRUE)
      
      for(i in 3:ncol(mytable1))
        mytable1[28,i]<-sum(mytable1[seq(2,27,2),i],na.rm=TRUE)
     
      
       
      mytable1
    }
  })
  
  output$hotable1<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30))
    row_highlight = c(27,26)
    
    rhandsontable(MyChanges1(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1000,height = 600)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)),readOnly = TRUE)%>%
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
    write.xlsx(hot_to_r(input$hotable1),"safety/major_accidents/major_accidents_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  
  values2 <- reactiveValues()
  
  
  previous2 <- reactive({
    read_excel("safety/minor_accidents/minor_accidents_2020.xlsx")
  })
  
  MyChanges2 <- reactive({
    if(is.null(input$hotable2)){return(previous2())}
    else if(!identical(previous2(),input$hotable2)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable2 <- as.data.frame(hot_to_r(input$hotable2))
      # here 2he second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable2 <- mytable2[1:nrow(previous2()),]
      
      #for(i in 3:ncol(mytable2))
      #  mytable2[27,i]<-sum(mytable2[seq(1,26,2),i],na.rm=TRUE)
      mytable2[28,3]<-sum(mytable2[seq(2,27,2),3],na.rm=TRUE)
      for(i in 3:(2+month(Sys.Date()-30)))
        if(i!=3)
        mytable2[28,i]<-mytable2[28,i-1]+sum(mytable2[seq(2,27,2),i],na.rm=TRUE)
      
      
      
      mytable2
    }
  })
  
  output$hotable2<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30))
    row_highlight = c(27,26)
    
    rhandsontable(MyChanges2(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1000,height = 600)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)),readOnly = TRUE)%>%
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
    write.xlsx(hot_to_r(input$hotable2),"safety/minor_accidents/minor_accidents_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  
  values3 <- reactiveValues()
  
  
  previous3 <- reactive({
    read_excel("safety/first_aid/first_aid_2020.xlsx")
  })
  
  MyChanges3 <- reactive({
    if(is.null(input$hotable3)){return(previous3())}
    else if(!identical(previous3(),input$hotable3)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable3 <- as.data.frame(hot_to_r(input$hotable3))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable3 <- mytable3[1:nrow(previous3()),]
      
      for(i in 3:ncol(mytable3))
        mytable3[27,i]<-sum(mytable3[seq(1,26,2),i],na.rm=TRUE)
      
      
      for(i in 3:ncol(mytable3))
        mytable3[28,i]<-sum(mytable3[seq(2,27,2),i],na.rm=TRUE)
      
      if(month(Sys.Date()-30)!=1)
      for(i in 4:(month(Sys.Date()-30)+2))
        mytable3[28,i]<-mytable3[28,i]+mytable3[28,i-1]
      
      
      mytable3
    }
  })
  
  output$hotable3<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30))
    row_highlight = c(27,26)
    
    rhandsontable(MyChanges3(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1000,height = 600)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)),readOnly = TRUE)%>%
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
    
    write.xlsx(hot_to_r(input$hotable3),"safety/first_aid/first_aid_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  
  values4 <- reactiveValues()
  
  
  previous4<- reactive({
    read_excel("safety/accidents_countermeasure/accidents_countermeasure_2020.xlsx")
  })
  
  MyChanges4 <- reactive({
    if(is.null(input$hotable4)){return(previous4())}
    else if(!identical(previous4(),input$hotable4)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable4 <- as.data.frame(hot_to_r(input$hotable4))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable4 <- mytable4[1:nrow(previous4()),]
      
     # mytable4[1:nrow(mytable4),3:ncol(mytable4)]<-as.numeric(mytable4[1:nrow(mytable4),3:ncol(mytable4)])
      
      for(i in 3:ncol(mytable4))
        mytable4[27,i]<-sum(mytable4[seq(1,26,2),i],na.rm=TRUE)
      
      for(i in 3:ncol(mytable4))
        mytable4[28,i]<-sum(mytable4[seq(2,27,2),i],na.rm=TRUE)
      
      mytable4
    }
  })
  
  output$hotable4<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30))
    row_highlight = c(27,26)
    
    rhandsontable(MyChanges4(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1000,height = 600)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)),readOnly = TRUE)%>%
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
    
    write.xlsx(hot_to_r(input$hotable4),"safety/accidents_countermeasure/accidents_countermeasure_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  
  
  values5 <- reactiveValues()
  
  
  previous5<- reactive({
    read_excel("safety/unsafe_acts/unsafe_acts_2020.xlsx")
  })
  
  MyChanges5 <- reactive({
    if(is.null(input$hotable5)){return(previous5())}
    else if(!identical(previous5(),input$hotable5)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable5 <- as.data.frame(hot_to_r(input$hotable5))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable5 <- mytable5[1:nrow(previous5()),]
      
      # mytable4[1:nrow(mytable4),3:ncol(mytable4)]<-as.numeric(mytable4[1:nrow(mytable4),3:ncol(mytable4)])
      
      for(i in 3:ncol(mytable5))
        mytable5[27,i]<-sum(mytable5[seq(1,26,2),i],na.rm=TRUE)
      
      for(i in 3:ncol(mytable5))
        mytable5[28,i]<-sum(mytable5[seq(2,27,2),i],na.rm=TRUE)
      
      mytable5
    }
  })
  
  output$hotable5<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30))
    row_highlight = c(27,26)
    
    rhandsontable(MyChanges5(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1000,height = 600)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)),readOnly = TRUE)%>%
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
  observeEvent(input$save_safety5,{
    
    write.xlsx(hot_to_r(input$hotable5),"safety/unsafe_acts/unsafe_acts_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  output$major_plot<-renderPlot({
    
    df2<-read.xlsx("safety/major_accidents/major_accidents_2020.xlsx",sheetIndex = 1)
    
    ggplot(data=df2, aes(x=dose, y=len, fill=supp)) +
      geom_bar(stat="identity")
  })
  
  
  
  
  
 
  
  output$ibox20 <- renderInfoBox({
    
    d1<-read.xlsx("safety/major_accidents/major_accidents_2020.xlsx",sheetIndex = 1)
    x1<-input$save_safety1
    if(sum(is.na(d1$Aug))==0){
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
    
    d1<-read.xlsx("safety/minor_accidents/minor_accidents_2020.xlsx",sheetIndex = 1)
    x1<-input$save_safety2
    if(sum(is.na(d1$Aug))==0){
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
    
    d1<-read.xlsx("safety/first_aid/first_aid_2020.xlsx",sheetIndex = 1)
    x1<-input$save_safety3
    if(sum(is.na(d1$Aug))==0){
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
    
    d1<-read.xlsx("safety/accidents_countermeasure/accidents_countermeasure_2020.xlsx",sheetIndex = 1)
    x1<-input$save_safety4
    if(sum(is.na(d1$Aug))==0){
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
  
  output$ibox24 <- renderInfoBox({
    
    d1<-read.xlsx("safety/unsafe_acts/unsafe_acts_2020.xlsx",sheetIndex = 1)
    x1<-input$save_safety5
    if(sum(is.na(d1$Aug))==0){
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
      title="Unsafe Acts",
      icon=icon(ic),
      color=col
    )
  })
  output$text_major<-renderText({
    d1<-read.xlsx("safety/major_accidents/major_accidents_2020.xlsx",sheetIndex = 1)
    te<-input$save_safety1
    if(sum(is.na(d1$Aug))==0)
      return(" ")
    text<-"Major Accidents:  "
    for(i in 1:nrow(d1)){
      if(is.na(d1$Aug[i]))
       text<- paste(text,d1$Department[i],"(",d1$Category[i],")  ",sep=" ")
    }
    text
  })
  
  output$text_minor<-renderText({
    d1<-read.xlsx("safety/minor_accidents/minor_accidents_2020.xlsx",sheetIndex = 1)
    te<-input$save_safety2
    if(sum(is.na(d1$Aug))==0)
      return(" ")
    text<-"Minor Accidents:  "
    for(i in 1:nrow(d1)){
      if(is.na(d1$Aug[i]))
        text<- paste(text,d1$Department[i],"(",d1$Category[i],")  ",sep=" ")
    }
    text
  })
  
  output$text_firstaid<-renderText({
    d1<-read.xlsx("safety/first_aid/first_aid_2020.xlsx",sheetIndex = 1)
    te<-input$save_safety3
    if(sum(is.na(d1$Aug))==0)
      return(" ")
    text<-"First Aid:  "
    for(i in 1:nrow(d1)){
      if(is.na(d1$Aug[i]))
        text<- paste0(text,d1$Department[i],"(",d1$Category[i],")  ",sep=" ")
    }
    text
  })
  output$text_counter<-renderText({
    d1<-read.xlsx("safety/accidents_countermeasure/accidents_countermeasure_2020.xlsx",sheetIndex = 1)
    te<-input$save_safety4
    if(sum(is.na(d1$Aug))==0)
      return(" ")
    text<-"Accident Counter Measure:  "
    for(i in 1:nrow(d1)){
      if(is.na(d1$Aug[i]))
        text<- paste(text,d1$Department[i],"(",d1$Category[i],")  ",sep=" ")
    }
    text
  })
  output$text_unsafe<-renderText({
    d1<-read.xlsx("safety/unsafe_acts/unsafe_acts_2020.xlsx",sheetIndex = 1)
    te<-input$save_safety5
    if(sum(is.na(d1$Aug))==0)
      return(" ")
    text<-"Unsafe Acts   :  "
    for(i in 1:nrow(d1)){
      if(is.na(d1$Aug[i]))
        text<- paste(text,d1$Department[i],"(",d1$Category[i],")  ",sep=" ")
    }
    text
  })
  output$plot_major<-renderPlotly({
    
    te<-input$save_safety1
    f<-paste("safety/major_accidents/major_accidents_",input$choose_plot_year_major,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[d$Department==input$choose_plot_major,]
    da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[1:1,3:14]))))
    da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[2:2,3:14]))))
    
    l<-list.files(path="safety/major_accidents/")
    
    for(z in l){
      na<-paste("safety/major_accidents/",z,sep='')
      dt<-read_excel(na)
      dt<-dt[dt$Department==input$choose_plot_major,]
      dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[1:1,3:14]))))
      dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[2:2,3:14]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=sum(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=sum(dta$yval,na.rm = TRUE))
    }
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
  })
  
  output$table_plot_major<-renderTable({
    
    te<-input$save_safety1
    f<-paste("safety/major_accidents/major_accidents_",input$choose_plot_year_major,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[d$Department==input$choose_plot_major,]
    d
  })
  
  output$plot_minor<-renderPlotly({
    
    te<-input$save_safety2
    f<-paste("safety/minor_accidents/minor_accidents_",input$choose_plot_year_minor,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[d$Department==input$choose_plot_minor,]
    da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[1:1,3:14]))))
    da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[2:2,3:14]))))
    da1$co<-NA
    
    l<-list.files(path="safety/minor_accidents/")
    
    for(z in l){
      na<-paste("safety/minor_accidents/",z,sep='')
      dt<-read_excel(na)
      dt<-dt[dt$Department==input$choose_plot_minor,]
      dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[1:1,3:14]))))
      dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[2:2,3:14]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
  })
  
  output$table_plot_minor<-renderTable({
    
    te<-input$save_safety2
    f<-paste("safety/minor_accidents/minor_accidents_",input$choose_plot_year_minor,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[d$Department==input$choose_plot_minor,]
    d
  })
  
  output$plot_firstaid<-renderPlotly({
    
    te<-input$save_safety3
    f<-paste("safety/first_aid/first_aid_",input$choose_plot_year_firstaid,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[d$Department==input$choose_plot_firstaid,]
    da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[1:1,3:14]))))
    da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[2:2,3:14]))))
   
    
    l<-list.files(path="safety/first_aid/")
    
    for(z in l){
      na<-paste("safety/first_aid/",z,sep='')
      dt<-read_excel(na)
      dt<-dt[dt$Department==input$choose_plot_firstaid,]
      dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[1:1,3:14]))))
      dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[2:2,3:14]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
    
    
  })
  output$table_plot_firstaid<-renderTable({
    
    te<-input$save_safety3
    f<-paste("safety/first_aid/first_aid_",input$choose_plot_year_firstaid,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[d$Department==input$choose_plot_firstaid,]
    d
  })
  
  output$plot_counter<-renderPlotly({
    
    te<-input$save_safety4
    f<-paste("safety/accidents_countermeasure/accidents_countermeasure_",input$choose_plot_year_counter,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[d$Department==input$choose_plot_counter,]
    da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[1:1,3:14]))))
    da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[2:2,3:14]))))
    
    l<-list.files(path="safety/accidents_countermeasure/")
    
    for(z in l){
      na<-paste("safety/accidents_countermeasure/",z,sep='')
      dt<-read_excel(na)
      dt<-dt[dt$Department==input$choose_plot_counter,]
      dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[1:1,3:14]))))
      dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[2:2,3:14]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval<da$yval,]
    da12<-da1[da1$yval>=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
    
  })
  output$table_plot_counter<-renderTable({
    te<-input$save_safety4
    f<-paste("safety/accidents_countermeasure/accidents_countermeasure_",input$choose_plot_year_counter,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[d$Department==input$choose_plot_counter,]
    d
  })
  output$plot_unsafe<-renderPlotly({
    
    te<-input$save_safety5
    f<-paste("safety/unsafe_acts/unsafe_acts_",input$choose_plot_year_unsafe,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[d$Department==input$choose_plot_unsafe,]
    da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[1:1,3:14]))))
    da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[2:2,3:14]))))
    
    l<-list.files(path="safety/unsafe_acts/")
    
    for(z in l){
      na<-paste("safety/unsafe_acts/",z,sep='')
      dt<-read_excel(na)
      dt<-dt[dt$Department==input$choose_plot_unsafe,]
      dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[1:1,3:14]))))
      dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[2:2,3:14]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    da11<-da1[da1$yval<da$yval,]
    da12<-da1[da1$yval>=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
    
  })
  
  output$table_plot_unsafe<-renderTable({
    te<-input$save_safety5
    f<-paste("safety/unsafe_acts/unsafe_acts_",input$choose_plot_year_unsafe,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[d$Department==input$choose_plot_unsafe,]
    d
  })
  
  
 
  #5 unsafe
  
  data_unsafe_act<-reactive({
    te<-input$save_safety5
    d <- read_excel("safety/unsafe_acts/unsafe_acts_2020.xlsx")
    d<-d[d$Department!="Plant level",]
    da<-d[d$Category=='Actual',]
    dt<-d[d$Category=='Target',]
    
    da<-da%>%gather(month,value,3:14)
    da$Category<-NULL
    da
  })
  data_unsafe_tar<-reactive({
    te<-input$save_safety5
    d <- read_excel("safety/unsafe_acts/unsafe_acts_2020.xlsx")
    d<-d[d$Department!="Plant level",]
    da<-d[d$Category=='Actual',]
    dt<-d[d$Category=='Target',]
    
    dt<-dt%>%gather(month,value,3:14)
    dt$Category<-NULL
    dt
  })
  data_dept_unsafe<-reactive({
    te<-input$save_safety5
    d <- read_excel("safety/unsafe_acts/unsafe_acts_2020.xlsx")
    d<-d[d$Department!="Plant level",]
    d<-d%>%gather(month,value,3:14)
    d$month<-as.yearmon(d$month,"%b %Y")
    d<-d[months(d$month)==input$choose_comp_unsafe,]
    d$Department<-factor(d$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
    
    d<-spread(d,Department,value )
    
    d$Category<-factor(d$Category,levels=c("Target","Actual"))
    d$month<-months(d$month)
     d
  })
  
  output$comp_unsafe<-renderPlotly({
    te<-input$save_safety5
    da<-data_unsafe_act()
    dt<-data_unsafe_tar()
    
    da$month<-as.yearmon(da$month,"%b %Y")
    dt$month<-as.yearmon(dt$month,"%b %Y")
    
    da1<-da[months(da$month)==input$choose_comp_unsafe,]
    da<-dt[months(dt$month)==input$choose_comp_unsafe,]
    
    da$Department<-factor(da$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
    da1$Department<-factor(da1$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
    
    da$yval<-as.numeric(da$value)
    da1$yval<-as.numeric(da1$value)
    
    da11<-da1[da1$value<da$value,]
    da12<-da1[da1$value>=da$value,]
    
    #if(nrow(da11)==0)
    #  da11<-da11%>%add_row(xval=na,yval=0)
    #if(nrow(da12)==0)
    #  da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$Department,y = da$value, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$Department,y=da11$value,text=round(da11$yval,digits=1), textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$Department,y=da12$value,text=round(da12$yval,digits=1), textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
  })
  
  output$table_comp_unsafe<-renderTable({
    data_dept_unsafe()
  })
  
  
 #4 counter
  
  data_counter_act<-reactive({
    te<-input$save_safety4
    d<-read_excel("safety/accidents_countermeasure/accidents_countermeasure_2020.xlsx")
    d<-d[d$Department!="Plant level",]
    da<-d[d$Category=='Actual',]
    dt<-d[d$Category=='Target',]
    
    da<-da%>%gather(month,value,3:14)
    da$Category<-NULL
    da
  })
  data_counter_tar<-reactive({
    te<-input$save_safety4
    d<-read_excel("safety/accidents_countermeasure/accidents_countermeasure_2020.xlsx")
    d<-d[d$Department!="Plant level",]
    da<-d[d$Category=='Actual',]
    dt<-d[d$Category=='Target',]
    
    dt<-dt%>%gather(month,value,3:14)
    dt$Category<-NULL
    dt
  })
  data_dept_counter<-reactive({
    te<-input$save_safety4
    d<-read_excel("safety/accidents_countermeasure/accidents_countermeasure_2020.xlsx")
    d<-d[d$Department!="Plant level",]
    
    d<-d%>%gather(month,value,3:14)
    d$month<-as.yearmon(d$month,"%b %Y")
    d<-d[months(d$month)==input$choose_comp_counter,]
    d$Department<-factor(d$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
    
    d<-spread(d,Department,value )
    
    d$Category<-factor(d$Category,levels=c("Target","Actual"))
    d$month<-months(d$month)
    d
  })
  
  output$comp_counter<-renderPlotly({
    te<-input$save_safety4
    da<-data_counter_act()
    dt<-data_counter_tar()
    
    da$month<-as.yearmon(da$month,"%b %Y")
    dt$month<-as.yearmon(dt$month,"%b %Y")
    
    da1<-da[months(da$month)==input$choose_comp_counter,]
    da<-dt[months(dt$month)==input$choose_comp_counter,]
    
    da$Department<-factor(da$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
    da1$Department<-factor(da1$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
    
    da$yval<-as.numeric(da$value)
    da1$yval<-as.numeric(da1$value)
    
    da11<-da1[da1$value<da$value,]
    da12<-da1[da1$value>=da$value,]
    
    #if(nrow(da11)==0)
    #  da11<-da11%>%add_row(xval=na,yval=0)
    #if(nrow(da12)==0)
    #  da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$Department,y = da$value, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$Department,y=da11$value,text=round(da11$yval,digits=1), textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$Department,y=da12$value,text=round(da12$yval,digits=1), textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
  })
  
  output$table_comp_counter<-renderTable({
    data_dept_counter()
  })
  
  #3 firstaid
  
  data_firstaid_act<-reactive({
    te<-input$save_safety3
    d<-read_excel("safety/first_aid/first_aid_2020.xlsx")
    d<-d[d$Department!="Plant level",]
    da<-d[d$Category=='Actual',]
    dt<-d[d$Category=='Target',]
    
    da<-da%>%gather(month,value,3:14)
    da$Category<-NULL
    da
  })
  data_firstaid_tar<-reactive({
    te<-input$save_safety5
    d<-read_excel("safety/first_aid/first_aid_2020.xlsx")
    d<-d[d$Department!="Plant level",]
    da<-d[d$Category=='Actual',]
    dt<-d[d$Category=='Target',]
    
    dt<-dt%>%gather(month,value,3:14)
    dt$Category<-NULL
    dt
  })
  data_dept_firstaid<-reactive({
    te<-input$save_safety3
    d<-read_excel("safety/first_aid/first_aid_2020.xlsx")
    d<-d[d$Department!="Plant level",]
    d<-d%>%gather(month,value,3:14)
    d$month<-as.yearmon(d$month,"%b %Y")
    d<-d[months(d$month)==input$choose_comp_firstaid,]
    d$Department<-factor(d$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
    
    d<-spread(d,Department,value )
    
    d$Category<-factor(d$Category,levels=c("Target","Actual"))
    d$month<-months(d$month)
    d
  })
  
  output$comp_firstaid<-renderPlotly({
    te<-input$save_safety3
    da<-data_firstaid_act()
    dt<-data_firstaid_tar()
    
    da$month<-as.yearmon(da$month,"%b %Y")
    dt$month<-as.yearmon(dt$month,"%b %Y")
    
    da1<-da[months(da$month)==input$choose_comp_firstaid,]
    da<-dt[months(dt$month)==input$choose_comp_firstaid,]
    
    da$Department<-factor(da$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
    da1$Department<-factor(da1$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
    
    da$yval<-as.numeric(da$value)
    da1$yval<-as.numeric(da1$value)
    
    da11<-da1[da1$value<da$value,]
    da12<-da1[da1$value>=da$value,]
    
    #if(nrow(da11)==0)
    #  da11<-da11%>%add_row(xval=na,yval=0)
    #if(nrow(da12)==0)
    #  da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$Department,y = da$value, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$Department,y=da11$value,text=round(da11$yval,digits=1), textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$Department,y=da12$value,text=round(da12$yval,digits=1), textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
  })
  
  output$table_comp_firstaid<-renderTable({
    data_dept_firstaid()
  })
  
  #2 minor
  
  data_minor_act<-reactive({
    te<-input$save_safety2
    d<-read_excel("safety/minor_accidents/minor_accidents_2020.xlsx")
    d<-d[d$Department!="Plant level",]
    da<-d[d$Category=='Actual',]
    dt<-d[d$Category=='Target',]
    
    da<-da%>%gather(month,value,3:14)
    da$Category<-NULL
    da
  })
  data_minor_tar<-reactive({
    te<-input$save_safety2
    d<-read_excel("safety/minor_accidents/minor_accidents_2020.xlsx")
    d<-d[d$Department!="Plant level",]
    da<-d[d$Category=='Actual',]
    dt<-d[d$Category=='Target',]
    
    dt<-dt%>%gather(month,value,3:14)
    dt$Category<-NULL
    dt
  })
  data_dept_minor<-reactive({
    te<-input$save_safety2
    d<-read_excel("safety/minor_accidents/minor_accidents_2020.xlsx")
    d<-d[d$Department!="Plant level",]
    d<-d%>%gather(month,value,3:14)
    d$month<-as.yearmon(d$month,"%b %Y")
    d<-d[months(d$month)==input$choose_comp_minor,]
    d$Department<-factor(d$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
    
    d<-spread(d,Department,value )
    
    d$Category<-factor(d$Category,levels=c("Target","Actual"))
    d$month<-months(d$month)
    d
  })
  
  output$comp_minor<-renderPlotly({
    te<-input$save_safety2
    da<-data_minor_act()
    dt<-data_minor_tar()
    
    da$month<-as.yearmon(da$month,"%b %Y")
    dt$month<-as.yearmon(dt$month,"%b %Y")
    
    da1<-da[months(da$month)==input$choose_comp_minor,]
    da<-dt[months(dt$month)==input$choose_comp_minor,]
    
    da$Department<-factor(da$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
    da1$Department<-factor(da1$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
    
    da$yval<-as.numeric(da$value)
    da1$yval<-as.numeric(da1$value)
    
    da11<-da1[da1$value<da$value,]
    da12<-da1[da1$value>=da$value,]
    
    #if(nrow(da11)==0)
    #  da11<-da11%>%add_row(xval=na,yval=0)
    #if(nrow(da12)==0)
    #  da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$Department,y = da$value, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$Department,y=da11$value,text=round(da11$yval,digits=1), textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$Department,y=da12$value,text=round(da12$yval,digits=1), textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
  })
  
  output$table_comp_minor<-renderTable({
    data_dept_minor()
  })
  
  #1 major
  
  data_major_act<-reactive({
    te<-input$save_safety1
    d<-read_excel("safety/major_accidents/major_accidents_2020.xlsx")
    d<-d[d$Department!="Plant level",]
    da<-d[d$Category=='Actual',]
    dt<-d[d$Category=='Target',]
    
    da<-da%>%gather(month,value,3:14)
    da$Category<-NULL
    da
  })
  data_major_tar<-reactive({
    te<-input$save_safety1
    d<-read_excel("safety/major_accidents/major_accidents_2020.xlsx")
    d<-d[d$Department!="Plant level",]
    da<-d[d$Category=='Actual',]
    dt<-d[d$Category=='Target',]
    
    dt<-dt%>%gather(month,value,3:14)
    dt$Category<-NULL
    dt
  })
  data_dept_major<-reactive({
    te<-input$save_safety1
    d<-read_excel("safety/major_accidents/major_accidents_2020.xlsx")
    d<-d[d$Department!="Plant level",]
    d<-d%>%gather(month,value,3:14)
    d$month<-as.yearmon(d$month,"%b %Y")
    d<-d[months(d$month)==input$choose_comp_major,]
    d$Department<-factor(d$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
    
    d<-spread(d,Department,value )
    
    d$Category<-factor(d$Category,levels=c("Target","Actual"))
    d$month<-months(d$month)
    d
  })
  
  output$comp_major<-renderPlotly({
    te<-input$save_safety1
    da<-data_major_act()
    dt<-data_major_tar()
    
    da$month<-as.yearmon(da$month,"%b %Y")
    dt$month<-as.yearmon(dt$month,"%b %Y")
    
    da1<-da[months(da$month)==input$choose_comp_major,]
    da<-dt[months(dt$month)==input$choose_comp_major,]
    
    da$Department<-factor(da$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
    da1$Department<-factor(da1$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
    
    da$yval<-as.numeric(da$value)
    da1$yval<-as.numeric(da1$value)
    
    da11<-da1[da1$value<da$value,]
    da12<-da1[da1$value>=da$value,]
    
 
    
    p<-plot_ly(x=da$Department,y = da$value, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$Department,y=da11$value,text=round(da11$yval,digits=1), textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$Department,y=da12$value,text=round(da12$yval,digits=1), textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
  })
  
  output$table_comp_major<-renderTable({
    data_dept_major()
  })
  
  
  
  
  
  
  
  #QM data
  
  values_qm_kpi <- reactiveValues()
  
  
  previous_qm_kpi <- reactive({
    d<-read_excel("QM/QM/QM_2020.xlsx",sheet="kpi")
    #d<-d[d$`KPI's`=='Quality',]
    d
  })
  
  MyChanges_qm_kpi <- reactive({
    if(is.null(input$hotable_qm_kpi)){return(previous_qm_kpi())}
    else if(!identical(previous_qm_kpi(),input$hotable_qm_kpi)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable_qm_kpi <- as.data.frame(hot_to_r(input$hotable_qm_kpi))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_qm_kpi <- mytable_qm_kpi[1:nrow(previous_qm_kpi()),]
      
      for(i in 4:ncol(mytable_qm_kpi))
        mytable_qm_kpi[4,i]<-mytable_qm_kpi[2,i]/mytable_qm_kpi[1,i]
      
      for(i in 4:ncol(mytable_qm_kpi))
        mytable_qm_kpi[55,i]<-sum(mytable_qm_kpi[seq(45,48),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_qm_kpi))
        mytable_qm_kpi[56,i]<-sum(mytable_qm_kpi[seq(45,54),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_qm_kpi))
        mytable_qm_kpi[58,i]<-mytable_qm_kpi[56,i]/mytable_qm_kpi[57,i]
      
      
      for(i in 4:ncol(mytable_qm_kpi))
        mytable_qm_kpi[61,i]<-mytable_qm_kpi[60,i]/(mytable_qm_kpi[66,i]+mytable_qm_kpi[67,i])
      
      for(i in 4:ncol(mytable_qm_kpi))
        mytable_qm_kpi[64,i]<-mytable_qm_kpi[63,i]/(mytable_qm_kpi[68,i])
      
      for(i in 4:ncol(mytable_qm_kpi))
        mytable_qm_kpi[69,i]<-sum(mytable_qm_kpi[c(66,67,68),i],na.rm = TRUE)
      
      mytable_qm_kpi
    }
  })
  
  output$hotable_qm_kpi<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30)+1)
    row_highlight = c(4,55,56,58,61,64,69)-1
    row_readonly=c(4,55,56,58,61,64,69)
    
    rhandsontable(MyChanges_qm_kpi(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1300,height = 600)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+2),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_qm_kpi,{
    #
    write.xlsx(hot_to_r(input$hotable_qm_kpi),"QM/QM/QM_2020.xlsx",sheetName = "kpi",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #QM DPU @ QFL4
  
  values_qm_dpu_qfl4 <- reactiveValues()
  
  
  previous_qm_dpu_qfl4 <- reactive({
    d<-read_excel("QM/QM_DPU_QFL4/QM_DPU_QFL4_2020.xlsx")
    #d<-d[d$`KPI's`=='Quality',]
    d
  })
  
  MyChanges_qm_dpu_qfl4 <- reactive({
    if(is.null(input$hotable_qm_dpu_qfl4)){return(previous_qm_dpu_qfl4())}
    else if(!identical(previous_qm_dpu_qfl4(),input$hotable_qm_dpu_qfl4)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable_qm_dpu_qfl4 <- as.data.frame(hot_to_r(input$hotable_qm_dpu_qfl4))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_qm_dpu_qfl4 <- mytable_qm_dpu_qfl4[1:nrow(previous_qm_dpu_qfl4()),]
      
      #for(i in 4:ncol(mytable_qm_dpu_qfl4))
      #  mytable_qm_dpu_qfl4[4,i]<-mytable_qm_dpu_qfl4[2,i]/mytable_qm_dpu_qfl4[1,i]
      
      
      
      mytable_qm_dpu_qfl4
    }
  })
  
  output$hotable_qm_dpu_qfl4<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30)+1)
    row_highlight = NA
    row_readonly=NA
    
    rhandsontable(MyChanges_qm_dpu_qfl4(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1250,height = 600)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+2),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_qm_dpu_qfl4,{
    #
    write.xlsx(hot_to_r(input$hotable_qm_dpu_qfl4),"QM/QM_DPU_QFL4/QM_DPU_QFL4_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
 
  
  #QM DPU @ QFL4- Ops related
  
  values_qm_dpu_qfl4_ops <- reactiveValues()
  
  
  previous_qm_dpu_qfl4_ops <- reactive({
    d<-read_excel("QM/QM_DPU_QFL4_OPS/QM_DPU_QFL4_OPS_2020.xlsx")
    #d<-d[d$`KPI's`=='Quality',]
    d
  })
  
  MyChanges_qm_dpu_qfl4_ops <- reactive({
    if(is.null(input$hotable_qm_dpu_qfl4_ops)){return(previous_qm_dpu_qfl4_ops())}
    else if(!identical(previous_qm_dpu_qfl4_ops(),input$hotable_qm_dpu_qfl4_ops)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable_qm_dpu_qfl4_ops <- as.data.frame(hot_to_r(input$hotable_qm_dpu_qfl4_ops))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_qm_dpu_qfl4_ops <- mytable_qm_dpu_qfl4_ops[1:nrow(previous_qm_dpu_qfl4_ops()),]
      
      #for(i in 4:ncol(mytable_qm_dpu_qfl4))
      #  mytable_qm_dpu_qfl4[4,i]<-mytable_qm_dpu_qfl4[2,i]/mytable_qm_dpu_qfl4[1,i]
      
      
      
      mytable_qm_dpu_qfl4_ops
    }
  })
  
  output$hotable_qm_dpu_qfl4_ops<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30)+1)
    row_highlight = NA
    row_readonly=NA
    
    rhandsontable(MyChanges_qm_dpu_qfl4_ops(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1250,height = 600)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+2),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_qm_dpu_qfl4_ops,{
    #
    write.xlsx(hot_to_r(input$hotable_qm_dpu_qfl4_ops),"QM/QM_DPU_QFL4_OPS/QM_DPU_QFL4_OPS_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  
  
  output$plot_qm_hdt_dpu_qfl4<-renderPlotly({
    
    te<-input$save_qm_dpu_qfl4
    f<-paste("QM/QM_DPU_QFL4/QM_DPU_QFL4_",input$choose_plot_year_qm_hdt_dpu_qfl4,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[1:1,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))))
    
    l<-list.files(path="QM/QM_DPU_QFL4/")
    
    for(z in l){
      na<-paste("QM/QM_DPU_QFL4/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[1:1,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[2:2,4:15]))))
      
      na<-strsplit(z,"_")[[1]][4]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="HDT DPU @ QFL4 (Overall)")
    p
  })

  output$table_plot_qm_hdt_dpu_qfl4<-renderTable({
    
    te<-input$save_qm_dpu_qfl4
    f<-paste("QM/QM_DPU_QFL4/QM_DPU_QFL4_",input$choose_plot_year_qm_hdt_dpu_qfl4,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[1:4,1:15]
    d
  })
  
  output$plot_qm_hdt_dpu_qfl4_ab<-renderPlotly({
    
    te<-input$save_qm_dpu_qfl4
    f<-paste("QM/QM_DPU_QFL4/QM_DPU_QFL4_",input$choose_plot_year_qm_hdt_dpu_qfl4,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[3:3,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[4:4,4:15]))))
    
    l<-list.files(path="QM/QM_DPU_QFL4/")
    
    for(z in l){
      na<-paste("QM/QM_DPU_QFL4/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[3:3,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[4:4,4:15]))))
      
      na<-strsplit(z,"_")[[1]][4]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="HDT DPU @ QFL4 [A+B]")
    p
  })
  
 
  
  
  # MDT
  output$plot_qm_mdt_dpu_qfl4<-renderPlotly({
    
    te<-input$save_qm_dpu_qfl4
    f<-paste("QM/QM_DPU_QFL4/QM_DPU_QFL4_",input$choose_plot_year_qm_mdt_dpu_qfl4,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[5:5,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[6:6,4:15]))))
    
    l<-list.files(path="QM/QM_DPU_QFL4/")
    
    for(z in l){
      na<-paste("QM/QM_DPU_QFL4/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[5:5,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[6:6,4:15]))))
      
      na<-strsplit(z,"_")[[1]][4]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="MDT DPU @ QFL4")
    p
  })
  
  output$table_plot_qm_mdt_dpu_qfl4<-renderTable({
    
    te<-input$save_qm_dpu_qfl4
    f<-paste("QM/QM_DPU_QFL4/QM_DPU_QFL4_",input$choose_plot_year_qm_mdt_dpu_qfl4,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[5:8,1:15]
    d
  })
  
  output$plot_qm_mdt_dpu_qfl4_ab<-renderPlotly({
    
    te<-input$save_qm_dpu_qfl4
    f<-paste("QM/QM_DPU_QFL4/QM_DPU_QFL4_",input$choose_plot_year_qm_mdt_dpu_qfl4,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[7:7,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[8:8,4:15]))))
    
    l<-list.files(path="QM/QM_DPU_QFL4/")
    
    for(z in l){
      na<-paste("QM/QM_DPU_QFL4/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[7:7,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[8:8,4:15]))))
      
      na<-strsplit(z,"_")[[1]][4]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
        title="MDT DPU @ QFL4 [A+B]")
    p
  })

  
  output$plot_qm_qfl4_eng<-renderPlotly({
    
    te<-input$save_qm_dpu_qfl4
    f<-paste("QM/QM/QM_",input$choose_plot_year_qm_qfl4_eng,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[15:15,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[16:16,4:15]))))
    
    l<-list.files(path="QM/QM/")
    
    for(z in l){
      na<-paste("QM/QM/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[15:15,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[16:16,4:15]))))
      
      na<-strsplit(z,"_")[[1]][2]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Engine DPU @ Teardown audit")
    p
  })
 
  output$table_plot_qm_qfl4_eng<-renderTable({
    
    te<-input$save_qm_dpu_qfl4
    f<-paste("QM/QM/QM_",input$choose_plot_year_qm_qfl4_eng,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[15:18,1:15]
    d
  })
  
  
  output$plot_qm_qfl4_eng_ab<-renderPlotly({
    
    te<-input$save_qm_dpu_qfl4
    f<-paste("QM/QM/QM_",input$choose_plot_year_qm_qfl4_eng,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[17:17,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[18:18,4:15]))))
    
    l<-list.files(path="QM/QM/")
    
    for(z in l){
      na<-paste("QM/QM/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[17:17,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[18:18,4:15]))))
      
      na<-strsplit(z,"_")[[1]][2]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Engine DPU @ Teardown audit [A+B]")
    p
  })
  
 
  
  
  output$plot_qm_qfl4_tra<-renderPlotly({
    
    te<-input$save_qm_dpu_qfl4
    f<-paste("QM/QM/QM_",input$choose_plot_year_qm_qfl4_tra,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[19:19,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[20:20,4:15]))))
    
    l<-list.files(path="QM/QM/")
    
    for(z in l){
      na<-paste("QM/QM/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[13:13,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[14:14,4:15]))))
      
      na<-strsplit(z,"_")[[1]][2]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Transmission DPU @ Teardown audit")
    p
  })
  
  
  
  output$plot_qm_qfl4_tra_ab<-renderPlotly({
    
    te<-input$save_qm_dpu_qfl4
    f<-paste("QM/QM/QM_",input$choose_plot_year_qm_qfl4_tra,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[21:21,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[22:22,4:15]))))
    
    l<-list.files(path="QM/QM/")
    
    for(z in l){
      na<-paste("QM/QM/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[21:21,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[22:22,4:15]))))
      
      na<-strsplit(z,"_")[[1]][2]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Transmission DPU @ Teardown audit [A+B]")
    p
  })
  
  output$table_plot_qm_qfl4_tra<-renderTable({
    
    te<-input$save_qm_dpu_qfl4
    f<-paste("QM/QM/QM_",input$choose_plot_year_qm_qfl4_tra,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[19:22,1:15]
    d
  })
  
  
  #DPU @ QFL4: HDT (Overall) - Ops related
  
  
  output$plot_qm_qfl4_hdt_ops<-renderPlotly({
    
    te<-input$save_qm_dpu_qfl4_ops
    f<-paste("QM/QM_DPU_QFL4_OPS/QM_DPU_QFL4_OPS_",input$choose_plot_year_qm_qfl4_hdt_ops,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[1:1,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))))
    
    l<-list.files(path="QM/QM_DPU_QFL4_OPS/")
    
    for(z in l){
      na<-paste("QM/QM_DPU_QFL4_OPS/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[1:1,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[2:2,4:15]))))
      
      na<-strsplit(z,"_")[[1]][5]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="DPU @ QFL4: HDT (Overall) - Ops related")
    p
  })
  
  
  #DPU @ QFL4: HDT (A+B) - Ops related
  
  
  output$plot_qm_qfl4_hdt_ops_ab<-renderPlotly({
    
    te<-input$save_qm_dpu_qfl4_ops
    f<-paste("QM/QM_DPU_QFL4_OPS/QM_DPU_QFL4_OPS_",input$choose_plot_year_qm_qfl4_hdt_ops_ab,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[3:3,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[4:4,4:15]))))
    
    l<-list.files(path="QM/QM_DPU_QFL4_OPS/")
    
    for(z in l){
      na<-paste("QM/QM_DPU_QFL4_OPS/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[3:3,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[4:4,4:15]))))
      
      na<-strsplit(z,"_")[[1]][5]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="DPU @ QFL4: HDT (A+B) - Ops related")
    p
  })
  
  output$table_plot_qm_qfl4_hdt_ops<-renderTable({
    
    te<-input$save_qm_dpu_qfl4_ops
    f<-paste("QM/QM_DPU_QFL4_OPS/QM_DPU_QFL4_OPS_",input$choose_plot_year_qm_qfl4_hdt_ops,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[1:2,1:15]
    d
  })
  output$table_plot_qm_qfl4_hdt_ops_ab<-renderTable({
    
    te<-input$save_qm_dpu_qfl4_ops
    f<-paste("QM/QM_DPU_QFL4_OPS/QM_DPU_QFL4_OPS_",input$choose_plot_year_qm_qfl4_hdt_ops_ab,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[3:4,1:15]
    d
  })
  
  
  #DPU @ QFL4: MDT (Overall) - Ops related
  
  
  output$plot_qm_qfl4_mdt_ops<-renderPlotly({
    
    te<-input$save_qm_dpu_qfl4_ops
    f<-paste("QM/QM_DPU_QFL4_OPS/QM_DPU_QFL4_OPS_",input$choose_plot_year_qm_qfl4_mdt_ops,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[5:5,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[6:6,4:15]))))
    
    l<-list.files(path="QM/QM_DPU_QFL4_OPS/")
    
    for(z in l){
      na<-paste("QM/QM_DPU_QFL4_OPS/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[5:5,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[6:6,4:15]))))
      
      na<-strsplit(z,"_")[[1]][5]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="DPU @ QFL4: MDT (Overall) - Ops related")
    p
  })
  
  
  #DPU @ QFL4: MDT (A+B) - Ops related
  
  
  output$plot_qm_qfl4_mdt_ops_ab<-renderPlotly({
    
    te<-input$save_qm_dpu_qfl4_ops
    f<-paste("QM/QM_DPU_QFL4_OPS/QM_DPU_QFL4_OPS_",input$choose_plot_year_qm_qfl4_mdt_ops_ab,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[7:7,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[8:8,4:15]))))
    
    l<-list.files(path="QM/QM_DPU_QFL4_OPS/")
    
    for(z in l){
      na<-paste("QM/QM_DPU_QFL4_OPS/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[7:7,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[8:8,4:15]))))
      
      na<-strsplit(z,"_")[[1]][5]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="DPU @ QFL4: MDT (A+B) - Ops related")
    p
  })
  
  output$table_plot_qm_qfl4_mdt_ops<-renderTable({
    
    te<-input$save_qm_dpu_qfl4_ops
    f<-paste("QM/QM_DPU_QFL4_OPS/QM_DPU_QFL4_OPS_",input$choose_plot_year_qm_qfl4_mdt_ops,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[5:6,1:15]
    d
  })
  
  output$table_plot_qm_qfl4_mdt_ops_ab<-renderTable({
    
    te<-input$save_qm_dpu_qfl4_ops
    f<-paste("QM/QM_DPU_QFL4_OPS/QM_DPU_QFL4_OPS_",input$choose_plot_year_qm_qfl4_mdt_ops_ab,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[7:8,1:15]
    d
  })
  
  #Vehicle DPU @QFL2 HDT
  
  
  output$plot_qm_qfl2_hdt<-renderPlotly({
    
    te<-input$save_qm_kpi
    
    
    f<-paste("Chassis/KPI/chassis_kpi_",input$choose_plot_year_qm_qfl2_hdt,".xlsx",sep="")
    d <- read_excel(f)
    
    d<-d[d$Description=="DPU @ QFL2 - HDT",]
    
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))),tar=c(t(array(d[1:1,4:15]))),co=c(t(array(d[2:2,4:15]+d[3:3,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[3:3,4:15]))),tar=c(t(array(d[1:1,4:15]))),co=c(t(array(d[2:2,4:15]+d[3:3,4:15]))))
    da$te<-"Veh_ass"
    da1$te<-"Scratches"
    da<-rbind(da,da1)
    
    da<-da%>%add_row(xval="2018",yval=2.87,tar=18,co=18,te="Veh_ass")
    da<-da%>%add_row(xval="2018",yval=15.13,tar=18,co=18,te="Scratches")
    
    da<-da%>%add_row(xval="2019",yval=1.61,tar=12,co=12,te="Veh_ass")
    da<-da%>%add_row(xval="2019",yval=9.03,tar=12,co=12,te="Scratches")
    
    da<-da%>%add_row(xval="2020",yval=rowMeans(d[2:2,4:15],na.rm=TRUE),tar=rowMeans(d[1:1,4:15],na.rm=TRUE),co=rowMeans(d[2:2,4:15],na.rm=TRUE)+rowMeans(d[3:3,4:15],na.rm=TRUE),te="Veh_ass")
    da<-da%>%add_row(xval="2020",yval=rowMeans(d[3:3,4:15],na.rm=TRUE),tar=rowMeans(d[1:1,4:15],na.rm=TRUE),co=rowMeans(d[2:2,4:15],na.rm=TRUE)+rowMeans(d[3:3,4:15],na.rm=TRUE),te="Scratches")
    
    
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=0,tar=12,co=0,te="Veh_ass")
    da<-da%>%add_row(xval=na,yval=0,tar=12,co=0,te="Scratches")
    
    
    
    for(i in 1:nrow(da))
      if(is.na(da$yval[i]))
        da$yval[i]<-0
    
    for(i in 1:nrow(da))
      if(is.na(da$co[i]))
        da$co[i]<-0
    
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$yval<-as.numeric(da$yval)
    
    da<-da[order(da$xval),]
    da$zval<-NA
    for(i in 1:nrow(da)){
      if(da$te[i]=="Scratches")
        da$zval[i]<-"grey"
      else if(da$co[i]>da$tar[i])
        da$zval[i]<-"red"
      else
        da$zval[i]<-"green4"
    }
    
    p<-ggplot(data=da, aes(x=xval, y=yval, fill=te)) +
      geom_bar(stat="identity",fill=da$zval)+
      geom_line(aes(y=tar,group = 1),color='black')+
      geom_text(aes(label=round(yval,2)), vjust=-1.5, color="white", size=2.5)+
      # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
      theme(axis.text.x = element_text(angle = 45, hjust = 1))+
      labs(x="",y="")
    p
  })
  
  output$table_plot_qm_qfl2_hdt<-renderTable({
    
    te<-input$save_qm_kpi
    f<-paste("Chassis/KPI/chassis_kpi_",input$choose_plot_year_qm_qfl2_mdt,".xlsx",sep="")
    d <- read_excel(f)
    
    d<-d[d$Description=="DPU @ QFL2 - HDT",]
    d<-d[1:3,1:15]
    d
  })
  
  
  #Vehicle DPU @QFL2 MDT
  
  
  output$plot_qm_qfl2_mdt<-renderPlotly({
    
    te<-input$save_qm_kpi
    
    f<-paste("Chassis/KPI/chassis_kpi_",input$choose_plot_year_qm_qfl2_mdt,".xlsx",sep="")
    d <- read_excel(f)
    
    d<-d[d$Description=="DPU @ QFL2 - MDT",]
    
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))),tar=c(t(array(d[1:1,4:15]))),co=c(t(array(d[2:2,4:15]+d[3:3,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[3:3,4:15]))),tar=c(t(array(d[1:1,4:15]))),co=c(t(array(d[2:2,4:15]+d[3:3,4:15]))))
    da$te<-"Veh_ass"
    da1$te<-"Scratches"
    da<-rbind(da,da1)
    
    da<-da%>%add_row(xval="2018",yval=1.74,tar=18,co=12.87,te="Veh_ass")
    da<-da%>%add_row(xval="2018",yval=11.13,tar=18,co=12.87,te="Scratches")
    
    da<-da%>%add_row(xval="2019",yval=2.38,tar=12,co=11.43,te="Veh_ass")
    da<-da%>%add_row(xval="2019",yval=9.05,tar=12,co=11.43,te="Scratches")
    
    da<-da%>%add_row(xval="2020",yval=rowMeans(d[2:2,4:15],na.rm=TRUE),tar=rowMeans(d[1:1,4:15],na.rm=TRUE),co=rowMeans(d[2:2,4:15],na.rm=TRUE)+rowMeans(d[3:3,4:15],na.rm=TRUE),te="Veh_ass")
    da<-da%>%add_row(xval="2020",yval=rowMeans(d[3:3,4:15],na.rm=TRUE),tar=rowMeans(d[1:1,4:15],na.rm=TRUE),co=rowMeans(d[2:2,4:15],na.rm=TRUE)+rowMeans(d[3:3,4:15],na.rm=TRUE),te="Scratches")
    
    
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=0,tar=12,co=0,te="Veh_ass")
    da<-da%>%add_row(xval=na,yval=0,tar=12,co=0,te="Scratches")
    
    
    
    for(i in 1:nrow(da))
      if(is.na(da$yval[i]))
        da$yval[i]<-0
    for(i in 1:nrow(da))
      if(is.na(da$co[i]))
        da$co[i]<-0
    
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$yval<-as.numeric(da$yval)
    
    da<-da[order(da$xval),]
    da$zval<-NA
    for(i in 1:nrow(da)){
      if(da$te[i]=="Scratches")
        da$zval[i]<-"grey"
      else if(da$co[i]>da$tar[i])
        da$zval[i]<-"red"
      else
        da$zval[i]<-"green4"
    }
    
    p<-ggplot(data=da, aes(x=xval, y=yval, fill=te)) +
      geom_bar(stat="identity",fill=da$zval)+
      geom_line(aes(y=tar,group = 1),color='black')+
      geom_text(aes(label=round(yval,2)), vjust=1.5, color="white", size=2.5)+
      # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
      theme(axis.text.x = element_text(angle = 45, hjust = 1))+
      labs(x="",y="")
    p
  })
  output$table_plot_qm_qfl2_mdt<-renderTable({
    
    te<-input$save_qm_kpi
    f<-paste("Chassis/KPI/chassis_kpi_",input$choose_plot_year_qm_qfl2_mdt,".xlsx",sep="")
    d <- read_excel(f)
    
    d<-d[d$Description=="DPU @ QFL2 - MDT",]
    d<-d[1:3,1:15]
    d
  })
  
  #Engine DPU @ QFL2
  
  
  output$plot_qm_qfl2_eng<-renderPlotly({
    
    te<-input$save_qm_kpi
    f<-paste("QM/QM/QM_",input$choose_plot_year_qm_qfl2_eng,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[35:35,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[36:36,4:15]))))
    
    l<-list.files(path="QM/QM/")
    
    for(z in l){
      na<-paste("QM/QM/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[35:35,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[36:36,4:15]))))
      
      na<-strsplit(z,"_")[[1]][2]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Engine DPU @ QFL2")
    p
  })
  output$table_plot_qm_qfl2_eng<-renderTable({
    
    te<-input$save_qm_kpi
    f<-paste("QM/QM/QM_",input$choose_plot_year_qm_qfl2_eng,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[35:36,1:15]
    d
  })
  
  #Transmission: DPU @ QFL2
  
  
  output$plot_qm_qfl2_tra<-renderPlotly({
    
    te<-input$save_qm_kpi
    f<-paste("QM/QM/QM_",input$choose_plot_year_qm_qfl2_tra,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[37:37,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[38:38,4:15]))))
    
    l<-list.files(path="QM/QM/")
    
    for(z in l){
      na<-paste("QM/QM/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[37:37,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[38:38,4:15]))))
      
      na<-strsplit(z,"_")[[1]][2]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Transmission: DPU @ QFL2")
    p
  })
  output$table_plot_qm_qfl2_tra<-renderTable({
    
    te<-input$save_qm_kpi
    f<-paste("QM/QM/QM_",input$choose_plot_year_qm_qfl2_tra,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[37:38,1:15]
    d
  })
  
  
  #Paint DPU @ QFL2
  
  
  output$plot_qm_qfl2_pai<-renderPlotly({
    
    te<-input$save_qm_kpi
    f<-paste("QM/QM/QM_",input$choose_plot_year_qm_qfl2_pai,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[39:39,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[40:40,4:15]))))
    
    l<-list.files(path="QM/QM/")
    
    for(z in l){
      na<-paste("QM/QM/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[39:39,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[40:40,4:15]))))
      
      na<-strsplit(z,"_")[[1]][2]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Paint DPU @ QFL2")
    p
  })
  
  output$table_plot_qm_qfl2_pai<-renderTable({
    
    te<-input$save_qm_kpi
    f<-paste("QM/QM/QM_",input$choose_plot_year_qm_qfl2_pai,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[39:40,1:15]
    d
  })
  
  #CiW DPU @ QFL2
  
  
  output$plot_qm_qfl2_ciw<-renderPlotly({
    
    te<-input$save_qm_kpi
    f<-paste("QM/QM/QM_",input$choose_plot_year_qm_qfl2_ciw,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[41:41,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[42:42,4:15]))))
    
    l<-list.files(path="QM/QM/")
    
    for(z in l){
      na<-paste("QM/QM/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[41:41,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[42:42,4:15]))))
      
      na<-strsplit(z,"_")[[1]][2]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="CiW DPU @ QFL2")
    p
  })
  
  output$table_plot_qm_qfl2_ciw<-renderTable({
    
    te<-input$save_qm_kpi
    f<-paste("QM/QM/QM_",input$choose_plot_year_qm_qfl2_ciw,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[41:42,1:15]
    d
  })
  
  #Frame DPU @ QFL2
  
  
  output$plot_qm_qfl2_fra<-renderPlotly({
    
    te<-input$save_qm_kpi
    f<-paste("QM/QM/QM_",input$choose_plot_year_qm_qfl2_fra,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[43:43,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[44:44,4:15]))))
    
    l<-list.files(path="QM/QM/")
    
    for(z in l){
      na<-paste("QM/QM/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[43:43,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[44:44,4:15]))))
      
      na<-strsplit(z,"_")[[1]][2]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Frame DPU @ QFL2")
    p
  })
  
  output$table_plot_qm_qfl2_fra<-renderTable({
    
    te<-input$save_qm_kpi
    f<-paste("QM/QM/QM_",input$choose_plot_year_qm_qfl2_fra,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[43:44,1:15]
    d
  })
  
  
  
  #FTT HDT
  
  
  output$plot_qm_ftt_hdt<-renderPlotly({
    
    te<-input$save_qm_kpi
    
    f<-paste("spr_ftt/spr_ftt_",input$choose_plot_year_qm_ftt_hdt,".xlsx",sep="")
    d <- read_excel(f)
    
    d<-d[d$KPI=="FTT",]
    d<-d[d$Description=="HDT",]
    
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[1:1,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))))
    
    
    l<-list.files(path="spr_ftt/")
    
    for(z in l){
      na<-paste("spr_ftt/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[1:1,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[2:2,4:15]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    
    da<-da[order(da$xval),]
    da1<-da1[order(da1$xval),]
    
    
    da$zval<-NA
    da$yval2<-da1$yval
    
    for(i in 1:nrow(da)){
      if(is.na(da$yval[i]))
        da$yval[i]<-0
      if(is.na(da$yval2[i]))
        da$yval2[i]<-0
    }
    
    for (i in 1:nrow(da)){
      if(da$yval[i]<=da$yval2[i])
        da$zval[i]<-"green4"
      else
        da$zval[i]<-"red"
    }
    
    
    p<-ggplot(da,aes(x=xval,y=yval2))+
      geom_bar(stat="identity", fill=da$zval, width = 0.75)+
      geom_text(aes(label=round(yval2)), vjust=1.5, color="black", size=3)+
      geom_line(aes(y=yval,group = 1),color='black')+
      # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
      theme(axis.text.x = element_text(angle = 45, hjust = 1))+
      labs(x="",y="")
    
    p
  })
  
  output$table_plot_qm_ftt_hdt<-renderTable({
    
    te<-input$save_qm_kpi
    f<-paste("spr_ftt/spr_ftt_",input$choose_plot_year_qm_ftt_hdt,".xlsx",sep="")
    d <- read_excel(f)
    
    d<-d[d$KPI=="FTT",]
    d<-d[d$Description=="HDT",]
    d<-d[1:2,1:15]
    d
  })
  
  #FTT MDT
  
  
  output$plot_qm_ftt_mdt<-renderPlotly({
    
    te<-input$save_qm_kpi
    
    f<-paste("spr_ftt/spr_ftt_",input$choose_plot_year_qm_ftt_mdt,".xlsx",sep="")
    d <- read_excel(f)
    
    d<-d[d$KPI=="FTT",]
    d<-d[d$Description=="MDT",]
    
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[1:1,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))))
    
    
    l<-list.files(path="spr_ftt/")
    
    for(z in l){
      na<-paste("spr_ftt/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[1:1,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[2:2,4:15]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    
    da<-da[order(da$xval),]
    da1<-da1[order(da1$xval),]
    
    
    da$zval<-NA
    da$yval2<-da1$yval
    
    for(i in 1:nrow(da)){
      if(is.na(da$yval[i]))
        da$yval[i]<-0
      if(is.na(da$yval2[i]))
        da$yval2[i]<-0
    }
    
    for (i in 1:nrow(da)){
      if(da$yval[i]<=da$yval2[i])
        da$zval[i]<-"green4"
      else
        da$zval[i]<-"red"
    }
    
    
    p<-ggplot(da,aes(x=xval,y=yval2))+
      geom_bar(stat="identity", fill=da$zval, width = 0.75)+
      geom_text(aes(label=round(yval2)), vjust=1.5, color="black", size=3)+
      geom_line(aes(y=yval,group = 1),color='black')+
      # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
      theme(axis.text.x = element_text(angle = 45, hjust = 1))+
      labs(x="",y="")
    
    p
  })
  
  output$table_plot_qm_ftt_mdt<-renderTable({
    
    te<-input$save_qm_kpi
    f<-paste("spr_ftt/spr_ftt_",input$choose_plot_year_qm_ftt_mdt,".xlsx",sep="")
    d <- read_excel(f)
    
    d<-d[d$KPI=="FTT",]
    d<-d[d$Description=="MDT",]
    d<-d[1:2,1:15]
    d
  })
  
  
  #FTT LDT
  
  
  output$plot_qm_ftt_ldt<-renderPlotly({
    
    te<-input$save_qm_kpi
    
    f<-paste("spr_ftt/spr_ftt_",input$choose_plot_year_qm_ftt_ldt,".xlsx",sep="")
    d <- read_excel(f)
    
    d<-d[d$KPI=="FTT",]
    d<-d[d$Description=="LDT",]
    
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[1:1,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))))
    
    
    l<-list.files(path="spr_ftt/")
    
    for(z in l){
      na<-paste("spr_ftt/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[1:1,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[2:2,4:15]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    
    da<-da[order(da$xval),]
    da1<-da1[order(da1$xval),]
    
    
    da$zval<-NA
    da$yval2<-da1$yval
    
    for(i in 1:nrow(da)){
      if(is.na(da$yval[i]))
        da$yval[i]<-0
      if(is.na(da$yval2[i]))
        da$yval2[i]<-0
    }
    
    for (i in 1:nrow(da)){
      if(da$yval[i]<=da$yval2[i])
        da$zval[i]<-"green4"
      else
        da$zval[i]<-"red"
    }
    
    
    p<-ggplot(da,aes(x=xval,y=yval2))+
      geom_bar(stat="identity", fill=da$zval, width = 0.75)+
      geom_text(aes(label=round(yval2)), vjust=1.5, color="black", size=3)+
      geom_line(aes(y=yval,group = 1),color='black')+
      # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
      theme(axis.text.x = element_text(angle = 45, hjust = 1))+
      labs(x="",y="")
    
    p
  })
  
  output$table_plot_qm_ftt_ldt<-renderTable({
    
    te<-input$save_qm_kpi
    f<-paste("spr_ftt/spr_ftt_",input$choose_plot_year_qm_ftt_ldt,".xlsx",sep="")
    d <- read_excel(f)
    
    d<-d[d$KPI=="FTT",]
    d<-d[d$Description=="LDT",]
    d<-d[1:2,1:15]
    d
  })
  
  #SPR HDT
  
  
  output$plot_qm_spr_hdt<-renderPlotly({
    
    te<-input$save_qm_kpi
    
    f<-paste("spr_ftt/spr_ftt_",input$choose_plot_year_qm_spr_hdt,".xlsx",sep="")
    d <- read_excel(f)
    
    d<-d[d$KPI=="SPR",]
    d<-d[d$Description=="HDT",]
    
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[1:1,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))))
    
    
    l<-list.files(path="spr_ftt/")
    
    for(z in l){
      na<-paste("spr_ftt/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[1:1,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[2:2,4:15]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    
    da<-da[order(da$xval),]
    da1<-da1[order(da1$xval),]
    
    
    da$zval<-NA
    da$yval2<-da1$yval
    
    for(i in 1:nrow(da)){
      if(is.na(da$yval[i]))
        da$yval[i]<-0
      if(is.na(da$yval2[i]))
        da$yval2[i]<-0
    }
    
    for (i in 1:nrow(da)){
      if(da$yval[i]<=da$yval2[i])
        da$zval[i]<-"green4"
      else
        da$zval[i]<-"red"
    }
    
    
    p<-ggplot(da,aes(x=xval,y=yval2))+
      geom_bar(stat="identity", fill=da$zval, width = 0.75)+
      geom_text(aes(label=round(yval2)), vjust=1.5, color="black", size=3)+
      geom_line(aes(y=yval,group = 1),color='black')+
      # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
      theme(axis.text.x = element_text(angle = 45, hjust = 1))+
      labs(x="",y="")
    
    p
  })
  
  output$table_plot_qm_spr_hdt<-renderTable({
    
    te<-input$save_qm_kpi
    f<-paste("spr_ftt/spr_ftt_",input$choose_plot_year_qm_spr_hdt,".xlsx",sep="")
    d <- read_excel(f)
    
    d<-d[d$KPI=="SPR",]
    d<-d[d$Description=="HDT",]
    d<-d[1:2,1:15]
    d
  })
  
  #SPR MDT
  
  
  output$plot_qm_spr_mdt<-renderPlotly({
    
    te<-input$save_qm_kpi
    
    f<-paste("spr_ftt/spr_ftt_",input$choose_plot_year_qm_spr_mdt,".xlsx",sep="")
    d <- read_excel(f)
    
    d<-d[d$KPI=="SPR",]
    d<-d[d$Description=="MDT",]
    
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[1:1,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))))
    
    
    l<-list.files(path="spr_ftt/")
    
    for(z in l){
      na<-paste("spr_ftt/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[1:1,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[2:2,4:15]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    
    da<-da[order(da$xval),]
    da1<-da1[order(da1$xval),]
    
    
    da$zval<-NA
    da$yval2<-da1$yval
    
    for(i in 1:nrow(da)){
      if(is.na(da$yval[i]))
        da$yval[i]<-0
      if(is.na(da$yval2[i]))
        da$yval2[i]<-0
    }
    
    for (i in 1:nrow(da)){
      if(da$yval[i]<=da$yval2[i])
        da$zval[i]<-"green4"
      else
        da$zval[i]<-"red"
    }
    
    
    p<-ggplot(da,aes(x=xval,y=yval2))+
      geom_bar(stat="identity", fill=da$zval, width = 0.75)+
      geom_text(aes(label=round(yval2)), vjust=1.5, color="black", size=3)+
      geom_line(aes(y=yval,group = 1),color='black')+
      # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
      theme(axis.text.x = element_text(angle = 45, hjust = 1))+
      labs(x="",y="")
    
    p
  })
  
  output$table_plot_qm_spr_mdt<-renderTable({
    
    te<-input$save_qm_kpi
    f<-paste("spr_ftt/spr_ftt_",input$choose_plot_year_qm_spr_mdt,".xlsx",sep="")
    d <- read_excel(f)
    
    d<-d[d$KPI=="SPR",]
    d<-d[d$Description=="MDT",]
    d<-d[1:2,1:15]
    d
  })
  
  #SPR LDT
  
  
  output$plot_qm_spr_ldt<-renderPlotly({
    
    te<-input$save_qm_kpi
    
    f<-paste("spr_ftt/spr_ftt_",input$choose_plot_year_qm_spr_ldt,".xlsx",sep="")
    d <- read_excel(f)
    
    d<-d[d$KPI=="SPR",]
    d<-d[d$Description=="LDT",]
    
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[1:1,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[2:2,4:15]))))
    
    
    l<-list.files(path="spr_ftt/")
    
    for(z in l){
      na<-paste("spr_ftt/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[1:1,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[2:2,4:15]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    
    da<-da[order(da$xval),]
    da1<-da1[order(da1$xval),]
    
    
    da$zval<-NA
    da$yval2<-da1$yval
    
    for(i in 1:nrow(da)){
      if(is.na(da$yval[i]))
        da$yval[i]<-0
      if(is.na(da$yval2[i]))
        da$yval2[i]<-0
    }
    
    for (i in 1:nrow(da)){
      if(da$yval[i]<=da$yval2[i])
        da$zval[i]<-"green4"
      else
        da$zval[i]<-"red"
    }
    
    
    p<-ggplot(da,aes(x=xval,y=yval2))+
      geom_bar(stat="identity", fill=da$zval, width = 0.75)+
      geom_text(aes(label=round(yval2)), vjust=1.5, color="black", size=3)+
      geom_line(aes(y=yval,group = 1),color='black')+
      # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
      theme(axis.text.x = element_text(angle = 45, hjust = 1))+
      labs(x="",y="")
    
    p
  })
  
  output$table_plot_qm_spr_ldt<-renderTable({
    
    te<-input$save_qm_kpi
    f<-paste("spr_ftt/spr_ftt_",input$choose_plot_year_qm_spr_ldt,".xlsx",sep="")
    d <- read_excel(f)
    
    d<-d[d$KPI=="SPR",]
    d<-d[d$Description=="LDT",]
    d<-d[1:2,1:15]
    d
  })
  
  
  #Chassis KPI input data
  
  values_chassis_kpi <- reactiveValues()
  
  
  previous_chassis_kpi <- reactive({
    d<-read_excel("Chassis/KPI/chassis_kpi_2020.xlsx")
    d
  })
  
  MyChanges_chassis_kpi <- reactive({
    if(is.null(input$hotable_chassis_kpi)){return(previous_chassis_kpi())}
    else if(!identical(previous_chassis_kpi(),input$hotable_chassis_kpi)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable_chassis_kpi <- as.data.frame(hot_to_r(input$hotable_chassis_kpi))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_chassis_kpi <- mytable_chassis_kpi[1:nrow(previous_chassis_kpi()),]
      
      for(i in 4:ncol(mytable_chassis_kpi))
        mytable_chassis_kpi[4,i]<-mytable_chassis_kpi[2,i]/mytable_chassis_kpi[1,i]
      
      for(i in 4:ncol(mytable_chassis_kpi))
        mytable_chassis_kpi[29,i]<-sum(mytable_chassis_kpi[c(23,25,27),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_chassis_kpi))
        mytable_chassis_kpi[30,i]<-sum(mytable_chassis_kpi[c(24,26,28),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_chassis_kpi))
        mytable_chassis_kpi[31,i]<-0.013*sum(mytable_chassis_kpi[c(13,15),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_chassis_kpi))
        mytable_chassis_kpi[32,i]<-sum(mytable_chassis_kpi[c(33,34,35),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_chassis_kpi))
        mytable_chassis_kpi[42,i]<-sum(mytable_chassis_kpi[38:41,i],na.rm=TRUE)
      
      
      for(i in 4:ncol(mytable_chassis_kpi))
        mytable_chassis_kpi[45,i]<-mytable_chassis_kpi[44,i]/(sum(mytable_chassis_kpi[c(60,61),i],na.rm=TRUE))
      
      for(i in 4:ncol(mytable_chassis_kpi))
        mytable_chassis_kpi[48,i]<-mytable_chassis_kpi[47,i]/(sum(mytable_chassis_kpi[c(62),i],na.rm=TRUE))
      
      for(i in 4:ncol(mytable_chassis_kpi))
        mytable_chassis_kpi[51,i]<-100*mytable_chassis_kpi[50,i]/(sum(mytable_chassis_kpi[c(59,60),i],na.rm=TRUE))
      
      for(i in 4:ncol(mytable_chassis_kpi))
        mytable_chassis_kpi[54,i]<-100*mytable_chassis_kpi[53,i]/(sum(mytable_chassis_kpi[c(61,62),i],na.rm=TRUE))
      
      for(i in 4:ncol(mytable_chassis_kpi))
        mytable_chassis_kpi[58,i]<-100*mytable_chassis_kpi[57,i]/mytable_chassis_kpi[56,i]
      
      for(i in 4:ncol(mytable_chassis_kpi))
        mytable_chassis_kpi[63,i]<-(sum(mytable_chassis_kpi[c(60,61,62),i],na.rm=TRUE))
      
      mytable_chassis_kpi
    }
  })   
  
  output$hotable_chassis_kpi<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30)+1)
    row_highlight = c(4,29,30,31,32,42,45,48,51,54,58,63)-1
    row_readonly=c(4,29,30,31,32,42,45,48,51,54,58,63)
    
    rhandsontable(MyChanges_chassis_kpi(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1250,height = 600)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+2),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_chassis_kpi,{
    #
    write.xlsx(hot_to_r(input$hotable_chassis_kpi),"Chassis/KPI/chassis_kpi_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  
  

  
  
  #Chassis Roll out input data
  
  values_chassis_rollout <- reactiveValues()
  
  
  previous_chassis_rollout <- reactive({
    d<-read_excel("Chassis/Chassis_rollout/chassis_rollout_2020.xlsx")
    d
  })
  
  MyChanges_chassis_rollout <- reactive({
    if(is.null(input$hotable_chassis_rollout)){return(previous_chassis_rollout())}
    else if(!identical(previous_chassis_rollout(),input$hotable_chassis_rollout)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable_chassis_rollout <- as.data.frame(hot_to_r(input$hotable_chassis_rollout))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_chassis_rollout <- mytable_chassis_rollout[1:nrow(previous_chassis_rollout()),]
      
      for(i in 3:ncol(mytable_chassis_rollout))
        mytable_chassis_rollout[5,i]<-sum(mytable_chassis_rollout[c(1,3),i],na.rm = TRUE)
      
      for(i in 3:ncol(mytable_chassis_rollout))
        mytable_chassis_rollout[6,i]<-sum(mytable_chassis_rollout[c(2,4),i],na.rm = TRUE)
      
      mytable_chassis_rollout
    }
  })
  
  output$hotable_chassis_rollout<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30))
    row_highlight = c(5,6)-1
    row_readonly=c(5,6)
    
    rhandsontable(MyChanges_chassis_rollout(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1250,height = 600)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+1),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_chassis_rollout,{
    #
    write.xlsx(hot_to_r(input$hotable_chassis_rollout),"Chassis/Chassis_rollout/chassis_rollout_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #Chassis QC OK input data
  
  values_chassis_qcok <- reactiveValues()
  
  
  previous_chassis_qcok <- reactive({
    d<-read_excel("Chassis/Chassis_qcok/chassis_qcok_2020.xlsx")
    d
  })
  
  MyChanges_chassis_qcok <- reactive({
    if(is.null(input$hotable_chassis_qcok)){return(previous_chassis_qcok())}
    else if(!identical(previous_chassis_qcok(),input$hotable_chassis_qcok)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable_chassis_qcok <- as.data.frame(hot_to_r(input$hotable_chassis_qcok))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_chassis_qcok <- mytable_chassis_qcok[1:nrow(previous_chassis_qcok()),]
      
      for(i in 3:ncol(mytable_chassis_qcok))
        mytable_chassis_qcok[5,i]<-sum(mytable_chassis_qcok[c(1,3),i],na.rm = TRUE)
      
      for(i in 3:ncol(mytable_chassis_qcok))
        mytable_chassis_qcok[6,i]<-sum(mytable_chassis_qcok[c(2,4),i],na.rm = TRUE)
      
      mytable_chassis_qcok
    }
  })
  
  output$hotable_chassis_qcok<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30))
    row_highlight = c(5,6)-1
    row_readonly=c(5,6)
    
    rhandsontable(MyChanges_chassis_qcok(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1250,height = 600)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+1),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_chassis_qcok,{
    #
    write.xlsx(hot_to_r(input$hotable_chassis_qcok),"Chassis/Chassis_qcok/chassis_qcok_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #Chassis Capacity utilization input data
  
  values_chassis_capacity <- reactiveValues()
  
  
  previous_chassis_capacity <- reactive({
    d<-read_excel("Chassis/Chassis_capacity/chassis_capacity_2020.xlsx")
    d
  })
  
  MyChanges_chassis_capacity <- reactive({
    if(is.null(input$hotable_chassis_capacity)){return(previous_chassis_capacity())}
    else if(!identical(previous_chassis_capacity(),input$hotable_chassis_capacity)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable_chassis_capacity <- as.data.frame(hot_to_r(input$hotable_chassis_capacity))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_chassis_capacity <- mytable_chassis_capacity[1:nrow(previous_chassis_capacity()),]
      
      #for(i in 3:ncol(mytable_chassis_capacity))
      #  mytable_chassis_capacity[5,i]<-sum(mytable_chassis_capacity[c(1,3),i],na.rm = TRUE)
      
      #for(i in 3:ncol(mytable_chassis_capacity))
      #  mytable_chassis_capacity[6,i]<-sum(mytable_chassis_capacity[c(2,4),i],na.rm = TRUE)
      
      mytable_chassis_capacity
    }
  })
  
  output$hotable_chassis_capacity<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30))
    #row_highlight = c(5,6)-1
    #row_readonly=c(5,6)
    
    rhandsontable(MyChanges_chassis_capacity(), col_highlight = col_highlight,width = 1250,height = 600)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+1),readOnly = TRUE)%>%
      #hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_chassis_capacity,{
    #
    write.xlsx(hot_to_r(input$hotable_chassis_capacity),"Chassis/Chassis_capacity/chassis_capacity_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #Delivery QC OK Vehicle
  
  
  output$plot_del_qcok_veh<-renderPlotly({
    
    te<-input$save_chassis_qcok
    f<-paste("Chassis/Chassis_qcok/chassis_qcok_",input$choose_plot_del_qcok_veh,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[5:5,3:14]))))
    da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[6:6,3:14]))))
    
    l<-list.files(path="Chassis/Chassis_qcok/")
    
    for(z in l){
      na<-paste("Chassis/Chassis_qcok/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[5:5,3:14]))))
      dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[6:6,3:14]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval<da$yval,]
    da12<-da1[da1$yval>=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Vehicle QC OK")
    p
  })
  
  #Delivery QC OK Bus
  
  
  output$plot_del_qcok_bus<-renderPlotly({
    
    te<-input$save_chassis_qcok
    f<-paste("Chassis/Chassis_qcok/chassis_qcok_",input$choose_plot_del_qcok_veh,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[7:7,3:14]))))
    da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[8:8,3:14]))))
    
    l<-list.files(path="Chassis/Chassis_qcok/")
    
    for(z in l){
      na<-paste("Chassis/Chassis_qcok/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[7:7,3:14]))))
      dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[8:8,3:14]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval<da$yval,]
    da12<-da1[da1$yval>=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Bus QC OK")
    p
  })
  
  output$table_plot_del_qcok_veh<-renderTable({
    
    te<-input$save_chassis_qcok
    f<-paste("Chassis/Chassis_qcok/chassis_qcok_",input$choose_plot_del_qcok_veh,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[5:8,1:14]
    d
  })
  
  #Delivery QC OK CKD
  
  
  output$plot_del_qcok_ckd<-renderPlotly({
    
    te<-input$save_chassis_qcok
    f<-paste("Chassis/Chassis_qcok/chassis_qcok_",input$choose_plot_del_qcok_ckd,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[9:9,3:14]))))
    da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[10:10,3:14]))))
    
    l<-list.files(path="Chassis/Chassis_qcok/")
    
    for(z in l){
      na<-paste("Chassis/Chassis_qcok/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[9:9,3:14]))))
      dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[10:10,3:14]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval<da$yval,]
    da12<-da1[da1$yval>=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="CKD QC OK")
    p
  })
  
  output$table_plot_del_qcok_ckd<-renderTable({
    
    te<-input$save_chassis_qcok
    f<-paste("Chassis/Chassis_qcok/chassis_qcok_",input$choose_plot_del_qcok_ckd,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[9:10,1:14]
    d
  })
  
  #Delivery Roll Out HDT
  
  
  output$plot_del_rollout_hdt<-renderPlotly({
    
    te<-input$save_chassis_rollout
    f<-paste("Chassis/Chassis_rollout/chassis_rollout_",input$choose_plot_del_rollout,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[1:1,3:14]))))
    da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[2:2,3:14]))))
    
    l<-list.files(path="Chassis/Chassis_rollout/")
    
    for(z in l){
      na<-paste("Chassis/Chassis_rollout/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[1:1,3:14]))))
      dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[2:2,3:14]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval<da$yval,]
    da12<-da1[da1$yval>=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Vehicle Roll Out- HDT")
    p
  })
  

  
  #Delivery Roll Out MDT
  
  
  output$plot_del_rollout_mdt<-renderPlotly({
    
    te<-input$save_chassis_rollout
    f<-paste("Chassis/Chassis_rollout/chassis_rollout_",input$choose_plot_del_rollout_mdt,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[3:3,3:14]))))
    da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[4:4,3:14]))))
    
    l<-list.files(path="Chassis/Chassis_rollout/")
    
    for(z in l){
      na<-paste("Chassis/Chassis_rollout/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[3:3,3:14]))))
      dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[4:4,3:14]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval<da$yval,]
    da12<-da1[da1$yval>=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Vehicle ROll Out- MDT")
    p
  })
  
  output$table_plot_del_rollout<-renderTable({
    
    te<-input$save_chassis_rollout
    f<-paste("Chassis/Chassis_rollout/chassis_rollout_",input$choose_plot_del_rollout,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[1:2,1:14]
    d
  })
  output$table_plot_del_rollout_mdt<-renderTable({
    
    te<-input$save_chassis_rollout
    f<-paste("Chassis/Chassis_rollout/chassis_rollout_",input$choose_plot_del_rollout,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[3:4,1:14]
    d
  })
  
  
  #Delivery Capacity utlization HDT
  
  output$plot_del_capacity_hdt<-renderPlotly({
    
    te<-input$save_chassis_capacity
   
    f<-paste("Chassis/Chassis_capacity/chassis_capacity_",input$choose_plot_del_capacity,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[1:1,3:14]))))
    da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[2:2,3:14]))))
    
    l<-list.files(path="Chassis/Chassis_capacity/")
    
    for(z in l){
      na<-paste("Chassis/Chassis_capacity/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[1:1,3:14]))))
      dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[2:2,3:14]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    da<-da[order(da$xval),]
    da1<-da1[order(da1$xval),]
    
    da$zval<-NA
    da$yval2<-da1$yval
    
    for(i in 1:nrow(da)){
      if(is.na(da$yval[i]))
        da$yval[i]<-0
      if(is.na(da$yval2[i]))
        da$yval2[i]<-0
    }
    
    for (i in 1:nrow(da)){
      if(da$yval[i]<=da$yval2[i])
        da$zval[i]<-"green4"
      else
        da$zval[i]<-"red"
    }
    
    
    p<-ggplot(da,aes(x=xval,y=yval2))+
      geom_bar(stat="identity", fill=da$zval, width = 0.75)+
      geom_text(aes(label=round(yval2)), vjust=1.5, color="black", size=3)+
      geom_line(aes(y=yval,group = 1),color='black')+
      # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
      theme(axis.text.x = element_text(angle = 45, hjust = 1))+
      labs(x="",y="")
    
    p
  })
  
  
  
  #Delivery Capacity utlization MDT
  
  
  output$plot_del_capacity_mdt<-renderPlotly({
    
    te<-input$save_chassis_capacity

    f<-paste("Chassis/Chassis_capacity/chassis_capacity_",input$choose_plot_del_capacity_mdt,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[3:3,3:14]))))
    da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[4:4,3:14]))))
    
    l<-list.files(path="Chassis/Chassis_capacity/")
    
    for(z in l){
      na<-paste("Chassis/Chassis_capacity/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[3:3,3:14]))))
      dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[4:4,3:14]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    
    da<-da[order(da$xval),]
    da1<-da1[order(da1$xval),]
    
    da$zval<-NA
    da$yval2<-da1$yval
    
    for(i in 1:nrow(da)){
      if(is.na(da$yval[i]))
        da$yval[i]<-0
      if(is.na(da$yval2[i]))
        da$yval2[i]<-0
    }
    
    for (i in 1:nrow(da)){
      if(da$yval[i]<=da$yval2[i])
        da$zval[i]<-"green4"
      else
        da$zval[i]<-"red"
    }
    
    
    p<-ggplot(da,aes(x=xval,y=yval2))+
      geom_bar(stat="identity", fill=da$zval, width = 0.75)+
      geom_text(aes(label=round(yval2)), vjust=1.5, color="black", size=2.5)+
      geom_line(aes(y=yval,group = 1),color='black')+
      # geom_text(aes(label=round(yval2), x=xval, y=yval2), colour="black",position = position_dodge(width = 0.8),size=3)+
      theme(axis.text.x = element_text(angle = 45, hjust = 1))+
      labs(x="",y="")
    
     p
  })
  
  output$table_plot_del_capacity<-renderTable({
    
    te<-input$save_chassis_capacity
    f<-paste("Chassis/Chassis_capacity/chassis_capacity_",input$choose_plot_del_capacity,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[1:2,1:14]
    d
  })
  output$table_plot_del_capacity_mdt<-renderTable({
    
    te<-input$save_chassis_capacity
    f<-paste("Chassis/Chassis_capacity/chassis_capacity_",input$choose_plot_del_capacity,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[3:4,1:14]
    d
  })
  
  
  #Delivery Non Forecast Shortages
  
  
  output$plot_del_nonforecast<-renderPlotly({
    
    te<-input$save_chassis_nonforecast
    f<-paste("Chassis/KPI/chassis_kpi_",input$choose_plot_del_nonforecast,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[21:21,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[22:22,4:15]))))
    
    l<-list.files(path="Chassis/KPI/")
    
    for(z in l){
      na<-paste("Chassis/KPI/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[21:21,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[22:22,4:15]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    
    p<-plot_ly(x=da1$xval,y = da1$yval,showlegend=FALSE,type = 'bar',name='Actual',textposition = 'auto',width=0.9,marker=list(color='red'))%>%
      layout(hovermode= 'compare',
             title="Non Forecasted Shortages")
    p
  })
  
  output$table_plot_del_nonforecast<-renderTable({
    
    te<-input$save_chassis_capacity
    f<-paste("Chassis/KPI/chassis_kpi_",input$choose_plot_del_nonforecast,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[21:22,1:15]
    d
  })
  
  #Delivery Vehicle losses due to operations
  
  
  output$plot_del_opp_loss<-renderPlotly({
    
    te<-input$save_chassis_opp_loss
    f<-paste("Chassis/KPI/chassis_kpi_",input$choose_plot_del_opp_loss,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[29:29,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[30:30,4:15]))))
    
    l<-list.files(path="Chassis/KPI/")
    
    for(z in l){
      na<-paste("Chassis/KPI/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[29:29,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[30:30,4:15]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    #da1$yval<-as.numeric(da1$yval)
    #da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Vehicle losses due to Operations")
    p
  })
  
  output$table_plot_del_opp_loss<-renderTable({
    
    te<-input$save_chassis_capacity
    f<-paste("Chassis/KPI/chassis_kpi_",input$choose_plot_del_opp_loss,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[c(29,30,24,26,28),3:15]
    d
  })
  
  #Delivery Vehicle losses due to aggregates
  
  
  output$plot_del_agg_loss<-renderPlotly({
    
    te<-input$save_chassis_agg_loss
    f<-paste("Chassis/KPI/chassis_kpi_",input$choose_plot_del_agg_loss,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[31:31,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[32:32,4:15]))))
    
    l<-list.files(path="Chassis/KPI/")
    
    for(z in l){
      na<-paste("Chassis/KPI/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[31:31,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[32:32,4:15]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Vehicle losses due to Aggregates")
    p
  })
  
  output$table_plot_del_agg_loss<-renderTable({
    
    te<-input$save_chassis_capacity
    f<-paste("Chassis/KPI/chassis_kpi_",input$choose_plot_del_agg_loss,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[31:35,3:15]
    d
  })
  
  
  #Cost
  
  #Cost HPU per capacity input data
  
  values_hpu_capacity <- reactiveValues()
  
  
  previous_hpu_capacity <- reactive({
    d<-read_excel("HPU_Capacity/HPU_Capacity/hpu_capacity_2020.xlsx")
    d
  })
  
  MyChanges_hpu_capacity <- reactive({
    if(is.null(input$hotable_hpu_capacity)){return(previous_hpu_capacity())}
    else if(!identical(previous_hpu_capacity(),input$hotable_hpu_capacity)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable_hpu_capacity <- as.data.frame(hot_to_r(input$hotable_hpu_capacity))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_hpu_capacity <- mytable_hpu_capacity[1:nrow(previous_hpu_capacity()),]
      
      for(i in 3:ncol(mytable_hpu_capacity))
        mytable_hpu_capacity[1,i]<-sum(mytable_hpu_capacity[seq(3,20,2),i],na.rm=TRUE)
      
      
      for(i in 3:ncol(mytable_hpu_capacity))
        mytable_hpu_capacity[2,i]<-sum(mytable_hpu_capacity[seq(4,20,2),i],na.rm=TRUE)
      
      
      mytable_hpu_capacity
    }
  })   
  
  output$hotable_hpu_capacity<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30))
    row_highlight = c(1,2)-1
    row_readonly=c(1,2)
    
    rhandsontable(MyChanges_hpu_capacity(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1250,height = 600)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+1),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_hpu_capacity,{
    #
    write.xlsx(hot_to_r(input$hotable_hpu_capacity),"HPU_Capacity/HPU_Capacity/hpu_capacity_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #Cost HPU per capacity
  
  
  output$plot_cos_hpu_capacity<-renderPlotly({
    
    te<-input$save_hpu_capacity
    f<-paste("HPU_Capacity/HPU_Capacity/hpu_capacity_",input$choose_plot_cos_hpu_capacity,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[1:1,3:14]))))
    da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[2:2,3:14]))))
    
    l<-list.files(path="HPU_Capacity/HPU_Capacity/")
    
    for(z in l){
      na<-paste("HPU_Capacity/HPU_Capacity/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[1:1,3:14]))))
      dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[2:2,3:14]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Plant Level - HPU per capacity")
    p
  })
  
  output$table_plot_cos_hpu_capacity<-renderTable({
    
    te<-input$save_hpu_capacity
    f<-paste("HPU_Capacity/HPU_Capacity/hpu_capacity_",input$choose_plot_cos_hpu_capacity,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[1:2,1:14]
    d
  })
  
  
  #Cost shop level HPU per capacity
  
  data_hpu_capacity_shop_act<-reactive({
    te<-input$save_hpu_capacity
    d<-read_excel("HPU_Capacity/HPU_Capacity/hpu_capacity_2020.xlsx")
    d<-d[d$Department!="Plant level",]
    da<-d[d$Category=='Actual',]
    dt<-d[d$Category=='Target',]
    
    da<-da%>%gather(month,value,3:14)
    da$Category<-NULL
    da
  })
  data_hpu_capacity_shop_tar<-reactive({
    te<-input$save_hpu_capacity
    d<-read_excel("HPU_Capacity/HPU_Capacity/hpu_capacity_2020.xlsx")
    d<-d[d$Department!="Plant level",]
    da<-d[d$Category=='Actual',]
    dt<-d[d$Category=='Target',]
    
    dt<-dt%>%gather(month,value,3:14)
    dt$Category<-NULL
    dt
  })
  data_dept_hpu_capacity_shop<-reactive({
    te<-input$save_hpu_capacity
    d<-read_excel("HPU_Capacity/HPU_Capacity/hpu_capacity_2020.xlsx")
    d<-d[d$Department!="Plant level",]
    
    d<-d%>%gather(month,value,3:14)
    d$month<-as.yearmon(d$month,"%b %Y")
    d<-d[months(d$month)==input$choose_comp_hpu_capacity_shop,]
    d$Department<-factor(d$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame","TOS"))
    
    d<-spread(d,Department,value )
    
    d$Category<-factor(d$Category,levels=c("Target","Actual"))
    d$month<-months(d$month)
    d
  })
  
  output$comp_hpu_capacity_shop<-renderPlotly({
    te<-input$save_hpu_capacity
    da<-data_hpu_capacity_shop_act()
    dt<-data_hpu_capacity_shop_tar()
    
    da$month<-as.yearmon(da$month,"%b %Y")
    dt$month<-as.yearmon(dt$month,"%b %Y")
    
    da1<-da[months(da$month)==input$choose_comp_hpu_capacity_shop,]
    da<-dt[months(dt$month)==input$choose_comp_hpu_capacity_shop,]
    
    da$Department<-factor(da$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","Frame","TOS"))
    da1$Department<-factor(da1$Department,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","Frame","TOS"))
    
    da$yval<-as.numeric(da$value)
    da1$yval<-as.numeric(da1$value)
    
    da11<-da1[da1$value>da$value,]
    da12<-da1[da1$value<=da$value,]
    
    #if(nrow(da11)==0)
    #  da11<-da11%>%add_row(xval=na,yval=0)
    #if(nrow(da12)==0)
    #  da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$Department,y = da$value, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$Department,y=da11$value,text=round(da11$yval,digits=1), textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$Department,y=da12$value,text=round(da12$yval,digits=1), textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
  })
  
  output$table_comp_hpu_capacity_shop<-renderTable({
    data_dept_hpu_capacity_shop()
  })
  
  #CabTrim
  
  
  values_cabtrim_kpi <- reactiveValues()
  
  
  previous_cabtrim_kpi <- reactive({
    d<-read_excel("Cabtrim/KPI/cabtrim_kpi_2020.xlsx")
    d
  })
  
  MyChanges_cabtrim_kpi <- reactive({
    if(is.null(input$hotable_cabtrim_kpi)){return(previous_cabtrim_kpi())}
    else if(!identical(previous_cabtrim_kpi(),input$hotable_cabtrim_kpi)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable_cabtrim_kpi <- as.data.frame(hot_to_r(input$hotable_cabtrim_kpi))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_cabtrim_kpi <- mytable_cabtrim_kpi[1:nrow(previous_cabtrim_kpi()),]
      
      for(i in 4:ncol(mytable_cabtrim_kpi))
        mytable_cabtrim_kpi[4,i]<-mytable_cabtrim_kpi[2,i]/mytable_cabtrim_kpi[1,i]
      
      
      for(i in 4:ncol(mytable_cabtrim_kpi))
        mytable_cabtrim_kpi[13,i]<-sum(mytable_cabtrim_kpi[c(9,10,11,12),i],na.rm=TRUE)
      
      
      for(i in 4:ncol(mytable_cabtrim_kpi))
        mytable_cabtrim_kpi[16,i]<-mytable_cabtrim_kpi[15,i]/sum(mytable_cabtrim_kpi[c(31,32),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_cabtrim_kpi))
        mytable_cabtrim_kpi[19,i]<-mytable_cabtrim_kpi[18,i]/sum(mytable_cabtrim_kpi[c(33),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_cabtrim_kpi))
        mytable_cabtrim_kpi[22,i]<-100*mytable_cabtrim_kpi[21,i]/sum(mytable_cabtrim_kpi[c(30,31),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_cabtrim_kpi))
        mytable_cabtrim_kpi[25,i]<-100*mytable_cabtrim_kpi[24,i]/sum(mytable_cabtrim_kpi[c(32,33),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_cabtrim_kpi))
        mytable_cabtrim_kpi[29,i]<-100*mytable_cabtrim_kpi[28,i]/mytable_cabtrim_kpi[27,i]
      
      for(i in 4:ncol(mytable_cabtrim_kpi))
        mytable_cabtrim_kpi[34,i]<-sum(mytable_cabtrim_kpi[c(31,32,33),i],na.rm=TRUE)
      
      mytable_cabtrim_kpi
    }
  })   
  
  output$hotable_cabtrim_kpi<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30)+1)
    row_highlight = c(4,13,16,19,22,25,29,34)-1
    row_readonly=c(4,13,16,19,22,25,29,34)
    
    rhandsontable(MyChanges_cabtrim_kpi(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1250,height = 650)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+2),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_cabtrim_kpi,{
    #
    write.xlsx(hot_to_r(input$hotable_cabtrim_kpi),"Cabtrim/KPI/cabtrim_kpi_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  

  #eol
  
  
  values_eol_kpi <- reactiveValues()
  
  
  previous_eol_kpi <- reactive({
    d<-read_excel("EOL/KPI/eol_kpi_2020.xlsx")
    d
  })
  
  MyChanges_eol_kpi <- reactive({
    if(is.null(input$hotable_eol_kpi)){return(previous_eol_kpi())}
    else if(!identical(previous_eol_kpi(),input$hotable_eol_kpi)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable_eol_kpi <- as.data.frame(hot_to_r(input$hotable_eol_kpi))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_eol_kpi <- mytable_eol_kpi[1:nrow(previous_eol_kpi()),]
      
      for(i in 4:ncol(mytable_eol_kpi))
        mytable_eol_kpi[4,i]<-mytable_eol_kpi[2,i]/mytable_eol_kpi[1,i]
      
      
      for(i in 4:ncol(mytable_eol_kpi))
        mytable_eol_kpi[19,i]<-sum(mytable_eol_kpi[c(15,16,17,18),i],na.rm=TRUE)
      
      
      for(i in 4:ncol(mytable_eol_kpi))
        mytable_eol_kpi[22,i]<-mytable_eol_kpi[21,i]/sum(mytable_eol_kpi[c(37,38),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_eol_kpi))
        mytable_eol_kpi[25,i]<-mytable_eol_kpi[24,i]/sum(mytable_eol_kpi[c(39),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_eol_kpi))
        mytable_eol_kpi[28,i]<-100*mytable_eol_kpi[27,i]/sum(mytable_eol_kpi[c(37,36),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_eol_kpi))
        mytable_eol_kpi[31,i]<-100*mytable_eol_kpi[30,i]/sum(mytable_eol_kpi[c(39,38),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_eol_kpi))
        mytable_eol_kpi[35,i]<-100*mytable_eol_kpi[34,i]/mytable_eol_kpi[33,i]
      
      for(i in 4:ncol(mytable_eol_kpi))
        mytable_eol_kpi[40,i]<-sum(mytable_eol_kpi[c(37,38,39),i],na.rm=TRUE)
        
      mytable_eol_kpi
    }
  })   
  
  output$hotable_eol_kpi<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30)+1)
    row_highlight = c(4,19,22,25,28,31,35,40)-1
    row_readonly=c(4,19,22,25,28,31,35,40)
    
    rhandsontable(MyChanges_eol_kpi(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1250,height = 650)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+2),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_eol_kpi,{
    #
    write.xlsx(hot_to_r(input$hotable_eol_kpi),"EOL/KPI/eol_kpi_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  
  #fbv kpi
  
  
  values_fbv_kpi <- reactiveValues()
  
  
  previous_fbv_kpi <- reactive({
    d<-read_excel("FBV/KPI/fbv_kpi_2020.xlsx")
    d
  })
  
  MyChanges_fbv_kpi <- reactive({
    if(is.null(input$hotable_fbv_kpi)){return(previous_fbv_kpi())}
    else if(!identical(previous_fbv_kpi(),input$hotable_fbv_kpi)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable_fbv_kpi <- as.data.frame(hot_to_r(input$hotable_fbv_kpi))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_fbv_kpi <- mytable_fbv_kpi[1:nrow(previous_fbv_kpi()),]
      
      for(i in 4:ncol(mytable_fbv_kpi))
        mytable_fbv_kpi[4,i]<-mytable_fbv_kpi[2,i]/mytable_fbv_kpi[1,i]
      
      
      for(i in 4:ncol(mytable_fbv_kpi))
        mytable_fbv_kpi[15,i]<-sum(mytable_fbv_kpi[c(11,12,13,14),i],na.rm=TRUE)
      
      
      for(i in 4:ncol(mytable_fbv_kpi))
        mytable_fbv_kpi[18,i]<-mytable_fbv_kpi[17,i]/sum(mytable_fbv_kpi[c(33,34),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_fbv_kpi))
        mytable_fbv_kpi[21,i]<-mytable_fbv_kpi[20,i]/sum(mytable_fbv_kpi[c(35),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_fbv_kpi))
        mytable_fbv_kpi[24,i]<-100*mytable_fbv_kpi[23,i]/sum(mytable_fbv_kpi[c(32,33),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_fbv_kpi))
        mytable_fbv_kpi[27,i]<-100*mytable_fbv_kpi[26,i]/sum(mytable_fbv_kpi[c(34,35),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_fbv_kpi))
        mytable_fbv_kpi[31,i]<-100*mytable_fbv_kpi[30,i]/mytable_fbv_kpi[29,i]
      
      for(i in 4:ncol(mytable_fbv_kpi))
        mytable_fbv_kpi[36,i]<-sum(mytable_fbv_kpi[c(33,34,35),i],na.rm=TRUE)
      
      mytable_fbv_kpi
    }
  })   
  
  output$hotable_fbv_kpi<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30)+1)
    row_highlight = c(4,15,18,21,24,27,31,36)-1
    row_readonly=c(4,15,18,21,24,27,31,36)
    
    rhandsontable(MyChanges_fbv_kpi(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1250,height = 650)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+2),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_fbv_kpi,{
    #
    write.xlsx(hot_to_r(input$hotable_fbv_kpi),"fbv/KPI/fbv_kpi_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  

  
  #ciw kpi
  
  
  values_ciw_kpi <- reactiveValues()
  
  
  previous_ciw_kpi <- reactive({
    d<-read_excel("CIW/KPI/ciw_kpi_2020.xlsx")
    d
  })
  
  MyChanges_ciw_kpi <- reactive({
    if(is.null(input$hotable_ciw_kpi)){return(previous_ciw_kpi())}
    else if(!identical(previous_ciw_kpi(),input$hotable_ciw_kpi)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable_ciw_kpi <- as.data.frame(hot_to_r(input$hotable_ciw_kpi))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_ciw_kpi <- mytable_ciw_kpi[1:nrow(previous_ciw_kpi()),]
      
      for(i in 4:ncol(mytable_ciw_kpi))
        mytable_ciw_kpi[4,i]<-mytable_ciw_kpi[2,i]/mytable_ciw_kpi[1,i]
      
      
      for(i in 4:ncol(mytable_ciw_kpi))
        mytable_ciw_kpi[15,i]<-sum(mytable_ciw_kpi[c(11,12,13,14),i],na.rm=TRUE)
      
      
      for(i in 4:ncol(mytable_ciw_kpi))
        mytable_ciw_kpi[18,i]<-mytable_ciw_kpi[17,i]/sum(mytable_ciw_kpi[c(33,34),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_ciw_kpi))
        mytable_ciw_kpi[21,i]<-mytable_ciw_kpi[20,i]/sum(mytable_ciw_kpi[c(35),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_ciw_kpi))
        mytable_ciw_kpi[24,i]<-100*mytable_ciw_kpi[23,i]/sum(mytable_ciw_kpi[c(32,33),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_ciw_kpi))
        mytable_ciw_kpi[27,i]<-100*mytable_ciw_kpi[26,i]/sum(mytable_ciw_kpi[c(34),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_ciw_kpi))
        mytable_ciw_kpi[31,i]<-100*mytable_ciw_kpi[30,i]/mytable_ciw_kpi[29,i]
      
      for(i in 4:ncol(mytable_ciw_kpi))
        mytable_ciw_kpi[36,i]<-sum(mytable_ciw_kpi[c(33,34,35),i],na.rm=TRUE)
        
      mytable_ciw_kpi
    }
  })   
  
  output$hotable_ciw_kpi<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30)+1)
    row_highlight = c(4,15,18,21,24,27,31,36)-1
    row_readonly=c(4,15,18,21,24,27,31,36)
    
    
    rhandsontable(MyChanges_ciw_kpi(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1250,height = 650)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+2),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_ciw_kpi,{
    #
    write.xlsx(hot_to_r(input$hotable_ciw_kpi),"ciw/KPI/ciw_kpi_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  
  
  #paint kpi
  
  
  values_paint_kpi <- reactiveValues()
  
  
  previous_paint_kpi <- reactive({
    d<-read_excel("PAINT/KPI/paint_kpi_2020.xlsx")
    d
  })
  
  MyChanges_paint_kpi <- reactive({
    if(is.null(input$hotable_paint_kpi)){return(previous_paint_kpi())}
    else if(!identical(previous_paint_kpi(),input$hotable_paint_kpi)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable_paint_kpi <- as.data.frame(hot_to_r(input$hotable_paint_kpi))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_paint_kpi <- mytable_paint_kpi[1:nrow(previous_paint_kpi()),]
      
      for(i in 4:ncol(mytable_paint_kpi))
        mytable_paint_kpi[4,i]<-mytable_paint_kpi[2,i]/mytable_paint_kpi[1,i]
      
      
      for(i in 4:ncol(mytable_paint_kpi))
        mytable_paint_kpi[15,i]<-sum(mytable_paint_kpi[c(11,12,13,14),i],na.rm=TRUE)
      
      
      for(i in 4:ncol(mytable_paint_kpi))
        mytable_paint_kpi[18,i]<-mytable_paint_kpi[17,i]/sum(mytable_paint_kpi[c(33,34),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_paint_kpi))
        mytable_paint_kpi[21,i]<-mytable_paint_kpi[20,i]/sum(mytable_paint_kpi[c(35),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_paint_kpi))
        mytable_paint_kpi[24,i]<-100*mytable_paint_kpi[23,i]/sum(mytable_paint_kpi[c(33,32),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_paint_kpi))
        mytable_paint_kpi[27,i]<-100*mytable_paint_kpi[26,i]/sum(mytable_paint_kpi[c(35,34),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_paint_kpi))
        mytable_paint_kpi[31,i]<-100*mytable_paint_kpi[30,i]/mytable_paint_kpi[29,i]
      
      for(i in 4:ncol(mytable_paint_kpi))
        mytable_paint_kpi[36,i]<-sum(mytable_paint_kpi[c(35,34,33),i],na.rm=TRUE)
      
      mytable_paint_kpi
    }
  })   
  
  output$hotable_paint_kpi<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30)+1)
    row_highlight = c(4,15,18,21,24,27,31,36)-1
    row_readonly=c(4,15,18,21,24,27,31,36)
    
    
    rhandsontable(MyChanges_paint_kpi(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1250,height = 650)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+2),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_paint_kpi,{
    #
    write.xlsx(hot_to_r(input$hotable_paint_kpi),"paint/KPI/paint_kpi_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  

  
  
  #engine kpi
  
  
  values_engine_kpi <- reactiveValues()
  
  
  previous_engine_kpi <- reactive({
    d<-read_excel("engine/KPI/engine_kpi_2020.xlsx")
    d
  })
  
  MyChanges_engine_kpi <- reactive({
    if(is.null(input$hotable_engine_kpi)){return(previous_engine_kpi())}
    else if(!identical(previous_engine_kpi(),input$hotable_engine_kpi)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable_engine_kpi <- as.data.frame(hot_to_r(input$hotable_engine_kpi))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_engine_kpi <- mytable_engine_kpi[1:nrow(previous_engine_kpi()),]
      
      for(i in 4:ncol(mytable_engine_kpi))
        mytable_engine_kpi[4,i]<-mytable_engine_kpi[2,i]/mytable_engine_kpi[1,i]
      
      
      for(i in 4:ncol(mytable_engine_kpi))
        mytable_engine_kpi[15,i]<-sum(mytable_engine_kpi[c(11,12,13,14),i],na.rm=TRUE)
      
      
      for(i in 4:ncol(mytable_engine_kpi))
        mytable_engine_kpi[18,i]<-mytable_engine_kpi[17,i]/sum(mytable_engine_kpi[c(33,34),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_engine_kpi))
        mytable_engine_kpi[21,i]<-mytable_engine_kpi[20,i]/sum(mytable_engine_kpi[c(35),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_engine_kpi))
        mytable_engine_kpi[24,i]<-100*mytable_engine_kpi[23,i]/sum(mytable_engine_kpi[c(33,32),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_engine_kpi))
        mytable_engine_kpi[27,i]<-100*mytable_engine_kpi[26,i]/sum(mytable_engine_kpi[c(35,34),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_engine_kpi))
        mytable_engine_kpi[31,i]<-100*mytable_engine_kpi[30,i]/mytable_engine_kpi[29,i]
      
      for(i in 4:ncol(mytable_engine_kpi))
        mytable_engine_kpi[36,i]<-sum(mytable_engine_kpi[c(35,34,33),i],na.rm=TRUE)
        
      mytable_engine_kpi
    }
  })   
  
  output$hotable_engine_kpi<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30)+1)
    row_highlight = c(4,15,18,21,24,27,31,36)-1
    row_readonly=c(4,15,18,21,24,27,31,36)
    
    
    rhandsontable(MyChanges_engine_kpi(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1250,height = 650)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+2),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_engine_kpi,{
    #
    write.xlsx(hot_to_r(input$hotable_engine_kpi),"engine/KPI/engine_kpi_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  

  
  
  #transmission kpi
  
  
  values_transmission_kpi <- reactiveValues()
  
  
  previous_transmission_kpi <- reactive({
    d<-read_excel("transmission/KPI/transmission_kpi_2020.xlsx")
    d
  })
  
  MyChanges_transmission_kpi <- reactive({
    if(is.null(input$hotable_transmission_kpi)){return(previous_transmission_kpi())}
    else if(!identical(previous_transmission_kpi(),input$hotable_transmission_kpi)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable_transmission_kpi <- as.data.frame(hot_to_r(input$hotable_transmission_kpi))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_transmission_kpi <- mytable_transmission_kpi[1:nrow(previous_transmission_kpi()),]
      
      for(i in 4:ncol(mytable_transmission_kpi))
        mytable_transmission_kpi[4,i]<-mytable_transmission_kpi[2,i]/mytable_transmission_kpi[1,i]
      
      
      for(i in 4:ncol(mytable_transmission_kpi))
        mytable_transmission_kpi[15,i]<-sum(mytable_transmission_kpi[c(11,12,13,14),i],na.rm=TRUE)
      
      
      for(i in 4:ncol(mytable_transmission_kpi))
        mytable_transmission_kpi[18,i]<-mytable_transmission_kpi[17,i]/sum(mytable_transmission_kpi[c(33,34),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_transmission_kpi))
        mytable_transmission_kpi[21,i]<-mytable_transmission_kpi[20,i]/sum(mytable_transmission_kpi[c(35),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_transmission_kpi))
        mytable_transmission_kpi[24,i]<-100*mytable_transmission_kpi[23,i]/sum(mytable_transmission_kpi[c(33,32),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_transmission_kpi))
        mytable_transmission_kpi[27,i]<-100*mytable_transmission_kpi[26,i]/sum(mytable_transmission_kpi[c(35,34),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_transmission_kpi))
        mytable_transmission_kpi[31,i]<-100*mytable_transmission_kpi[30,i]/mytable_transmission_kpi[29,i]
      
      for(i in 4:ncol(mytable_transmission_kpi))
        mytable_transmission_kpi[36,i]<-sum(mytable_transmission_kpi[c(33,35,34),i],na.rm=TRUE)
      
      mytable_transmission_kpi
    }
  })   
  
  output$hotable_transmission_kpi<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30)+1)
    row_highlight = c(4,15,18,21,24,27,31,36)-1
    row_readonly=c(4,15,18,21,24,27,31,36)
    
    
    rhandsontable(MyChanges_transmission_kpi(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1250,height = 650)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+2),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_transmission_kpi,{
    #
    write.xlsx(hot_to_r(input$hotable_transmission_kpi),"transmission/KPI/transmission_kpi_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  

  
  #frame kpi
  
  
  values_frame_kpi <- reactiveValues()
  
  
  previous_frame_kpi <- reactive({
    d<-read_excel("frame/KPI/frame_kpi_2020.xlsx")
    d
  })
  
  MyChanges_frame_kpi <- reactive({
    
    if(is.null(input$hotable_frame_kpi)){return(previous_frame_kpi())}
    else if(!identical(previous_frame_kpi(),input$hotable_frame_kpi)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable_frame_kpi <- as.data.frame(hot_to_r(input$hotable_frame_kpi))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_frame_kpi <- mytable_frame_kpi[1:nrow(previous_frame_kpi()),]
      
      for(i in 4:ncol(mytable_frame_kpi))
        mytable_frame_kpi[4,i]<-mytable_frame_kpi[2,i]/mytable_frame_kpi[1,i]
      
      
      for(i in 4:ncol(mytable_frame_kpi))
        mytable_frame_kpi[15,i]<-sum(mytable_frame_kpi[c(11,12,13,14),i],na.rm=TRUE)
      
      
      for(i in 4:ncol(mytable_frame_kpi))
        mytable_frame_kpi[18,i]<-100*mytable_frame_kpi[17,i]/sum(mytable_frame_kpi[c(30,31),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_frame_kpi))
        mytable_frame_kpi[21,i]<-100*mytable_frame_kpi[20,i]/sum(mytable_frame_kpi[c(29,30),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_frame_kpi))
        mytable_frame_kpi[24,i]<-100*mytable_frame_kpi[23,i]/sum(mytable_frame_kpi[c(31),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_frame_kpi))
        mytable_frame_kpi[28,i]<-100*mytable_frame_kpi[27,i]/mytable_frame_kpi[26,i]
      
      mytable_frame_kpi
    }
  })   
  
  output$hotable_frame_kpi<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30)+1)
    row_highlight = c(4,15,18,21,24,28)-1
    row_readonly=c(4,15,18,21,24,28)
    
    
    rhandsontable(MyChanges_frame_kpi(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1250,height = 650)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+2),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_frame_kpi,{
    #
    write.xlsx(hot_to_r(input$hotable_frame_kpi),"frame/KPI/frame_kpi_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  

  
  #ipl kpi
  
  
  values_ipl_kpi <- reactiveValues()
  
  
  previous_ipl_kpi <- reactive({
    d<-read_excel("ipl/KPI/ipl_kpi_2020.xlsx")
    d
  })
  
  MyChanges_ipl_kpi <- reactive({
    if(is.null(input$hotable_ipl_kpi)){return(previous_ipl_kpi())}
    else if(!identical(previous_ipl_kpi(),input$hotable_ipl_kpi)){
      # hot.to.df function will convert your updated table into the dataipl
      mytable_ipl_kpi <- as.data.frame(hot_to_r(input$hotable_ipl_kpi))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_ipl_kpi <- mytable_ipl_kpi[1:nrow(previous_ipl_kpi()),]
      
      for(i in 4:ncol(mytable_ipl_kpi))
        mytable_ipl_kpi[4,i]<-mytable_ipl_kpi[2,i]/mytable_ipl_kpi[1,i]
      
      
      for(i in 4:ncol(mytable_ipl_kpi))
        mytable_ipl_kpi[11,i]<-sum(mytable_ipl_kpi[c(7,8,9,10),i],na.rm=TRUE)
      
      
      
      for(i in 4:ncol(mytable_ipl_kpi))
        mytable_ipl_kpi[14,i]<-mytable_ipl_kpi[13,i]/sum(mytable_ipl_kpi[c(29,30),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_ipl_kpi))
        mytable_ipl_kpi[17,i]<-mytable_ipl_kpi[16,i]/sum(mytable_ipl_kpi[c(31),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_ipl_kpi))
        mytable_ipl_kpi[20,i]<-100*mytable_ipl_kpi[19,i]/sum(mytable_ipl_kpi[c(28,29),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_ipl_kpi))
        mytable_ipl_kpi[23,i]<-100*mytable_ipl_kpi[22,i]/sum(mytable_ipl_kpi[c(30,31),i],na.rm=TRUE)
      
      for(i in 4:ncol(mytable_ipl_kpi))
        mytable_ipl_kpi[27,i]<-100*mytable_ipl_kpi[26,i]/mytable_ipl_kpi[25,i]
      
      for(i in 4:ncol(mytable_ipl_kpi))
        mytable_ipl_kpi[32,i]<-sum(mytable_ipl_kpi[c(29,30,31),i],na.rm=TRUE)
      
      mytable_ipl_kpi
    }
  })   
  
  output$hotable_ipl_kpi<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30)+1)
    row_highlight = c(4,11,14,17,20,23,27,32)-1
    row_readonly=c(4,11,14,17,20,23,27,32)
    
    
    rhandsontable(MyChanges_ipl_kpi(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1250,height = 650)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+2),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_ipl_kpi,{
    #
    write.xlsx(hot_to_r(input$hotable_ipl_kpi),"ipl/KPI/ipl_kpi_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  

  
  #fm kpi
  
  
  values_fm_kpi <- reactiveValues()
  
  
  previous_fm_kpi <- reactive({
    d<-read_excel("fm/KPI/fm_kpi_2020.xlsx")
    d
  })
  
  MyChanges_fm_kpi <- reactive({
    if(is.null(input$hotable_fm_kpi)){return(previous_fm_kpi())}
    else if(!identical(previous_fm_kpi(),input$hotable_fm_kpi)){
      # hot.to.df function will convert your updated table into the datafm
      mytable_fm_kpi <- as.data.frame(hot_to_r(input$hotable_fm_kpi))
      # here the second column is a function of the first and it will be multfmed by 100 given the values in the first column
      mytable_fm_kpi <- mytable_fm_kpi[1:nrow(previous_fm_kpi()),]
      
    for(i in 4:ncol(mytable_fm_kpi))
      mytable_fm_kpi[6,i]<-100*mytable_fm_kpi[5,i]/mytable_fm_kpi[4,i]
    
    
    
    for(i in 4:ncol(mytable_fm_kpi))
      mytable_fm_kpi[19,i]<-mytable_fm_kpi[18,i]/sum(mytable_fm_kpi[c(24,25),i],na.rm=TRUE)
    
    for(i in 4:ncol(mytable_fm_kpi))
      mytable_fm_kpi[22,i]<-100*mytable_fm_kpi[21,i]/sum(mytable_fm_kpi[c(23,24),i],na.rm=TRUE)
    
      mytable_fm_kpi
    }
  })   
  
  output$hotable_fm_kpi<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30)+1)
    row_highlight = c(6,19,22)-1
    row_readonly=c(6,15,19,22)
    
    rhandsontable(MyChanges_fm_kpi(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1250,height = 650)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+2),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_fm_kpi,{
    
    
    write.xlsx(hot_to_r(input$hotable_fm_kpi),"fm/KPI/fm_kpi_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #wc kpi
  
  
  values_wc_kpi <- reactiveValues()
  
  
  previous_wc_kpi <- reactive({
    d<-read_excel("white_collar/white_collar_2020.xlsx")
    d
  })
  
  MyChanges_wc_kpi <- reactive({
    if(is.null(input$hotable_wc_kpi)){return(previous_wc_kpi())}
    else if(!identical(previous_wc_kpi(),input$hotable_wc_kpi)){
      # hot.to.df function will convert your updated table into the datafm
      mytable_wc_kpi <- as.data.frame(hot_to_r(input$hotable_wc_kpi))
      # here the second column is a function of the first and it will be multfmed by 100 given the values in the first column
      mytable_wc_kpi <- mytable_wc_kpi[1:nrow(previous_wc_kpi()),]
     
      for(j in 1:14) 
      for(i in 3:ncol(mytable_wc_kpi))
        mytable_wc_kpi[3*j,i]<-100*mytable_wc_kpi[3*j-1,i]/mytable_wc_kpi[3*j-2,i]
      
      for(j in 1:14) 
        for(i in 4:ncol(mytable_wc_kpi))
          mytable_wc_kpi[3*j-2,i]<-sum(mytable_wc_kpi[3*j-2,i-1],-mytable_wc_kpi[3*j-1,i-1],na.rm = TRUE)
      
      for(i in 4:ncol(mytable_wc_kpi))
        mytable_wc_kpi[43,i]<-sum(mytable_wc_kpi[c(1:14)*3-2,i],na.rm = TRUE)
      
      mytable_wc_kpi[44,i]<-sum(mytable_wc_kpi[c(1:14)*3-1,3],na.rm = TRUE)
      
      for(i in 4:ncol(mytable_wc_kpi))
        mytable_wc_kpi[44,i]<-sum(mytable_wc_kpi[c(1:14)*3-1,i],mytable_wc_kpi[44,i-1],na.rm = TRUE)
      
      for(i in 3:ncol(mytable_wc_kpi))
        mytable_wc_kpi[46,i]<-100*mytable_wc_kpi[44,i]/mytable_wc_kpi[43,i]
      
      mytable_wc_kpi
    }
  })   
  
  output$hotable_wc_kpi<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30))
    row_highlight = c(c(1:14)*3,43,44,46)-1
    row_readonly=c(c(1:14)*3,43,44,46)
    
    rhandsontable(MyChanges_wc_kpi(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1250,height = 650)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+1),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_wc_kpi,{
    
    
    write.xlsx(hot_to_r(input$hotable_wc_kpi),"white_collar/white_collar_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #Morale White collar
  
  
  output$plot_mor_white_collar<-renderPlotly({
    
    te<-input$save_wc_kpi
    f<-paste("white_collar/white_collar_",input$choose_plot_mor_white_collar,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[45:45,3:14]))))
    da1<-data.frame(xval=colnames(d)[3:14],yval=c(t(array(d[46:46,3:14]))))
    
    l<-list.files(path="white_collar/")
    
    for(z in l){
      na<-paste("white_collar/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[45:45,3:14]))))
      dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[46:46,3:14]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:14]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Plant Level - White Collar Attrition Rate")
    p
  })
  
  output$table_plot_mor_white_collar<-renderTable({
    
    te<-input$save_wc_kpi
    f<-paste("white_collar/white_collar_",input$choose_plot_mor_white_collar,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[45:46,1:14]
    d
  })
  
  
  #Cost shop level white_collar
  
  data_white_collar_shop_act<-reactive({
    te<-input$save_wc_kpi
    d<-read_excel("white_collar/white_collar_2020.xlsx")
    d<-d[d$KPI!="Plant level",]
    da<-d[d$Description=='Ratio',]
    dt<-d[d$Description=='Target',]
    
    da<-da%>%gather(month,value,3:14)
    da$Description<-NULL
    da
  })
  data_white_collar_shop_tar<-reactive({
    te<-input$save_wc_kpi
    d<-read_excel("white_collar/white_collar_2020.xlsx")
    d<-d[d$KPI!="Plant level",]
    da<-d[d$Description=='Ratio',]
    dt<-da
    
    dt<-dt%>%gather(month,value,3:14)
    dt$value<-3.3
    dt$Description<-NULL
    dt
  })
  data_dept_white_collar_shop<-reactive({
    te<-input$save_wc_kpi
    d<-read_excel("white_collar/white_collar_2020.xlsx")
    d<-d[d$KPI!="Plant level",]
    
    d<-d%>%gather(month,value,3:14)
    d$month<-as.yearmon(d$month,"%b %Y")
    d<-d[months(d$month)==input$choose_comp_white_collar_shop,]
    d$KPI<-factor(d$KPI,levels=c("Chassis","Cabtrim","EOL/ FBV","CiW","Paint","Engine","Transmn.","QM","FM","IPL","ME","Frame","VP office","TOS & OMCD"))
    
    d<-spread(d,KPI,value )
    
    d$month<-months(d$month)
    d
  })
  
  output$comp_white_collar_shop<-renderPlotly({
    te<-input$save_wc_kpi
    da<-data_white_collar_shop_act()
    dt<-data_white_collar_shop_tar()
    
    da$month<-as.yearmon(da$month,"%b %Y")
    dt$month<-as.yearmon(dt$month,"%b %Y")
    
    da1<-da[months(da$month)==input$choose_comp_white_collar_shop,]
    da<-dt[months(dt$month)==input$choose_comp_white_collar_shop,]
    
    da$KPI<-factor(da$KPI,levels=c("Chassis","Cabtrim","EOL/ FBV","CiW","Paint","Engine","Transmn.","QM","FM","IPL","ME","Frame","VP office","TOS & OMCD"))
    da1$KPI<-factor(da1$KPI,levels=c("Chassis","Cabtrim","EOL/ FBV","CiW","Paint","Engine","Transmn.","QM","FM","IPL","ME","Frame","VP office","TOS & OMCD"))
    
    da$yval<-as.numeric(da$value)
    da1$yval<-as.numeric(da1$value)
    
    da11<-da1[da1$value>da$value,]
    da12<-da1[da1$value<=da$value,]
    
    #if(nrow(da11)==0)
    #  da11<-da11%>%add_row(xval=na,yval=0)
    #if(nrow(da12)==0)
    #  da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$KPI,y = da$value, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$KPI,y=da11$value,text=round(da11$yval,digits=1),width=0.5, textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$KPI,y=da12$value,text=round(da12$yval,digits=1), width=0.5, textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
  })
  
  output$table_comp_white_collar_shop<-renderTable({
    data_dept_white_collar_shop()
  })
  
  morale_data<-reactive({
    
    d1<-read_excel("Chassis/KPI/chassis_kpi_2020.xlsx")
    d1<-d1[d1$KPI=="Morale",]
    d1$KPI<-"Chassis"
    
    d<-d1
    
    d1<-read_excel("Cabtrim/KPI/cabtrim_kpi_2020.xlsx")
    d1<-d1[d1$KPI=="Morale",]
    d1$KPI<-"Cabtrim"
    
    d<-rbind(d,d1)
    
    d1<-read_excel("EOL/KPI/eol_kpi_2020.xlsx")
    d1<-d1[d1$KPI=="Morale",]
    d1$KPI<-"EOL"
    
    d<-rbind(d,d1)
    
    d1<-read_excel("FBV/KPI/fbv_kpi_2020.xlsx")
    d1<-d1[d1$KPI=="Morale",]
    d1$KPI<-"FBV"
    
    d<-rbind(d,d1)
    
    d1<-read_excel("CIW/KPI/ciw_kpi_2020.xlsx")
    d1<-d1[d1$KPI=="Morale",]
    d1$KPI<-"CiW"
    
    d<-rbind(d,d1)
    
    d1<-read_excel("PAINT/KPI/paint_kpi_2020.xlsx")
    d1<-d1[d1$KPI=="Morale",]
    d1$KPI<-"Paint"
    
    d<-rbind(d,d1)
    
    d1<-read_excel("engine/KPI/engine_kpi_2020.xlsx")
    d1<-d1[d1$KPI=="Morale",]
    d1$KPI<-"Engine"
    
    d<-rbind(d,d1)
    
    d1<-read_excel("transmission/KPI/transmission_kpi_2020.xlsx")
    d1<-d1[d1$KPI=="Morale",]
    d1$KPI<-"Transmission"
    
    d<-rbind(d,d1)
    
    d1<-read_excel("ipl/KPI/ipl_kpi_2020.xlsx")
    d1<-d1[d1$KPI=="Morale",]
    d1$KPI<-"IPL"
    
    d<-rbind(d,d1)
    
    
    d1<-read_excel("QM/QM/QM_2020.xlsx")
    
    d1<-d1[d1$KPI=="Morale",]
    d1$KPI<-"QM"
    
    d<-rbind(d,d1)
    
    d1<-read_excel("fm/KPI/fm_kpi_2020.xlsx")
    
    d1<-d1[d1$KPI=="Morale",]
    d1$KPI<-"FM"
    
    d<-rbind(d,d1)
    
    
    d1<-read_excel("frame/KPI/frame_kpi_2020.xlsx")
    
    d1<-d1[d1$KPI=="Morale",]
    d1$KPI<-"Frame"
    
    d<-rbind(d,d1)
    
    d
  })
  
  
  #Morale bca_participation
  
  
  output$plot_mor_bca_participation<-renderPlotly({
    
    
    d <- morale_data()
    
    d<-d[d$Description=="Participation in AOM - BCA/T & Engineers",]
    
    dd<-d[d$Category=="Actual",]
    dd2<-d[d$Category=="Target",]
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd[,4:15],mean,na.rm=TRUE)))))
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd2[,4:15],mean,na.rm=TRUE)))))
    
    da<-da%>%add_row(xval="2020",yval=mean(da$yval))
    da1<-da1%>%add_row(xval="2020",yval=mean(da1$yval))
    
    da<-da%>%add_row(xval="2017",yval=0.38)
    da1<-da1%>%add_row(xval="2017",yval=0.39)
    
    da<-da%>%add_row(xval="2018",yval=0.42)
    da1<-da1%>%add_row(xval="2018",yval=0.42)
    
    da<-da%>%add_row(xval="2019",yval=0.42)
    da1<-da1%>%add_row(xval="2019",yval=0.43)
    
  # l<-list.files(path="bca_participation/")
  # 
  # for(z in l){
  #   na<-paste("bca_participation/",z,sep='')
  #   dt<-read_excel(na)
  #   dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[45:45,3:14]))))
  #   dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[46:46,3:14]))))
  # 
  #   na<-strsplit(z,"_")[[1]][3]
  #   na<-strsplit(na,"[.]")[[1]][1]
  #   da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
  #   da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
  # }
  # 
  # na=toString(year(Sys.Date())+1)
  # 
  # da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
  # da1<-da1%>%add_row(xval=na,yval=0)

  da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:15]))
  da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:15]))
  da1$yval<-as.numeric(da1$yval)
  da$yval<-as.numeric(da$yval)


    da11<-da1[da1$yval<da$yval,]
    da12<-da1[da1$yval>=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=2), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=2), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Plant Level - Kaizen per BCA")
    p
  })
  
  output$table_plot_mor_bca_participation<-renderTable({
    
    d <- morale_data()
    
    d<-d[d$Description=="Participation in AOM - BCA/T & Engineers",]
    
    dd<-d[d$Category=="Actual",]
    dd2<-d[d$Category=="Target",]
    da1<-sapply(dd[,4:15],mean,na.rm=TRUE)
    da<-sapply(dd2[,4:15],mean,na.rm=TRUE)
    
    d<-rbind(da,da1)
  })
  
  
  #Morale shop level bca_participation
  
  data_bca_participation_shop_act<-reactive({
    d <- morale_data()
    
    d<-d[d$Description=="Participation in AOM - BCA/T & Engineers",]
    
    dd<-d[d$Category=="Actual",]
    dd2<-d[d$Category=="Target",]
    
    da<-dd%>%gather(month,value,4:15)
    da$Category<-NULL
    da
  })
  data_bca_participation_shop_tar<-reactive({
    d <- morale_data()
    
    d<-d[d$Description=="Participation in AOM - BCA/T & Engineers",]
    
    dd<-d[d$Category=="Actual",]
    dd2<-d[d$Category=="Target",]
    
    dt<-dd2%>%gather(month,value,4:15)
    
    dt$Category<-NULL
    dt
  })
  data_dept_bca_participation_shop<-reactive({
    d <- morale_data()
    
    d<-d[d$Description=="Participation in AOM - BCA/T & Engineers",]
    
    
    d<-d%>%gather(month,value,4:15)
    d$month<-as.yearmon(d$month,"%b %Y")
    d<-d[months(d$month)==input$choose_comp_bca_participation_shop,]
    d$KPI<-factor(d$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL"))
    
    d<-spread(d,KPI,value )
    
    d$month<-months(d$month)
    d
  })
  
  output$comp_bca_participation_shop<-renderPlotly({
    te<-input$save_wc_kpi
    da<-data_bca_participation_shop_act()
    dt<-data_bca_participation_shop_tar()
    
    da$month<-as.yearmon(da$month,"%b %Y")
    dt$month<-as.yearmon(dt$month,"%b %Y")
    
    da1<-da[months(da$month)==input$choose_comp_bca_participation_shop,]
    da<-dt[months(dt$month)==input$choose_comp_bca_participation_shop,]
    
    da$KPI<-factor(da$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL"))
    da1$KPI<-factor(da1$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL"))
    
    da$yval<-as.numeric(da$value)
    da1$yval<-as.numeric(da1$value)
    
    da11<-da1[da1$value<=da$value,]
    da12<-da1[da1$value>da$value,]
    
    #if(nrow(da11)==0)
    #  da11<-da11%>%add_row(xval=na,yval=0)
    #if(nrow(da12)==0)
    #  da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$KPI,y = da$value, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$KPI,y=da11$value,text=round(da11$yval,digits=2),width=0.5, textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$KPI,y=da12$value,text=round(da12$yval,digits=2), width=0.5, textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
  })
  
  output$table_comp_bca_participation_shop<-renderTable({
    data_dept_bca_participation_shop()
  })
  
  
  #Morale caba_participation
  
  
  output$plot_mor_caba_participation<-renderPlotly({
    
    
    d <- morale_data()
    d<-d[d$KPI!="IPL",]
    d<-d[d$Description=="Participation in AOM - CA/BA",]
    
    dd<-d[d$Category=="Actual",]
    dd2<-d[d$Category=="Target",]
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd[,4:15],mean,na.rm=TRUE)))))
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd2[,4:15],mean,na.rm=TRUE)))))
    
    da<-da%>%add_row(xval="2020",yval=mean(da$yval))
    da1<-da1%>%add_row(xval="2020",yval=mean(da1$yval))
    
    
    da<-da%>%add_row(xval="2019",yval=0.42)
    da1<-da1%>%add_row(xval="2019",yval=0.39)
    
    
    # l<-list.files(path="caba_participation/")
    # 
    # for(z in l){
    #   na<-paste("caba_participation/",z,sep='')
    #   dt<-read_excel(na)
    #   dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[45:45,3:14]))))
    #   dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[46:46,3:14]))))
    # 
    #   na<-strsplit(z,"_")[[1]][3]
    #   na<-strsplit(na,"[.]")[[1]][1]
    #   da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    #   da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    # }
    # 
    # na=toString(year(Sys.Date())+1)
    # 
    # da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    # da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval<da$yval,]
    da12<-da1[da1$yval>=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=2), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=2), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Plant Level - Participation in AOM - CA/BA")
    p
  })
  
  output$table_plot_mor_caba_participation<-renderTable({
    
    d <- morale_data()
    d<-d[d$KPI!="IPL",]
    d<-d[d$Description=="Participation in AOM - CA/BA",]
    
    dd<-d[d$Category=="Actual",]
    dd2<-d[d$Category=="Target",]
    da1<-sapply(dd[,4:15],mean,na.rm=TRUE)
    da<-sapply(dd2[,4:15],mean,na.rm=TRUE)
    
    d<-rbind(da,da1)
  })
  
  
  #Morale shop level caba_participation
  
  data_caba_participation_shop_act<-reactive({
    d <- morale_data()
    d<-d[d$KPI!="IPL",]
    d<-d[d$Description=="Participation in AOM - CA/BA",]
    
    dd<-d[d$Category=="Actual",]
    dd2<-d[d$Category=="Target",]
    
    da<-dd%>%gather(month,value,4:15)
    da$Category<-NULL
    da
  })
  data_caba_participation_shop_tar<-reactive({
    d <- morale_data()
    d<-d[d$KPI!="IPL",]
    d<-d[d$Description=="Participation in AOM - CA/BA",]
    
    dd<-d[d$Category=="Actual",]
    dd2<-d[d$Category=="Target",]
    
    dt<-dd2%>%gather(month,value,4:15)
    
    dt$Category<-NULL
    dt
  })
  data_dept_caba_participation_shop<-reactive({
    d <- morale_data()
    d<-d[d$KPI!="IPL",]
    d<-d[d$Description=="Participation in AOM - CA/BA",]
    
    
    d<-d%>%gather(month,value,4:15)
    d$month<-as.yearmon(d$month,"%b %Y")
    d<-d[months(d$month)==input$choose_comp_caba_participation_shop,]
    d$KPI<-factor(d$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","QM","FM"))
    
    d<-spread(d,KPI,value )
    
    d$month<-months(d$month)
    d
  })
  
  output$comp_caba_participation_shop<-renderPlotly({
    te<-input$save_wc_kpi
    da<-data_caba_participation_shop_act()
    dt<-data_caba_participation_shop_tar()
    
    da$month<-as.yearmon(da$month,"%b %Y")
    dt$month<-as.yearmon(dt$month,"%b %Y")
    
    da1<-da[months(da$month)==input$choose_comp_caba_participation_shop,]
    da<-dt[months(dt$month)==input$choose_comp_caba_participation_shop,]
    
    da$KPI<-factor(da$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","QM","FM"))
    da1$KPI<-factor(da1$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","QM","FM"))
    
    da$yval<-as.numeric(da$value)
    da1$yval<-as.numeric(da1$value)
    
    da11<-da1[da1$value<da$value,]
    da12<-da1[da1$value>=da$value,]
    
    #if(nrow(da11)==0)
    #  da11<-da11%>%add_row(xval=na,yval=0)
    #if(nrow(da12)==0)
    #  da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$KPI,y = da$value, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$KPI,y=da11$value,text=round(da11$yval,digits=2),width=0.5, textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$KPI,y=da12$value,text=round(da12$yval,digits=2), width=0.5, textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
  })
  
  output$table_comp_caba_participation_shop<-renderTable({
    data_dept_caba_participation_shop()
  })
  
  
  
  
  
  #Morale man_attrition
  
  
  output$plot_mor_man_attrition<-renderPlotly({
    
    
    d <- morale_data()
    
    d<-d[d$Description=="Attrition rate of Managers + Engineers",]
    
    dd<-d[d$Category=="Actual",]
    dd2<-d[d$Category=="Target",]
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd[,4:15],mean,na.rm=TRUE)))))
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd2[,4:15],mean,na.rm=TRUE)))))
    
    # l<-list.files(path="man_attrition/")
    # 
    # for(z in l){
    #   na<-paste("man_attrition/",z,sep='')
    #   dt<-read_excel(na)
    #   dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[45:45,3:14]))))
    #   dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[46:46,3:14]))))
    # 
    #   na<-strsplit(z,"_")[[1]][3]
    #   na<-strsplit(na,"[.]")[[1]][1]
    #   da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    #   da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    # }
    # 
    # na=toString(year(Sys.Date())+1)
    # 
    # da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    # da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=2), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=2), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Plant Level - Attrition rate of Managers + Engineers")
    p
  })
  
  output$table_plot_mor_man_attrition<-renderTable({
    
    d <- morale_data()
    
    d<-d[d$Description=="Attrition rate of Managers + Engineers",]
    
    dd<-d[d$Category=="Actual",]
    dd2<-d[d$Category=="Target",]
    da1<-sapply(dd[,4:15],mean,na.rm=TRUE)
    da<-sapply(dd2[,4:15],mean,na.rm=TRUE)
    
    d<-rbind(da,da1)
  })
  
  
  #Morale shop level man_attrition
  
  data_man_attrition_shop_act<-reactive({
    d <- morale_data()
    
    d<-d[d$Description=="Attrition rate of Managers + Engineers",]
    
    dd<-d[d$Category=="Actual",]
    dd2<-d[d$Category=="Target",]
    
    da<-dd%>%gather(month,value,4:15)
    da$Category<-NULL
    da
  })
  data_man_attrition_shop_tar<-reactive({
    d <- morale_data()
    
    d<-d[d$Description=="Attrition rate of Managers + Engineers",]
    
    dd<-d[d$Category=="Actual",]
    dd2<-d[d$Category=="Target",]
    
    dt<-dd2%>%gather(month,value,4:15)
    
    dt$Category<-NULL
    dt
  })
  data_dept_man_attrition_shop<-reactive({
    d <- morale_data()
    
    d<-d[d$Description=="Attrition rate of Managers + Engineers",]
    
    
    d<-d%>%gather(month,value,4:15)
    d$month<-as.yearmon(d$month,"%b %Y")
    d<-d[months(d$month)==input$choose_comp_man_attrition_shop,]
    d$KPI<-factor(d$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL"))
    
    d<-spread(d,KPI,value )
    
    d$month<-months(d$month)
    d
  })
  
  output$comp_man_attrition_shop<-renderPlotly({
    te<-input$save_wc_kpi
    da<-data_man_attrition_shop_act()
    dt<-data_man_attrition_shop_tar()
    
    da$month<-as.yearmon(da$month,"%b %Y")
    dt$month<-as.yearmon(dt$month,"%b %Y")
    
    da1<-da[months(da$month)==input$choose_comp_man_attrition_shop,]
    da<-dt[months(dt$month)==input$choose_comp_man_attrition_shop,]
    
    da$KPI<-factor(da$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","IPL","QM","FM"))
    da1$KPI<-factor(da1$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","IPL","QM","FM"))
    
    da$yval<-as.numeric(da$value)
    da1$yval<-as.numeric(da1$value)
    
    da11<-da1[da1$value>da$value,]
    da12<-da1[da1$value<=da$value,]
    
    #if(nrow(da11)==0)
    #  da11<-da11%>%add_row(xval=na,yval=0)
    #if(nrow(da12)==0)
    #  da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$KPI,y = da$value, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$KPI,y=da11$value,text=round(da11$yval,digits=2),width=0.5, textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$KPI,y=da12$value,text=round(da12$yval,digits=2), width=0.5, textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
  })
  
  output$table_comp_man_attrition_shop<-renderTable({
    data_dept_man_attrition_shop()
  })
  
  
  
  
  #Morale bca_attrition
  
  
  output$plot_mor_bca_attrition<-renderPlotly({
    
    
    d <- morale_data()
    
    d<-d[d$Description=="Attrition rate of BCA/BCAT/CA",]
    
    dd<-d[d$Category=="Actual",]
    dd2<-d[d$Category=="Target",]
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd[,4:15],mean,na.rm=TRUE)))))
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd2[,4:15],mean,na.rm=TRUE)))))
    
    # l<-list.files(path="bca_attrition/")
    # 
    # for(z in l){
    #   na<-paste("bca_attrition/",z,sep='')
    #   dt<-read_excel(na)
    #   dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[45:45,3:14]))))
    #   dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[46:46,3:14]))))
    # 
    #   na<-strsplit(z,"_")[[1]][3]
    #   na<-strsplit(na,"[.]")[[1]][1]
    #   da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    #   da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    # }
    # 
    # na=toString(year(Sys.Date())+1)
    # 
    # da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    # da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[3:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=2), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=2), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Plant Level - Attrition rate of BCA/BCAT/CA")
    p
  })
  
  output$table_plot_mor_bca_attrition<-renderTable({
    
    d <- morale_data()
    
    d<-d[d$Description=="Attrition rate of BCA/BCAT/CA",]
    
    dd<-d[d$Category=="Actual",]
    dd2<-d[d$Category=="Target",]
    da1<-sapply(dd[,4:15],mean,na.rm=TRUE)
    da<-sapply(dd2[,4:15],mean,na.rm=TRUE)
    
    d<-rbind(da,da1)
  })
  
  
  #Morale shop level bca_attrition
  
  data_bca_attrition_shop_act<-reactive({
    d <- morale_data()
    
    d<-d[d$Description=="Attrition rate of BCA/BCAT/CA",]
    
    dd<-d[d$Category=="Actual",]
    dd2<-d[d$Category=="Target",]
    
    da<-dd%>%gather(month,value,4:15)
    da$Category<-NULL
    da
  })
  data_bca_attrition_shop_tar<-reactive({
    d <- morale_data()
    
    d<-d[d$Description=="Attrition rate of BCA/BCAT/CA",]
    
    dd<-d[d$Category=="Actual",]
    dd2<-d[d$Category=="Target",]
    
    dt<-dd2%>%gather(month,value,4:15)
    
    dt$Category<-NULL
    dt
  })
  data_dept_bca_attrition_shop<-reactive({
    d <- morale_data()
    
    d<-d[d$Description=="Attrition rate of BCA/BCAT/CA",]
    
    
    d<-d%>%gather(month,value,4:15)
    d$month<-as.yearmon(d$month,"%b %Y")
    d<-d[months(d$month)==input$choose_comp_bca_attrition_shop,]
    d$KPI<-factor(d$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL"))
    
    d<-spread(d,KPI,value )
    
    d$month<-months(d$month)
    d
  })
  
  output$comp_bca_attrition_shop<-renderPlotly({
    te<-input$save_wc_kpi
    da<-data_bca_attrition_shop_act()
    dt<-data_bca_attrition_shop_tar()
    
    da$month<-as.yearmon(da$month,"%b %Y")
    dt$month<-as.yearmon(dt$month,"%b %Y")
    
    da1<-da[months(da$month)==input$choose_comp_bca_attrition_shop,]
    da<-dt[months(dt$month)==input$choose_comp_bca_attrition_shop,]
    
    da$KPI<-factor(da$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","IPL","QM","FM"))
    da1$KPI<-factor(da1$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","IPL","QM","FM"))
    
    da$yval<-as.numeric(da$value)
    da1$yval<-as.numeric(da1$value)
    
    da11<-da1[da1$value>da$value,]
    da12<-da1[da1$value<=da$value,]
    
    #if(nrow(da11)==0)
    #  da11<-da11%>%add_row(xval=na,yval=0)
    #if(nrow(da12)==0)
    #  da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$KPI,y = da$value, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$KPI,y=da11$value,text=round(da11$yval,digits=2),width=0.5, textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$KPI,y=da12$value,text=round(da12$yval,digits=2), width=0.5, textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
  })
  
  output$table_comp_bca_attrition_shop<-renderTable({
    data_dept_bca_attrition_shop()
  })
  
  
  
  
  
  #Morale con_attrition
  
  
  output$plot_mor_con_attrition<-renderPlotly({
    
    
    d <- morale_data()
    
    d<-d[d$Description=="Attrition rate of Contractors",]
    
    dd<-d[d$Category=="Required",]
    dd2<-d[d$Category=="Left",]
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd[,4:15],sum,na.rm=TRUE)))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd2[,4:15],sum,na.rm=TRUE)))))
    
    da<-da%>%add_row(xval="2020",yval=mean(da$yval))
    da1<-da1%>%add_row(xval="2020",yval=mean(da1$yval))
    
    
    
    
    da1$yval<-100*da1$yval/(da$yval)
    da$yval<-2
    
    
    da<-da%>%add_row(xval="2019",yval=2)
    da1<-da1%>%add_row(xval="2019",yval=7.1)
    
    da<-da%>%add_row(xval="2018",yval=0.8)
    da1<-da1%>%add_row(xval="2018",yval=3.6)
    
    da<-da%>%add_row(xval="2017",yval=0.8)
    da1<-da1%>%add_row(xval="2017",yval=17.3)
    
    da<-da%>%add_row(xval="2021",yval=2)
    da1<-da1%>%add_row(xval="2021",yval=0)
    
    # l<-list.files(path="con_attrition/")
    # 
    # for(z in l){
    #   na<-paste("con_attrition/",z,sep='')
    #   dt<-read_excel(na)
    #   dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[45:45,3:14]))))
    #   dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[46:46,3:14]))))
    # 
    #   na<-strsplit(z,"_")[[1]][3]
    #   na<-strsplit(na,"[.]")[[1]][1]
    #   da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    #   da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    # }
    # 
    # na=toString(year(Sys.Date())+1)
    # 
    # da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    # da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=NA,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=NA,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=2), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=2), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Plant Level - Attrition rate of Contractors")
    p
  })
  
  output$table_plot_mor_con_attrition<-renderTable({
    
    d <- morale_data()
    
    d<-d[d$Description=="Attrition rate of Contractors",]
    
    dd<-d[d$Category=="Required",]
    dd2<-d[d$Category=="Left",]
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd[,4:15],sum,na.rm=TRUE)))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd2[,4:15],sum,na.rm=TRUE)))))
    
    
    
    da1$yval<-100*da1$yval/(da$yval)
    da$yval<-2
    
    rownames(da)<-da$xval
    da$xval<-NULL
    da<-t(da)
    rownames(da)<-"Target"
    
    
    rownames(da1)<-da1$xval
    da1$xval<-NULL
    da1<-t(da1)
    rownames(da1)<-"Actual"
    
    
    
    d<-rbind((da),(da1))
    
    d
  })
  
  
  #Morale shop level con_attrition
  
  data_con_attrition_shop_act<-reactive({
    d <- morale_data()
    
    d<-d[d$Description=="Attrition rate of Contractors",]
    
    dd<-d[d$Category=="Rate",]
    dd2<-d[d$Category=="Target",]
    
    da<-dd%>%gather(month,value,4:15)
    da$Category<-NULL
    da
  })
  data_con_attrition_shop_tar<-reactive({
    d <- morale_data()
    
    d<-d[d$Description=="Attrition rate of Contractors",]
    
    dd<-d[d$Category=="Rate",]
    dd2<-d[d$Category=="Target",]
    
    dt<-dd2%>%gather(month,value,4:15)
    
    dt$Category<-NULL
    dt
  })
  data_dept_con_attrition_shop<-reactive({
    d <- morale_data()
    
    d<-d[d$Description=="Attrition rate of Contractors",]
    
    
    d<-d%>%gather(month,value,4:15)
    d$month<-as.yearmon(d$month,"%b %Y")
    d<-d[months(d$month)==input$choose_comp_con_attrition_shop,]
    d$KPI<-factor(d$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame"))
    
    d<-spread(d,KPI,value )
    
    d$month<-months(d$month)
    d
  })
  
  output$comp_con_attrition_shop<-renderPlotly({
    te<-input$save_wc_kpi
    da<-data_con_attrition_shop_act()
    dt<-data_con_attrition_shop_tar()
    
    da$month<-as.yearmon(da$month,"%b %Y")
    dt$month<-as.yearmon(dt$month,"%b %Y")
    
    da1<-da[months(da$month)==input$choose_comp_con_attrition_shop,]
    da<-dt[months(dt$month)==input$choose_comp_con_attrition_shop,]
    
    da$KPI<-factor(da$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","IPL","QM","FM","Frame"))
    da1$KPI<-factor(da1$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","IPL","QM","FM","Frame"))
    
    da$yval<-as.numeric(da$value)
    da1$yval<-as.numeric(da1$value)
    
    da11<-da1[da1$value<da$value,]
    da12<-da1[da1$value>=da$value,]
    
    #if(nrow(da11)==0)
    #  da11<-da11%>%add_row(xval=na,yval=0)
    #if(nrow(da12)==0)
    #  da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$KPI,y = da$value, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$KPI,y=da11$value,text=round(da11$yval,digits=2),width=0.5, textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$KPI,y=da12$value,text=round(da12$yval,digits=2), width=0.5, textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
  })
  
  output$table_comp_con_attrition_shop<-renderTable({
    data_dept_con_attrition_shop()
  })
  
  
  
  
  
  # Cost data
  
  cost_data<-reactive({
    
    d1<-read_excel("Chassis/KPI/chassis_kpi_2020.xlsx")
    d1<-d1[d1$KPI=="Cost",]
    d1$KPI<-"Chassis"
    
    d<-d1
    
    d1<-read_excel("Cabtrim/KPI/cabtrim_kpi_2020.xlsx")
    d1<-d1[d1$KPI=="Cost",]
    d1$KPI<-"Cabtrim"
    
    d<-rbind(d,d1)
    
    d1<-read_excel("EOL/KPI/eol_kpi_2020.xlsx")
    d1<-d1[d1$KPI=="Cost",]
    d1$KPI<-"EOL"
    
    d<-rbind(d,d1)
    
    d1<-read_excel("FBV/KPI/fbv_kpi_2020.xlsx")
    d1<-d1[d1$KPI=="Cost",]
    d1$KPI<-"FBV"
    
    d<-rbind(d,d1)
    
    d1<-read_excel("CIW/KPI/ciw_kpi_2020.xlsx")
    d1<-d1[d1$KPI=="Cost",]
    d1$KPI<-"CiW"
    
    d<-rbind(d,d1)
    
    d1<-read_excel("PAINT/KPI/paint_kpi_2020.xlsx")
    d1<-d1[d1$KPI=="Cost",]
    d1$KPI<-"Paint"
    
    d<-rbind(d,d1)
    
    d1<-read_excel("engine/KPI/engine_kpi_2020.xlsx")
    d1<-d1[d1$KPI=="Cost",]
    d1$KPI<-"Engine"
    
    d<-rbind(d,d1)
    
    d1<-read_excel("transmission/KPI/transmission_kpi_2020.xlsx")
    d1<-d1[d1$KPI=="Cost",]
    d1$KPI<-"Transmission"
    
    d<-rbind(d,d1)
    
    d1<-read_excel("ipl/KPI/ipl_kpi_2020.xlsx")
    d1<-d1[d1$KPI=="Cost",]
    d1$KPI<-"IPL"
    
    d<-rbind(d,d1)
    
    
    d1<-read_excel("QM/QM/QM_2020.xlsx")
    
    d1<-d1[d1$KPI=="Cost",]
    d1$KPI<-"QM"
    
    d<-rbind(d,d1)
    
    d1<-read_excel("fm/KPI/fm_kpi_2020.xlsx")
    
    d1<-d1[d1$KPI=="Cost",]
    d1$KPI<-"FM"
    
    d<-rbind(d,d1)
    
    
    d1<-read_excel("frame/KPI/frame_kpi_2020.xlsx")
    
    d1<-d1[d1$KPI=="Cost",]
    d1$KPI<-"Frame"
    
    d<-rbind(d,d1)
    
    d
  })
  
  
  
  
  #COst Indirect consumables
  
  
  output$plot_cos_indirect_cons<-renderPlotly({
    
    
    d <- cost_data()
    
    d<-d[d$Description=="Indirect Consumables",]
    
    dd<-d[d$Category=="Target",]
    dd2<-d[d$Category=="Actual",]
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd[,4:15],sum,na.rm=TRUE)))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd2[,4:15],sum,na.rm=TRUE)))))
    
    
    # l<-list.files(path="con_attrition/")
    # 
    # for(z in l){
    #   na<-paste("con_attrition/",z,sep='')
    #   dt<-read_excel(na)
    #   dtt<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[45:45,3:14]))))
    #   dta<-data.frame(xval=colnames(dt)[3:14],yval=c(t(array(dt[46:46,3:14]))))
    # 
    #   na<-strsplit(z,"_")[[1]][3]
    #   na<-strsplit(na,"[.]")[[1]][1]
    #   da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    #   da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    # }
    # 
    # na=toString(year(Sys.Date())+1)
    # 
    # da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    # da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=NA,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=NA,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=2), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=2), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Plant Level - Indirect Consumables")
    p
  })
  
  output$table_plot_cos_indirect_cons<-renderTable({
    
    d <- cost_data()
    
    d<-d[d$Description=="Indirect Consumables",]
    
    dd<-d[d$Category=="Target",]
    dd2<-d[d$Category=="Actual",]
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd[,4:15],sum,na.rm=TRUE)))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(sapply(dd2[,4:15],sum,na.rm=TRUE)))))
    
    
    rownames(da)<-da$xval
    da$xval<-NULL
    da<-t(da)
    rownames(da)<-"Target"
    
    
    rownames(da1)<-da1$xval
    da1$xval<-NULL
    da1<-t(da1)
    rownames(da1)<-"Actual"
    
    
    
    d<-rbind((da),(da1))
    
    d
  })
  
  
  #Morale shop level con_attrition
  
  data_cos_indirect_cons_shop_act<-reactive({
    d <- cost_data()
    
    d<-d[d$Description=="Indirect Consumables",]
    
    dd<-d[d$Category=="Actual",]
    dd2<-d[d$Category=="Target",]
    
    da<-dd%>%gather(month,value,4:15)
    da$Category<-NULL
    da
  })
  data_cos_indirect_cons_shop_tar<-reactive({
    d <- cost_data()
    
    d<-d[d$Description=="Indirect Consumables",]
    
    dd<-d[d$Category=="Actual",]
    dd2<-d[d$Category=="Target",]
    
    dt<-dd2%>%gather(month,value,4:15)
    
    dt$Category<-NULL
    dt
  })
  data_dept_cos_indirect_cons_shop<-reactive({
    d <- cost_data()
    
    d<-d[d$Description=="Indirect Consumables",]
    
    
    d<-d%>%gather(month,value,4:15)
    d$month<-as.yearmon(d$month,"%b %Y")
    d<-d[months(d$month)==input$choose_comp_cos_indirect_cons_shop,]
    d$KPI<-factor(d$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","QM","FM","IPL","Frame"))
    
    d<-spread(d,KPI,value )
    
    d$month<-months(d$month)
    d
  })
  
  output$comp_cos_indirect_cons_shop<-renderPlotly({
    te<-input$save_wc_kpi
    da<-data_cos_indirect_cons_shop_act()
    dt<-data_cos_indirect_cons_shop_tar()
    
    da$month<-as.yearmon(da$month,"%b %Y")
    dt$month<-as.yearmon(dt$month,"%b %Y")
    
    da1<-da[months(da$month)==input$choose_cos_indirect_cons_shop,]
    da<-dt[months(dt$month)==input$choose_cos_indirect_cons_shop,]
    
    da$KPI<-factor(da$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","IPL","QM","FM","Frame"))
    da1$KPI<-factor(da1$KPI,levels=c("Chassis","Cabtrim","EOL", "FBV","CiW","Paint","Engine","Transmission","IPL","QM","FM","Frame"))
    
    da$yval<-as.numeric(da$value)
    da1$yval<-as.numeric(da1$value)
    
    da11<-da1[da1$value>da$value,]
    da12<-da1[da1$value<=da$value,]
    
    #if(nrow(da11)==0)
    #  da11<-da11%>%add_row(xval=na,yval=0)
    #if(nrow(da12)==0)
    #  da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$KPI,y = da$value, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$KPI,y=da11$value,text=round(da11$yval,digits=2),width=0.5, textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$KPI,y=da12$value,text=round(da12$yval,digits=2), width=0.5, textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
  })
  
  output$table_comp_cos_indirect_cons_shop<-renderTable({
    data_dept_cos_indirect_cons_shop()
  })
  
  
  
  
  
  
  #Cost Rejection cost/truck
  
  
  output$plot_cos_rej_cost<-renderPlotly({
    
    te<-input$save_rej_cost
    f<-paste("QM/QM/QM_",input$choose_plot_cos_rej_cost,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[57:57,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[58:58,4:15]))))
    da$yval<-150
    
    l<-list.files(path="QM/QM/")
    
    for(z in l){
      na<-paste("QM/QM/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[57:57,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[58:58,4:15]))))
      dtt$yval<-150
      
      na<-strsplit(z,"_")[[1]][2]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Plant Level - Rejection cost/truck")
    p
  })
  
  output$table_plot_cos_rej_cost<-renderTable({
    
    te<-input$save_rej_cost
    f<-paste("QM/QM/QM_",input$choose_plot_cos_rej_cost,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[55:58,1:15]
    d
  })
  
  
  
  #Cost shop level rej/cost
  
  data_rej_cost_shop_act<-reactive({
    te<-input$save_rej_cost
    d<-read_excel("QM/QM/QM_2020.xlsx")
    d<-d[d$Description=="Rejection cost",]
    da<-d[d$Category!='Cost/truck',]
    da<-da[da$Category!='Volume',]
    da<-da[da$Category!='Plant level',]
    da<-da[da$Category!='Veh prod',]
    
    da<-da%>%gather(month,value,4:15)
    
    da
  })
  data_rej_cost_shop_tar<-reactive({
    te<-input$save_rej_cost
    d<-read_excel("QM/QM/QM_2020.xlsx")
    d<-d[d$Description=="Rejection cost",]
    dt<-d[d$Category!='Cost/truck',]
    dt<-dt[dt$Category!='Volume',]
    dt<-dt[dt$Category!='Plant level',]
    dt<-dt[dt$Category!='Veh prod',]
    
    
    dt<-dt%>%gather(month,value,4:15)
    
    dt
  })
  data_dept_rej_cost_shop<-reactive({
    te<-input$save_qm_kpi
    d<-read_excel("QM/QM/QM_2020.xlsx")
    d<-d[d$Description=="Rejection cost",]
    d<-d[d$Category!='Cost/truck',]
    d<-d[d$Category!='Volume',]
    d<-d[d$Category!='Plant level',]
    d<-d[d$Category!='Veh prod',]
    
    
    d<-d%>%gather(month,value,4:15)
    d$month<-as.yearmon(d$month,"%b %Y")
    d<-d[months(d$month)==input$choose_comp_rej_cost_shop,]
    d$Category<-factor(d$Category,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QA-Vehicle","QA-PTI"))
    dd<-d
    dd$value<-c(38000,19000,7000,0,1000,1000,32000,4000,1000,47000)
    d<-spread(d,Category,value )
    dd<-spread(dd,Category,value )
    d<-rbind(d,dd)
    d$month<-months(d$month)
    d
  })
  
  output$comp_rej_cost_shop<-renderPlotly({
    te<-input$save_qm_kpi
    da<-data_rej_cost_shop_act()
    dt<-data_rej_cost_shop_tar()
    
    da$month<-as.yearmon(da$month,"%b %Y")
    dt$month<-as.yearmon(dt$month,"%b %Y")
    
    da1<-da[months(da$month)==input$choose_comp_rej_cost_shop,]
    da<-dt[months(dt$month)==input$choose_comp_rej_cost_shop,]
    
    da$value<-c(38000,19000,7000,0,1000,1000,32000,4000,1000,47000)
    
    da$Category<-factor(da$Category,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QA-Vehicle","QA-PTI"))
    da1$Category<-factor(da1$Category,levels=c("Chassis","Cabtrim","EOL","FBV","CiW","Paint","Engine","Transmission","QA-Vehicle","QA-PTI"))
    
    da$yval<-as.numeric(da$value)
    da1$yval<-as.numeric(da1$value)
    
    da11<-da1[da1$value>da$value,]
    da12<-da1[da1$value<=da$value,]
    
    #if(nrow(da11)==0)
    #  da11<-da11%>%add_row(xval=na,yval=0)
    #if(nrow(da12)==0)
    #  da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$Category,y = da$value, showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$Category,y=da11$value,text=round(da11$yval,digits=1), textposition = 'auto',marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$Category,y=da12$value,text=round(da12$yval,digits=1), textposition = 'auto',marker=list(color='green'),name='Actual')%>%
      layout(hovermode = 'compare')
    p
  })
  
  output$table_comp_rej_cost_shop<-renderTable({
    data_dept_rej_cost_shop()
  })
  
  
  #Cost Electricity
  
  
  output$plot_cos_ele<-renderPlotly({
    
    te<-input$save_fm_kpi
    f<-paste("fm/KPI/fm_kpi_",input$choose_plot_cos_ele,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[13:13,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[14:14,4:15]))))
    
    l<-list.files(path="fm/KPI/")
    
    for(z in l){
      na<-paste("fm/KPI/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[13:13,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[14:14,4:15]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="Electricity Consumption")
    p
  })
  
  output$table_plot_cos_ele<-renderTable({
    
    te<-input$save_ele
    f<-paste("fm/KPI/fm_kpi_",input$choose_plot_cos_ele,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[13:14,1:15]
    d
  })
  
  #Cost proctricity
  
  
  output$plot_cos_pro<-renderPlotly({
    
    te<-input$save_fm_kpi
    f<-paste("fm/KPI/fm_kpi_",input$choose_plot_cos_pro,".xlsx",sep="")
    d <- read_excel(f)
    da<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[15:15,4:15]))))
    da1<-data.frame(xval=colnames(d)[4:15],yval=c(t(array(d[16:16,4:15]))))
    
    l<-list.files(path="fm/KPI/")
    
    for(z in l){
      na<-paste("fm/KPI/",z,sep='')
      dt<-read_excel(na)
      dtt<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[15:15,4:15]))))
      dta<-data.frame(xval=colnames(dt)[4:15],yval=c(t(array(dt[16:16,4:15]))))
      
      na<-strsplit(z,"_")[[1]][3]
      na<-strsplit(na,"[.]")[[1]][1]
      da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
      da1<-da1%>%add_row(xval=na,yval=mean(dta$yval,na.rm = TRUE))
    }
    
    na=toString(year(Sys.Date())+1)
    
    da<-da%>%add_row(xval=na,yval=mean(dtt$yval,na.rm = TRUE))
    da1<-da1%>%add_row(xval=na,yval=0)
    
    da1$xval<-factor(da1$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da$xval<-factor(da$xval,levels=c(2015:(year(Sys.Date())+1),colnames(d)[4:15]))
    da1$yval<-as.numeric(da1$yval)
    da$yval<-as.numeric(da$yval)
    
    
    da11<-da1[da1$yval>da$yval,]
    da12<-da1[da1$yval<=da$yval,]
    da<-da[order(da$xval),]
    
    if(nrow(da11)==0)
      da11<-da11%>%add_row(xval=na,yval=0)
    if(nrow(da12)==0)
      da12<-da12%>%add_row(xval=na,yval=0)
    
    p<-plot_ly(x=da$xval,y = da$yval,showlegend=FALSE,type = 'scatter',mode='lines+markers',name='Target')%>%
      add_trace(type='bar',x=da11$xval,y=da11$yval,text=round(da11$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='red'),name='Actual')%>%
      add_trace(type='bar',x=da12$xval,y=da12$yval,text=round(da12$yval,digits=1), textposition = 'auto',width=0.9,marker=list(color='green'),name='Actual')%>%
      layout(hovermode= 'compare',
             title="propane Consumption")
    p
  })
  
  output$table_plot_cos_pro<-renderTable({
    
    te<-input$save_pro
    f<-paste("fm/KPI/fm_kpi_",input$choose_plot_cos_pro,".xlsx",sep="")
    d <- read_excel(f)
    d<-d[13:14,1:15]
    d
  })
  
  
  
  
  
  #HPU
  
  #Cabtrim HPU input data
  
  values_hpu_cabtrim <- reactiveValues()
  
  
  previous_hpu_cabtrim <- reactive({
    d<-read_excel("HPU/Cabtrim/cabtrim_hpu_hdt_shift1_2020.xlsx")
    d
  })
  
  MyChanges_hpu_cabtrim <- reactive({
    if(is.null(input$hotable_hpu_cabtrim)){return(previous_hpu_cabtrim())}
    else if(!identical(previous_hpu_cabtrim(),input$hotable_hpu_cabtrim)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable_hpu_cabtrim <- as.data.frame(hot_to_r(input$hotable_hpu_cabtrim))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_hpu_cabtrim <- mytable_hpu_cabtrim[1:nrow(previous_hpu_cabtrim()),]
      
       for(i in 7:ncol(mytable_hpu_cabtrim))
        mytable_hpu_cabtrim[15,i]<-sum(mytable_hpu_cabtrim[3:14,i],na.rm=TRUE)
      
      
      for(i in 7:ncol(mytable_hpu_cabtrim))
        mytable_hpu_cabtrim[16,i]<-sum(mytable_hpu_cabtrim[3:14,i]*mytable_hpu_cabtrim[3:14,6],na.rm=TRUE)
      
      for(i in 7:ncol(mytable_hpu_cabtrim))
        mytable_hpu_cabtrim[22,i]<-mytable_hpu_cabtrim[1,i]*sum(mytable_hpu_cabtrim[17:20,i],na.rm=TRUE)
      
      for(i in 7:ncol(mytable_hpu_cabtrim))
        mytable_hpu_cabtrim[23,i]<-100*(mytable_hpu_cabtrim[22,i]-mytable_hpu_cabtrim[21,i])/mytable_hpu_cabtrim[22,i]
      
      for(i in 7:ncol(mytable_hpu_cabtrim))
        mytable_hpu_cabtrim[24,i]<-mytable_hpu_cabtrim[1,i]*mytable_hpu_cabtrim[17,i]*9*mytable_hpu_cabtrim[23,i]/(mytable_hpu_cabtrim[15,i]*100)
      
      for(i in 7:ncol(mytable_hpu_cabtrim))
        mytable_hpu_cabtrim[25,i]<-mytable_hpu_cabtrim[2,i]*mytable_hpu_cabtrim[17,i]*9*mytable_hpu_cabtrim[23,i]/(mytable_hpu_cabtrim[15,i]*100)
      
      for(i in 7:ncol(mytable_hpu_cabtrim))
        mytable_hpu_cabtrim[26,i]<-mytable_hpu_cabtrim[2,i]*(mytable_hpu_cabtrim[18,i]+mytable_hpu_cabtrim[19,i])*9*mytable_hpu_cabtrim[23,i]/(mytable_hpu_cabtrim[15,i]*100)
      
      for(i in 7:ncol(mytable_hpu_cabtrim))
        mytable_hpu_cabtrim[27,i]<-mytable_hpu_cabtrim[2,i]*mytable_hpu_cabtrim[17,i]*9*mytable_hpu_cabtrim[23,i]/(mytable_hpu_cabtrim[16,i]*100)
      
      for(i in 7:ncol(mytable_hpu_cabtrim))
        mytable_hpu_cabtrim[28,i]<-mytable_hpu_cabtrim[2,i]*(mytable_hpu_cabtrim[18,i]+mytable_hpu_cabtrim[19,i])*9*mytable_hpu_cabtrim[23,i]/(mytable_hpu_cabtrim[16,i]*100)
      
      
      mytable_hpu_cabtrim
    }
  })   
  
  output$hotable_hpu_cabtrim<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30)+3)
    row_highlight = c(15,16,22:28)-1
    row_readonly=c(15,16,22:28)
    
    rhandsontable(MyChanges_hpu_cabtrim(), col_highlight = col_highlight, row_highlight = row_highlight,width = 1250,height = 600)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+4),readOnly = TRUE)%>%
      hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_hpu_cabtrim,{
    #
    write.xlsx(hot_to_r(input$hotable_hpu_cabtrim),"HPU/Cabtrim/cabtrim_hpu_hdt_shift1_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #FTT SPR input data
  
  values_fttspr <- reactiveValues()
  
  
  previous_fttspr <- reactive({
    d<-read_excel("spr_ftt/spr_ftt_2020.xlsx")
    d
  })
  
  MyChanges_fttspr <- reactive({
    if(is.null(input$hotable_fttspr)){return(previous_fttspr())}
    else if(!identical(previous_fttspr(),input$hotable_fttspr)){
      # hot.to.df function will convert your updated table into the dataframe
      mytable_fttspr <- as.data.frame(hot_to_r(input$hotable_fttspr))
      # here the second column is a function of the first and it will be multipled by 100 given the values in the first column
      mytable_fttspr <- mytable_fttspr[1:nrow(previous_fttspr()),]
      
   
      
      mytable_fttspr
    }
  })   
  
  output$hotable_fttspr<-renderRHandsontable({
    col_highlight = 0:(month(Sys.Date()-30)+1)
    row_highlight = c()
    row_readonly=c()
    
    rhandsontable(MyChanges_fttspr(), col_highlight = col_highlight, width = 1250,height = 600)%>%
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE)%>%
      hot_col(1:(month(Sys.Date()-30)+2),readOnly = TRUE)%>%
      #hot_row(row_readonly, readOnly = TRUE) %>%
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
            }", fixedColumnsLeft=3)
    
  })
  observeEvent(input$save_fttspr,{
    #
    write.xlsx(hot_to_r(input$hotable_fttspr),"spr_ftt/spr_ftt_2020.xlsx",row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  
  
  #major
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="major")
  updateTextInput(session,"enab_major",value=d$enabler[n])
  updateTextInput(session,"task_major",value=d$keytask[n])
  
  observeEvent(input$save_comm_major,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="major")
    d$enabler[n]<-input$enab_major
    d$keytask[n]<-input$task_major
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  }) 
  
  #major_shop
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="major_shop")
  updateTextInput(session,"enab_major_shop",value=d$enabler[n])
  updateTextInput(session,"task_major_shop",value=d$keytask[n])
  
  observeEvent(input$save_comm_major_shop,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="major_shop")
    d$enabler[n]<-input$enab_major_shop
    d$keytask[n]<-input$task_major_shop
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  }) 
  
  
  #minor
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="minor")
  updateTextInput(session,"enab_minor",value=d$enabler[n])
  updateTextInput(session,"task_minor",value=d$keytask[n])
  
  observeEvent(input$save_comm_minor,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="minor")
    d$enabler[n]<-input$enab_minor
    d$keytask[n]<-input$task_minor
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  }) 
  
  #minor_shop
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="minor_shop")
  updateTextInput(session,"enab_minor_shop",value=d$enabler[n])
  updateTextInput(session,"task_minor_shop",value=d$keytask[n])
  
  observeEvent(input$save_comm_minor_shop,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="minor_shop")
    d$enabler[n]<-input$enab_minor_shop
    d$keytask[n]<-input$task_minor_shop
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  }) 
  
  #firstaid
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="firstaid")
  updateTextInput(session,"enab_firstaid",value=d$enabler[n])
  updateTextInput(session,"task_firstaid",value=d$keytask[n])
  
  observeEvent(input$save_comm_firstaid,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="firstaid")
    d$enabler[n]<-input$enab_firstaid
    d$keytask[n]<-input$task_firstaid
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  }) 
  
  #firstaid_shop
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="firstaid_shop")
  updateTextInput(session,"enab_firstaid_shop",value=d$enabler[n])
  updateTextInput(session,"task_firstaid_shop",value=d$keytask[n])
  
  observeEvent(input$save_comm_firstaid_shop,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="firstaid_shop")
    d$enabler[n]<-input$enab_firstaid_shop
    d$keytask[n]<-input$task_firstaid_shop
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  }) 
  
  
  #counter
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="counter")
  updateTextInput(session,"enab_counter",value=d$enabler[n])
  updateTextInput(session,"task_counter",value=d$keytask[n])
  
  observeEvent(input$save_comm_counter,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="counter")
    d$enabler[n]<-input$enab_counter
    d$keytask[n]<-input$task_counter
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  }) 
  
  #counter_shop
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="counter_shop")
  updateTextInput(session,"enab_counter_shop",value=d$enabler[n])
  updateTextInput(session,"task_counter_shop",value=d$keytask[n])
  
  observeEvent(input$save_comm_counter_shop,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="counter_shop")
    d$enabler[n]<-input$enab_counter_shop
    d$keytask[n]<-input$task_counter_shop
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  }) 
  
  #unsafe
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="unsafe")
  updateTextInput(session,"enab_unsafe",value=d$enabler[n])
  updateTextInput(session,"task_unsafe",value=d$keytask[n])
  
  observeEvent(input$save_comm_unsafe,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="unsafe")
    d$enabler[n]<-input$enab_unsafe
    d$keytask[n]<-input$task_unsafe
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  }) 
  
  #unsafe_shop
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="unsafe_shop")
  updateTextInput(session,"enab_unsafe_shop",value=d$enabler[n])
  updateTextInput(session,"task_unsafe_shop",value=d$keytask[n])
  
  observeEvent(input$save_comm_unsafe_shop,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="unsafe_shop")
    d$enabler[n]<-input$enab_unsafe_shop
    d$keytask[n]<-input$task_unsafe_shop
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  }) 
  
  #dpu_hdt_ops
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="dpu_hdt_ops")
  updateTextInput(session,"enab_dpu_hdt_ops",value=d$enabler[n])
  updateTextInput(session,"task_dpu_hdt_ops",value=d$keytask[n])
  
  observeEvent(input$save_comm_dpu_hdt_ops,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="dpu_hdt_ops")
    d$enabler[n]<-input$enab_dpu_hdt_ops
    d$keytask[n]<-input$task_dpu_hdt_ops
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #dpu_hdt_ops_ab
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="dpu_hdt_ops_ab")
  updateTextInput(session,"enab_dpu_hdt_ops_ab",value=d$enabler[n])
  updateTextInput(session,"task_dpu_hdt_ops_ab",value=d$keytask[n])
  
  observeEvent(input$save_comm_dpu_hdt_ops_ab,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="dpu_hdt_ops_ab")
    d$enabler[n]<-input$enab_dpu_hdt_ops_ab
    d$keytask[n]<-input$task_dpu_hdt_ops_ab
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #dpu_mdt_ops
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="dpu_mdt_ops")
  updateTextInput(session,"enab_dpu_mdt_ops",value=d$enabler[n])
  updateTextInput(session,"task_dpu_mdt_ops",value=d$keytask[n])
  
  observeEvent(input$save_comm_dpu_mdt_ops,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="dpu_mdt_ops")
    d$enabler[n]<-input$enab_dpu_mdt_ops
    d$keytask[n]<-input$task_dpu_mdt_ops
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
 
  #dpu_mdt_ops_ab
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="dpu_mdt_ops_ab")
  updateTextInput(session,"enab_dpu_mdt_ops_ab",value=d$enabler[n])
  updateTextInput(session,"task_dpu_mdt_ops_ab",value=d$keytask[n])
  
  observeEvent(input$save_comm_dpu_mdt_ops_ab,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="dpu_mdt_ops_ab")
    d$enabler[n]<-input$enab_dpu_mdt_ops_ab
    d$keytask[n]<-input$task_dpu_mdt_ops_ab
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  }) 
  
  #dpu_hdt_qfl2
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="dpu_hdt_qfl2")
  updateTextInput(session,"enab_dpu_hdt_qfl2",value=d$enabler[n])
  updateTextInput(session,"task_dpu_hdt_qfl2",value=d$keytask[n])
  
  observeEvent(input$save_comm_dpu_hdt_qfl2,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="dpu_hdt_qfl2")
    d$enabler[n]<-input$enab_dpu_hdt_qfl2
    d$keytask[n]<-input$task_dpu_hdt_qfl2
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #dpu_mdt_qfl2
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="dpu_mdt_qfl2")
  updateTextInput(session,"enab_dpu_mdt_qfl2",value=d$enabler[n])
  updateTextInput(session,"task_dpu_mdt_qfl2",value=d$keytask[n])
  
  observeEvent(input$save_comm_dpu_mdt_qfl2,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="dpu_mdt_qfl2")
    d$enabler[n]<-input$enab_dpu_mdt_qfl2
    d$keytask[n]<-input$task_dpu_mdt_qfl2
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #dpu_eng
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="dpu_eng")
  updateTextInput(session,"enab_dpu_eng",value=d$enabler[n])
  updateTextInput(session,"task_dpu_eng",value=d$keytask[n])
  
  observeEvent(input$save_comm_dpu_eng,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="dpu_eng")
    d$enabler[n]<-input$enab_dpu_eng
    d$keytask[n]<-input$task_dpu_eng
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #dpu_tra
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="dpu_tra")
  updateTextInput(session,"enab_dpu_tra",value=d$enabler[n])
  updateTextInput(session,"task_dpu_tra",value=d$keytask[n])
  
  observeEvent(input$save_comm_dpu_tra,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="dpu_tra")
    d$enabler[n]<-input$enab_dpu_tra
    d$keytask[n]<-input$task_dpu_tra
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #dpu_ciw
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="dpu_ciw")
  updateTextInput(session,"enab_dpu_ciw",value=d$enabler[n])
  updateTextInput(session,"task_dpu_ciw",value=d$keytask[n])
  
  observeEvent(input$save_comm_dpu_ciw,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="dpu_ciw")
    d$enabler[n]<-input$enab_dpu_ciw
    d$keytask[n]<-input$task_dpu_ciw
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #dpu_pai
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="dpu_pai")
  updateTextInput(session,"enab_dpu_pai",value=d$enabler[n])
  updateTextInput(session,"task_dpu_pai",value=d$keytask[n])
  
  observeEvent(input$save_comm_dpu_pai,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="dpu_pai")
    d$enabler[n]<-input$enab_dpu_pai
    d$keytask[n]<-input$task_dpu_pai
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #dpu_fra
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="dpu_fra")
  updateTextInput(session,"enab_dpu_fra",value=d$enabler[n])
  updateTextInput(session,"task_dpu_fra",value=d$keytask[n])
  
  observeEvent(input$save_comm_dpu_fra,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="dpu_fra")
    d$enabler[n]<-input$enab_dpu_fra
    d$keytask[n]<-input$task_dpu_fra
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #ftt_hdt
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="ftt_hdt")
  updateTextInput(session,"enab_ftt_hdt",value=d$enabler[n])
  updateTextInput(session,"task_ftt_hdt",value=d$keytask[n])
  
  observeEvent(input$save_comm_ftt_hdt,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="ftt_hdt")
    d$enabler[n]<-input$enab_ftt_hdt
    d$keytask[n]<-input$task_ftt_hdt
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #ftt_mdt
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="ftt_mdt")
  updateTextInput(session,"enab_ftt_mdt",value=d$enabler[n])
  updateTextInput(session,"task_ftt_mdt",value=d$keytask[n])
  
  observeEvent(input$save_comm_ftt_mdt,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="ftt_mdt")
    d$enabler[n]<-input$enab_ftt_mdt
    d$keytask[n]<-input$task_ftt_mdt
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  #ftt_ldt
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="ftt_ldt")
  updateTextInput(session,"enab_ftt_ldt",value=d$enabler[n])
  updateTextInput(session,"task_ftt_ldt",value=d$keytask[n])
  
  observeEvent(input$save_comm_ftt_ldt,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="ftt_ldt")
    d$enabler[n]<-input$enab_ftt_ldt
    d$keytask[n]<-input$task_ftt_ldt
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #spr_hdt
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="spr_hdt")
  updateTextInput(session,"enab_spr_hdt",value=d$enabler[n])
  updateTextInput(session,"task_spr_hdt",value=d$keytask[n])
  
  observeEvent(input$save_comm_spr_hdt,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="spr_hdt")
    d$enabler[n]<-input$enab_spr_hdt
    d$keytask[n]<-input$task_spr_hdt
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #spr_mdt
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="spr_mdt")
  updateTextInput(session,"enab_spr_mdt",value=d$enabler[n])
  updateTextInput(session,"task_spr_mdt",value=d$keytask[n])
  
  observeEvent(input$save_comm_spr_mdt,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="spr_mdt")
    d$enabler[n]<-input$enab_spr_mdt
    d$keytask[n]<-input$task_spr_mdt
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  #spr_ldt
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="spr_ldt")
  updateTextInput(session,"enab_spr_ldt",value=d$enabler[n])
  updateTextInput(session,"task_spr_ldt",value=d$keytask[n])
  
  observeEvent(input$save_comm_spr_ldt,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="spr_ldt")
    d$enabler[n]<-input$enab_spr_ldt
    d$keytask[n]<-input$task_spr_ldt
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #roll_hdt
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="roll_hdt")
  updateTextInput(session,"enab_roll_hdt",value=d$enabler[n])
  updateTextInput(session,"task_roll_hdt",value=d$keytask[n])
  
  observeEvent(input$save_comm_roll_hdt,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="roll_hdt")
    d$enabler[n]<-input$enab_roll_hdt
    d$keytask[n]<-input$task_roll_hdt
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #roll_mdt
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="roll_mdt")
  updateTextInput(session,"enab_roll_mdt",value=d$enabler[n])
  updateTextInput(session,"task_roll_mdt",value=d$keytask[n])
  
  observeEvent(input$save_comm_roll_mdt,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="roll_mdt")
    d$enabler[n]<-input$enab_roll_mdt
    d$keytask[n]<-input$task_roll_mdt
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #cap_hdt
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="cap_hdt")
  updateTextInput(session,"enab_cap_hdt",value=d$enabler[n])
  updateTextInput(session,"task_cap_hdt",value=d$keytask[n])
  
  observeEvent(input$save_comm_cap_hdt,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="cap_hdt")
    d$enabler[n]<-input$enab_cap_hdt
    d$keytask[n]<-input$task_cap_hdt
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #cap_mdt
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="cap_mdt")
  updateTextInput(session,"enab_cap_mdt",value=d$enabler[n])
  updateTextInput(session,"task_cap_mdt",value=d$keytask[n])
  
  observeEvent(input$save_comm_cap_mdt,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="cap_mdt")
    d$enabler[n]<-input$enab_cap_mdt
    d$keytask[n]<-input$task_cap_mdt
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #forecasted
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="forecasted")
  updateTextInput(session,"enab_forecasted",value=d$enabler[n])
  updateTextInput(session,"task_forecasted",value=d$keytask[n])
  
  observeEvent(input$save_comm_forecasted,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="forecasted")
    d$enabler[n]<-input$enab_forecasted
    d$keytask[n]<-input$task_forecasted
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #loss_ope
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="loss_ope")
  updateTextInput(session,"enab_loss_ope",value=d$enabler[n])
  updateTextInput(session,"task_loss_ope",value=d$keytask[n])
  
  observeEvent(input$save_comm_loss_ope,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="loss_ope")
    d$enabler[n]<-input$enab_loss_ope
    d$keytask[n]<-input$task_loss_ope
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #loss_agg
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="loss_agg")
  updateTextInput(session,"enab_loss_agg",value=d$enabler[n])
  updateTextInput(session,"task_loss_agg",value=d$keytask[n])
  
  observeEvent(input$save_comm_loss_agg,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="loss_agg")
    d$enabler[n]<-input$enab_loss_agg
    d$keytask[n]<-input$task_loss_agg
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  
  #cons_plant
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="cons_plant")
  updateTextInput(session,"enab_cons_plant",value=d$enabler[n])
  updateTextInput(session,"task_cons_plant",value=d$keytask[n])
  
  observeEvent(input$save_comm_cons_plant,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="cons_plant")
    d$enabler[n]<-input$enab_cons_plant
    d$keytask[n]<-input$task_cons_plant
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #cons_shop
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="cons_shop")
  updateTextInput(session,"enab_cons_shop",value=d$enabler[n])
  updateTextInput(session,"task_cons_shop",value=d$keytask[n])
  
  observeEvent(input$save_comm_cons_shop,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="cons_shop")
    d$enabler[n]<-input$enab_cons_shop
    d$keytask[n]<-input$task_cons_shop
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #hpu_cap_plant
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="hpu_cap_plant")
  updateTextInput(session,"enab_hpu_cap_plant",value=d$enabler[n])
  updateTextInput(session,"task_hpu_cap_plant",value=d$keytask[n])
  
  observeEvent(input$save_comm_hpu_cap_plant,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="hpu_cap_plant")
    d$enabler[n]<-input$enab_hpu_cap_plant
    d$keytask[n]<-input$task_hpu_cap_plant
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #hpu_cap_shop
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="hpu_cap_shop")
  updateTextInput(session,"enab_hpu_cap_shop",value=d$enabler[n])
  updateTextInput(session,"task_hpu_cap_shop",value=d$keytask[n])
  
  observeEvent(input$save_comm_hpu_cap_shop,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="hpu_cap_shop")
    d$enabler[n]<-input$enab_hpu_cap_shop
    d$keytask[n]<-input$task_hpu_cap_shop
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #electricity
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="electricity")
  updateTextInput(session,"enab_electricity",value=d$enabler[n])
  updateTextInput(session,"task_electricity",value=d$keytask[n])
  
  observeEvent(input$save_comm_electricity,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="electricity")
    d$enabler[n]<-input$enab_electricity
    d$keytask[n]<-input$task_electricity
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #propane
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="propane")
  updateTextInput(session,"enab_propane",value=d$enabler[n])
  updateTextInput(session,"task_propane",value=d$keytask[n])
  
  observeEvent(input$save_comm_propane,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="propane")
    d$enabler[n]<-input$enab_propane
    d$keytask[n]<-input$task_propane
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #rej_plant
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="rej_plant")
  updateTextInput(session,"enab_rej_plant",value=d$enabler[n])
  updateTextInput(session,"task_rej_plant",value=d$keytask[n])
  
  observeEvent(input$save_comm_rej_plant,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="rej_plant")
    d$enabler[n]<-input$enab_rej_plant
    d$keytask[n]<-input$task_rej_plant
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #rej_shop
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="rej_shop")
  updateTextInput(session,"enab_rej_shop",value=d$enabler[n])
  updateTextInput(session,"task_rej_shop",value=d$keytask[n])
  
  observeEvent(input$save_comm_rej_shop,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="rej_shop")
    d$enabler[n]<-input$enab_rej_shop
    d$keytask[n]<-input$task_rej_shop
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #att_white_plant
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="att_white_plant")
  updateTextInput(session,"enab_att_white_plant",value=d$enabler[n])
  updateTextInput(session,"task_att_white_plant",value=d$keytask[n])
  
  observeEvent(input$save_comm_att_white_plant,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="att_white_plant")
    d$enabler[n]<-input$enab_att_white_plant
    d$keytask[n]<-input$task_att_white_plant
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #att_white_shop
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="att_white_shop")
  updateTextInput(session,"enab_att_white_shop",value=d$enabler[n])
  updateTextInput(session,"task_att_white_shop",value=d$keytask[n])
  
  observeEvent(input$save_comm_att_white_shop,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="att_white_shop")
    d$enabler[n]<-input$enab_att_white_shop
    d$keytask[n]<-input$task_att_white_shop
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #att_bca_plant
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="att_bca_plant")
  updateTextInput(session,"enab_att_bca_plant",value=d$enabler[n])
  updateTextInput(session,"task_att_bca_plant",value=d$keytask[n])
  
  observeEvent(input$save_comm_att_bca_plant,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="att_bca_plant")
    d$enabler[n]<-input$enab_att_bca_plant
    d$keytask[n]<-input$task_att_bca_plant
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #att_bca_shop
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="att_bca_shop")
  updateTextInput(session,"enab_att_bca_shop",value=d$enabler[n])
  updateTextInput(session,"task_att_bca_shop",value=d$keytask[n])
  
  observeEvent(input$save_comm_att_bca_shop,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="att_bca_shop")
    d$enabler[n]<-input$enab_att_bca_shop
    d$keytask[n]<-input$task_att_bca_shop
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #att_baca_plant
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="att_baca_plant")
  updateTextInput(session,"enab_att_baca_plant",value=d$enabler[n])
  updateTextInput(session,"task_att_baca_plant",value=d$keytask[n])
  
  observeEvent(input$save_comm_att_baca_plant,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="att_baca_plant")
    d$enabler[n]<-input$enab_att_baca_plant
    d$keytask[n]<-input$task_att_baca_plant
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #att_baca_shop
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="att_baca_shop")
  updateTextInput(session,"enab_att_baca_shop",value=d$enabler[n])
  updateTextInput(session,"task_att_baca_shop",value=d$keytask[n])
  
  observeEvent(input$save_comm_att_baca_shop,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="att_baca_shop")
    d$enabler[n]<-input$enab_att_baca_shop
    d$keytask[n]<-input$task_att_baca_shop
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #att_con_plant
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="att_con_plant")
  updateTextInput(session,"enab_att_con_plant",value=d$enabler[n])
  updateTextInput(session,"task_att_con_plant",value=d$keytask[n])
  
  observeEvent(input$save_comm_att_con_plant,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="att_con_plant")
    d$enabler[n]<-input$enab_att_con_plant
    d$keytask[n]<-input$task_att_con_plant
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  
  #att_con_shop
  f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
  d<-read_excel(f)
  n<-which(d$KPI=="att_con_shop")
  updateTextInput(session,"enab_att_con_shop",value=d$enabler[n])
  updateTextInput(session,"task_att_con_shop",value=d$keytask[n])
  
  observeEvent(input$save_comm_att_con_shop,{
    f<-paste("comments/comments_",year(Sys.Date()-30),"_",month(Sys.Date()-30),".xlsx",sep="")
    d<-read_excel(f)
    n<-which(d$KPI=="att_con_shop")
    d$enabler[n]<-input$enab_att_con_shop
    d$keytask[n]<-input$task_att_con_shop
    write.xlsx(as.data.frame(d),f,row.names = FALSE)
    shinyalert(title = "Successfully saved", type = "success",closeOnClickOutside=TRUE)
  })
  
  #Morale
  
  output$slickr_morale_att_con <- renderSlickR({
    x_morale_att_con<-input$myFile_morale_att_con
    x_morale_att_con1<-input$delete_slickr_morale_att_con
    imgs_morale_att_con <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/att_con/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_morale_att_con)
  })
  observeEvent(input$delete_slickr_morale_att_con, {
    imgs_morale_att_con <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/att_con/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_morale_att_con)
  })
  observeEvent(input$myFile_morale_att_con, {
    inFile <- input$myFile_morale_att_con
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/att_con/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/att_con/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_morale_att_bca <- renderSlickR({
    x_morale_att_bca<-input$myFile_morale_att_bca
    x_morale_att_bca1<-input$delete_slickr_morale_att_bca
    imgs_morale_att_bca <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/att_bca/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_morale_att_bca)
  })
  observeEvent(input$delete_slickr_morale_att_bca, {
    imgs_morale_att_bca <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/att_bca/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_morale_att_bca)
  })
  observeEvent(input$myFile_morale_att_bca, {
    inFile <- input$myFile_morale_att_bca
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/att_bca/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/att_bca/",sep=""), nam) ,overwrite = TRUE)
  })
  
  
  output$slickr_morale_kai_ca <- renderSlickR({
    x_morale_kai_ca<-input$myFile_morale_kai_ca
    x_morale_kai_ca1<-input$delete_slickr_morale_kai_ca
    imgs_morale_kai_ca <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/kai_ca/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_morale_kai_ca)
  })
  observeEvent(input$delete_slickr_morale_kai_ca, {
    imgs_morale_kai_ca <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/kai_ca/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_morale_kai_ca)
  })
  observeEvent(input$myFile_morale_kai_ca, {
    inFile <- input$myFile_morale_kai_ca
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/kai_ca/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/kai_ca/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_morale_kai_bca <- renderSlickR({
    x_morale_kai_bca<-input$myFile_morale_kai_bca
    x_morale_kai_bca1<-input$delete_slickr_morale_kai_bca
    imgs_morale_kai_bca <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/kai_bca/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_morale_kai_bca)
  })
  observeEvent(input$delete_slickr_morale_kai_bca, {
    imgs_morale_kai_bca <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/kai_bca/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_morale_kai_bca)
  })
  observeEvent(input$myFile_morale_kai_bca, {
    inFile <- input$myFile_morale_kai_bca
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/kai_bca/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/kai_bca/",sep=""), nam) ,overwrite = TRUE)
  })
  
  
  
  output$slickr_morale_white <- renderSlickR({
    x_morale_white<-input$myFile_morale_white
    x_morale_white1<-input$delete_slickr_morale_white
    imgs_morale_white <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/white/",sep=""), pattern=".png", full.names = TRUE)

    slickR(imgs_morale_white)
  })
  observeEvent(input$delete_slickr_morale_white, {
    imgs_morale_white <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/white/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_morale_white)
  })
  observeEvent(input$myFile_morale_white, {
    inFile <- input$myFile_morale_white
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/white/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/morale/white/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_cost_hpucapacity <- renderSlickR({
    x_cost_hpucapacity<-input$myFile_cost_hpucapacity
    x_cost_hpucapacity1<-input$delete_slickr_cost_hpucapacity
    imgs_cost_hpucapacity <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/hpucapacity/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_cost_hpucapacity)
  })
  observeEvent(input$delete_slickr_cost_hpucapacity, {
    imgs_cost_hpucapacity <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/hpucapacity/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_cost_hpucapacity)
  })
  observeEvent(input$myFile_cost_hpucapacity, {
    inFile <- input$myFile_cost_hpucapacity
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/hpucapacity/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/hpucapacity/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_cost_indirect <- renderSlickR({
    x_cost_indirect<-input$myFile_cost_indirect
    x_cost_indirect1<-input$delete_slickr_cost_indirect
    imgs_cost_indirect <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/indirect/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_cost_indirect)
  })
  observeEvent(input$delete_slickr_cost_indirect, {
    imgs_cost_indirect <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/indirect/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_cost_indirect)
  })
  observeEvent(input$myFile_cost_indirect, {
    inFile <- input$myFile_cost_indirect
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/indirect/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/indirect/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_cost_rej <- renderSlickR({
    x_cost_rej<-input$myFile_cost_rej
    x_cost_rej1<-input$delete_slickr_cost_rej
    imgs_cost_rej <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/rej/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_cost_rej)
  })
  observeEvent(input$delete_slickr_cost_rej, {
    imgs_cost_rej <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/rej/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_cost_rej)
  })
  observeEvent(input$myFile_cost_rej, {
    inFile <- input$myFile_cost_rej
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/rej/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/rej/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_cost_ele <- renderSlickR({
    x_cost_ele<-input$myFile_cost_ele
    x_cost_ele<-input$delete_slickr_cost_ele
    imgs_cost_ele <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/ele/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_cost_ele)
  })
  observeEvent(input$delete_slickr_cost_ele, {
    imgs_cost_ele <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/ele/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_cost_ele)
  })
  observeEvent(input$myFile_cost_ele, {
    inFile <- input$myFile_cost_ele
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/ele/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/cost/ele/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_del_qcok <- renderSlickR({
    x_del_qcok<-input$myFile_del_qcok
    x_del_qcok1<-input$delete_slickr_del_qcok
    imgs_del_qcok <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/qcok/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_del_qcok)
  })
  observeEvent(input$delete_slickr_del_qcok, {
    imgs_del_qcok <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/qcok/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_del_qcok)
  })
  observeEvent(input$myFile_del_qcok, {
    inFile <- input$myFile_del_qcok
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/qcok/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/qcok/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_del_roll <- renderSlickR({
    x_del_roll<-input$myFile_del_roll
    x_del_roll1<-input$delete_slickr_del_roll
    imgs_del_roll <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/roll/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_del_roll)
  })
  observeEvent(input$delete_slickr_del_roll, {
    imgs_del_roll <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/roll/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_del_roll)
  })
  observeEvent(input$myFile_del_roll, {
    inFile <- input$myFile_del_roll
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/roll/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/roll/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_del_cap <- renderSlickR({
    x_del_cap<-input$myFile_del_cap
    x_del_cap1<-input$delete_slickr_del_cap
    imgs_del_cap <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/cap/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_del_cap)
  })
  observeEvent(input$delete_slickr_del_cap, {
    imgs_del_cap <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/cap/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_del_cap)
  })
  observeEvent(input$myFile_del_cap, {
    inFile <- input$myFile_del_cap
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/cap/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/cap/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_del_fore <- renderSlickR({
    x_del_fore<-input$myFile_del_fore
    x_del_fore1<-input$delete_slickr_del_fore
    imgs_del_fore <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/fore/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_del_fore)
  })
  observeEvent(input$delete_slickr_del_fore, {
    imgs_del_fore <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/fore/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_del_fore)
  })
  observeEvent(input$myFile_del_fore, {
    inFile <- input$myFile_del_fore
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/fore/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/fore/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_del_loss <- renderSlickR({
    x_del_loss<-input$myFile_del_loss
    x_del_loss1<-input$delete_slickr_del_loss
    imgs_del_loss <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/loss/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_del_loss)
  })
  observeEvent(input$delete_slickr_del_loss, {
    imgs_del_loss <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/loss/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_del_loss)
  })
  observeEvent(input$myFile_del_loss, {
    inFile <- input$myFile_del_loss
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/loss/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/delivery/loss/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_qua_qfl4 <- renderSlickR({
    x_qua_qfl4<-input$myFile_qua_qfl4
    x_qua_qfl41<-input$delete_slickr_qua_qfl4
    imgs_qua_qfl4 <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/qfl4/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_qua_qfl4)
  })
  observeEvent(input$delete_slickr_qua_qfl4, {
    imgs_qua_qfl4 <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/qfl4/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_qua_qfl4)
  })
  observeEvent(input$myFile_qua_qfl4, {
    inFile <- input$myFile_qua_qfl4
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/qfl4/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/qfl4/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_qua_tear <- renderSlickR({
    x_qua_tear<-input$myFile_qua_tear
    x_qua_tear<-input$delete_slickr_qua_tear
    imgs_qua_tear <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/tear/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_qua_tear)
  })
  observeEvent(input$delete_slickr_qua_tear, {
    imgs_qua_tear <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/tear/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_qua_tear)
  })
  observeEvent(input$myFile_qua_tear, {
    inFile <- input$myFile_qua_tear
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/tear/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/tear/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_qua_ops_hdt <- renderSlickR({
    x_qua_ops_hdt<-input$myFile_qua_ops_hdt
    x_qua_ops_hdt1<-input$delete_slickr_qua_ops_hdt
    imgs_qua_ops_hdt <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/ops_hdt/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_qua_ops_hdt)
  })
  observeEvent(input$delete_slickr_qua_ops_hdt, {
    imgs_qua_ops_hdt <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/ops_hdt/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_qua_ops_hdt)
  })
  observeEvent(input$myFile_qua_ops_hdt, {
    inFile <- input$myFile_qua_ops_hdt
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/ops_hdt/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/ops_hdt/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_qua_ops_mdt <- renderSlickR({
    x_qua_ops_mdt<-input$myFile_qua_ops_mdt
    x_qua_ops_mdt1<-input$delete_slickr_qua_ops_mdt
    imgs_qua_ops_mdt <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/ops_mdt/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_qua_ops_mdt)
  })
  observeEvent(input$delete_slickr_qua_ops_mdt, {
    imgs_qua_ops_mdt <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/ops_mdt/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_qua_ops_mdt)
  })
  
  observeEvent(input$myFile_qua_ops_mdt, {
    inFile <- input$myFile_qua_ops_mdt
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/ops_mdt/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/ops_mdt/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_qua_qfl2 <- renderSlickR({
    x_qua_qfl2<-input$myFile_qua_qfl2
    x_qua_qfl21<-input$delete_slickr_qua_qfl2
    imgs_qua_qfl2 <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/qfl2/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_qua_qfl2)
  })
  observeEvent(input$delete_slickr_qua_qfl2, {
    imgs_qua_qfl2 <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/qfl2/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_qua_qfl2)
  })
  observeEvent(input$myFile_qua_qfl2, {
    inFile <- input$myFile_qua_qfl2
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/qfl2/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/qfl2/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_qua_ftt <- renderSlickR({
    x_qua_ftt<-input$myFile_qua_ftt
    x_qua_ftt1<-input$delete_slickr_qua_ftt
    imgs_qua_ftt <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/ftt/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_qua_ftt)
  })
  observeEvent(input$delete_slickr_qua_ftt, {
    imgs_qua_ftt <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/ftt/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_qua_ftt)
  })
  
  observeEvent(input$myFile_qua_ftt, {
    inFile <- input$myFile_qua_ftt
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/ftt/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/ftt/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_qua_spr <- renderSlickR({
    x_qua_spr<-input$myFile_qua_spr
    x_qua_spr1<-input$delete_slickr_qua_spr
    imgs_qua_spr <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/spr/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_qua_spr)
  })
  observeEvent(input$delete_slickr_qua_spr, {
    imgs_qua_spr <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/spr/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_qua_spr)
  })
  observeEvent(input$myFile_qua_spr, {
    inFile <- input$myFile_qua_spr
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/spr/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/quality/spr/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_major <- renderSlickR({
    x_major<-input$myFile_major
    xx<-input$delete_slickr_major
    imgs_major <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/major/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_major)
  })
  
  observeEvent(input$myFile_major, {
    inFile <- input$myFile_major
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/major/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/major/",sep=""), nam) ,overwrite = TRUE)
  })
  
  observeEvent(input$delete_slickr_major, {
    imgs_major <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/major/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_major)
  })
  
  output$slickr_minor <- renderSlickR({
    x_minor<-input$myFile_minor
    x_minor<-input$delete_slickr_minor
    imgs_minor <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/minor/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_minor)
  })
  observeEvent(input$delete_slickr_minor, {
    imgs_minor <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/minor/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_minor)
  })
  observeEvent(input$myFile_minor, {
    inFile <- input$myFile_minor
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/minor/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/minor/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_firstaid <- renderSlickR({
    x_firstaid<-input$myFile_firstaid
    x_firstaid<-input$delete_slickr_firstaid
    imgs_firstaid <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/firstaid/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_firstaid)
  })
  observeEvent(input$delete_slickr_firstaid, {
    imgs_firstaid <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/firstaid/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_firstaid)
  })
  observeEvent(input$myFile_firstaid, {
    inFile <- input$myFile_firstaid
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/firstaid/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/firstaid/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_unsafe <- renderSlickR({
    x_unsafe<-input$myFile_unsafe
    x_unsafe<-input$delete_slickr_unsafe
    imgs_unsafe <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/unsafe/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_unsafe)
  })
  observeEvent(input$delete_slickr_unsafe, {
    imgs_unsafe <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/unsafe/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_unsafe)
  })
  observeEvent(input$myFile_unsafe, {
    inFile <- input$myFile_unsafe
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/unsafe/",sep="")))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/unsafe/",sep=""), nam) ,overwrite = TRUE)
  })
  
  output$slickr_counter <- renderSlickR({
    x_counter<-input$myFile_counter
    x_counter1<-input$delete_slickr_counter
    imgs_counter <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/counter/",sep=""), pattern=".png", full.names = TRUE)
    
    slickR(imgs_counter)
  })
  observeEvent(input$delete_slickr_counter, {
    imgs_counter <- list.files(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/counter/",sep=""), pattern=".png", full.names = TRUE)
    file.remove(imgs_counter)
  })
  observeEvent(input$myFile_counter, {
    inFile <- input$myFile_counter
    if (is.null(inFile))
      return()
    le<-length(list.files(path=paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/counter/"),sep=""))
    nam<-paste(le+1,".png",sep="")
    file.copy(inFile$datapath, file.path(paste("one_pager/c_",year(Sys.Date()-30),"_",month(Sys.Date()-30),"/safety/counter/",sep=""), nam) ,overwrite = TRUE)
  })
  
  
  
  
}



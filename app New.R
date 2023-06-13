options(shiny.maxRequestSize=100*1024^2)
options(
  gargle_oauth_email = TRUE,
  gargle_oauth_cache = ".secrets"
)

library(shiny)
library(semantic.dashboard)
library(dplyr)
library(readr)
library(readxl)
library(dplyr)
library(stringr)
library(tidyr)
library(openxlsx)
library(shinydashboard)
library(DT)
library(xlsx)
library(plotly)
library("webshot")
library("googledrive")
library(easycsv)
library("tools")
library(shinyTime)
library("shinyjs")
hours_in_day=8

get_mapping <- function(){
  Mapping <- read_excel("Input Files/Mapping.xlsx")
  Mapping <- Mapping[,c("JIRA Projects Name","Cluster Group Name","PMO Projects Name")]
  colnames(Mapping)[colnames(Mapping)=="Cluster Group Name"] = "Groups"
  
  return(Mapping)
}
get_table_style<- function(wb, sheet_number, row_number,col_number){
  openxlsx::setColWidths(wb, sheet = sheet_number, cols = LETTERS[0:col_number+1], widths = c(20, rep(10, col_number)) )
  # header style
  header_col <-
    createStyle(fontSize = 12, fontColour = "black", halign = "left",  textDecoration = "bold")
  addStyle(wb, sheet = sheet_number, header_col, rows = 1, cols = 0:col_number+1, gridExpand = TRUE)
  table_body_style1 <-
    createStyle(fontSize = 10, fgFill = "#DAEEF3")
  addStyle(wb, sheet = sheet_number, table_body_style1, rows = seq(2, row_number+1, by = 2), cols = 0:col_number+1, gridExpand = TRUE,stack=TRUE)
  table_body_style2 <-
    createStyle(fontSize = 10, fgFill = "#8DB4E2")
  addStyle(wb, sheet = sheet_number, table_body_style2, rows = seq(1, row_number+1, by = 2), cols = 0:col_number+1, gridExpand = TRUE,stack=TRUE)
  
  return(wb)
}
pre_process_jira_data  <- function(jira_data){
  if (nrow(jira_data[jira_data$`Activity Logs`=='',])>0) jira_data[jira_data$`Activity Logs`=='',]$`Activity Logs`="No"
  if (nrow(jira_data[!str_detect(jira_data$`Activity Logs`, "No"),])>0) jira_data[!str_detect(jira_data$`Activity Logs`, "No"),]$`Activity Logs` = "Yes"
  if (nrow(jira_data[str_detect(jira_data$`Activity Logs`, "itemId"),])>0) jira_data[str_detect(jira_data$`Activity Logs`, "itemId"),]$`Activity Logs` = "Yes"
  
  if (nrow(jira_data[which(jira_data$`Activity Type` == "Admin- Leave" ),])>0) jira_data[which(jira_data$`Activity Type` == "Admin- Leave" ),]$`Activity Logs`= "Yes"
  if (nrow(jira_data[which(jira_data$`Activity Type` == "Admin- Leave" ),])>0) jira_data[which(jira_data$`Activity Type` == "Admin- Leave" ),]$`Activity Logs`= "Yes"
  
  if ( sum(is.na(jira_data$`Work Description`))>0) jira_data[is.na(jira_data$`Work Description`),]$`Work Description` <- ""
  if (nrow(jira_data[str_detect(jira_data$`Work Description`, "JIRALog:"),])>0) jira_data[str_detect(jira_data$`Work Description`, "JIRALog:"),]$`Activity Logs` = "Yes"
  
  jira_data$`Emp Code` = apply(data.frame(jira_data$`Full name`),1,function(x)gsub('.*-', '',x))
  jira_data$`Emp Code` = trimws(jira_data$`Emp Code`)
  
  colnames(jira_data)[colnames(jira_data)=="Activity Logs"] = "JLA Activity Log"
  
  return(jira_data)
}
pre_process_pmo_data<- function(alloc_data){
  ######Remove duplicate 100 % allocations 
  temp= alloc_data %>% group_by( `Emp Code`,)%>%
    summarise(`%Allocated` = sum(`%Allocated`))
  if( length(temp[temp$`%Allocated`>100,]$`Emp Code`) > 0)
  {
    alloc_100=filter( alloc_data,`Emp Code` %in% temp[temp$`%Allocated`<=100,]$`Emp Code` )
    
    alloc_more_100=alloc_dups = filter( alloc_data,`Emp Code` %in% temp[temp$`%Allocated`>100,]$`Emp Code` )
    alloc_dups$rn=1:nrow(alloc_dups)
    emp_list=unique(alloc_dups$`Emp Code`)
    for (emp in emp_list )
    { 
      emp_df= filter( alloc_dups,`Emp Code` %in% emp)
      emp_df= emp_df[order(emp_df$`Finish Date`, decreasing = TRUE),]
      alloc_=0
      for( e in emp_df$rn)
      {
        if((alloc_+alloc_dups[e,]$`%Allocated`) <= 100)
        {
          alloc_=alloc_+alloc_dups[e,]$`%Allocated`
        }else{
          alloc_more_100[e,]=NA
        }
      }
      
      
    }
    
    alloc_more_100=alloc_more_100[complete.cases(alloc_more_100$`Emp Code`),]
    
    alloc_data = rbind(alloc_more_100, alloc_100)
    
  }
  
  alloc_data =alloc_data[,c("Groups", "Resource Name" , "Emp Code","Team","Reporting TL","%Allocated","Project Name","Email","Product")]
  alloc_data$`Emp Code` =  apply(data.frame(alloc_data$`Emp Code`),1,function(x)gsub('.*-', '',x))
  ######Adding Allocation % for allocation records to same clusters
  alloc_data = alloc_data %>% group_by(`Groups`, `Emp Code`,Team,`Project Name` ,`Resource Name`,`Email`, `Product` )%>% 
    summarise(`%Allocated` = sum(`%Allocated`))
  colnames(alloc_data)[colnames(alloc_data)=="Project Name"] = "PMO Projects Name"
  
  return(alloc_data)
}
calculate_resource_allocation <- function(cluster_wise_resource_logging,JLA_total_hours,Hours){
  ##Calculation of `Hours required per allocation`
  cluster_wise_resource_logging$`Available Hours` = (cluster_wise_resource_logging$`%Allocated` /100)*Hours
  
  # Hours logged on another cluster
  total_hours= cluster_wise_resource_logging %>% group_by( `Emp Code`) %>% 
    summarise( `Sum of Worked-Total`= sum(Jira_hours, na.rm=TRUE))
  
  cluster_wise_resource_logging = merge(cluster_wise_resource_logging,total_hours, all.x = TRUE)
  
  # Total JLA hours logged on another cluster
  colnames(JLA_total_hours)[colnames(JLA_total_hours)=="JLA_hours"] = "JLA Hours-Total"
  cluster_wise_resource_logging = merge(cluster_wise_resource_logging,JLA_total_hours, all.x = TRUE)
  
  #remove jira info and only keep PMO allocated cluster info
  cluster_wise_resource_logging=cluster_wise_resource_logging[complete.cases(cluster_wise_resource_logging$`%Allocated`),]
  
  
  
  ########Here `Sum of Worked - Total`=x field cannot be summed up for summary because
  ########for resources who have multiple allocation they would x field duplicated for each row so for that 
  ########duplicate rows would have to be removed first
  
  return(cluster_wise_resource_logging)
}

calculate_cluster_summary <- function(cluster_wise_resource_logging,Hours )
{
  
  ##total hours logged on jira calculation; dividing hours per alloc
  temp=cluster_wise_resource_logging
  temp$`Sum of Worked-Total per alloc`= ((temp$`%Allocated`)/100)*temp$`Sum of Worked-Total`
  temp$`JLA Hours-Total per alloc`= ((temp$`%Allocated`)/100)*temp$`JLA Hours-Total`
  temp[is.na(temp)]=0
  ##Capping
  # temp$`JLA_hours_cp` = temp$`JLA_hours`
  # 
  # temp$`Jira_hours_cp` = temp$`Jira_hours`
  # # 
  # temp$`JLA Hours-Total per alloc_cp` = temp$`JLA Hours-Total per alloc`
  # 
  # temp$`Sum of Worked-Total per alloc_cp` = temp$`Sum of Worked-Total per alloc`
  
  temp$`JLA_hours_cp` = ifelse (temp$`JLA_hours`>temp$`Available Hours`
                                ,temp$`Available Hours`
                                ,temp$`JLA_hours`)
  
  temp$`Jira_hours_cp` = ifelse (temp$`Jira_hours`>temp$`Available Hours`
                                 ,temp$`Available Hours`,
                                 temp$`Jira_hours`)
  #
  temp$`JLA Hours-Total per alloc_cp` = ifelse (temp$`JLA Hours-Total per alloc`>temp$`Available Hours`, temp$`Available Hours`
                                                ,temp$`JLA Hours-Total per alloc`)
  
  temp$`Sum of Worked-Total per alloc_cp` = ifelse (temp$`Sum of Worked-Total per alloc`>temp$`Available Hours`,
                                                    temp$`Available Hours`,
                                                    temp$`Sum of Worked-Total per alloc`)
  
  cluster_wise_summary = temp %>% group_by(`Groups`)%>%
    summarise(`JLA Hours-Cluster` = sum(`JLA_hours_cp`, na.rm=TRUE),
              `Sum of Worked-Cluster` = sum(`Jira_hours_cp`, na.rm=TRUE),
              `Sum of Worked-Total` = sum(`Sum of Worked-Total per alloc_cp`, na.rm=TRUE),
              `JLA Hours-Total` = sum(`JLA Hours-Total per alloc_cp`, na.rm=TRUE),
              `Available Hours` = sum(`Available Hours`, na.rm=TRUE),
              `Head Count` = sum(`%Allocated`, na.rm=TRUE)/100
              
    )
  # cluster_wise_summary$`Available Hours` =  round(cluster_wise_summary$`Head Count`)*Hours
  
  cluster_wise_summary$`Jira Delta` =  cluster_wise_summary$`Sum of Worked-Total`-cluster_wise_summary$`Sum of Worked-Cluster`
  cluster_wise_summary$`JLA Delta` =  cluster_wise_summary$`JLA Hours-Total`-cluster_wise_summary$`JLA Hours-Cluster`
  
  # cluster_wise_summary$`Daily %age` = 
  #   round((cluster_wise_summary$`Sum of Worked-Total`/cluster_wise_summary$`Available Hours`)*100,0)
  # cluster_wise_summary$`JLA %age` = round((cluster_wise_summary$`JLA Hours-Total`/cluster_wise_summary$`Available Hours`)*100,0)
  # 
  cluster_wise_summary$`Daily %age` = 
    (cluster_wise_summary$`Sum of Worked-Total`/cluster_wise_summary$`Available Hours`)*100
  cluster_wise_summary$`JLA %age` = (cluster_wise_summary$`JLA Hours-Total`/cluster_wise_summary$`Available Hours`)*100
  
  
  #Grand Total
  cluster_wise_summary[nrow(cluster_wise_summary)+1,] = NA
  cluster_wise_summary$Groups[nrow(cluster_wise_summary)] = "Grand Total"
  cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",-1] =
    as.list(colSums(cluster_wise_summary[,-1], na.rm = TRUE))
  
  
  cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",]$`Daily %age`=100*(
    as.numeric(cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",]$`Sum of Worked-Total`)/as.numeric(cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",]$`Available Hours`))
  
  cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",]$`JLA %age` =100*(
    as.numeric(cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",]$`JLA Hours-Total`)/as.numeric(cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",]$`Available Hours`))
  
  # cluster_wise_summary[,-1]= round(cluster_wise_summary[,-1],0)
  return(cluster_wise_summary)
}  

calculate_team_summary<- function(cluster_wise_resource_logging,Hours){
  temp=cluster_wise_resource_logging
  temp$`Sum of Worked-Total per alloc`= ((temp$`%Allocated`)/100)*temp$`Sum of Worked-Total`
  temp$`JLA Hours-Total per alloc`= ((temp$`%Allocated`)/100)*temp$`JLA Hours-Total`
  temp[is.na(temp)]=0
  ##Capping
  temp$`JLA_hours_cp` = ifelse (temp$`JLA_hours`>temp$`Available Hours`
                                ,temp$`Available Hours`
                                ,temp$`JLA_hours`)
  
  temp$`Jira_hours_cp` = ifelse (temp$`Jira_hours`>temp$`Available Hours`
                                 ,temp$`Available Hours`,
                                 temp$`Jira_hours`)
  # 
  temp$`JLA Hours-Total per alloc_cp` = ifelse (temp$`JLA Hours-Total per alloc`>temp$`Available Hours`, temp$`Available Hours`
                                                ,temp$`JLA Hours-Total per alloc`)
  
  temp$`Sum of Worked-Total per alloc_cp` = ifelse (temp$`Sum of Worked-Total per alloc`>temp$`Available Hours`, 
                                                    temp$`Available Hours`,
                                                    temp$`Sum of Worked-Total per alloc`)
  
  team_wise_summary = temp %>% group_by(`Team`)%>%
    summarise(`JLA Hours-Cluster` = sum(`JLA_hours_cp`, na.rm=TRUE),
              `Sum of Worked-Cluster` = sum(`Jira_hours_cp`, na.rm=TRUE),
              `Sum of Worked-Total` = sum(`Sum of Worked-Total per alloc_cp`, na.rm=TRUE),
              `JLA Hours-Total` = sum(`JLA Hours-Total per alloc_cp`, na.rm=TRUE),
              `Available Hours` = sum(`Available Hours`, na.rm=TRUE),
              `Head Count` = sum(`%Allocated`, na.rm=TRUE)/100
              
    )
  # team_wise_summary$`Head Count` =  round(team_wise_summary$`Available Hours`/Hours)
  # team_wise_summary$`Available Hours` =  round(team_wise_summary$`Head Count`)*Hours
  
  
  team_wise_summary$`Daily %age` = 
    round((team_wise_summary$`Sum of Worked-Total`/team_wise_summary$`Available Hours`)*100,0)
  team_wise_summary$`JLA %age` =
    round((team_wise_summary$`JLA Hours-Total`/team_wise_summary$`Available Hours`)*100,0)
  
  
  
  #Grand Total
  team_wise_summary[nrow(team_wise_summary)+1,] = NA
  team_wise_summary$Team[nrow(team_wise_summary)] = "Grand Total"
  
  
  team_wise_summary[team_wise_summary$Team=="Grand Total",-1] = 
    as.list(colSums(team_wise_summary[,-1], na.rm = TRUE))
  # as.numeric(apply(team_wise_summary[,-1], 2,FUN = function(x) sum(as.numeric(x), na.rm = TRUE)))
  
  
  team_wise_summary[team_wise_summary$Team=="Grand Total",]$`Daily %age`=100*(
    as.numeric(team_wise_summary[team_wise_summary$Team=="Grand Total",]$`Sum of Worked-Total`)/as.numeric(team_wise_summary[team_wise_summary$Team=="Grand Total",]$`Available Hours`))
  
  team_wise_summary[team_wise_summary$Team=="Grand Total",]$`JLA %age` =100*(
    as.numeric(team_wise_summary[team_wise_summary$Team=="Grand Total",]$`JLA Hours-Total`)/as.numeric(team_wise_summary[team_wise_summary$Team=="Grand Total",]$`Available Hours`))
  
  # team_wise_summary[,-1]= round(team_wise_summary[,-1],0)
  
  return(team_wise_summary)
}

calculate_project_summary<- function(cluster_wise_resource_logging,Hours){
  # Project Summary Statistics -------------------------------------------------
  temp=cluster_wise_resource_logging
  temp$`Sum of Worked-Total per alloc`= ((temp$`%Allocated`)/100)*temp$`Sum of Worked-Total`
  temp$`JLA Hours-Total per alloc`= ((temp$`%Allocated`)/100)*temp$`JLA Hours-Total`
  temp[is.na(temp)]=0
  ##Capping
  temp$`JLA_hours_cp` = ifelse (temp$`JLA_hours`>temp$`Available Hours`
                                ,temp$`Available Hours`
                                ,temp$`JLA_hours`)
  
  temp$`Jira_hours_cp` = ifelse (temp$`Jira_hours`>temp$`Available Hours`
                                 ,temp$`Available Hours`,
                                 temp$`Jira_hours`)
  # 
  temp$`JLA Hours-Total per alloc_cp` = ifelse (temp$`JLA Hours-Total per alloc`>temp$`Available Hours`, temp$`Available Hours`
                                                ,temp$`JLA Hours-Total per alloc`)
  
  temp$`Sum of Worked-Total per alloc_cp` = ifelse (temp$`Sum of Worked-Total per alloc`>temp$`Available Hours`, 
                                                    temp$`Available Hours`,
                                                    temp$`Sum of Worked-Total per alloc`)
  
  project_wise_summary = temp %>% group_by(Groups,`PMO Projects Name`)%>%
    summarise(`JLA Hours-Cluster` = sum(`JLA_hours_cp`, na.rm=TRUE),
              `Sum of Worked-Cluster` = sum(`Jira_hours_cp`, na.rm=TRUE),
              `Sum of Worked-Total` = sum(`Sum of Worked-Total per alloc_cp`, na.rm=TRUE),
              `JLA Hours-Total` = sum(`JLA Hours-Total per alloc_cp`, na.rm=TRUE),
              `Available Hours` = sum(`Available Hours`, na.rm=TRUE),
              `Head Count` = sum(`%Allocated`, na.rm=TRUE)/100
              
    )
  # project_wise_summary$`Head Count` =  round(project_wise_summary$`Available Hours`/Hours)
  # project_wise_summary$`Available Hours` =  round(project_wise_summary$`Head Count`)*Hours
  
  project_wise_summary$`Daily %age` = 
    round((project_wise_summary$`Sum of Worked-Total`/project_wise_summary$`Available Hours`)*100,0)
  project_wise_summary$`JLA %age` = round((project_wise_summary$`JLA Hours-Total`/project_wise_summary$`Available Hours`)*100,0)
  
  
  
  #Grand Total
  project_wise_summary[nrow(project_wise_summary)+1,] = NA
  project_wise_summary$`PMO Projects Name`[nrow(project_wise_summary)] = "Grand Total"
  
  
  project_wise_summary[project_wise_summary$`PMO Projects Name`=="Grand Total",-(1:2)] = 
    as.list(colSums(project_wise_summary[,-(1:2)], na.rm = TRUE))
  # as.numeric(apply(project_wise_summary[,-1], 2,FUN = function(x) sum(as.numeric(x), na.rm = TRUE)))
  
  
  project_wise_summary[project_wise_summary$`PMO Projects Name`=="Grand Total",]$`Daily %age`=100*(
    as.numeric(project_wise_summary[project_wise_summary$`PMO Projects Name`=="Grand Total",]$`Sum of Worked-Total`)/as.numeric(project_wise_summary[project_wise_summary$`PMO Projects Name`=="Grand Total",]$`Available Hours`))
  
  project_wise_summary[project_wise_summary$`PMO Projects Name`=="Grand Total",]$`JLA %age` =100*(
    as.numeric(project_wise_summary[project_wise_summary$`PMO Projects Name`=="Grand Total",]$`JLA Hours-Total`)/as.numeric(project_wise_summary[project_wise_summary$`PMO Projects Name`=="Grand Total",]$`Available Hours`))
  
  # project_wise_summary[,-(1:2)]= round(project_wise_summary[,-(1:2)],0)
  project_wise_summary$Groups[nrow(project_wise_summary)]="Grand Total"
  
  return(project_wise_summary)
}
get_cluster_activity <- function(jira_activity_data, type, alloc_data ){
  
  ref_df <- read_excel("Input Files/Jira Reference sheet.xlsx")
  
  proj_group= data.frame( matrix(nrow=nrow(ref_df),ncol=2))
  proj_group =ref_df[,c("Projects","Group")]
  proj_group=proj_group[complete.cases(proj_group),]
  
  CLUSTERS = unique(proj_group$Group) 
  CLUSTERS = CLUSTERS[CLUSTERS!="JIRA Service Desk"]
  
  Activity_df= data.frame( matrix(nrow=nrow(ref_df),ncol=2))
  Activity_df =ref_df[,c("Activity Type","Heads")]
  Activity_df=Activity_df[complete.cases(Activity_df),]
  
  df =  merge(jira_activity_data,proj_group, by.x = "Project Name", by.y="Projects")
  df= merge(df ,Activity_df )
  
  
  if(type=="emp_code")
  {
    df$Hours = as.numeric(df$Hours)
    temp=df %>% group_by(Group, `Emp Code`,Heads) %>%summarise(Hours = sum(Hours, na.rm = TRUE))
    cluster_wise_activity_seg=temp %>% 
      mutate(rn = row_number()) %>%
      spread(Heads,Hours) %>%
      select(-rn)
    cluster_wise_activity_seg[is.na(cluster_wise_activity_seg)] <- 0
    
    cluster_wise_activity_seg=cluster_wise_activity_seg %>% group_by(Group, `Emp Code`) %>%summarise(across(everything(), list(sum)))
    colnames(cluster_wise_activity_seg)=gsub("_1", "", colnames(cluster_wise_activity_seg))

    non_leaves = cluster_wise_activity_seg[cluster_wise_activity_seg$Group!="Leaves",]
    # # non_leaves_ =merge(non_leaves[,-1], alloc_data[,c("Emp Code", "Group")],by=c("Emp Code") )
    # 
    # for (i in 1:nrow(non_leaves)) {
    #   emp_code=non_leaves$`Emp Code`[i]
    #   non_leaves$Group[i]=alloc_data[alloc_data$`Emp Code` == emp_code,c("Group")]
    #   
    # }
    # 
    cluster_mapping=read.csv("Input Files/cluster_mapping.csv")
    leaves =cluster_wise_activity_seg[cluster_wise_activity_seg$Group=="Leaves",]
    colnames(alloc_data)[colnames(alloc_data)=="Groups"] = "Group"
    leaves_ =merge(leaves[,-1], alloc_data[,c("Emp Code", "Group")],by=c("Emp Code") )
    for (i in 1:nrow(cluster_mapping)) {
      # emp_code=leaves_$`Emp Code`[i]
      cluster=cluster_mapping$`PMOCluster`[i]
      # print(leaves_[leaves_$Group == cluster,c("Group")])
      # print( cluster_mapping[cluster_mapping$`PMOCluster` == cluster,c("JiraCluster")])
      leaves_[leaves_$Group == cluster,c("Group")] = cluster_mapping[cluster_mapping$`PMOCluster` == cluster,c("JiraCluster")]

    }

    
    
    cluster_wise_activity_seg=rbind(non_leaves, leaves_)
  }else{
    temp=df %>% group_by(Group,Heads) %>%summarise(Hours = sum(Hours, na.rm = TRUE))
    cluster_wise_activity_seg=temp %>% 
      mutate(rn = row_number()) %>%
      spread(Heads,Hours) %>%
      select(-rn)
    cluster_wise_activity_seg[is.na(cluster_wise_activity_seg)] <- 0
    
    cluster_wise_activity_seg=cluster_wise_activity_seg %>% group_by(Group) %>%summarise(across(everything(), list(sum)))
    colnames(cluster_wise_activity_seg)=gsub("_1", "", colnames(cluster_wise_activity_seg))
    ###Grand total per activity
    cluster_wise_activity_seg[(nrow(cluster_wise_activity_seg)+1),]=NA
    cluster_wise_activity_seg$Group[nrow(cluster_wise_activity_seg)] = c("Grand Total")
    cluster_wise_activity_seg[nrow(cluster_wise_activity_seg),2:ncol(cluster_wise_activity_seg)]= 
      lapply(cluster_wise_activity_seg[-nrow(cluster_wise_activity_seg),2:ncol(cluster_wise_activity_seg)],sum)
    
    cluster_wise_activity_seg[,2:ncol(cluster_wise_activity_seg)]=round(cluster_wise_activity_seg[,2:ncol(cluster_wise_activity_seg)])
    ###Grand total per Cluster
    cluster_wise_activity_seg$`Grand Total` = rowSums( cluster_wise_activity_seg[,-1])
    ###per activity percentages 
    cluster_wise_activity_seg[(nrow(cluster_wise_activity_seg)+1),]=NA
    cluster_wise_activity_seg$Group[nrow(cluster_wise_activity_seg)] = "%"
    
    grand_total = cluster_wise_activity_seg[cluster_wise_activity_seg$Group=="Grand Total",]$`Grand Total`
    # activity="Admin Activities"
    for (activity in colnames(cluster_wise_activity_seg)[-1])
    {
      cluster_wise_activity_seg[cluster_wise_activity_seg$Group=="%",activity] =
        (cluster_wise_activity_seg[cluster_wise_activity_seg$Group=="Grand Total",activity]/grand_total)*100
      
      
    }
    cluster_wise_activity_seg[,-1]= round(cluster_wise_activity_seg[,-1],2)
    
  }
  
  return(cluster_wise_activity_seg)
}

week_summary<- function(days,hours_in_day,HC, weekly_summary ){
  
  #Grand Total
  weekly_summary[nrow(weekly_summary)+1,] = NA
  weekly_summary$Groups[nrow(weekly_summary)] = "Grand Total"
  
  
  weekly_summary[weekly_summary$Groups=="Grand Total",-1] = 
    as.list(colSums(weekly_summary[,-1], na.rm = TRUE))
  
  for (i in 1:length(days)){
    
    d=days[i]
    
    jira_total=paste("Sum of Worked-Total", d,sep="|")
    jla_total=paste("JLA Hours-Total", d,sep="|")
    jira_perc=paste("Daily %age", d,sep="|")
    jla_perc= paste("JLA %age", d,sep="|")
    
    work_done = as.numeric(weekly_summary[weekly_summary$Groups=="Grand Total",jira_total])
    head_count = as.numeric(weekly_summary[weekly_summary$Groups=="Grand Total",]$`Available Hours`)
    weekly_summary[weekly_summary$Groups=="Grand Total",jira_perc]=100*(work_done/head_count)
    
    weekly_summary[weekly_summary$Groups=="Grand Total",jla_perc] =100*(
      as.numeric(weekly_summary[weekly_summary$Groups=="Grand Total",jla_total])/as.numeric(weekly_summary[weekly_summary$Groups=="Grand Total",]$`Available Hours`))
    
    
  }
  
  
  
  # Cluster Summary Statistics -------------------------------------------------
  jira_total=paste("Sum of Worked-Total", days,sep="|")
  jla_total=paste("JLA Hours-Total", days,sep="|")
  
  weekly_summary$`Sum of Worked-Total` =rowSums(weekly_summary[ ,jira_total],na.rm = TRUE)
  weekly_summary$`JLA Hours-Total` =rowSums(weekly_summary[ ,jla_total],na.rm = TRUE)
  
  
  cluster_wise_summary=weekly_summary[,c("Groups","Sum of Worked-Total","JLA Hours-Total")]
  cluster_wise_summary = merge(HC,cluster_wise_summary, by=c("Groups"))
  # HC
  
  # cluster_wise_summary$`Available Hours` =  cluster_wise_summary$`Head Count` * hours_in_day*4
  cluster_wise_summary$`Available Hours` =  cluster_wise_summary$`Head Count` * hours_in_day*5
  
  cluster_wise_summary$`Weekly %age` = 
    (cluster_wise_summary$`Sum of Worked-Total`/cluster_wise_summary$`Available Hours`)*100
  cluster_wise_summary$`JLA %age` = (cluster_wise_summary$`JLA Hours-Total`/cluster_wise_summary$`Available Hours`)*100
  
  
  #Grand Total
  cluster_wise_summary[nrow(cluster_wise_summary)+1,] = NA
  cluster_wise_summary$Groups[nrow(cluster_wise_summary)] = "Grand Total"
  cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",-1] =
    as.list(colSums(cluster_wise_summary[,-1], na.rm = TRUE))
  
  
  cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",]$`Weekly %age`=100*(
    as.numeric(cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",]$`Sum of Worked-Total`)/as.numeric(cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",]$`Available Hours`))
  
  cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",]$`JLA %age` =100*(
    as.numeric(cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",]$`JLA Hours-Total`)/as.numeric(cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",]$`Available Hours`))
  
  
  cws_order=c("Groups","Head Count","Available Hours","Sum of Worked-Total","JLA Hours-Total","Weekly %age","JLA %age")
  cluster_wise_summary= cluster_wise_summary[,cws_order]
  
  ##format weekly_summary data a bit
  weekly_summary[is.na(weekly_summary)]=0
  weekly_summary=weekly_summary[,1:(ncol(weekly_summary)-2)]
  
  colnames(weekly_summary)[4:ncol(weekly_summary)]=
    substr(colnames(weekly_summary)[4:ncol(weekly_summary)], 1,  
           nchar(colnames(weekly_summary)[4:ncol(weekly_summary)])-11)
  
  return(list("weekly_summary"=weekly_summary,"cluster_wise_summary"=cluster_wise_summary))
  
  
  
}

print_CWS<- function(cluster_wise_summary,report_type){
  if(report_type=="Mid-day")
  {
    colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="Sum of Worked-Total"] = "Sum of Worked"
    colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="JLA Hours-Total"] = "JLA Hours"
    
    colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="Daily %age"] = "Mid-Day %age"
    cws_order=c("Groups","Head Count","Available Hours","Sum of Worked","JLA Hours","Mid-Day %age","JLA %age")
  }
  else{
    cws_order=c("Groups","Head Count","Available Hours","Sum of Worked-Cluster","Sum of Worked-Total", "Jira Delta",
                "JLA Hours-Cluster","JLA Hours-Total","JLA Delta" ,"Daily %age","JLA %age")
    
  }
  
  cluster_wise_summary =cluster_wise_summary[,cws_order]
  
  return(cluster_wise_summary)
  
}
print_TWS<- function(cluster_wise_summary,Hours){
  # if(Hours!=8)
  # {
  #   colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="Sum of Worked-Total"] = "Sum of Worked"
  #   colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="JLA Hours-Total"] = "JLA Hours"
  #   
  #   colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="Daily %age"] = "Mid-Day %age"
  #   cws_order=c("Groups","Head Count","Available Hours","Sum of Worked","JLA Hours","Mid-Day %age","JLA %age")
  # }
  # else
  {
    cws_order=c("Team","Head Count","Available Hours","Sum of Worked-Total",
                "JLA Hours-Total","Daily %age","JLA %age")
    
  }
  
  cluster_wise_summary =cluster_wise_summary[,cws_order]
  
  return(cluster_wise_summary)
  
}
print_PWS<- function(cluster_wise_summary,Hours){
  # if(Hours!=8)
  # {
  # colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="Sum of Worked-Total"] = "Sum of Worked"
  # colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="JLA Hours-Total"] = "JLA Hours"
  #   
  #   colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="Daily %age"] = "Mid-Day %age"
  #   cws_order=c("Groups","Head Count","Available Hours","Sum of Worked","JLA Hours","Mid-Day %age","JLA %age")
  # }
  # else
  {
    colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="Sum of Worked-Cluster"] = "Sum of Worked-Project"
    colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="JLA Hours-Cluster"] = "JLA Hours-Project"
    colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="PMO Projects Name"] = "Project"
    colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="Groups"] = "Group"
    
    cws_order=c("Group","Project","Head Count","Available Hours","Sum of Worked-Project","Sum of Worked-Total",
                "JLA Hours-Project","JLA Hours-Total","Daily %age","JLA %age")
    
  }
  
  cluster_wise_summary =cluster_wise_summary[,cws_order]
  
  return(cluster_wise_summary)
  
}

print_CWRL<- function(cluster_wise_resource_logging){
  
  CWRL=cluster_wise_resource_logging
  colnames(CWRL)[colnames(CWRL)=="Jira_hours"] = "Sum of Worked-Cluster"
  colnames(CWRL)[colnames(CWRL)=="JLA_hours"] = "JLA Hours-Cluster"
  
  
  cwrl_order=c("Emp Code","Resource Name","Product","Groups","PMO Projects Name", 
               "Team","%Allocated","Available Hours","Sum of Worked-Cluster",
               "Sum of Worked-Total","JLA Hours-Cluster","JLA Hours-Total")
  CWRL=CWRL[,cwrl_order]
}
print_NC<- function(cluster_wise_resource_logging){
  
  CWRL=cluster_wise_resource_logging
  colnames(CWRL)[colnames(CWRL)=="Jira_hours"] = "Sum of Worked-Cluster"
  colnames(CWRL)[colnames(CWRL)=="JLA_hours"] = "JLA Hours-Cluster"
  
  
  cwrl_order=c("Emp Code","Resource Name","Groups","PMO Projects Name", 
               "Team","%Allocated","Available Hours","Sum of Worked-Cluster",
               "Sum of Worked-Total","JLA Hours-Cluster","JLA Hours-Total","Email")
  CWRL=CWRL[,cwrl_order]
}

print_CWRL_hr<- function(cluster_wise_resource_logging){
  
  CWRL=cluster_wise_resource_logging
  colnames(CWRL)[colnames(CWRL)=="Sum of Worked-Total"] = "Total Hours"
  colnames(CWRL)[colnames(CWRL)=="JLA Hours-Total"] = "JLA Hours"
  CWRL$`Jira Manual Hours` = round(CWRL$`Total Hours` - CWRL$`JLA Hours`,2)
  
  cwrl_order=c("Emp Code","Resource Name","Product","Groups","PMO Projects Name", 
               "Team","%Allocated","Available Hours","Total Hours",
               "JLA Hours","Jira Manual Hours")
  CWRL=CWRL[,cwrl_order]
}

print_bar_plot <- function(cluster_wise_summary){
  
  
  colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="Daily %age"] = "day_percentage"
  colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="JLA %age"] = "jla_percentage"
  
  
  
  fig <- cluster_wise_summary %>% plot_ly()
  fig <- fig %>% add_trace(x = ~Groups, y = ~day_percentage, type = 'bar', name ="Jira %",
                           text = cluster_wise_summary$day_percentage, textposition = 'auto',
                           marker = list(color = 'rgb(158,202,225)',
                                         line = list(color = 'rgb(8,48,107)', width = 1.5)))
  fig <- fig %>% add_trace(x = ~Groups, y = ~jla_percentage, type = 'bar',name ="JLA %",
                           text = cluster_wise_summary$jla_percentage, textposition = 'auto',
                           marker = list(color = 'rgb(58,200,225)',
                                         line = list(color = 'rgb(8,48,107)', width = 1.5)))
  
  fig <- fig %>% layout(title = title,
                        barmode = 'group',
                        xaxis = list(title = ""),
                        yaxis = list(title = ""))
  
  return(fig)
}

print_horizontal_bar <- function(cluster_wise_activity_seg){
  df= (cluster_wise_activity_seg[cluster_wise_activity_seg$Group=="Grand Total",colnames(cluster_wise_activity_seg)[!colnames(cluster_wise_activity_seg) %in% c("Group","Grand Total")]])
  
  df=data.frame(t(df))
  df$Type=row.names(df)
  row.names(df) =1:nrow(df)
  colnames(df)=c("Hours","Type")
  df = df[order(df$Hours , decreasing=TRUE),]
  fig <-  plot_ly()
  fig <- fig %>% add_trace(x = df$Hours, y = df$Type, type = 'bar', name ="Jira %",
                           text = df$Hours, textposition = 'auto',
                           marker = list(color = 'rgb(160,202,225)',
                                         line = list(color = 'rgb(8,48,107)', width = 1.5)))
  fig <- fig %>% layout(title = "Activity Wise Jira Logging",
                        orientation = "h",
                        xaxis = list(title = "Hours"),
                        yaxis = list(title = "Activity",categoryorder = "total descending"))
  
  return(fig)
  
  
}
ui <- dashboardPage(
  dashboardHeader(title = "NFS Jira Compliance"),
  dashboardSidebar(
    # size = "wide",
    sidebarMenu(
      menuItem(tabName = "upload_data", text = "Upload Data"),
      menuItem(tabName = "dashboard", text = "Dashboard"),
      menuItem(tabName = "hr_view", text = "HR")
      
    )
  ),
  dashboardBody(
    shinyjs::useShinyjs(),  
    
    tabItems(
      # selected = 1,
      tabItem(
        tabName = "upload_data",
        
        box( width=16,status = "primary",# solidHeader = TRUE, title="Table",
             fileInput("pmo_input", "Upload PMO File", accept = c( "xls/xlsx") ),
             
             br(),
             # tableOutput("table1")
        ),
        box( width=16, status = "primary" ,# solidHeader = TRUE, title="Table",
             fileInput("jira_input", "Upload Jira File" , accept = c( "csv")),
             br(),
             # tableOutput("table1")
        ),
        
        # radioButtons("day_type","Select Day ",
        #              choices = list("Mid Day" = 4,
        #                             "Daily" = 8,
        #                             "Weekly" = 40),
        #              selected = character(0)
        # ),
        selectInput("report_type", "Select Report Type", c("Mid-day", "Daily","Weekly")),
        
        
        conditionalPanel(
          condition = "input.report_type == 'Mid-day' ||  input.report_type == 'Daily'",
          dateInput(inputId='date',
                    label = 'Select Date',
                    value = Sys.Date()
          ),
        ),
        
        conditionalPanel(
          condition = "input.report_type == 'Weekly' ",
          dateRangeInput("date_range", 
                         "Date range",
                         # start = as.character(as.Date("2023-02-27")),
                         # end = as.character(as.Date("2023-03-03")),
                         start = as.character(Sys.Date()-7),
                         end = as.character(Sys.Date()-3)
          ),
        ),
        
        
        
        # numericInput("report_type", "Select Hours:", 4, min = 4, max = 40),
        actionButton("analysis","Analyze"),
        downloadButton('downloadData', 'Download Results')
        
      ),
      tabItem(
        tabName = "dashboard",
        plotlyOutput("plot"),
        plotlyOutput("activity_plot")
        
      ),
      tabItem(
        tabName = "hr_view",
        dateInput(inputId='date_hr',
                  label = 'Select Date',
                  value = Sys.Date()-1
        ),
        actionButton("analysis_hr","Check Data"),
        downloadButton('download_hr', "Generate Report")
      )
      
      
    )
  )
)

server <- function(input, output) { 
  
  ####### Initializing Variables
  alloc_data = jira_data = NULL
  val <- reactiveValues()
  val$cluster_wise_summary = val$non_compliant=data.frame()
  val$cluster_wise_resource_logging = data.frame()
  val$hours_df= data.frame()
  val$cluster_wise_activity_seg = data.frame
  val$team_wise_summary = data.frame()
  val$project_wise_summary = data.frame()
  val$pmo_input_file_name =val$jira_file_name = NULL
  val$emp_wise_activity_seg = data.frame()
  val$hr= NULL
  val$hr_cwd=NULL
  
  #Empty Directory before uploading new files
  
  ###### Save PMO input File
  observeEvent(input$pmo_input, {
    if ( file_ext(input$pmo_input$name) =="xls" || file_ext(input$pmo_input$name) =="xlsx")
    {
      val$pmo_input_file_name = input$pmo_input$name
      drive_upload(media = input$pmo_input$datapath,name = input$pmo_input$name,overwrite=TRUE)
    }
    else{
      showModal(modalDialog(
        title = "Error",
        paste("Please upload PMO file with extension xls or xlsx"),
        easyClose = TRUE,
        footer = NULL
      ))
    }
  })
  ###### Save Jira input File
  observeEvent(input$jira_input, {
    if ( file_ext(input$jira_input$name) =="csv" )
    {
      val$jira_file_name =input$jira_input$name
      showModal(modalDialog("Uploading File on drive", footer=NULL))
      drive_upload(media = input$jira_input$datapath,name = input$jira_input$name,overwrite=TRUE)
      removeModal()
    }
    else{
      showModal(modalDialog(
        title = "Error",
        paste("Please upload Jira file with extension csv"),
        easyClose = TRUE,
        footer = NULL
      ))
    }
  })
  
  observeEvent(input$analysis, {
    jira_data= jira_users= alloc_data =NULL
    
    if (!is.null(val$pmo_input_file_name)  & !is.null(val$jira_file_name))
    {
      showModal(modalDialog("Downloading Data from drive....", footer=NULL))
      #Download Jira File from drive
      drive_download(val$jira_file_name,overwrite = TRUE)
      jira_data= read.csv(val$jira_file_name)
      names(jira_data) <- gsub("."," ",colnames(jira_data),fixed=TRUE )
      
      
      #Download PMO File from drive
      drive_download(val$pmo_input_file_name,overwrite = TRUE)
      alloc_data= read_excel(val$pmo_input_file_name, sheet=1)
      names(alloc_data) <- gsub(x = names(alloc_data),pattern = "\\.",replacement = " ")
      
      jira_users= read_excel(val$pmo_input_file_name, sheet=2)
      names(jira_users) <- gsub(x = names(jira_users),pattern = "\\.",replacement = " ")
      removeModal()
    }
    else{
      showModal(modalDialog(
        title = "Error",
        paste("Please upload both PMO file and Jira file"),
        easyClose = TRUE,
        footer = NULL
      ))
    }
    
    ################################
    if (!is.null(jira_data)  & !is.null(alloc_data))
    {
      
      
      if (nrow(jira_data) > 0 ){
        
        showModal(modalDialog("Processing....", footer=NULL))
        if(input$report_type=="Daily" || input$report_type=="Mid-day")
        {
          jira_data$`Work date` = as.Date(as.Date(input$date),format='%Y-%m-%d')
        }
        
        Mapping = get_mapping()
        jira_data = pre_process_jira_data(jira_data)
        # jira_data=jira_data[-1,]
        # jira_data$Hours= as.numeric( jira_data$Hours)
        # jira_data=jira_data[!is.na(jira_data$Hours),]
        jira_activity_data= jira_data[,c("Emp Code","Hours", "Project Name","Activity Type" )]
        jira_data= jira_data[,c("Emp Code","Full name","Hours", "Team", "JLA Activity Log", "Project Name","Work date" )]
        
        
        alloc_data <- pre_process_pmo_data(alloc_data)
        alloc_email = alloc_data[,c("Emp Code","Email"  )]
        alloc_data_prod= alloc_data[,c("Emp Code", "Product")]
        
        alloc_data = alloc_data[,c("Groups","Emp Code","Team","PMO Projects Name","Resource Name","%Allocated")]
        
        
        
        #######
        
        if (input$report_type != "Weekly") 
        {
          jira_data =merge(Mapping, jira_data, by.x="JIRA Projects Name",by.y="Project Name",all.y=TRUE)
          
          # jira_data=jira_data[grepl('{',  jira_data$Hours),]
          
          JLA_data = jira_data[jira_data$`JLA Activity Log`=="Yes",]
          
          JLA_team= JLA_data %>% group_by( `PMO Projects Name`, Groups,`Emp Code`) %>% summarise(JLA_hours = sum(as.numeric(Hours)))
          JLA_total_hours=JLA_data %>% group_by(`Emp Code`) %>% summarise(JLA_hours = sum(as.numeric(Hours)))
          Jira_team= jira_data %>% group_by(`PMO Projects Name`, Groups,`Emp Code`) %>% summarise( Jira_hours= sum(as.numeric(Hours)))
          
          
          team =merge(Jira_team, JLA_team,  all=TRUE )
          
          rm(JLA_data,JLA_team,Mapping)
          # Resource Allocation Wise Statistics -------------------------------------------------
          cluster_wise_resource_logging =merge( team , alloc_data, by=c("Emp Code","Groups","PMO Projects Name")#,"PMO Projects Name"
                                                , all=TRUE )
          
          # val$cluster_wise_resource_logging = calculate_resource_allocation(cluster_wise_resource_logging,JLA_total_hours,val$hr)
          df = calculate_resource_allocation(cluster_wise_resource_logging,JLA_total_hours,val$hr)
          df$`Product`= alloc_data_prod$Product[match(df$`Emp Code`,alloc_data_prod$`Emp Code` )]
          val$cluster_wise_resource_logging = df
          
          # Resource with ZERO jira Compliance -------------------------------------------------
          non_compliant = val$cluster_wise_resource_logging[val$cluster_wise_resource_logging$`Sum of Worked-Total`==0,]
          alloc_email = alloc_email[!duplicated(alloc_email$`Emp Code`),]
          non_compliant[is.na(non_compliant)]= 0
          val$non_compliant =merge(non_compliant, alloc_email,by="Emp Code", all.x = TRUE )
          
          # Cluster Summary Statistics -------------------------------------------------
          val$cluster_wise_summary = calculate_cluster_summary(val$cluster_wise_resource_logging,val$hr)
          # Team Summary Statistics -------------------------------------------------
          val$team_wise_summary = calculate_team_summary(val$cluster_wise_resource_logging,val$hr)
          # Project Summary Statistics -------------------------------------------------
          val$project_wise_summary = calculate_project_summary(val$cluster_wise_resource_logging,val$hr)
          
          # --------------------# Handling Un-allocated Resources
          # Calculating HOURS Summary Statistics -------------------------------------------------
          jira_users=jira_users[complete.cases(jira_users$User),]
          hours_df = data.frame(matrix(nrow=4, ncol=2))
          colnames(hours_df) = c("Type","Count")
          hours_df$Type=c('>= 10 hours','>= 8 hours','>= 1 and < 4 hours','Did Not Log')
          
          hours_df[hours_df$Type==">= 10 hours",]$Count =length(jira_users[jira_users$Logged>=10, ]$Logged)
          hours_df[hours_df$Type==">= 8 hours",]$Count=length(jira_users[jira_users$Logged>=8, ]$Logged)
          hours_df[hours_df$Type==">= 1 and < 4 hours",]$Count =length(jira_users[(jira_users$Logged>=1 & jira_users$Logged<4), ]$Logged)
          hours_df[hours_df$Type=="Did Not Log",]$Count =length(jira_users[jira_users$Logged ==0,]$Logged)
          hours_df[nrow(hours_df)+1,] = "NA"
          hours_df$Type[nrow(hours_df)] = "Total Jira Users"
          hours_df[hours_df$Type=="Total Jira Users",]$Count  = nrow(jira_users)
          
          val$hours_df = hours_df
          if(input$report_type=="Daily")
          {
            val$emp_wise_activity_seg = get_cluster_activity(jira_activity_data,"emp_code", alloc_data)
            # val$cluster_wise_resource_logging[duplicated(cluster_wise_resource_logging),c("Emp Code", "Resource Name")]
            val$emp_wise_activity_seg = merge(val$emp_wise_activity_seg, val$cluster_wise_resource_logging[,c("Emp Code", "Resource Name")],by="Emp Code", x.all= TRUE)
            val$emp_wise_activity_seg  =  val$emp_wise_activity_seg[!duplicated(val$emp_wise_activity_seg),] 
            
            l=colnames(val$emp_wise_activity_seg[,!(names(val$emp_wise_activity_seg)  %in% c("Group","Emp Code","Resource Name"))])
            l= append(c("Group","Emp Code","Resource Name"), l)
            val$emp_wise_activity_seg = val$emp_wise_activity_seg[,l]
          }
          # write.csv(val$cluster_wise_summary,"temp.csv")
        }
        else{
          # jira_data$`Work date`=as.Date(jira_data$`Work date`,  "%Y-%m-%d")
          jira_data$`Work date` = as.Date( format(as.POSIXct(jira_data$`Work date`,format='%m/%d/%Y ')))
          
          CLUSTERS =unique(alloc_data$Groups)
          # HC calculation -------------------------------------------------
          alloc_data$HC = alloc_data$`%Allocated`/100
          HC = alloc_data %>% group_by(Groups) %>% summarize(`Head Count` = sum(HC))
          
          weekly_summary = data.frame(matrix(nrow=(length(CLUSTERS)),ncol=0))
          # weekly_summary$Groups = c(CLUSTERS, "Grand Total")
          weekly_summary$Groups = CLUSTERS
          weekly_summary = merge(HC,weekly_summary, by=c("Groups"))
          weekly_summary$`Available Hours` =  weekly_summary$`Head Count` * hours_in_day
          weekly_summary[nrow(weekly_summary)+1,]$Groups=c("Grand Total")
          
          days=unique(jira_data[order(jira_data$`Work date`),]$`Work date` ,na.rm=TRUE)
          
          
          for (i in 1:length(days)){
            
            day=days[i]
            jira_data_1= jira_data[jira_data$`Work date` == day,]
            Mapping = get_mapping()
            jira_data_1 =merge(Mapping, jira_data_1, by.x="JIRA Projects Name",by.y="Project Name",all.y=TRUE)
            
            JLA_data = jira_data_1[jira_data_1$`JLA Activity Log`=="Yes",]
            JLA_team= JLA_data %>% group_by( `PMO Projects Name`, Groups,`Emp Code`) %>% summarise(JLA_hours = sum(as.numeric(Hours)))
            JLA_total_hours=JLA_data %>% group_by(`Emp Code`) %>% summarise(JLA_hours = sum(as.numeric(Hours)))
            Jira_team= jira_data_1 %>% group_by(`PMO Projects Name`, Groups,`Emp Code`) %>% summarise( Jira_hours= sum(as.numeric(Hours)))
            team =merge(Jira_team, JLA_team,  all=TRUE )
            rm(JLA_data,JLA_team,Mapping)
            # Resource Allocation Wise Statistics -------------------------------------------------
            # cluster_wise_resource_logging =merge( team , alloc_data, by=c("Emp Code","Groups","PMO Projects Name")#,"PMO Projects Name", all=TRUE )# 
            cluster_wise_resource_logging =merge( team , alloc_data, by=c("Emp Code","Groups","PMO Projects Name"), all=TRUE )
            cluster_wise_resource_logging = calculate_resource_allocation(cluster_wise_resource_logging,JLA_total_hours,hours_in_day)
            # write.csv(cluster_wise_resource_logging,paste("cwrl ", day,".csv",sep=""))
            # # Cluster Summary Statistics -------------------------------------------------
            day_summary = calculate_cluster_summary(cluster_wise_resource_logging,hours_in_day)
            # write.csv(day_summary,paste("cws ", day,".csv",sep=""))
            day_summary =day_summary[-nrow(day_summary),]
            ############
            # day_summary[nrow(day_summary)+1,]=NA
            # day_summary$Groups[nrow(day_summary)]="Cluster - Web 3.0"
            cols=c("Sum of Worked-Total","JLA Hours-Total","Daily %age","JLA %age")
            
            weekly_summary = merge(weekly_summary,day_summary[,c(cols,"Groups")], by="Groups",all.y=TRUE )
            
            
            new_cols <- paste(colnames(weekly_summary[,cols]), day,sep="|")
            
            index=(ncol(weekly_summary)-3):ncol(weekly_summary)
            colnames(weekly_summary)[index]<- paste(colnames(weekly_summary[,index]), day,sep="|")
            
            # #Grand Total
            # day_summary[nrow(day_summary)+1,]=NA
            # day_summary$Groups[nrow(day_summary)]="Grand Total"
            # day_summary[day_summary$Groups=="Grand Total",-1] =
            #   as.list(colSums(day_summary[,-1], na.rm = TRUE))
            
            
          }
          # 
          
          
          # JLA_data = jira_data[jira_data$`JLA Activity Log`=="Yes",]
          # JLA_team= JLA_data %>% group_by( `Emp Code`,`Work date`) %>% summarise(JLA_hours = sum(Hours,na.rm=TRUE))
          # JLA_total_hours=JLA_data %>% group_by(`Emp Code`,`Work date`) %>% summarise(JLA_hours = sum(Hours))
          # Jira_team= jira_data %>% group_by(`Emp Code`,`Work date`) %>% summarise( Jira_hours= sum(Hours))
          # team =merge(  Jira_team, JLA_team,  all=TRUE )
          # 
          # rm(Mapping,JLA_data,JLA_team,Jira_team,jira_data)
          
          out = week_summary(days,hours_in_day,HC, weekly_summary )
          val$cluster_wise_summary =out$cluster_wise_summary
          val$weekly_summary = out$weekly_summary
          
        }
        # Activity Type Statistics -------------------------------------------------
        
        val$cluster_wise_activity_seg = get_cluster_activity(jira_activity_data,"group")
        removeModal()
        
      }
      else{
        showModal(modalDialog(
          title = "Error",
          paste("Please upload file for date", input$date, "OR Update date according to uploaded Jira data"),
          easyClose = TRUE,
          footer = NULL
        ))
      }
      # file_name = paste("JLA Stats. ",input$date,".xlsx", sep="" )
      # 
      # list_of_datasets <- list("Summary" = cluster_wise_summary,
      #                          "Cluster wise resource logging" = cluster_wise_resource_logging)
      # write.xlsx(list_of_datasets, file = file_name)
      
    }
    else{
      showModal(modalDialog(
        title = "Error",
        paste("Please upload both PMO file and Jira file"),
        easyClose = TRUE,
        footer = NULL
      ))
    }
    
  })
  
  output$downloadData <- downloadHandler(
    filename = function() { 
      # filename=
      paste("NFS Weekly Compliance.xlsx", sep="" )
      if(input$report_type=="Mid-day")
      {
        paste("NFS Mid-day Compliance_",as.character(input$date),".xlsx", sep="" )
      }
      else if (input$report_type=="Daily")
      {
        paste("NFS Daily Compliance_",as.character(input$date),".xlsx", sep="" )
      }
      else if (input$report_type=="Weekly")
      {
        
        paste("NFS Weekly Compliance " , as.character(input$date_range[1]), " to ", as.character(input$date_range[2]),".xlsx",sep="")
        
        
      }
      
      
    },
    content = function(file){
      
      # set path
      temp <- setwd(tempdir())
      on.exit(setwd(temp))
      
      
      if(input$report_type =="Mid-day")
      {
        CWS = print_CWS(val$cluster_wise_summary, input$report_type)
        CWRL =print_CWRL(val$cluster_wise_resource_logging)
        wb <- openxlsx::createWorkbook()
        addWorksheet(wb,"Summary" )
        writeDataTable(wb, "Summary", startCol = 1,   startRow = 1, x = as.data.frame(CWS), tableStyle = "TableStyleMedium9")
        
        addWorksheet(wb, "Cluster wise resource logging")
        writeDataTable(wb, "Cluster wise resource logging", startCol = 1,   startRow = 1, x = as.data.frame(CWRL), tableStyle = "TableStyleMedium9")
        
        
        openxlsx::saveWorkbook(wb, file = file, overwrite = TRUE)
        
        
      }
      
      else if(input$report_type =="Daily")
      {
        #########For HR 
        dc_file_name =  paste("NFS Daily Compliance_",as.character(input$date),".xlsx", sep="" )
        CWRL_hr =print_CWRL_hr(val$cluster_wise_resource_logging)
        
        wb_hr <- openxlsx::createWorkbook()
        addWorksheet(wb_hr, "Cluster wise resource logging")
        writeDataTable(wb_hr, "Cluster wise resource logging", startCol = 1,   startRow = 1, x = as.data.frame(CWRL_hr), tableStyle = "TableStyleMedium9")
        openxlsx::saveWorkbook(wb_hr, file = dc_file_name, overwrite = TRUE)
        drive_upload(media = dc_file_name,name = dc_file_name,overwrite=TRUE)
        
        
        ##########
        
        CWS = print_CWS(val$cluster_wise_summary, input$report_type)
        CWRL =print_CWRL(val$cluster_wise_resource_logging)
        TWS = print_TWS(val$team_wise_summary)
        PWS = print_PWS(val$project_wise_summary)
        NC = print_NC(val$non_compliant)
        
        wb <- openxlsx::createWorkbook()
        
        addWorksheet(wb,"Summary" )
        writeDataTable(wb, "Summary", startCol = 1,   startRow = 1, x = as.data.frame(CWS), tableStyle = "TableStyleMedium9")
        writeDataTable(wb, "Summary",startCol = 15,   startRow = 1, x = as.data.frame(val$hours_df), tableStyle = "TableStyleMedium9")
        writeDataTable(wb, "Summary",startCol = 1,   startRow = 18, x = as.data.frame(TWS), tableStyle = "TableStyleMedium9")
        writeDataTable(wb, "Summary",startCol = 1,   startRow = 50, x = as.data.frame(PWS), tableStyle = "TableStyleMedium9")
        
        addWorksheet(wb, "Cluster wise resource logging")
        writeDataTable(wb, "Cluster wise resource logging", startCol = 1,   startRow = 1, x = as.data.frame(CWRL), tableStyle = "TableStyleMedium9")
        
        addWorksheet(wb, "Cluster wise activity")
        writeDataTable(wb, "Cluster wise activity", startCol = 1,   startRow = 1,
                       x = as.data.frame(val$cluster_wise_activity_seg), tableStyle = "TableStyleMedium9")
        
        addWorksheet(wb, "Employee wise activity")
        writeDataTable(wb, "Employee wise activity", startCol = 1,   startRow = 1,
                       x = as.data.frame(val$emp_wise_activity_seg), tableStyle = "TableStyleMedium9")
        
        
        addWorksheet(wb, "Non Compliant Resources")
        writeDataTable(wb, "Non Compliant Resources", startCol = 1,   startRow = 1,
                       x = as.data.frame(NC), tableStyle = "TableStyleMedium9")
        openxlsx::saveWorkbook(wb, file = file, overwrite = TRUE)
      }
      else if (input$report_type =="Weekly")
      {
        days = seq( input$date_range[1],input$date_range[2],by="days")
        
        
        wb <- openxlsx::createWorkbook()
        addWorksheet(wb,"Summary" )
        
        writeDataTable(wb, "Summary", startCol = 1,   startRow = 1, x = as.data.frame(val$cluster_wise_summary), tableStyle = "TableStyleMedium9")
        writeDataTable(wb, "Summary",startCol = 15,   startRow = 1, x = as.data.frame(val$cluster_wise_activity_seg), tableStyle = "TableStyleLight11")
        
        
        
        writeData(wb, "Summary",   startRow = 22,startCol = 1, val$weekly_summary)
        
        
        writeData(wb, "Summary",   startRow = 21,startCol = 1, x = "Groups" )
        writeData(wb, "Summary",   startRow = 21,startCol = 2, x = "Head Count" )
        writeData(wb, "Summary",   startRow = 21,startCol = 3, x = "Available Hours" )
        openxlsx ::mergeCells(wb, "Summary", cols = 1, rows = 21:22)
        openxlsx ::mergeCells(wb, "Summary", cols = 2, rows = 21:22)
        openxlsx ::mergeCells(wb, "Summary", cols = 3, rows = 21:22)
        
        dates= as.Date(sort(rep(days,4)), format="%Y-%m-%d")
        writeData(wb, "Summary",   startRow = 21,startCol = 4, x = days[1])
        writeData(wb, "Summary",   startRow = 21,startCol = 8, x = days[2] )
        writeData(wb, "Summary",   startRow = 21,startCol = 12, x = days[3] )
        writeData(wb, "Summary",   startRow = 21,startCol = 16, x = days[4] )
        writeData(wb, "Summary",   startRow = 21,startCol = 20, x = days[5] )
        openxlsx ::mergeCells(wb, "Summary",  rows = 21, cols = 4:7)
        openxlsx ::mergeCells(wb, "Summary",  rows = 21, cols = 8:11)
        openxlsx ::mergeCells(wb, "Summary",  rows = 21, cols = 12:15)
        openxlsx ::mergeCells(wb, "Summary",  rows = 21, cols = 16:19)
        openxlsx ::mergeCells(wb, "Summary",  rows = 21, cols = 20:23)
        
        
        openxlsx::saveWorkbook(wb, file = file, overwrite = TRUE)
        # filename=paste("NFS Weekly Compliance " , as.character(input$date_range[1]), "to", as.character(input$date_range[2]),".xlsx",sep="")
        
        # openxlsx::saveWorkbook(wb, file = paste(getwd( ),"/",filename,sep=""), overwrite = TRUE)
        
        print("write saveWorkbook")
      }
      
    }
  )
  
  output$plot <- renderPlotly({
    req(  val$cluster_wise_summary)
    if( nrow( val$cluster_wise_summary) > 1){
      
      cluster_wise_summary =val$cluster_wise_summary
      cluster_wise_summary[,-1]= round(cluster_wise_summary[,-1],0)
      cluster_wise_summary$Groups <- factor(cluster_wise_summary$Groups,
                                            levels = cluster_wise_summary$Groups)
      # colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="Daily %age"] = "day_percentage"
      # colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="JLA %age"] = "jla_percentage"
      
      if (input$report_type == "Daily")
      {
        title = paste("Daily Jira Compliance ", as.character(input$date), sep =
                        "")
        jira_perc = cluster_wise_summary$"Daily %age"
        jla_prec = cluster_wise_summary$"JLA %age"
        
      }  
      else if (input$report_type == "Mid-day") {
        title = paste("Mid-day Jira Compliance ", as.character(input$date), sep ="")
        jira_perc = cluster_wise_summary$"Daily %age"
        jla_prec = cluster_wise_summary$"JLA %age"
        
      } 
      else if (input$report_type == "Weekly") {
        title = paste(
          "Weekly Jira Compliance ",
          as.character(input$date_range[1]),
          "-",
          as.character(input$date_range[2]),
          sep = ""
        )
        jira_perc = cluster_wise_summary$"Weekly %age"
        jla_prec = cluster_wise_summary$"JLA %age"
        
      }
      
      
      
      
      fig <- cluster_wise_summary %>% plot_ly()
      fig <- fig %>% add_trace(x = ~Groups, y = jira_perc, type = 'bar', name ="Jira %",
                               text = jira_perc, textposition = 'auto',
                               marker = list(color = 'rgb(158,202,225)',
                                             line = list(color = 'rgb(8,48,107)', width = 1.5)))
      fig <- fig %>% add_trace(x = ~Groups, y = jla_prec, type = 'bar',name ="JLA %",
                               text = jla_prec, textposition = 'auto',
                               marker = list(color = 'rgb(58,200,225)',
                                             line = list(color = 'rgb(8,48,107)', width = 1.5)))
      fig <- fig %>% layout(title = title,
                            barmode = 'group',
                            xaxis = list(title = ""),
                            yaxis = list(title = ""))
    }
    
  })
  
  output$activity_plot <- renderPlotly({
    req(  val$cluster_wise_activity_seg)
    if( nrow( val$cluster_wise_activity_seg) > 1 & input$report_type != "Mid-day"){
      print_horizontal_bar(val$cluster_wise_activity_seg)
    }
  })
  
  observeEvent(input$summary_data_type,{
    
    
  })
  
  # output$summary_table <- DT::renderDataTable(print_CWS(val$cluster_wise_summary, val$hr),
  #                                             options = list(paging = F, dom = 't'),
  #                                             #,scrollX = TRUE
  #                                             rownames = FALSE)
  output$head_count <- renderText({ 
    req(val$cluster_wise_summary)
    round(sum(val$cluster_wise_summary$`Head Count`),0) })
  
  
  observeEvent(input$report_type,{
    if(input$report_type=="Mid-day") val$hr=4
    else if(input$report_type=="Daily") val$hr=8
    else if(input$report_type=="Weekly") val$hr=40
    # if(input$report_type=="Mid-day") val$hr=3.5
    # else if(input$report_type=="Daily") val$hr=7
    # else if(input$report_type=="Weekly") val$hr=35
    
    
  })
  
  observeEvent(input$analysis_hr, {
    showModal(modalDialog("Checking if data exists ", footer=NULL))
    print(  paste("NFS Daily Compliance_",as.character(input$date_hr),".xlsx",sep=""))
    dc_file_name = paste("NFS Daily Compliance_",as.character(input$date_hr),".xlsx",sep="")
    result <- tryCatch({
      drive_download(dc_file_name,overwrite = TRUE)
      compliance_=read_excel(dc_file_name)
      if (!is.null(compliance_))
      {
        val$hr_cwd = compliance_
        removeModal()
        # downloadButton("download_hr", "Download Data")
        showModal(modalDialog(
          paste("Data is available. Please proceed with download"),
          easyClose = TRUE,
          footer = NULL
        ))
        shinyjs::enable("download_hr")
        
      }
    }, error = function(err) {
      showModal(modalDialog(
        title = "Error",
        paste("Data not available for selected date."),
        easyClose = TRUE,
        footer = NULL
      ))
      val$hr_cwd = NULL
      # disable("download_hr")
      shinyjs::disable("download_hr")
      
    })
    
    # removeModal()
    
  })
  
  output$download_hr <- downloadHandler(
    filename = function() {
      
      paste("NFS Daily Compliance_",as.character(input$date_hr),".xlsx",sep="")
      
    },
    content = function(file){
      if (nrow(val$hr_cwd)>1)
      {
        CWRL_hr =print_CWRL_hr(val$hr_cwd)
        wb_hr <- openxlsx::createWorkbook()
        addWorksheet(wb_hr, "Cluster wise resource logging")
        writeDataTable(wb_hr, "Cluster wise resource logging", startCol = 1,   startRow = 1, x = as.data.frame(CWRL_hr), tableStyle = "TableStyleMedium9")
        openxlsx::saveWorkbook(wb_hr, file = file, overwrite = TRUE)
      }
      
      
    }
  )
  
  
  
}

shinyApp(ui, server)
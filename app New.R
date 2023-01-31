options(shiny.maxRequestSize=30*1024^2)
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
  jira_data[!is.na(jira_data$`Activity Logs`),]$`Activity Logs` <- "Yes"
  jira_data[is.na(jira_data$`Activity Logs`),]$`Activity Logs` <- "No"
  jira_data[which(jira_data$`Activity Type` == "Admin- Leave" ),]$`Activity Logs`= "Yes"
  jira_data[is.na(jira_data$`Work Description`),]$`Work Description` <- ""
  jira_data[str_detect(jira_data$`Work Description`, "JIRALog:"),]$`Activity Logs` = "Yes"
  jira_data$`Emp Code` = apply(data.frame(jira_data$`Full name`),1,function(x)gsub('.*-', '',x))
  jira_data$`Emp Code` = trimws(jira_data$`Emp Code`)
  
  colnames(jira_data)[colnames(jira_data)=="Activity Logs"] = "JLA Activity Log"
  
  return(jira_data)
}
pre_process_pmo_data<- function(alloc_data){
  alloc_data =alloc_data[,c("Groups", "Resource Name" , "Emp Code","Team","Reporting TL","%Allocated","Project Name")]
  alloc_data$`Id` =alloc_data$`Emp Code`
  alloc_data$`Emp Code` =  apply(data.frame(alloc_data$`Emp Code`),1,function(x)gsub('.*-', '',x))
  ######Adding Allocation % for allocation records to same clusters
  alloc_data = alloc_data %>% group_by(`Groups`, `Emp Code`,Team,`Project Name` ,`Resource Name` )%>% 
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
  
  cluster_wise_summary = temp %>% group_by(`Groups`)%>%
    summarise(`JLA Hours-Cluster` = sum(`JLA_hours`, na.rm=TRUE),
              `Sum of Worked-Cluster` = sum(`Jira_hours`, na.rm=TRUE),
              `Available Hours` = sum(`Available Hours`, na.rm=TRUE),
              `Available Hours` = sum(`Available Hours`, na.rm=TRUE),
              `Sum of Worked-Total` = sum(`Sum of Worked-Total per alloc`, na.rm=TRUE),
              `JLA Hours-Total` = sum(`JLA Hours-Total per alloc`, na.rm=TRUE)
    )
  cluster_wise_summary$`Head Count` =  cluster_wise_summary$`Available Hours`/Hours
  
  # 
  # 
  # #[!duplicated(cluster_wise_resource_logging$`Emp Code`),]
  # temp2  = temp %>% group_by(`Groups`)%>%
  #   summarise(`Sum of Worked-Total` = sum(`Sum of Worked-Total per alloc`, na.rm=TRUE),
  #             `JLA Hours-Total` = sum(`JLA Hours-Total per alloc`, na.rm=TRUE),)
  # cluster_wise_summary= merge(cluster_wise_summary,temp2)  
  # 
  
  # cluster_wise_summary$`Daily %age` = 
  #   round((cluster_wise_summary$`Sum of Worked-Total`/cluster_wise_summary$`Available Hours`)*100,0)
  # cluster_wise_summary$`JLA %age` = round((cluster_wise_summary$`JLA Hours-Cluster`/cluster_wise_summary$`Available Hours`)*100,0)
  
  cluster_wise_summary$`Daily %age` = 
    round((cluster_wise_summary$`Sum of Worked-Total`/cluster_wise_summary$`Available Hours`)*100,0)
  cluster_wise_summary$`JLA %age` = round((cluster_wise_summary$`JLA Hours-Total`/cluster_wise_summary$`Available Hours`)*100,0)
  
  #Grand Total
  cluster_wise_summary[nrow(cluster_wise_summary)+1,] = NA
  cluster_wise_summary$Groups[nrow(cluster_wise_summary)] = "Grand Total"
  cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",-1] =
    as.list(colSums(cluster_wise_summary[,-1], na.rm = TRUE))
  
  
  cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",]$`Daily %age`=100*(
    as.numeric(cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",]$`Sum of Worked-Total`)/as.numeric(cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",]$`Available Hours`))
  
  cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",]$`JLA %age` =100*(
    as.numeric(cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",]$`JLA Hours-Total`)/as.numeric(cluster_wise_summary[cluster_wise_summary$Groups=="Grand Total",]$`Available Hours`))
  
  cluster_wise_summary[,-1]= round(cluster_wise_summary[,-1],0)
  return(cluster_wise_summary)
}  

calculate_team_summary<- function(cluster_wise_resource_logging,Hours){
  temp=cluster_wise_resource_logging
  temp$`Sum of Worked-Total per alloc`= ((temp$`%Allocated`)/100)*temp$`Sum of Worked-Total`
  temp$`JLA Hours-Total per alloc`= ((temp$`%Allocated`)/100)*temp$`JLA Hours-Total`
  
  team_wise_summary = temp %>% group_by(`Team`)%>%
    summarise(`JLA Hours-Cluster` = sum(`JLA_hours`, na.rm=TRUE),
              `Sum of Worked-Cluster` = sum(`Jira_hours`, na.rm=TRUE),
              `Available Hours` = sum(`Available Hours`, na.rm=TRUE),
              `Available Hours` = sum(`Available Hours`, na.rm=TRUE),
              `Sum of Worked-Total` = sum(`Sum of Worked-Total per alloc`, na.rm=TRUE),
              `JLA Hours-Total` = sum(`JLA Hours-Total per alloc`, na.rm=TRUE)
    )
  team_wise_summary$`Head Count` =  team_wise_summary$`Available Hours`/Hours
  
  team_wise_summary$`Daily %age` = 
    round((team_wise_summary$`Sum of Worked-Total`/team_wise_summary$`Available Hours`)*100,0)
  team_wise_summary$`JLA %age` = round((team_wise_summary$`JLA Hours-Cluster`/team_wise_summary$`Available Hours`)*100,0)
  
  
  
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
  
  team_wise_summary[,-1]= round(team_wise_summary[,-1],0)
  
  return(team_wise_summary)
}

calculate_project_summary<- function(cluster_wise_resource_logging,Hours){
  # Project Summary Statistics -------------------------------------------------
  temp=cluster_wise_resource_logging
  temp$`Sum of Worked-Total per alloc`= ((temp$`%Allocated`)/100)*temp$`Sum of Worked-Total`
  temp$`JLA Hours-Total per alloc`= ((temp$`%Allocated`)/100)*temp$`JLA Hours-Total`
  
  project_wise_summary = temp %>% group_by(Groups,`PMO Projects Name`)%>%
    summarise(`JLA Hours-Cluster` = sum(`JLA_hours`, na.rm=TRUE),
              `Sum of Worked-Cluster` = sum(`Jira_hours`, na.rm=TRUE),
              `Available Hours` = sum(`Available Hours`, na.rm=TRUE),
              `Available Hours` = sum(`Available Hours`, na.rm=TRUE),
              `Sum of Worked-Total` = sum(`Sum of Worked-Total per alloc`, na.rm=TRUE),
              `JLA Hours-Total` = sum(`JLA Hours-Total per alloc`, na.rm=TRUE)
    )
  project_wise_summary$`Head Count` =  project_wise_summary$`Available Hours`/Hours
  
  project_wise_summary$`Daily %age` = 
    round((project_wise_summary$`Sum of Worked-Total`/project_wise_summary$`Available Hours`)*100,0)
  project_wise_summary$`JLA %age` = round((project_wise_summary$`JLA Hours-Cluster`/project_wise_summary$`Available Hours`)*100,0)
  
  
  
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
  
  project_wise_summary[,-(1:2)]= round(project_wise_summary[,-(1:2)],0)
  project_wise_summary$Groups[nrow(project_wise_summary)]="Grand Total"
  
  return(project_wise_summary)
}
get_cluster_activity <- function(jira_activity_data){
  
  ref_df <- read_excel("Input Files/Jira Reference sheet.xlsx")
  proj_group= data.frame( matrix(nrow=nrow(ref_df),ncol=2))
  proj_group =ref_df[,c("Projects","Group")]
  
  Activity_df= data.frame( matrix(nrow=nrow(ref_df),ncol=2))
  Activity_df =ref_df[,c("Activity Type","Heads")]
  Activity_df=Activity_df[complete.cases(Activity_df),]
  
  
  df =  merge(jira_activity_data,proj_group, by.x = "Project Name", by.y="Projects")
  df= merge(df ,Activity_df )
  
  temp=df %>% group_by(Group,Heads) %>%summarise(Hours = sum(Hours, na.rm = TRUE))
  
  cluster_wise_activity_seg=temp %>% 
    mutate(rn = row_number()) %>%
    spread(Heads,Hours) %>%
    select(-rn)
  cluster_wise_activity_seg[is.na(cluster_wise_activity_seg)] <- 0
  
  cluster_wise_activity_seg=cluster_wise_activity_seg %>% group_by(Group) %>%summarise(across(everything(), list(sum)))
  colnames(cluster_wise_activity_seg)=gsub("_1", "", colnames(cluster_wise_activity_seg))
  
  cluster_wise_activity_seg[(nrow(cluster_wise_activity_seg)+1),]=NA
  cluster_wise_activity_seg$Group[nrow(cluster_wise_activity_seg)] = c("Grand Total")
  cluster_wise_activity_seg[nrow(cluster_wise_activity_seg),2:ncol(cluster_wise_activity_seg)]= 
    lapply(cluster_wise_activity_seg[-nrow(cluster_wise_activity_seg),2:ncol(cluster_wise_activity_seg)],sum)
  
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
  cluster_wise_activity_seg[,2:ncol(cluster_wise_activity_seg)]=round(cluster_wise_activity_seg[,2:ncol(cluster_wise_activity_seg)])
  
  return(cluster_wise_activity_seg)
}

print_CWS<- function(cluster_wise_summary,Hours){
  if(Hours!=8)
  {
    colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="Sum of Worked-Total"] = "Sum of Worked"
    colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="JLA Hours-Total"] = "JLA Hours"
    
    colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="Daily %age"] = "Mid-Day %age"
    cws_order=c("Groups","Head Count","Available Hours","Sum of Worked","JLA Hours","Mid-Day %age","JLA %age")
  }
  else{
    cws_order=c("Groups","Head Count","Available Hours","Sum of Worked-Cluster","Sum of Worked-Total",
                "JLA Hours-Cluster","JLA Hours-Total","Daily %age","JLA %age")
    
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
    cws_order=c("Team","Head Count","Available Hours","Sum of Worked-Cluster","Sum of Worked-Total",
                "JLA Hours-Cluster","JLA Hours-Total","Daily %age","JLA %age")
    
  }
  
  cluster_wise_summary =cluster_wise_summary[,cws_order]
  
  return(cluster_wise_summary)
  
}
print_PWS<- function(cluster_wise_summary,Hours){
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
    colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="PMO Projects Name"] = "Project"
    colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="Groups"] = "Group"
    
    cws_order=c("Group","Project","Head Count","Available Hours","Sum of Worked-Cluster","Sum of Worked-Total",
                "JLA Hours-Cluster","JLA Hours-Total","Daily %age","JLA %age")
    
  }
  
  cluster_wise_summary =cluster_wise_summary[,cws_order]
  
  return(cluster_wise_summary)
  
}

print_CWRL<- function(cluster_wise_resource_logging){
  
  CWRL=cluster_wise_resource_logging
  colnames(CWRL)[colnames(CWRL)=="Jira_hours"] = "Sum of Worked-Cluster"
  colnames(CWRL)[colnames(CWRL)=="JLA_hours"] = "JLA Hours-Cluster"
  
  
  cwrl_order=c("Emp Code","Resource Name","Groups","PMO Projects Name", 
               "Team","%Allocated","Available Hours","Sum of Worked-Cluster",
               "Sum of Worked-Total","JLA Hours-Cluster","JLA Hours-Total")
  CWRL=CWRL[,cwrl_order]
}
print_bar_plot <- function(cluster_wise_summary){
  # x <- c('Product A', 'Product B', 'Product C')
  # y <- c(20, 14, 23)
  # y2 <- c(16,12,27)
  # text <- c('27% market share', '24% market share', '19% market share')
  # data <- data.frame(x, y, y2, text)
  colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="Daily %age"] = "day_percentage"
  colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="JLA %age"] = "jla_percentage"
  
  # colnames(cluster_wise_summary)[colnames(cluster_wise_summary)=="Groups"] = "Group"
  
  
  fig <- cluster_wise_summary %>% plot_ly()
  fig <- fig %>% add_trace(x = ~Groups, y = ~day_percentage, type = 'bar', name ="Jira %",
                           text = cluster_wise_summary$day_percentage, textposition = 'auto',
                           marker = list(color = 'rgb(158,202,225)',
                                         line = list(color = 'rgb(8,48,107)', width = 1.5)))
  fig <- fig %>% add_trace(x = ~Groups, y = ~jla_percentage, type = 'bar',name ="JLA %",
                           text = cluster_wise_summary$jla_percentage, textposition = 'auto',
                           marker = list(color = 'rgb(58,200,225)',
                                         line = list(color = 'rgb(8,48,107)', width = 1.5)))
  fig <- fig %>% layout(title = "Mid-day JLA Report",
                        barmode = 'group',
                        xaxis = list(title = ""),
                        yaxis = list(title = ""))
  
  return(fig)
}

ui <- dashboardPage(
  dashboardHeader(title = "NFS Daily Jira Compliance"),
  dashboardSidebar(
    # size = "wide",
    sidebarMenu(
      menuItem(tabName = "upload_data", text = "Upload Data"),
      menuItem(tabName = "dashboard", text = "Dashboard"),
      menuItem(tabName = "detailed_view", text = "Detailed view")
      
    )
  ),
  dashboardBody(
    tabItems(
      # selected = 1,
      tabItem(
        tabName = "upload_data",
        
        box( width=16,status = "primary",# solidHeader = TRUE, title="Table",
             fileInput("pmo_input", "Upload PMO File"),
             br(),
             # tableOutput("table1")
        ),
        box( width=16, status = "primary" ,# solidHeader = TRUE, title="Table",
             fileInput("jira_input", "Upload Jira File"),
             br(),
             # tableOutput("table1")
        ),
        dateInput('date',
                  label = 'Select Date',
                  value = Sys.Date()-1
        ),
        # radioButtons("day_type","Select Day ",
        #              choices = list("Half Day" = 4,
        #                             "Full Day" = 8),
        #              selected = character(0)
        # ),
        numericInput("day_type", "Select Hours:", 4, min = 4, max = 8),
        actionButton("analysis","Analyze"),
        # downloadButton('downloadData', 'Download Results')
        actionButton('downloadData', 'Download Results')
        
      ),
      tabItem(
        tabName = "dashboard", dataTableOutput("summary_table")
        # fluidRow( dataTableOutput("summary_table")),
        # fluidRow(
        # selectInput("display", "Display data by :", c("Cluster","Project","Team")),
        # dataTableOutput("type_table")
        # )
        
      )
      ,
      tabItem(
        tabName = "detailed_view",
        fluidRow(
          box(
            width = 4, background = "light-blue",
            "A box with a solid black background" ,
            selectInput("summary_data_type", "Display data by :", c("Cluster","Project","Team")),
            
          ),
          box(
            title = "Title 5", width = 4, background = "light-blue",
            "A box with a solid light-blue background",dataTableOutput("type_table")
          )
        )
      )
    )
  )
)

server <- function(input, output) { 
  ####### Initializing Variables
  alloc_data = jira_data = NULL
  val <- reactiveValues()
  val$cluster_wise_summary = data.frame()
  val$cluster_wise_resource_logging = data.frame()
  val$hours_df= data.frame()
  val$cluster_wise_activity_seg = data.frame
  val$team_wise_summary = data.frame()
  val$project_wise_summary = data.frame()
  observeEvent(input$analysis, {
    jira_data= jira_users= alloc_data =NULL
    
    ###### Save PMO input File
    
    inFile <- input$pmo_input
    if (!is.null(inFile)) {   
      dataFile <- read_excel(inFile$datapath, sheet=1)
      dat <- data.frame(dataFile, stringsAsFactors=FALSE)
      colnames(dat) <- colnames(dataFile)
      alloc_data = dat
      # write.xlsx(dat, file = "pmo_input.xlsx")
    }
    
    ###### Save Jira input File
    inFile <- input$jira_input
    if (!is.null(inFile)) {   
      Worklogs <- read_excel(inFile$datapath, sheet=1)
      users <- read_excel(inFile$datapath, sheet=2)
      
      dat <- data.frame(Worklogs, stringsAsFactors=FALSE)
      names(dat) <- gsub(x = names(dat),pattern = "\\.",replacement = " ")
      jira_data = dat
      
      dat <- data.frame(users, stringsAsFactors=FALSE)
      names(dat) <- gsub(x = names(dat),pattern = "\\.",replacement = " ")
      jira_users = dat
      # write.xlsx(dat, file = "jira_input.xlsx")
    }
    
    ################################
    
    if (!is.null(jira_data)  & !is.null(alloc_data))
    {
      
      Mapping = get_mapping()
      
      jira_data$`Work date`=as.Date(jira_data$`Work date`, origin = "1899-12-30")
      jira_data$`Work date`=strftime(jira_data$`Work date`, format="%Y-%m-%d")
      jira_data=jira_data[jira_data$`Work date`==input$date,]
      if (nrow(jira_data) > 0 ){
        showModal(modalDialog("Processing....", footer=NULL))
        jira_data = pre_process_jira_data(jira_data)
        jira_activity_data= jira_data[,c("Emp Code","Hours", "Project Name","Activity Type" )]
        jira_data =jira_data[,c("Emp Code","Full name","Hours", "Team", "JLA Activity Log", "Project Name" )]
        
        alloc_data = pre_process_pmo_data(alloc_data)
        
        jira_data =merge(Mapping, jira_data, by.x="JIRA Projects Name",by.y="Project Name",all.y=TRUE)
        
        JLA_data = jira_data[jira_data$`JLA Activity Log`=="Yes",]
        JLA_team= JLA_data %>% group_by( `PMO Projects Name`, Groups,`Emp Code`) %>% summarise(JLA_hours = sum(Hours))
        JLA_total_hours=JLA_data %>% group_by(`Emp Code`) %>% summarise(JLA_hours = sum(Hours))
        Jira_team= jira_data %>% group_by(`PMO Projects Name`, Groups,`Emp Code`) %>% summarise( Jira_hours= sum(Hours))
        
        
        team =merge(Jira_team, JLA_team,  all=TRUE )
        
        rm(JLA_data,JLA_team,Mapping)
        # Resource Allocation Wise Statistics -------------------------------------------------
        cluster_wise_resource_logging =merge( team , alloc_data, by=c("Emp Code","Groups","PMO Projects Name")#,"PMO Projects Name"
                                              , all=TRUE )
        
        val$cluster_wise_resource_logging = calculate_resource_allocation(cluster_wise_resource_logging,JLA_total_hours,as.integer(input$day_type))
        # Cluster Summary Statistics -------------------------------------------------
        val$cluster_wise_summary = calculate_cluster_summary(val$cluster_wise_resource_logging,as.integer(input$day_type))
        # Team Summary Statistics -------------------------------------------------
        val$team_wise_summary = calculate_team_summary(val$cluster_wise_resource_logging,as.integer(input$day_type))
        # Project Summary Statistics -------------------------------------------------
        val$project_wise_summary = calculate_project_summary(val$cluster_wise_resource_logging,as.integer(input$day_type))
        
        # --------------------# Handling Un-allocated Resources
        # Activity Type Statistics -------------------------------------------------
        
        val$cluster_wise_activity_seg = get_cluster_activity (jira_activity_data)
        
        # Calculating HOURS Summary Statistics -------------------------------------------------
        jira_users=jira_users[complete.cases(jira_users$`Full name`),]
        hours_df = data.frame(matrix(nrow=4, ncol=2))
        colnames(hours_df) = c("Type","Count")
        hours_df$Type=c('>= 10 hours','>= 8 hours','>= 1 and < 4 hours','Did Not Log')
        
        hours_df[hours_df$Type==">= 10 hours",]$Count =length(jira_users[jira_users$`Worked`>=10, ]$`Worked`)
        hours_df[hours_df$Type==">= 8 hours",]$Count=length(jira_users[jira_users$`Worked`>=8, ]$`Worked`)
        hours_df[hours_df$Type==">= 1 and < 4 hours",]$Count =length(jira_users[(jira_users$`Worked`>=1 & jira_users$`Worked`<4), ]$`Worked`)
        hours_df[hours_df$Type=="Did Not Log",]$Count =length(jira_users[jira_users$Worked ==0,]$`Worked`)
        hours_df[nrow(hours_df)+1,] = "NA"
        hours_df$Type[nrow(hours_df)] = "Total Jira Users"
        hours_df[hours_df$Type=="Total Jira Users",]$Count  = nrow(jira_users)
        # hours_df[hours_df$Type=="Total Head Count",]$Count  = nrow(jira_users)
        
        
        
        val$hours_df = hours_df
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
  
  observeEvent(input$downloadData, {
    path=choose.dir()
    showModal(modalDialog("saving Data!"))
    file_name = paste(path,"/","NFS Daily Compliance_",as.character(input$date),".xlsx", sep="" )
    # r_CWS_cols=
    
    CWS = print_CWS(val$cluster_wise_summary, as.integer(input$day_type))
    CWRL =print_CWRL(val$cluster_wise_resource_logging)
    TWS = print_TWS(val$team_wise_summary)
    PWS = print_PWS(val$project_wise_summary)
    if(as.integer(input$day_type)!=8)
    {
      # list_of_datasets <- list("Summary" = CWS, 
      #                          "Cluster wise resource logging" = CWRL)
      # file_name = paste(path,"/","NFS Mid-day Compliance_",as.character(input$date),".xlsx", sep="" )
      # openxlsx::write.xlsx(list_of_datasets, file = file_name)
      wb <- openxlsx::createWorkbook("Summary")
      openxlsx::addWorksheet(wb, sheet = "Summary", gridLines = TRUE)
      openxlsx::addWorksheet(wb, "Cluster wise resource logging", gridLines = TRUE)
      
      openxlsx::writeData(wb, sheet = "Summary", CWS, rowNames = FALSE)
      openxlsx::writeData(wb, sheet = "Cluster wise resource logging" , CWRL, rowNames = FALSE)
      
      wb=get_table_style(wb, 1, nrow(CWS),ncol(CWS))
      wb=get_table_style(wb, 2, nrow(CWRL),ncol(CWRL))
      openxlsx::saveWorkbook(wb, file_name, overwrite = TRUE)
      
      
      fig=print_bar_plot(val$cluster_wise_summary)
      export(fig, file = paste(path,"/","chart.png", sep="" ))
      
      
    }
    else
    {
      wb = createWorkbook()
      sheet = createSheet(wb, "Summary")
      
      addDataFrame(as.data.frame(CWS), sheet=sheet, startColumn=1, row.names=FALSE)
      addDataFrame(as.data.frame(val$hours_df), sheet=sheet, startColumn=15, row.names=FALSE)
      addDataFrame(as.data.frame(TWS), sheet=sheet, startRow = 18, startColumn=1, row.names=FALSE)
      addDataFrame(as.data.frame(PWS), sheet=sheet, startRow = 50, startColumn=1, row.names=FALSE)
      
      sheet = createSheet(wb, "Cluster wise resource logging")
      addDataFrame(as.data.frame(CWRL), sheet=sheet, startColumn=1, row.names=FALSE)
      
      sheet = createSheet(wb, "Cluster wise activity")
      addDataFrame(as.data.frame(val$cluster_wise_activity_seg), sheet=sheet, startColumn=1, row.names=FALSE)
      
      saveWorkbook(wb, file_name)
      
      # list_of_datasets <- list("Summary" = CWS,
      #                        "Cluster wise resource logging" = CWRL,
      #                        "Cluster wise activity"=val$cluster_wise_activity_seg,
      #                        "Time Log. Head Count"= val$hours_df)
    }
    
    showModal(modalDialog("saving finished!"))
  })
  
  observeEvent(input$summary_data_type,{
    
    
  })
  
  output$summary_table <- DT::renderDataTable(val$cluster_wise_summary,
                                              options = list(paging = F, dom = 't'),
                                              #,scrollX = TRUE
                                              rownames = FALSE)
  output$head_count <- renderText({ 
    req(val$cluster_wise_summary)
    round(sum(val$cluster_wise_summary$`Head Count`),0) })
  
  
  
}

shinyApp(ui, server)
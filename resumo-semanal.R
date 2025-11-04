library(AzureGraph)
library(Microsoft365R)
library(RPostgreSQL)
library(dplyr)
library(writexl)
library(lubridate)


source('param.R')
setwd(wd)

# Miscellaneous Functions
source('miscellaneous.R')


# Initialize a connection object outside the tryCatch block
con <- NULL

# Use a tryCatch block to handle potential connection errors
tryCatch({
  
  # Create a connection to the PostgreSQL database
  write_log(log_message = "########################################## Resumo Semanal #############################################", log_file = log_file_resumo)
  log_msg_db_con         <-  paste0(Sys.time(), "  [sidmat] - Conecting to postgresql server...")
  write_log(log_message  <-  log_msg_db_con,log_file = log_file_resumo)
  
  con <- dbConnect(
    PostgreSQL(),
    host = db_host,
    port = db_port,
    dbname = db_name,
    user = db_user,
    password = db_password
  )
  
  # Get the current date
  current_date <- Sys.Date() 
  #current_date <- as.Date('2024-10-28')  # Testing with October 2024 data
  
  
  # Get the first day of the month
  first_day_of_month <- floor_date(current_date, unit = "month")
  # Get the last day of the month
  last_day_of_month <- ceiling_date(current_date, unit = "month") - 1
  
  monday_of_week <- current_date - (wday(current_date) - 2) %% 7
  friday_of_week <- monday_of_week + 4
  
  period_semanal <- paste0(monday_of_week, " a ", friday_of_week)
  period_mensal <- paste0(first_day_of_month, " a ", last_day_of_month)
  
  
  # Read data from the table
  query_resumo_mensal <- paste("

select req.unidade_sanitaria, us.nome, req.data_requisicao, a.area,  gs.nr_guia, gs.data_guia, s.name
from  api.requisicao req
inner join api.material mat on mat.id = req.material
inner join api.area a on a.id = mat.area
left join api.guia_saida gs on gs.id = req.nr_guia
left join api.status s on s.id = gs.status
left join api.unidade_sanitaria us on us.id = req.unidade_sanitaria

where req.canceled = 'No' and req.data_requisicao::date between '" , first_day_of_month, "' and '",  last_day_of_month, "' ;")
  
  # Read data from the table
  query_resumo_semanal <- paste("

select req.unidade_sanitaria, us.nome, req.data_requisicao, a.area,  gs.nr_guia, gs.data_guia, s.name
from  api.requisicao req
inner join api.material mat on mat.id = req.material
inner join api.area a on a.id = mat.area
left join api.guia_saida gs on gs.id = req.nr_guia
left join api.status s on s.id = gs.status
left join api.unidade_sanitaria us on us.id = req.unidade_sanitaria

where req.canceled = 'No' and req.data_requisicao::date between '" , monday_of_week, "' and '",  friday_of_week, "' ;")
  
  df_requisicoes_mensal <- dbGetQuery(con, query_resumo_mensal)
  df_requisicoes_semanal <- dbGetQuery(con, query_resumo_semanal)
  
  log_msg_data_check <- paste0(Sys.time(), "  [sidmat] - Found ", nrow(df_requisicoes_mensal), " monthly records and ", nrow(df_requisicoes_semanal), " weekly records")
  write_log(log_message = log_msg_data_check, log_file = log_file_resumo)
  message(log_msg_data_check)
  
  if( nrow(df_requisicoes_mensal ) > 0 | nrow(df_requisicoes_semanal) > 0 ){
    
    query_colaborador_area <- "select c.nome, c.email, a.id, a.area  from api.colaborador c  inner join api.colaborador_area ca on c.id = ca.colaborador inner join api.area a on a.id = ca.area  inner join api.usuario u on u.colaborador = c.id where u.status = 'Active';"
    df_colaborador_area <- dbGetQuery(con,query_colaborador_area )
    
    
    df_total_requisicoes_semanal     <- df_requisicoes_semanal %>% group_by(area) %>% summarise(total_requisicoes = n())
    df_total_pendentes_semanal       <- df_requisicoes_semanal %>% filter(is.na(nr_guia)) %>% group_by(area) %>% summarise(pendentes = n())
    df_total_processadas_semanal     <- df_requisicoes_semanal %>%   filter(!is.na(nr_guia) & name=="NOVA") %>% group_by(area) %>% summarise(processadas = n())
    df_total_entregues_semanal       <- df_requisicoes_semanal %>%  filter(!is.na(nr_guia) & name=="ENTREGUE") %>% group_by(area) %>% summarise(entregues = n())
    df_resumo_semanal                <- df_total_requisicoes_semanal %>% left_join(df_total_pendentes_semanal, by ='area') %>% left_join(df_total_processadas_semanal, by ='area')  %>% left_join(df_total_entregues_semanal, by ='area')
    
    df_total_requisicoes_mensal     <- df_requisicoes_mensal %>% group_by(area) %>% summarise(total_requisicoes = n())
    df_total_pendentes_mensal       <- df_requisicoes_mensal %>% filter(is.na(nr_guia)) %>% group_by(area) %>% summarise(pendentes = n())
    df_total_processadas_mensal     <- df_requisicoes_mensal %>%   filter(!is.na(nr_guia) & name=="NOVA") %>% group_by(area) %>% summarise(processadas = n())
    df_total_entregues_mensal       <- df_requisicoes_mensal %>%  filter(!is.na(nr_guia) & name=="ENTREGUE") %>% group_by(area) %>% summarise(entregues = n())
    df_resumo_mensal                <- df_total_requisicoes_mensal %>% left_join(df_total_pendentes_mensal, by ='area') %>% left_join(df_total_processadas_mensal, by ='area')  %>% left_join(df_total_entregues_mensal, by ='area')
    
    
    df_resumo_semanal[is.na(df_resumo_semanal)] <- 0
    df_resumo_mensal[is.na(df_resumo_mensal)] <- 0
    
    # Create a Microsoft Graph login
    # TODO Run only once...
    log_msg_graph_login<- paste0(Sys.time(), "  [sidmat] - Create a Microsoft Graph login ...")
    write_log(log_message = log_msg_graph_login, log_file = log_file_resumo)
    
    gr <- create_graph_login(tenant, app, password=pwd, auth_type="client_credentials")
    
    if(is.environment(gr)){
         
        email_addr <- 'nunomoura@ccsaude.org.mz'
      
        user <- gr$get_user(email=email_addr)
        outlook <- user$get_outlook()
        
        # get all areas
        areas <- unique(df_resumo_semanal$area)
        
        for (area in areas) {
          write_log(log_message = "----------------------------------------------------------------------------------------------------------------------- ", log_file = log_file_resumo)
          log_msg_process_area<- paste0(Sys.time(), "  [sidmat] - Processando dados da area : { ", area," } ")
          write_log(log_message = log_msg_process_area, log_file = log_file_resumo)
          write_log(log_message = "----------------------------------------------------------------------------------------------------------------------- ", log_file = log_file_resumo)
          message(log_msg_process_area)
          
          area_name <- area
          emails_responsavel_area <- df_colaborador_area[which(df_colaborador_area$area==area_name),c("email")]
          
          # This should not happen
          if(length(emails_responsavel_area)==0){
            log_msg_email_missing <- paste0(Sys.time(), "  [sidmat] - Error e-mail do responsavel da area { ", area," } nao foi encontrado. ")
            write_log(log_message = log_msg_email_missing, log_file = log_file_resumo)
            break
          }
          temp_df <- df_resumo_semanal %>% filter(area==area_name)
          
          # remove special character from area_name 
          if(grepl(pattern = '/',x = area,ignore.case = TRUE)){
            # remove 
            area_name <- gsub(pattern = '/',replacement = '_',ignore.case = TRUE,x = area_name)
          }
          
         
          # write resumo semanal to disk
          temp_path <- paste0(xls_file_dir,"/resumo_",area_name,"_",current_date,".xlsx")
          assign(paste0("df_",area_name), value = temp_df, envir = .GlobalEnv)
          # write file to disk and use as attachment 
          # write_xlsx(
          #   x         = temp_df,
          #   path      = temp_path,
          #   col_names = TRUE,
          #   format_headers = TRUE,
          #   use_zip64 = FALSE)
          
          
          # Mais de um responsavel da area
          if(length(emails_responsavel_area) >1) {
            for (email in emails_responsavel_area) {
            
              response <- microsoft_365r_notify_resumo_semanal(outlook = outlook,recipient = email ,df.resumo = temp_df,area.name = area_name, period = period_semanal, cc_recipients = c('nunomoura@ccsaude.org.mz'), bcc_recipients = c('agnaldosamuel@ccsaude.org.mz'))
              received_date <- response$properties$receivedDateTime
              
              # Message received--> update notification status in material table
              if(substr(response$properties$receivedDateTime,1,4)==substr(Sys.Date(),1,4)){ 
                log_msg_notification<- paste0(Sys.time(), "  [sidmat] - Notificacao enviada para : { ", area_name," - ",  email, " } ")
                write_log(log_message = log_msg_notification, log_file = log_file_resumo)
                message(log_msg_notification)
              
              }
              
            }
            
            
          }
           else {
             
             response <- microsoft_365r_notify_resumo_semanal(outlook = outlook,recipient =  emails_responsavel_area,df.resumo = temp_df,area.name = area_name, period = period_semanal, cc_recipients = c('nunomoura@ccsaude.org.mz'), bcc_recipients = c('agnaldosamuel@ccsaude.org.mz'))
             received_date <- response$properties$receivedDateTime
             
             # Message received--> update notification status in material table
             if(substr(response$properties$receivedDateTime,1,4)==substr(Sys.Date(),1,4)){ 
               log_msg_notification<- paste0(Sys.time(), "  [sidmat] - Notificacao enviada para : { ", area_name," - ",  emails_responsavel_area, " } ")
               write_log(log_message = log_msg_notification, log_file = log_file_resumo)
               message(log_msg_notification)
               
             }
           }          
          write_log(log_message = "########################################################################################################################", log_file = log_file_resumo)
        

        }
        
        #TODO Personalizar envio automatico de emails
        #Send Resumo Semanal to nuno and Dra. Shital
        response <- microsoft_365r_notify_resumo_semanal_mensal(outlook = outlook,recipient =  "shitalmobaracaly@ccsaude.org.mz",df.resumo.mensal = df_resumo_mensal,df.resumo.semanal = df_resumo_semanal,
                                                               area.name =array_to_str(areas),period.semanal = period_semanal, period.mensal = period_mensal )
        #Send Resumo Semanal to Nuno
        response <- microsoft_365r_notify_resumo_semanal_mensal(outlook = outlook,recipient =  "nunomoura@ccsaude.org.mz",df.resumo.mensal = df_resumo_mensal,df.resumo.semanal = df_resumo_semanal,
                                                              area.name =array_to_str(areas),period.semanal = period_semanal, period.mensal = period_mensal )
        #
        #Send Resumo Semanal to Hugo
        #response <- microsoft_365r_notify_resumo_semanal_mensal(outlook = outlook,recipient =  "hugoazevedo@ccsaude.org.mz",df.resumo.mensal = df_resumo_mensal,df.resumo.semanal = df_resumo_semanal,
        #                                                        area.name =array_to_str(areas),period.semanal = period_semanal, period.mensal = period_mensal )
        
        #Send Resumo Semanal to Agnaldo
        response <- microsoft_365r_notify_resumo_semanal_mensal(outlook = outlook,recipient =  "agnaldosamuel@ccsaude.org.mz",df.resumo.mensal = df_resumo_mensal,df.resumo.semanal = df_resumo_semanal,
                                                                area.name =array_to_str(areas),period.semanal = period_semanal, period.mensal = period_mensal )

 
        
        if (!is.null(con)) {
          # Close the database connection in the finally block
          message("Closing database connection")
          dbDisconnect(con)
        }
        
    } else {
      
      log_msg_graph_login<- paste0(Sys.time(), "  [sidmat] - Error : Failed to Create a Microsoft Graph login ...")
      write_log(log_message = log_msg_graph_login, log_file = log_file_resumo)
      
      if (!is.null(con)) {
        # Close the database connection in the finally block
        message("Closing database connection")
        dbDisconnect(con)
      }
      
    }
  } else {
    
    log_msg_no_req <- paste0(Sys.time(), "  [sidmat] - Error : Nao existem requisicoes neste periodo")
    write_log(log_message = log_msg_no_req, log_file = log_file_resumo) 
    message(log_msg_no_req)
    if (!is.null(con)) {
      # Close the database connection in the finally block
      message("closing  database connection ")
      dbDisconnect(con)
    }
    
    
  }
  
  
}, error = function(e) {
  
   log_msg_error <- paste0(Sys.time(), "  [sidmat] - Unknown error ...")
   log_msg_error_message <- paste0(Sys.time(), "  [sidmat] - Error message: ", e$message)
   message(log_msg_error_message)

  if (!is.null(con)) {
    # Close the database connection in the finally block
    message("ErrorL closing  database connection:", )
    dbDisconnect(con)
  }
  write_log(log_message = log_msg_error, log_file = log_file_resumo)
  write_log(log_message = log_msg_error_message, log_file = log_file_resumo)
  cat(paste("Error message: ", e$message, "\n"))
  
},finally = {
  if (!is.null(con)) {
    # Close the database connection in the finally block
    dbDisconnect(con)
  }
})



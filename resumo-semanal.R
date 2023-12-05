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
 # current_date <- Sys.Date()
  # Calculate Monday and Friday of the current week
  current_date = Sys.Date()
  monday_of_week <- current_date - (wday(current_date) - 2) %% 7
  friday_of_week <- monday_of_week + 4
  
  period <- paste0(monday_of_week, " a ", friday_of_week)
  
  
  # Read data from the table
  query_resumo<- paste("

select req.unidade_sanitaria, us.nome, req.data_requisicao, a.area,  gs.nr_guia, gs.data_guia, s.name
from  api.requisicao req
inner join api.material mat on mat.id = req.material
inner join api.area a on a.id = mat.area
left join api.guia_saida gs on gs.id = req.nr_guia
left join api.status s on s.id = gs.status
left join api.unidade_sanitaria us on us.id = req.unidade_sanitaria

where req.canceled = 'No' and req.data_requisicao::date between '" , monday_of_week, "' and '",  friday_of_week, "' ;")
  
  df_requisicoes <- dbGetQuery(con, query_resumo)
  
  if( nrow(df_requisicoes) > 0 ){
    
    query_colaborador_area <- "select c.nome, c.email, a.id, a.area from api.colaborador c inner join api.colaborador_area ca on c.id = ca.colaborador inner join api.area a on a.id = ca.area;"
    df_colaborador_area <- dbGetQuery(con,query_colaborador_area )
    
    df_total_requisicoes     <- df_requisicoes %>% group_by(area) %>% summarise(total_requisicoes = n())
    df_total_pendentes       <- df_requisicoes %>% filter(is.na(nr_guia)) %>% group_by(area) %>% summarise(pendentes = n())
    df_total_processadas     <- df_requisicoes %>%   filter(!is.na(nr_guia) & name=="NOVA") %>% group_by(area) %>% summarise(processadas = n())
    df_total_entregues       <- df_requisicoes %>%  filter(!is.na(nr_guia) & name=="ENTREGUE") %>% group_by(area) %>% summarise(entregues = n())
    df_resumo                <- df_total_requisicoes %>% left_join(df_total_pendentes, by ='area') %>% left_join(df_total_processadas, by ='area')  %>% left_join(df_total_entregues, by ='area')
    
    df_resumo[is.na(df_resumo)] <- 0
    
    # Create a Microsoft Graph login
    # TODO Run only once...
    log_msg_graph_login<- paste0(Sys.time(), "  [sidmat] - Create a Microsoft Graph login ...")
    write_log(log_message = log_msg_graph_login, log_file = log_file_resumo)
    
    gr <- create_graph_login(tenant, app, password=pwd, auth_type="client_credentials")
    
    if(is.environment(gr )){
         
        email_addr <- 'nunomoura@ccsaude.org.mz'
      
        user <- gr$get_user(email=email_addr)
        outlook <- user$get_outlook()
        
        # get all areas
        areas <- unique(df_resumo$area)
        
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
          temp_df <- df_resumo %>% filter(area==area_name)
          
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
            
              response <- microsoft_365r_notify_resumo_semanal(outlook = outlook,recipient = email ,df.resumo = temp_df,area.name = area_name, period = period)
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
            
            response <- microsoft_365r_notify_resumo_semanal(outlook = outlook,recipient =  emails_responsavel_area,df.resumo = temp_df,area.name = area_name, period = period)
            received_date <- response$properties$receivedDateTime
            
            # Message received--> update notification status in material table
            if(substr(response$properties$receivedDateTime,1,4)==substr(Sys.Date(),1,4)){ 
              log_msg_notification<- paste0(Sys.time(), "  [sidmat] - Notificacao enviada para : { ", area_name," - ",  email, " } ")
              write_log(log_message = log_msg_notification, log_file = log_file_resumo)
              message(log_msg_notification)
              
            }
          }
          
          
          write_log(log_message = "########################################################################################################################", log_file = log_file_resumo)
        

        }
        
        #TODO Personalizar envio automatico de emails
        #Send Resumo Semanal to nuno and Dra. Shital
        response <- microsoft_365r_notify_resumo_semanal(outlook = outlook,recipient =  "shitalmobaracaly@ccsaude.org.mz",df.resumo = df_resumo,area.name =  array_to_str(areas), period = period)
        #Send Resumo Semanal to nuno and Dra. Shital
        response <- microsoft_365r_notify_resumo_semanal(outlook = outlook,recipient =  "nunomoura@ccsaude.org.mz",df.resumo = df_resumo,area.name =  array_to_str(areas), period = period)
        #Send Resumo Semanal to Hugo
        response <- microsoft_365r_notify_resumo_semanal(outlook = outlook,recipient =  "hugoazevedo@ccsaude.org.mz",df.resumo = df_resumo,area.name =  array_to_str(areas), period = period)
        
        
        if (!is.null(con)) {
          # Close the database connection in the finally block
          message("Closing database connection")
          dbDisconnect(con)
        }
        
        
    else {
      
      log_msg_graph_login<- paste0(Sys.time(), "  [sidmat] - Error : Failed to Create a Microsoft Graph login ...")
      write_log(log_message = log_msg_graph_login, log_file = log_file_resumo)
      
      if (!is.null(con)) {
        # Close the database connection in the finally block
        message("Closing database connection")
        dbDisconnect(con)
      }
      
    }
  }
   }
  else {
    
    log_msg_no_req <- paste0(Sys.time(), "  [sidmat] - Error : Nao existem requisicoes neste periodo")
    write_log(log_message = log_msg_no_req, log_file = log_file_resumo) 
    
    if (!is.null(con)) {
      # Close the database connection in the finally block
      message("closing  database connection ")
      dbDisconnect(con)
    }
    
    
  }
  
  
}, error = function(e) {
  
  log_msg_error <- paste0(Sys.time(), "  [sidmat] - Unknown error ...")
  log_msg_error_message <- paste0(Sys.time(), "  [sidmat] - Error message: ", e$message)
  if (!is.null(con)) {
    # Close the database connection in the finally block
    message("closing  database connection ")
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



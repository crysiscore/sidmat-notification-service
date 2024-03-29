library(AzureGraph)
library(Microsoft365R)
library(RPostgreSQL)
library(dplyr)
library(writexl)

source('param.R')
setwd(wd)

# Miscellaneous Functions
source('miscellaneous.R')


#Not run
# Create auth token cache directory, otherwise it will prompt the the console for input
# create_AzureR_dir()
#### The AzureR packages save your login sessions so that you don’t need to reauthenticate each time. If you’re experiencing authentication failures, you can try clearing the saved data by running the following code:
#   
# AzureAuth::clean_token_directory()
# AzureGraph::delete_graph_login(tenant="mytenant")


# Initialize a connection object outside the tryCatch block
con <- NULL

# Use a tryCatch block to handle potential connection errors
tryCatch({
  
  # Create a connection to the PostgreSQL database
  write_log(log_message = "########################################################################################################################", log_file = log_file_guias)
  log_msg_db_con         <-  paste0(Sys.time(), "  [sidmat] - Conecting to postgresql server...")
  write_log(log_message  <-  log_msg_db_con,log_file = log_file_guias)
  
  con <- dbConnect(
    PostgreSQL(),
    host = db_host,
    port = db_port,
    dbname = db_name,
    user = db_user,
    password = db_password
  )
  
  log_msg_con_sucess <- paste0(Sys.time(), "  [sidmat] - acquired connection to postgres server")
  log_msg_db_query <- paste0(Sys.time(), "  [sidmat] - Quering database ...")
  
  write_log(log_message = log_msg_con_sucess, log_file = log_file_guias)
  write_log(log_message = log_msg_db_query, log_file = log_file_guias)
  
  # Specify the table you want to read from
  table_name <- "api.vw_confirmacao_guia"
  
  # Read data from the table
  query_guia<- paste("SELECT * FROM ", table_name)
  query_colaborador_area <- "select c.nome, c.email, a.id, a.area from api.colaborador c inner join api.colaborador_area ca on c.id = ca.colaborador inner join api.area a on a.id = ca.area;"
  df_guia_confirmada <- dbGetQuery(con, query_guia)
  df_colaborador_area <- dbGetQuery(con,query_colaborador_area )


  # NOT RUN (just for tests)
  # df_guia_confirmada <- df_guia_confirmada %>% filter(area %in% c("APSS","SMI/PTV"))
  ##df_colaborador_area$email="agnaldosamuel@ccsaude.org.mz"
  # df_guia_confirmada  <- filter(df_guia_confirmada, area %in% c("M&A"))
  #df_colaborador_area <- subset(df_colaborador_area, ! nome %in% c("Mauricio Timecane") )
  
  # only runs it there are new  confirmed guia delivery (notification_status ='P')
  if(nrow(df_guia_confirmada) > 0 && nrow(df_colaborador_area) > 0 ){ 
  
    removeOlderXlsxFiles(folder_path = xls_file_dir)
    # create a Microsoft Graph login
    #TODO Run only once...
    log_msg_graph_login<- paste0(Sys.time(), "  [sidmat] - Create a Microsoft Graph login ...")
    write_log(log_message = log_msg_graph_login, log_file = log_file_guias)
    
    gr <- create_graph_login(tenant, app, password=pwd, auth_type="client_credentials")
  
    
    if(is.environment(gr )){
      

      # get all areas
      areas <- unique(df_guia_confirmada$area)
      
      
      for (area in areas) {
        write_log(log_message = "-------------------------------------------------------------------------------", log_file = log_file_guias)
        log_msg_process_area<- paste0(Sys.time(), "  [sidmat] - Processando dados da area : { ", area," } ")
        write_log(log_message = log_msg_process_area, log_file = log_file_guias)
        write_log(log_message = "-------------------------------------------------------------------------------", log_file = log_file_guias)
        
        area_name <- area

        emails_responsavel_area <- df_colaborador_area[which(df_colaborador_area$area==area_name),c("email")]
        
        # This should not happen
        if(length(emails_responsavel_area)==0){
          log_msg_email_missing <- paste0(Sys.time(), "  [sidmat] - Error e-mail do responsavel da area { ", area," } nao foi encontrado. ")
          write_log (log_message = log_msg_email_missing, log_file = log_file_guias)
          break
        }
        temp_df <- df_guia_confirmada %>% filter(area==area_name)
        
        # retrieving another user's details
        tmp_confirmed_name <-  df_guia_confirmada$confirmed_by[1]
        tmp_confirmed_email <- df_colaborador_area$email[which(df_colaborador_area$nome==tmp_confirmed_name)]
        user <- gr$get_user(email=tmp_confirmed_email)
        outlook <- user$get_outlook()
        

        # remove special character from area_name 
        if(grepl(pattern = '/',x = area,ignore.case = TRUE)){
          # remove 
          area_name <- gsub(pattern = '/',replacement = '_',ignore.case = TRUE,x = area_name)
        }
        current_date <- Sys.Date()
        temp_path <- paste0(xls_file_dir,"/guias_",area_name,"_",current_date,".xlsx")
        assign(paste0("df_",area_name), value = temp_df, envir = .GlobalEnv)
        # write file to disk and use as attachment 
        write_xlsx(
          x=temp_df,
          path = temp_path,
          col_names = TRUE,
          format_headers = TRUE,
          use_zip64 = FALSE)
      
       # Mais de um responsavel da area
       if(length(emails_responsavel_area) >1) {
         
         for (email in emails_responsavel_area) {
            print(email)
           # get the ids of guia_Saida
           vec_guias_ids <- ""
           for (i in 1:nrow(temp_df)) {
             
             id <- temp_df$id[i]
             if(i== nrow(temp_df)){
               vec_guias_ids <- paste0(vec_guias_ids, id)
             } else {
               
               vec_guias_ids <- paste0(vec_guias_ids, id," , ")
             } 
             
             
           }
           
           response <- microsoft_365r_notify_guia_confirmation(outlook = outlook,recipient = email ,attachment = temp_path, vec_guias = array_to_str(temp_df$nr_guia))
           received_date <- response$properties$receivedDateTime
           
           # Message received--> update notification status in guia_saida table
           if(substr(response$properties$receivedDateTime,1,4)==substr(Sys.Date(),1,4)){
             log_msg_notification<- paste0(Sys.time(), "  [sidmat] - Notificacao enviada para : { ", area," - ",  email, " } ")
             write_log(log_message = log_msg_notification, log_file = log_file_guias)

             sql_update_notification <- paste0("update api.guia_saida set notification_status = 'S' where id in (", vec_guias_ids, ") ;")
             
             log_msg_notification_status_update<- paste0(Sys.time(), "  [sidmat] - Materiais actualizados  ids: { ",vec_guias_ids, " } ")
             write_log(log_message = log_msg_notification_status_update, log_file = log_file_guias)
             
             result <- dbSendQuery(con, sql_update_notification)
             dbClearResult(result)
             
           }
           
         }
        
      }
       else {
         
         # get the ids of guia_saidas
         vec_guias_ids <- ""
         for (i in 1:nrow(temp_df)) {
           
           id <- temp_df$id[i]
           if(i== nrow(temp_df)){
             
             vec_guias_ids <- paste0(vec_guias_ids, id)
             
           } else {
             
             vec_guias_ids <- paste0(vec_guias_ids, id," , ")
             
           } 
           
           
         }
         
        
        response <- microsoft_365r_notify_guia_confirmation(outlook = outlook,recipient = emails_responsavel_area ,attachment = temp_path, vec_guias =  array_to_str(temp_df$nr_guia))
        received_date <- response$properties$receivedDateTime
        
        # Message received--> update notification status in guia_saida table
        if(substr(response$properties$receivedDateTime,1,4)==substr(Sys.Date(),1,4)){
          
          log_msg_notification<- paste0(Sys.time(), "  [sidmat] - Notificacao enviada para : { ", area," - ",  emails_responsavel_area, " } ")
          write_log(log_message = log_msg_notification, log_file = log_file_guias)

          sql_update_notification <- paste0("update api.guia_saida set notification_status = 'S' where id in (", vec_guias_ids, ") ;")
          result <- dbSendQuery(con, sql_update_notification)
          log_msg_notification_status_update<- paste0(Sys.time(), "  [sidmat] - Guias de saida actualizados  ids: { ",vec_guias_ids, " } ")
          write_log(log_message = log_msg_notification_status_update, log_file = log_file_guias)
          dbClearResult(result)
          
        }
        
      }
       
       if (!is.null(con)) {
          # Close the database connection in the finally block
          dbDisconnect(con)
      }
        write_log(log_message = "########################################################################################################################", log_file = log_file_guias)
    
  }
  
    } else {
      
      # DO nothing 
      # The script will try again in the next cycle
      log_msg_graph_login<- paste0(Sys.time(), "  [sidmat] - Error : Failed to Create a Microsoft Graph login ...")
      write_log(log_message = log_msg_graph_login, log_file = log_file_guias)
    }
  }
  
  
}, error = function(e) {
  
  log_msg_error <- paste0(Sys.time(), "  [sidmat] - Unknown error ...")
  log_msg_error_message <- paste0(Sys.time(), "  [sidmat] - Error message: ", e$message)
  
  write_log(log_message = log_msg_error, log_file = log_file_guias)
  write_log(log_message = log_msg_error_message, log_file = log_file_guias)
  cat(paste("Error message: ", e$message, "\n"))
  
}, finally = {
  if (!is.null(con)) {
    # Close the database connection in the finally block
    dbDisconnect(con)
  }
})


  
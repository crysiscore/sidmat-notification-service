library(AzureGraph)
library(Microsoft365R)
library(RPostgreSQL)
library(dplyr)
library(writexl)

source('param.R')
setwd(wd)

# Miscellaneous Functions
source('miscellaneous.R')

# Initialize a connection object outside the tryCatch block
con <- NULL

# Use a tryCatch block to handle potential connection errors
tryCatch({
  
  # Create a connection to the PostgreSQL database
  write_log(log_message = "########################################################################################################################", log_file = log_file_material)
  log_msg_db_con         <-  paste0(Sys.time(), "  [sidmat] - Conecting to postgresql server...")
  write_log(log_message  <-  log_msg_db_con,log_file = log_file_material)
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
  
  write_log(log_message = log_msg_con_sucess, log_file = log_file_material)
  write_log(log_message = log_msg_db_query, log_file = log_file_material)
  message(log_msg_con_sucess)
  message(log_msg_db_query)
  
  
  # Specify the table you want to read from
  table_name <- "api.vw_novo_material"
  
  # Read data from the table
  query_material <- paste("SELECT * FROM ", table_name)
  query_colaborador_area <- "select c.nome, c.email, a.id, a.area  from api.colaborador c  inner join api.colaborador_area ca on c.id = ca.colaborador inner join api.area a on a.id = ca.area  inner join api.usuario u on u.colaborador = c.id where u.status = 'Active';"
  
  df_novo_material <- dbGetQuery(con, query_material)
  
  df_colaborador_area <- dbGetQuery(con,query_colaborador_area )
  
  # NOT RUN
  # df_novo_material <- df_novo_material %>% filter(area %in% c("APSS","SMI/PTV"))
  # df_novo_material  <- filter(df_novo_material, area %in% c("SMI/PTV","VBG","ACT","PCT"))
  
  
  # only runs it there are new materials (notification_status ='P')
  if(nrow(df_novo_material) > 0 && nrow(df_colaborador_area) > 0 ){ # Novo material importado
    
    removeOlderXlsxFiles(folder_path = xls_file_dir)
    # create a Microsoft Graph login
    #TODO Run only once...
    log_msg_graph_login<- paste0(Sys.time(), "  [sidmat] - Create a Microsoft Graph login ...")
    write_log(log_message = log_msg_graph_login, log_file = log_file_material)
    message(log_msg_graph_login)
    
    gr <- create_graph_login(tenant, app, password=pwd, auth_type="client_credentials")
    # retrieving another user's details
    user <- gr$get_user(email="nunomoura@ccsaude.org.mz")
    outlook <- user$get_outlook()
    
    if(is.environment(gr ) && is.environment(outlook) ){
      
      
      # get all areas
      areas <- unique(df_novo_material$area)
      
      for (area in areas) {
        write_log(log_message = "----------------------------------------------------------------------------------------------------------------------- ", log_file = log_file_material)
        log_msg_process_area<- paste0(Sys.time(), "  [sidmat] - Processando dados da area : { ", area," } ")
        write_log(log_message = log_msg_process_area, log_file = log_file_material)
        write_log(log_message = "----------------------------------------------------------------------------------------------------------------------- ", log_file = log_file_material)
        message(log_msg_process_area)
        
        area_name <- area
        emails_responsavel_area <- df_colaborador_area[which(df_colaborador_area$area==area_name),c("email")]
        
        # This should not happen
        if(length(emails_responsavel_area)==0){
          log_msg_email_missing <- paste0(Sys.time(), "  [sidmat] - Error e-mail do responsavel da area { ", area," } nao foi encontrado. ")
          write_log(log_message = log_msg_email_missing, log_file = log_file_material)
          break
        }
        temp_df <- df_novo_material %>% filter(area==area_name)
        
        # remove special character from area_name 
        if(grepl(pattern = '/',x = area,ignore.case = TRUE)){
          # remove 
          area_name <- gsub(pattern = '/',replacement = '_',ignore.case = TRUE,x = area_name)
        }
        current_date <- Sys.Date()
        temp_path <- paste0(xls_file_dir,"/material_",area_name,"_",current_date,".xlsx")
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
            
            response <- microsoft_365r_notify_new_material(outlook = outlook,recipient = email ,attachment = temp_path)
            received_date <- response$properties$receivedDateTime
            
            # Message received--> update notification status in material table
            if(substr(response$properties$receivedDateTime,1,4)==substr(Sys.Date(),1,4)){
              log_msg_notification<- paste0(Sys.time(), "  [sidmat] - Notificacao enviada para : { ", area," - ",  email, " } ")
              write_log(log_message = log_msg_notification, log_file = log_file_material)
              message(log_msg_notification)
              # get the ids of material
              vec_material_ids <- ""
              for (i in 1:nrow(temp_df)) {
                
                id <- temp_df$id[i]
                if(i== nrow(temp_df)){
                  vec_material_ids <- paste0(vec_material_ids, id)
                } else {
                  
                  vec_material_ids <- paste0(vec_material_ids, id," , ")
                } 
                
                
              }
              sql_udopate_notification <- paste0("update api.material set notification_status = 'S' where id in (", vec_material_ids, ") ;")
              
              log_msg_notification_status_update<- paste0(Sys.time(), "  [sidmat] - Materiais actualizados  ids: { ",vec_material_ids, " } ")
              write_log(log_message = log_msg_notification_status_update, log_file = log_file_material)
              message(log_msg_notification_status_update)
              result <- dbSendQuery(con, sql_udopate_notification)
              dbClearResult(result)
              
            }
            
          }
          
        } else {
          
          response <- microsoft_365r_notify_new_material(outlook = outlook,recipient = emails_responsavel_area ,attachment = temp_path)
          received_date <- response$properties$receivedDateTime
          
          # Message received--> update notification status in material table
          if(substr(response$properties$receivedDateTime,1,4)==substr(Sys.Date(),1,4)){
            
            log_msg_notification<- paste0(Sys.time(), "  [sidmat] - Notificacao enviada para : { ", area," - ",  emails_responsavel_area, " } ")
            write_log(log_message = log_msg_notification, log_file = log_file_material)
            message(log_msg_notification)
            # get the ids of material
            vec_material_ids <- ""
            for (i in 1:nrow(temp_df)) {
              
              id <- temp_df$id[i]
              if(i== nrow(temp_df)){
                
                vec_material_ids <- paste0(vec_material_ids, id)
                
              } else {
                
                vec_material_ids <- paste0(vec_material_ids, id," , ")
                
              } 
              
              
            }
            sql_udopate_notification <- paste0("update api.material set notification_status = 'S' where id in (", vec_material_ids, ") ;")
            result <- dbSendQuery(con, sql_udopate_notification)
            log_msg_notification_status_update<- paste0(Sys.time(), "  [sidmat] - Materiais actualizados  ids: { ",vec_material_ids, " } ")
            write_log(log_message = log_msg_notification_status_update, log_file = log_file_material)
            message(log_msg_notification_status_update)
            dbClearResult(result)
            
          }
          
        }
        
        
        write_log(log_message = "########################################################################################################################", log_file = log_file_material)
        
      }
      
      
      # Close the database connection in the finally block
      dbDisconnect(con)
      #rm(con)
      
    } else {
      
      # DO nothing 
      # The script will try again in the next cycle
      log_msg_graph_login<- paste0(Sys.time(), "  [sidmat] - Error : Failed to Create a Microsoft Graph login ...")
      write_log(log_message = log_msg_graph_login, log_file = log_file_material)
      message(log_msg_graph_login)
      
    }
  } else {
    #Close the database connection in the finally block
    log_msg <- paste0(Sys.time(), " [sidmat] - Nao ha materiais novos adicionados. ")
    message(log_msg)
    write_log(log_message = log_msg, log_file = log_file_material)
    dbDisconnect(con)
    #rm(con)
  }
  
  
}, error = function(e) {
  
  log_msg_error <- paste0(Sys.time(), "  [sidmat] - Unknown error ...")
  log_msg_error_message <- paste0(Sys.time(), "  [sidmat] - Error message: ", e$message)
  write_log(log_message = log_msg_error, log_file = log_file_material)
  write_log(log_message = log_msg_error_message, log_file = log_file_material)
  cat(paste("Error message: ", e$message, "\n"))
  # Close the database connection in the finally block
  dbDisconnect(con)
  #rm(con)
}, finally = {
  
  # Close the database connection in the finally block
  if (!is.null(con)) {
    # Close the database connection in the finally block
    dbDisconnect(con)
  }
  
})



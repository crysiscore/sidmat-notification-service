library(AzureGraph)
library(Microsoft365R)
library(RPostgreSQL)
library(dplyr)
library(writexl)


# Miscellaneous Functions
source('miscellaneous.R')

# PostgreSQL connection parameters
db_host <- Sys.getenv("POSTGRES_HOST")
db_port <- Sys.getenv("POSTGRES_PORT")
db_name <- Sys.getenv("POSTGRES_DB_NAME")
db_user <-Sys.getenv("POSTGRES_USER")
db_password <- Sys.getenv("POSTGRES_PASSWORD")

#directory to store new material disponivel file and logs for each area programatica
xls_file_dir <- "~/Documents/tmp" 
#xls_file_dir <- "/home/ccsadmin/Projects/"
log_file_sidmat <- paste0(xls_file_dir,"/", "sidmat_new_material_logs.txt")

#Not run
# Create auth token cache directory, otherwise it will prompt the the console for input
# create_AzureR_dir()
#### The AzureR packages save your login sessions so that you don’t need to reauthenticate each time. If you’re experiencing authentication failures, you can try clearing the saved data by running the following code:
#   
# AzureAuth::clean_token_directory()
# AzureGraph::delete_graph_login(tenant="mytenant")

# Microsoft App Registration
tenant <- "339b661c-15dc-4fdc-9c47-66f74d5eb137"
# the application/client ID of the app registration you created in AAD
app <- "be5c0edd-fa3d-4ba8-8bad-9a0a0d367ee9"
# retrieve the client secret (password) from an environment variable
pwd <- Sys.getenv("MS365R_CLIENT_SECRET")
# retrieve the user whose OneDrive we want to access
user <- Sys.getenv("MS365R_TARGET_USER")

# Initialize a connection object outside the tryCatch block
con <- NULL

# Use a tryCatch block to handle potential connection errors
tryCatch({
  
  # Create a connection to the PostgreSQL database
  write_log(log_message = "########################################################################################################################", log_file = log_file_sidmat)
  log_msg_db_con         <-  paste0(Sys.time(), "  [sidmat] - Conecting to postgresql server...")
  write_log(log_message  <-  log_msg_db_con,log_file = log_file_sidmat)
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
  
  write_log(log_message = log_msg_con_sucess, log_file = log_file_sidmat)
  write_log(log_message = log_msg_db_query, log_file = log_file_sidmat)
  
  # Specify the table you want to read from
  table_name <- "api.vw_novo_material"
  
  # Read data from the table
  query_material <- paste("SELECT * FROM ", table_name)
  query_colaborador_area <- "select c.nome, c.email, a.id, a.area from api.colaborador c inner join api.colaborador_area ca on c.id = ca.colaborador inner join api.area a on a.id = ca.area;"
  
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
    write_log(log_message = log_msg_graph_login, log_file = log_file_sidmat)
    
    gr <- create_graph_login(tenant, app, password=pwd, auth_type="client_credentials")
    # retrieving another user's details
    user <- gr$get_user(email="nunomoura@ccsaude.org.mz")
    outlook <- user$get_outlook()
    
    if(is.environment(gr ) && is.environment(outlook) ){
      

      # get all areas
      areas <- unique(df_novo_material$area)
      
      for (area in areas) {
        write_log(log_message = "----------------------------------------------------------------------------------------------------------------------- ", log_file = log_file_sidmat)
        log_msg_process_area<- paste0(Sys.time(), "  [sidmat] - Processando dados da area : { ", area," } ")
        write_log(log_message = log_msg_process_area, log_file = log_file_sidmat)
        write_log(log_message = "----------------------------------------------------------------------------------------------------------------------- ", log_file = log_file_sidmat)
        
        area_name <- area
        emails_responsavel_area <- df_colaborador_area[which(df_colaborador_area$area==area_name),c("email")]
        
        # This should not happen
        if(length(emails_responsavel_area)==0){
          log_msg_email_missing <- paste0(Sys.time(), "  [sidmat] - Error e-mail do responsavel da area { ", area," } nao foi encontrado. ")
          write_log(log_message = log_msg_email_missing, log_file = log_file_sidmat)
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
             write_log(log_message = log_msg_notification, log_file = log_file_sidmat)
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
             write_log(log_message = log_msg_notification_status_update, log_file = log_file_sidmat)
             
             result <- dbSendQuery(con, sql_udopate_notification)
             dbClearResult(result)
             
           }
           
         }
        
      }
       else {
        
        response <- microsoft_365r_notify_new_material(outlook = outlook,recipient = emails_responsavel_area ,attachment = temp_path)
        received_date <- response$properties$receivedDateTime
        
        # Message received--> update notification status in material table
        if(substr(response$properties$receivedDateTime,1,4)==substr(Sys.Date(),1,4)){
          
          log_msg_notification<- paste0(Sys.time(), "  [sidmat] - Notificacao enviada para : { ", area," - ",  emails_responsavel_area, " } ")
          write_log(log_message = log_msg_notification, log_file = log_file_sidmat)
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
          write_log(log_message = log_msg_notification_status_update, log_file = log_file_sidmat)
          dbClearResult(result)
          
        }
        
      }
       
       if (!is.null(con)) {
          # Close the database connection in the finally block
          dbDisconnect(con)
      }
      write_log(log_message = "########################################################################################################################", log_file = log_file_sidmat)
    
  }
  
    } else {
      
      # DO nothing 
      # The script will try again in the next cycle
      log_msg_graph_login<- paste0(Sys.time(), "  [sidmat] - Error : Failed to Create a Microsoft Graph login ...")
      write_log(log_message = log_msg_graph_login, log_file = log_file_sidmat)
    }
  }
  
  
}, error = function(e) {
  
  log_msg_error <- paste0(Sys.time(), "  [sidmat] - Unknown error ...")
  log_msg_error_message <- paste0(Sys.time(), "  [sidmat] - Error message: ", e$message)
  
  write_log(log_message = log_msg_error, log_file = log_file_sidmat)
  write_log(log_message = log_msg_error_message, log_file = log_file_sidmat)
  cat(paste("Error message: ", e$message, "\n"))
  
}, finally = {
  if (!is.null(con)) {
    # Close the database connection in the finally block
    dbDisconnect(con)
  }
})


  
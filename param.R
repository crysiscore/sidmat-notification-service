# Change here when deploying on server
#set working dir
wd <- '/home/ccsadmin/Projects/sidmat-notification-service/'

#Local tests
# wd <-getwd()

# PostgreSQL connection parameters
db_host <- Sys.getenv("POSTGRES_HOST")
db_port <- Sys.getenv("POSTGRES_PORT")
db_name <- Sys.getenv("POSTGRES_DB_NAME")
db_user <-Sys.getenv("POSTGRES_USER")
db_password <- Sys.getenv("POSTGRES_PASSWORD")

# Directory to store new guia_saida  file and logs for each area programatica
# Change here when deploying on server
xls_file_dir <- "/home/ccsadmin/Projects/sidmat-notification-service/tmp" 

# Local tests
#xls_file_dir <- "~/Projects/sidmat-notification-service" 

log_file_guias <- paste0(xls_file_dir,"/", "sidmat_guia_confirmation_logs.txt")
log_file_material <- paste0(xls_file_dir,"/", "sidmat_new_material_logs.txt")
log_file_resumo <- paste0(xls_file_dir,"/", "sidmat_resumo_logs.txt")
# Microsoft App Registration
tenant <-  Sys.getenv("MS365_TENANT")
# the application/client ID of the app registration you created in AAD
app <- Sys.getenv("MS365_APP")
# retrieve the client secret (password) from an environment variable
pwd <- Sys.getenv("MS365R_CLIENT_SECRET")
# retrieve the user whose OneDrive we want to access
user <- Sys.getenv("MS365R_TARGET_USER")

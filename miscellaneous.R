# compose an email with blastula
library(blastula)



microsoft_365r_notify_new_material <- function(outlook , recipient, attachment){ 
  bl_body <- "## Prezado(a) 

  Recebeu este email porque tem material disponivel para sua área  no armazém,
  Para fazer o plano de distribuição visite o endereco https://ccs-sidmat.vercel.app/auth/login.

  Cpts,

  CCS Logistica"
  
  
  bl_em <- compose_email(
    body=md(bl_body),
    footer=md("sent via Microsoft365R")
  )
  em <- outlook$create_email(bl_em, subject="Material disponivel no Armazem", to=recipient)
  
  # add an attachment and send it
  em$add_attachment(attachment)
  em$send()
  
  
}

microsoft_365r_notify_guia_confirmation <- function(outlook , recipient, attachment, vec_guias="" ){ 
  bl_body <- paste0("## Prezado(a) 

  Recebeu este email porque tem requisições com confirmação de entrega. Veja as Guias: ", vec_guias ,"
  no endereco https://ccs-sidmat.vercel.app/auth/login .

  Cpts,

  CCS Logistica")
  
  
  bl_em <- compose_email(
    body=md(bl_body),
    footer=md("sent via Microsoft365R")
  )
  em <- outlook$create_email(bl_em, subject="Confirmação de Entrega de Material", to=recipient)
  
  # add an attachment and send it
  em$add_attachment(attachment)
  em$send()
  
  
}


# Define a function to write logs to a text file
write_log <- function(log_message, log_file = "log.txt") {
  # Get the current timestamp
  timestamp <- format(Sys.time(), format = "%Y-%m-%d %H:%M:%S")
  
  
  # Append the log entry to the log file
  cat(log_message,"\n", file = log_file, append = TRUE)
}


removeOlderXlsxFiles <- function(folder_path){
  
  # List all XLSX files in the folder
  xlsx_files <- list.files(path = folder_path, pattern = "\\.xlsx$", full.names = TRUE)
  
  # Check if there are any XLSX files in the folder
  if (length(xlsx_files) > 0) {
    # Delete each XLSX file
    file.remove(xlsx_files)
    cat(paste("Deleted the following XLSX files:\n", xlsx_files, collapse = "\n"))
  } else {
    cat("No XLSX files found in the specified folder.")
  }
  
  
  
}



array_to_str <- function(vec){
  
  if(length(vec)==1){
    return(paste0("[", vec , "]"))
  } else if (length(vec)>1){
    
    str_tmp <- ''
    for (i in 1:length(vec)) {
      if(i==1){
        
        str_tmp <- paste0( "[", vec[i])
        
      } else if(i==length(vec) ){
        
        str_tmp <- paste0( str_tmp,","  ,vec[i],  "]" )
        
      } else {
        
        str_tmp <- paste0( str_tmp,",", vec[i] )
        
      }
      
    }
    return(str_tmp)
    
  }
}
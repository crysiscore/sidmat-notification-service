# compose an email with blastula
library(blastula)



send_by_microsoft_365r <- function(outlook , recipient, attachment){ 
  bl_body <- "## Prezado(a) 

  Recebeu este email porque tem material disponivel para sua área  no armazém
  Para o efeito visite o endereco http://localhost:3001 e efecture uma requisição.

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

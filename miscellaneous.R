# compose an email with blastula
# Install and load required packages
# install.packages("blastula")
# install.packages("kableExtra")
# install.packages("htmltools")
# install.packages("rmarkdown")
# install.packages("glue")
library(blastula)
library(kableExtra)
library(htmltools)
library(blastula)
library(glue)
library(rmarkdown)

microsoft_365r_notify_new_material <- function(outlook , recipient, attachment, cc_recipients = c('nunomoura@ccsaude.org.mz'), bcc_recipients = c('agnaldosamuel@ccsaude.org.mz')){ 
  bl_body <- "## Prezado(a) 

  Recebeu este email porque tem material disponivel para sua área  no armazém,
  Para fazer o plano de distribuição visite o endereco https://ccs-sidmat.vercel.app/auth/login.

  Cpts,

  CCS Logistica"
  
  
  bl_em <- compose_email(
    body=md(bl_body),
    footer=md("sent via Microsoft365R")
  )
  
  # Create email with main recipient and default CC/BCC
  em <- outlook$create_email(bl_em, subject="Material disponivel no Armazem", to=recipient, cc=cc_recipients, bcc=bcc_recipients)
  
  # add an attachment and send it
  em$add_attachment(attachment)
  em$send()
  
  
}


microsoft_365r_notify_resumo_semanal <- function(outlook , recipient, df.resumo, area.name, period, cc_recipients = c('nunomoura@ccsaude.org.mz'), bcc_recipients = c('agnaldosamuel@ccsaude.org.mz')){ 
  
  # Convert the data frame to an HTML table
    names(df.resumo)[1] <- "Area"
    names(df.resumo)[2] <-  " Total Requisicoes"
    names(df.resumo)[3] <-  " Pendentes"
    names(df.resumo)[4] <-  " Processadas"
    names(df.resumo)[5] <-  " Entregues"
    
  html_table <- kable(df.resumo , format = "html", escape = FALSE) %>%
    kable_styling(
      full_width = FALSE,
      bootstrap_options = c("striped", "hover", "condensed")) %>%
    column_spec(1, bold = TRUE)  # Apply bold formatting to the first column
  
  
  bl_body <- md (glue( "Prezado(a) ,

  Segue no anexo o resumo semanal do plano de distribuição de materiais para {area.name}: {period}  \\
 
  {html_table}

  Para mais informação visite https://ccs-sidmat.vercel.app/auth/login.

  Cpts,

  CCS Logistica" ))
  
  
  bl_em <- compose_email(
    body=md(bl_body),
    footer=md("Enviado através de Microsoft365R")
  )
  
  # Create email with main recipient and default CC/BCC
  em <- outlook$create_email(bl_em, subject="Resumo semanal do plano de distribuição de materiais", to=recipient, cc=cc_recipients, bcc=bcc_recipients)
  
  # add an attachment and send it
  # em$add_attachment(attachment)
  em$send()
  
  
}


microsoft_365r_notify_resumo_semanal_mensal <- function(outlook , recipient, df.resumo.mensal, df.resumo.semanal , area.name, period.semanal, period.mensal, cc_recipients = c('nunomoura@ccsaude.org.mz'), bcc_recipients = c('agnaldosamuel@ccsaude.org.mz')){ 
  
  # Convert the data frame to an HTML table
  names(df.resumo.mensal)[1] <- "Area"
  names(df.resumo.mensal)[2] <-  " Total Requisicoes"
  names(df.resumo.mensal)[3] <-  " Pendentes"
  names(df.resumo.mensal)[4] <-  " Processadas"
  names(df.resumo.mensal)[5] <-  " Entregues"
  
  names(df.resumo.semanal)[1] <- "Area"
  names(df.resumo.semanal)[2] <-  " Total Requisicoes"
  names(df.resumo.semanal)[3] <-  " Pendentes"
  names(df.resumo.semanal)[4] <-  " Processadas"
  names(df.resumo.semanal)[5] <-  " Entregues"
  
  html_table_mensal <- kable(df.resumo.mensal , format = "html", escape = FALSE) %>%
    kable_styling(
      full_width = FALSE,
      bootstrap_options = c("striped", "hover", "condensed")) %>%
    column_spec(1, bold = TRUE)  # Apply bold formatting to the first column
  
  html_table_semanal<- kable(df.resumo.semanal , format = "html", escape = FALSE) %>%
    kable_styling(
      full_width = FALSE,
      bootstrap_options = c("striped", "hover", "condensed")) %>%
    column_spec(1, bold = TRUE)  # Apply bold formatting to the first column
  
  
  bl_body <- md (glue( "Prezado(a) ,

  Resumo Mensal do plano de distribuição de materiais para {area.name}: {period.mensal}  \\
 
  {html_table_mensal}

  Resumo Semanal do plano de distribuição de materiais para {area.name}: {period.semanal}  \\
  
  {html_table_semanal}
   
  Para mais informação visite https://ccs-sidmat.vercel.app/auth/login.

  Cpts,

  CCS Logistica" ))
  
  
  bl_em <- compose_email(
    body=md(bl_body),
    footer=md("Enviado através de Microsoft365R")
  )
  
  # Create email with main recipient and default CC/BCC
  em <- outlook$create_email(bl_em, subject="Resumo semanal do plano de distribuição de materiais", to=recipient, cc=cc_recipients, bcc=bcc_recipients)
  
  # add an attachment and send it
  # em$add_attachment(attachment)
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
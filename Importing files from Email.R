#Installing RDCOMClient package
#devtools::install_github("omegahat/RDCOMClient")

# Load the RDCOMClient package
library(RDCOMClient)

# Connect to Outlook
OutApp <- COMCreate("Outlook.Application")#This allows R to interact with 
#Outlook as if it were any other COM object and allows to access different 
#features of Outlook like emails, calendar, contacts, etc.


ns <- OutApp$GetNamespace("MAPI")#retrieves the namespace of
#the Outlook application, this namespace allows the application to access 
#different functionality of Outlook such as email, calendar, contacts, etc.
#`MAPI` is the Messaging Application Programming Interface that is used to 
#access the Outlook data.

# Get the inbox folder
inbox <- ns$GetDefaultFolder(6)
#Choosing folder
inbox <- inbox$Folders.Item("Folder 1")
#the number 6 corresponds to the inbox folder, 
#the number 4 corresponds to the sent items folder

# List all folders
folders <- inbox$Folders
for(i in 1:folders$Count()) {
  print(folders.Item(i)$Name)
}

# Search for emails with the specific date in the inbox
search_date <- "09/08/2021"
emails <- inbox$Items()$Restrict("[ReceivedTime] >= '09/07/2022'
                                 and [ReceivedTime] <= '09/07/2022 23:59'")

emails$Count()

# Iterate through the emails and attachments
for (i in 1:emails$Count()) {
  email <- emails[[i]]
  Files=NULL
  # Check if the email has attachments
  if (email$Attachments()$Count() > 0) {
    for (j in 1:email$Attachments()$Count()) {
      attachment <- email$Attachments(j)
      
      # Check if the attachment is a csv file
      if (grepl(".xls", attachment$FileName())) {
        Files=rbind(Files,attachment$FileName())
      }
    }
  }
  #select the file you want to pull
  File=Files[2]
  
  for(k in 1:email$Attachments()$Count()){
    attachment <- email$Attachments(k)
    #Checking if file exist
    if (grepl(File, attachment$FileName())) {
      attachment_file <- tempfile()
      #Saving the attachment to current working directory as a temporary file
      attachment$SaveAsFile(attachment_file)
      #Read the csv file and assign it
      data <- readxl::read_excel(attachment_file,skip = 2)
    }
  }
}

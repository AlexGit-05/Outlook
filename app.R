#library("devtools")
#install_github('omegahat/RDCOMClient')
library(shiny)
library(RDCOMClient)

# Define UI for application that draws a histogram
ui <- fluidPage(theme = bslib::bs_theme(bootswatch = "darkly"),
                titlePanel("Select a csv File from Outlook Inbox"),
                tabsetPanel(
                  tabPanel("Select date and file",
                           numericInput("email","Select email address",
                                        value = 1),
                           textOutput("Email"),
                           textInput("folder","Select Folder",
                                     value = "Folder 1"),
                           textOutput("Folder"),
                           dateInput("date", "Select a date", 
                                     value = Sys.Date()),
                           numericInput("file","Select file",
                                        value = 1),
                           textOutput("File"),
                           downloadButton("download","Download",
                                          style="color: #fff; 
                                 background-color: green; 
                                 border-color: Black;")
                  ),
                  tabPanel("Files Available",
                           tableOutput("files")),
                  tabPanel("Imported Data",
                           DT::dataTableOutput("data"))
                )
)

# Define server logic required to draw a histogram
server <- function(input, output) {
  mail=reactive({
    #Installing RDCOMClient package
    #devtools::install_github("omegahat/RDCOMClient")
    
    # Load the RDCOMClient package
    library(RDCOMClient)
    
    # Connect to Outlook
    OutApp <- COMCreate("Outlook.Application")#This allows R to interact with 
    #Outlook as if it were any other COM object and allows to access different 
    #features of Outlook like emails, calendar, contacts, etc.
    
    
    outlookNameSpace <- OutApp$GetNamespace("MAPI")#retrieves the namespace of
    #the Outlook application, this namespace allows the application to access 
    #different functionality of Outlook such as email, calendar, contacts, etc.
    #`MAPI` is the Messaging Application Programming Interface that is used to 
    #access the Outlook data.
    #connect to an email address
    mailbox <- outlookNameSpace$Folders(input$email)
    mailbox
  })
  
  Folder=reactive({
    # Get the inbox folder
    inbox <- mail()$Folders("Inbox")
    #the number 6 corresponds to the inbox folder, 
    #the number 4 corresponds to the sent items folder
    #Choosing folder
    folder <- inbox$folders(input$folder)
    #File path for the folder
    #path=folder$folderpath()
    #path
    #Email extraction using date
    #search <- OutApp$AdvancedSearch(
    #  paste0("'", path, "'"), "")
  })  
  
  Data_file=reactive({
    
    Emails=Folder()$items()
    
    Date=input$date
    # Iterate through the emails and attachments
    for (i in 1:Emails$count()) {
      select_date=
        as.Date("1899-12-30") + floor(Emails[[i]]$ReceivedTime())==as.Date(Date)
      if (select_date==T) {
        email <- Emails[[i]]
        Files=NULL
        # Check if the email has attachments
        if (email$Attachments()$Count() > 0) {
          for (j in 1:email$Attachments()$Count()) {
            attachment <- email$Attachments(j)
            
            # Check if the attachment is a csv file
            if (grepl(".csv", attachment$FileName())) {
              Files=rbind(Files,attachment$FileName())
            }
          }
        }
      }
    }
    
    #email$attachments(1)$filename()
    
    #select the file you want to pull
    File=Files[input$file]
    
    for(k in 1:email$Attachments()$Count()){
      attachment <- email$Attachments(k)
      #Checking if file exist
      if (grepl(File, attachment$FileName())==T) {
        attachment_file <- tempfile()
        #Saving the attachment to current working directory as a temporary file
        attachment$SaveAsFile(attachment_file)
        #Read the csv file and assign it
        data <- readr::read_csv(attachment_file)
      }
    }
    
    out=list(File=File,Files=Files,Data=data)
    out
  })
    output$Email <- renderText({
      mail()$Name()
    })
    output$Folder <- renderText({
      Folder()$Name()
    })
    output$File <- renderText({
      Data_file()$File
    })
    #Attached csv files 
    output$files <- renderTable({
      Data_file()$Files
      })
    #Selected file
    output$data <- DT::renderDataTable({
      Data_file()$Data
    })
    output$download <- downloadHandler(
      filename = function(){
        paste(Data_file()$File)},
      content = function(file){
        write.csv(Data_file()$Data,file)
      },contentType = "csv"
    ) 
    
}

# Run the application 
shinyApp(ui = ui, server = server)

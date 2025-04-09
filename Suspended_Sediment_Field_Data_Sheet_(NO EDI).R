# Check and install required packages
required_packages <- c("shiny", "shinythemes", "DT", "openxlsx", "shinyjs")
for(pkg in required_packages) {
  if(!require(pkg, character.only = TRUE, quietly = TRUE)) {
    install.packages(pkg)
    library(pkg, character.only = TRUE)
  }
}

# Corrected sampler types list
sampler_types <- c(
  "DH-48",
  "DH-59",
  "DH-81",
  "DH-95",
  "D-49",
  "D-74",
  "D-74AL",
  "P-61",
  "P-72",
  "D-95",
  "D-96",
  "Other"
)

# Stage options
stage_options <- c(
  "Steady",
  "Rising",
  "Falling",
  "Peak"
)

# Sampling methods (removed EDI options)
sampling_methods <- c(
  "EWI Iso",
  "EWI Non-Iso",
  "Single Vertical",
  "Multi-Vertical",
  "Point",
  "Grab",
  "ISCO",
  "Box Single"
)

# Create a user guide HTML
user_guide_html <- HTML('
<div style="background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin-bottom: 20px; border-left: 5px solid #4682B4;">
  <h4 style="color: #4682B4;"><i class="fa fa-info-circle"></i> User Guide: Suspended Sediment Field Data Sheet</h4>

  <p><strong>Purpose:</strong> This application helps you document suspended sediment sampling activities in the field and generates properly formatted reports.</p>

  <div style="margin-left: 15px;">
    <p><strong>Important Notes:</strong></p>
    <ol>
      <li><strong>Channel Information:</strong> Enter the left and right edge measurements - the channel width will be calculated automatically.</li>
      <li><strong>Sampling Methods:</strong> Select one or more sampling methods. Each method will have its own panel for data entry.</li>
      <li><strong>The output excel file is editable.</strong></li>
      <li><strong>Actions:</strong>
        <ul>
          <li><span style="color: #337ab7;"><strong>Save Data</strong></span> - Saves the current form data to memory (visible in the Saved Data section)</li>
          <li><span style="color: #5cb85c;"><strong>Download Excel</strong></span> - Generates a formatted Excel file containing all entered information</li>
          <li><span style="color: #f0ad4e;"><strong>Clear Form</strong></span> - Resets all form fields to start over</li>
        </ul>
      </li>
    </ol>
  </div>
')

# UI definition
ui <- fluidPage(
  theme = shinythemes::shinytheme("flatly"),
  tags$head(
    tags$link(rel = "stylesheet", href = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css")
  ),
  
  titlePanel("Suspended Sediment Field Data Sheet"),
  
  # User Guide Section
  fluidRow(
    column(12,
           wellPanel(
             tags$a(
               id = "toggle_guide",
               class = "btn btn-info btn-sm pull-right",
               href = "#",
               style = "margin-top: -5px;",
               HTML('<i class="fa fa-question-circle"></i> Toggle User Guide')
             ),
             tags$div(id = "user_guide", user_guide_html)
           )
    )
  ),
  
  fluidRow(
    column(6,
           wellPanel(
             h4("Station Information"),
             textInput("station_number", "Station Number:", ""),
             dateInput("date", "Date:", Sys.Date()),
             textInput("station_name", "Station Name:", ""),
             textInput("party", "Party:", ""),
             numericInput("water_temp", "Water Temperature (°C):", value = NA, min = -10, max = 50, step = 0.1)
           )
    )
  ),
  
  fluidRow(
    column(12,
           wellPanel(
             h4("Channel Information"),
             fluidRow(
               column(4,
                      numericInput("left_edge", "Sediment Channel Left Edge of Water (ft):", value = 0, step = 0.1)
               ),
               column(4,
                      numericInput("right_edge", "Sediment Channel Right Edge of Water (ft):", value = 0, step = 0.1)
               ),
               column(4,
                      div(
                        style = "margin-top: 25px;",
                        textOutput("channel_width")
                      )
               )
             )
           )
    )
  ),
  
  fluidRow(
    column(12,
           wellPanel(
             h4("Sampling Methods"),
             checkboxGroupInput("selected_methods", "Select Sampling Methods:",
                                choices = sampling_methods,
                                selected = "EWI Iso"),
             hr(),
             
             # Container for method-specific panels (will be filled dynamically)
             uiOutput("method_panels")
           )
    )
  ),
  
  fluidRow(
    column(12,
           wellPanel(
             h4("Equipment Information"),
             selectInput("sampler_type", "Sampler Type:", sampler_types),
             conditionalPanel(
               condition = "input.sampler_type == 'Other'",
               textAreaInput("other_sampler", "Please explain other sampler:", rows = 3)
             )
           )
    )
  ),
  
  fluidRow(
    column(12,
           wellPanel(
             actionButton("save_btn", "Save Data", icon = icon("save"),
                          class = "btn-primary", width = "48%"),
             downloadButton("download_excel", "Download Excel",
                            class = "btn-success", style = "width:48%; float:right;"),
             br(), br(),
             actionButton("clear_btn", "Clear Form", icon = icon("eraser"),
                          class = "btn-warning", width = "100%")
           )
    )
  ),
  
  fluidRow(
    column(12,
           wellPanel(
             h4("Saved Data"),
             verbatimTextOutput("data_output")
           )
    )
  )
)

# Server logic
server <- function(input, output, session) {
  # Hide/show user guide
  observeEvent(input$toggle_guide, {
    shinyjs::toggle("user_guide")
  })
  
  # Create reactive values to store saved entries
  saved_data <- reactiveVal(data.frame())
  current_record <- reactiveVal(NULL)
  
  # Calculate channel width
  channel_width <- reactive({
    abs(input$right_edge - input$left_edge)
  })
  
  # Display channel width
  output$channel_width <- renderText({
    paste("Channel Width:", channel_width(), "ft")
  })
  
  # Dynamic method panels
  output$method_panels <- renderUI({
    req(input$selected_methods)
    
    # Add your sampling method panels here
    tagList()
  })
  
  # Display saved data
  output$data_output <- renderPrint({
    if(nrow(saved_data()) > 0) {
      saved_data()
    } else {
      "No data saved yet."
    }
  })
  
  # Clear form button action
  observeEvent(input$clear_btn, {
    # Reset all inputs (base values)
    updateTextInput(session, "station_number", value = "")
    updateDateInput(session, "date", value = Sys.Date())
    updateTextInput(session, "station_name", value = "")
    updateTextInput(session, "party", value = "")
    updateNumericInput(session, "water_temp", value = NA)
    updateNumericInput(session, "left_edge", value = 0)
    updateNumericInput(session, "right_edge", value = 0)
    updateCheckboxGroupInput(session, "selected_methods", selected = "EWI Iso")
    updateSelectInput(session, "sampler_type", selected = sampler_types[1])
    updateTextAreaInput(session, "other_sampler", value = "")
  })
}

# Run the application
shinyApp(ui = ui, server = server)

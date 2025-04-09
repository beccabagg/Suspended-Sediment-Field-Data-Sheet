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

# Sampling methods
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

# User Guide
user_guide_html <- HTML('
<div style="background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin-bottom: 20px; border-left: 5px solid #4682B4;">
  <h4 style="color: #4682B4;"><i class="fa fa-info-circle"></i> User Guide: Suspended Sediment Field Data Sheet</h4>
  <p><strong>Purpose:</strong> This application helps you document suspended sediment sampling activities in the field and generates properly formatted reports.</p>
  <p><strong>How to Use:</strong></p>
  <ul>
    <li>Enter station details, measurement type, and stage conditions.</li>
    <li>Select and configure sampling methods (e.g., EWI, Grab).</li>
    <li>Generate a formatted Excel report with average times and sampler types.</li>
  </ul>
</div>
')

# UI definition
ui <- fluidPage(
  theme = shinythemes::shinytheme("flatly"),
  tags$head(tags$link(rel = "stylesheet", href = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css")),
  titlePanel("Suspended Sediment Field Data Sheet"),
  
  # User Guide Section
  fluidRow(
    column(12, wellPanel(
      tags$a(
        id = "toggle_guide",
        class = "btn btn-info btn-sm pull-right",
        href = "#",
        style = "margin-top: -5px;",
        HTML('<i class="fa fa-question-circle"></i> Toggle User Guide')
      ),
      tags$div(id = "user_guide", user_guide_html)
    ))
  ),
  
  # Station Information
  fluidRow(
    column(6, wellPanel(
      h4("Station Information"),
      textInput("station_number", "Station Number:", ""),
      dateInput("date", "Date:", Sys.Date()),
      textInput("station_name", "Station Name:", ""),
      textInput("party", "Party:", ""),
      numericInput("water_temp", "Water Temperature (°C):", value = NA, min = -10, max = 50, step = 0.1)
    ))
  ),
  
  # Channel Information
  fluidRow(
    column(12, wellPanel(
      h4("Channel Information"),
      fluidRow(
        column(4, numericInput("left_edge", "Sediment Channel Left Edge of Water (ft):", value = 0, step = 0.1)),
        column(4, numericInput("right_edge", "Sediment Channel Right Edge of Water (ft):", value = 0, step = 0.1)),
        column(4, div(style = "margin-top: 25px;", textOutput("channel_width")))
      )
    ))
  ),
  
  # Sampling Methods
  fluidRow(
    column(12, wellPanel(
      h4("Sampling Methods"),
      checkboxGroupInput("selected_methods", "Select Sampling Methods:", choices = sampling_methods, selected = "EWI Iso"),
      hr(),
      uiOutput("method_panels")
    ))
  ),
  
  # Equipment Information
  fluidRow(
    column(12, wellPanel(
      h4("Equipment Information"),
      selectInput("sampler_type", "Sampler Type:", sampler_types),
      conditionalPanel(
        condition = "input.sampler_type == 'Other'",
        textAreaInput("other_sampler", "Please explain other sampler:", rows = 3)
      )
    ))
  ),
  
  # Actions: Save, Download, Clear
  fluidRow(
    column(12, wellPanel(
      actionButton("save_btn", "Save Data", icon = icon("save"), class = "btn-primary", width = "48%"),
      downloadButton("download_excel", "Download Excel", class = "btn-success", style = "width:48%; float:right;"),
      br(), br(),
      actionButton("clear_btn", "Clear Form", icon = icon("eraser"), class = "btn-warning", width = "100%")
    ))
  ),
  
  # Saved Data Section
  fluidRow(
    column(12, wellPanel(
      h4("Saved Data"),
      verbatimTextOutput("data_output")
    ))
  )
)

# Server logic
server <- function(input, output, session) {
  # Hide/show user guide
  observeEvent(input$toggle_guide, {
    shinyjs::toggle("user_guide")
  })
  
  # Reactive values for saved data
  saved_data <- reactiveVal(data.frame())
  current_record <- reactiveVal(NULL)
  
  # Channel width calculation
  channel_width <- reactive({
    abs(input$right_edge - input$left_edge)
  })
  output$channel_width <- renderText({
    paste("Channel Width:", channel_width(), "ft")
  })
  
  # Dynamic UI for sampling methods
  output$method_panels <- renderUI({
    req(input$selected_methods)
    
    lapply(input$selected_methods, function(method) {
      method_id <- gsub(" ", "_", tolower(method))
      
      if (method == "Grab") {
        tagList(
          h4("Grab Sampling"),
          fluidRow(
            column(4, textInput(paste0(method_id, "_start_time_a"), "Start Time (A):")),
            column(4, textInput(paste0(method_id, "_end_time_a"), "End Time (A):")),
            column(4, selectInput(paste0(method_id, "_sampler_a"), "Sampler Type (A):", sampler_types))
          ),
          fluidRow(
            column(4, textInput(paste0(method_id, "_start_time_b"), "Start Time (B):")),
            column(4, textInput(paste0(method_id, "_end_time_b"), "End Time (B):")),
            column(4, selectInput(paste0(method_id, "_sampler_b"), "Sampler Type (B):", sampler_types))
          )
        )
      } else if (method %in% c("EWI Iso", "EWI Non-Iso")) {
        tagList(
          h4(paste(method, "Sampling")),
          lapply(c("A", "B", "C", "D"), function(set) {
            fluidRow(
              column(3, numericInput(paste0(method_id, "_num_verticals_", tolower(set)), paste("Number of Verticals (", set, "):", sep = ""), value = 10, min = 1, max = 40)),
              column(3, textInput(paste0(method_id, "_start_time_", tolower(set)), paste("Start Time (", set, "):", sep = ""))),
              column(3, textInput(paste0(method_id, "_end_time_", tolower(set)), paste("End Time (", set, "):", sep = ""))),
              column(3, selectInput(paste0(method_id, "_sampler_", tolower(set)), paste("Sampler Type (", set, "):", sep = ""), sampler_types))
            )
          })
        )
      }
    })
  })
  
  # Save button action
  observeEvent(input$save_btn, {
    # Save logic
  })
  
  # Generate Excel file
  output$download_excel <- downloadHandler(
    filename = function() {
      paste0("Suspended_Sediment_Data_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      # Excel generation logic
    }
  )
  
  # Clear form
  observeEvent(input$clear_btn, {
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

# Run the Shiny app
shinyApp(ui = ui, server = server)

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

# Measurement types - Change order to make Wading the first (default) option
measurement_types <- c(
  "Wading",
  "Bridge Upstream",
  "Bridge Downstream",
  "Cableway",
  "Boat",
  "Ice"
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
    <p><strong>How to Use This App:</strong></p>
    <ol>
      <li><strong>Station Information:</strong> Enter basic station details including number, name, date, and water temperature.</li>
      <li><strong>Measurement Information:</strong> Select the measurement type (wading, bridge, etc.) and document stage conditions.</li>
      <li><strong>Channel Information:</strong> Enter the left and right edge measurements - the channel width will be calculated automatically.</li>
      <li><strong>Sampling Methods:</strong> Select one or more sampling methods. Each method will have its own panel for data entry.</li>
      <li><strong>EWI Sampling:</strong> For EWI sampling, you can enter data up to 3 sets (A, B, C) at the same section widths.</li>
      <li><strong>Times:</strong> Record start and end times for each sampling method or set. Average time will be calculated automatically.</li>
      <li><strong>Equipment Information:</strong> Select the sampler type used during collection.</li>
      <li><strong>Actions:</strong>
        <ul>
          <li><span style="color: #337ab7;"><strong>Save Data</strong></span> - Saves the current form data to memory (visible in the Saved Data section)</li>
          <li><span style="color: #5cb85c;"><strong>Download Excel</strong></span> - Generates a formatted Excel file containing all entered information</li>
          <li><span style="color: #f0ad4e;"><strong>Clear Form</strong></span> - Resets all form fields to start over</li>
        </ul>
      </li>
    </ol>
  </div>

  <p><strong>Sampling Method Details:</strong></p>
  <ul style="font-size: 0.9em;">
    <li><strong>EWI Iso</strong> - Equal Width Increment with isokinetic sampling - divides channel into equal widths</li>
    <li><strong>EWI Non-Iso</strong> - Equal Width Increment with non-isokinetic equipment</li>
    <li><strong>Single Vertical</strong> - Sample collected at one representative location</li>
    <li><strong>Multi-Vertical</strong> - Samples at multiple selected points across the channel</li>
    <li><strong>Point</strong> - Sample at specific fixed point</li>
    <li><strong>Grab</strong> - Simple grab sample from surface or specific depth</li>
    <li><strong>ISCO</strong> - Automated sampling</li>
    <li><strong>Box Single</strong> - Box sampler at a single location</li>
  </ul>

  <p><small>Note: Required fields will be highlighted if left blank. For best results, complete all relevant sections before generating an Excel report.</small></p>
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
    ),

    column(6,
           wellPanel(
             h4("Measurement Information"),
             selectInput("measurement_type", "Measurement Type:", measurement_types),
             radioButtons("stage", "Stage:", stage_options),
             fluidRow(
               column(6,
                      numericInput("location_value", "Location (feet):", value = 0, step = 1)
               ),
               column(6,
                      radioButtons("location_direction", "Direction:", c("Upstream", "Downstream"))
               )
             )
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

    method_panels <- lapply(input$selected_methods, function(method) {
      method_id <- gsub(" ", "_", tolower(method))

# In the method_panels renderUI function, update the EWI methods section:
      if(method %in% c("EWI Iso", "EWI Non-Iso")) {
        # EWI methods with A, B, and C sets
        tagList(
          div(class = "panel panel-default",
              div(class = "panel-heading", h4(method)),
              div(class = "panel-body",
                  tabsetPanel(id = paste0(method_id, "_tabs"),
                              # Set A tab
                              tabPanel("Set A",
                                       fluidRow(
                                         column(4, numericInput(paste0(method_id, "_num_verticals_a"),
                                                                "Number of Verticals:", value = 10, min = 1, max = 40, step = 1)),
                                         column(4, textInput(paste0(method_id, "_start_time_a"), "Start Time:", value = "")),
                                         column(4, textInput(paste0(method_id, "_end_time_a"), "End Time:", value = ""))
                                       ),
                                       textOutput(paste0(method_id, "_avg_time_a")),
                                       hr(),
                                       h4("Sampling Sections"),
                                       DTOutput(paste0(method_id, "_section_table_a"))
                              ),
                              # Set B tab
                              tabPanel("Set B",
                                       fluidRow(
                                         column(4, numericInput(paste0(method_id, "_num_verticals_b"),
                                                                "Number of Verticals:", value = 10, min = 1, max = 40, step = 1)),
                                         column(4, textInput(paste0(method_id, "_start_time_b"), "Start Time:", value = "")),
                                         column(4, textInput(paste0(method_id, "_end_time_b"), "End Time:", value = ""))
                                       ),
                                       textOutput(paste0(method_id, "_avg_time_b")),
                                       hr(),
                                       h4("Sampling Sections"),
                                       DTOutput(paste0(method_id, "_section_table_b"))
                              ),
                              # Set C tab (new)
                              tabPanel("Set C",
                                       fluidRow(
                                         column(4, numericInput(paste0(method_id, "_num_verticals_c"),
                                                                "Number of Verticals:", value = 10, min = 1, max = 40, step = 1)),
                                         column(4, textInput(paste0(method_id, "_start_time_c"), "Start Time:", value = "")),
                                         column(4, textInput(paste0(method_id, "_end_time_c"), "End Time:", value = ""))
                                       ),
                                       textOutput(paste0(method_id, "_avg_time_c")),
                                       hr(),
                                       h4("Sampling Sections"),
                                       DTOutput(paste0(method_id, "_section_table_c"))
                              )
                  )
              )
          )
        )

      } else if(method == "Multi-Vertical") {
        # Multi-Vertical method
        tagList(
          div(class = "panel panel-default",
              div(class = "panel-heading", h4(method)),
              div(class = "panel-body",
                  fluidRow(
                    column(4, numericInput(paste0(method_id, "_num_verticals"),
                                           "Number of Verticals:", value = 3, min = 1, max = 40, step = 1)),
                    column(4, textInput(paste0(method_id, "_start_time"), "Start Time:", value = "")),
                    column(4, textInput(paste0(method_id, "_end_time"), "End Time:", value = ""))
                  ),
                  textOutput(paste0(method_id, "_avg_time")),
                  hr(),
                  h4("Sampling Sections"),
                  DTOutput(paste0(method_id, "_section_table"))
              )
          )
        )
      } else if(method %in% c("Single Vertical", "Point")) {
        # Single location methods
        tagList(
          div(class = "panel panel-default",
              div(class = "panel-heading", h4(method)),
              div(class = "panel-body",
                  fluidRow(
                    column(3, textInput(paste0(method_id, "_start_time"), "Start Time:", value = "")),
                    column(3, textInput(paste0(method_id, "_end_time"), "End Time:", value = "")),
                    column(3, numericInput(paste0(method_id, "_location"), "Location (ft from LEW):",
                                           value = 0, step = 0.1)),
                    column(3, numericInput(paste0(method_id, "_depth"), "Depth (ft):",
                                           value = 0, step = 0.1))
                  ),
                  textOutput(paste0(method_id, "_avg_time"))
              )
          )
        )
      } else {
        # Other methods (ISCO, Grab, Box Single)
        tagList(
          div(class = "panel panel-default",
              div(class = "panel-heading", h4(method)),
              div(class = "panel-body",
                  fluidRow(
                    column(4, textInput(paste0(method_id, "_start_time"), "Start Time:", value = "")),
                    column(4, textInput(paste0(method_id, "_end_time"), "End Time:", value = "")),
                    column(4, textAreaInput(paste0(method_id, "_notes"), "Notes:", rows = 2))
                  ),
                  textOutput(paste0(method_id, "_avg_time"))
              )
          )
        )
      }
    })

    # Return all panels in a tagList
    do.call(tagList, method_panels)
  })

  # Calculate sampling sections for EWI methods (Set A)
  observe({
    req(input$selected_methods)

    for(method in input$selected_methods) {
      method_id <- gsub(" ", "_", tolower(method))

      if(method %in% c("EWI Iso", "EWI Non-Iso")) {
        # For Set A
        local({
          local_method <- method
          local_method_id <- method_id

          # Create reactive for this specific method's num_verticals
          num_verticals_id <- paste0(local_method_id, "_num_verticals_a")

          # Update section table when num verticals or edges change
          observeEvent(c(input[[num_verticals_id]], input$left_edge, input$right_edge), {
            req(input[[num_verticals_id]], input$left_edge, input$right_edge)

            # Calculate sections
            width <- channel_width()
            num_sections <- input[[num_verticals_id]]
            section_width <- width / num_sections

            data <- data.frame(
              Section = integer(),
              Distance_from_LEW = numeric(),
              Section_Width = numeric(),
              stringsAsFactors = FALSE
            )

            for(i in 1:num_sections) {
              section_distance <- input$left_edge + (i - 0.5) * section_width
              data <- rbind(data, data.frame(
                Section = i,
                Distance_from_LEW = round(section_distance, 2),
                Section_Width = round(section_width, 2)
              ))
            }

            # Render the table for this specific method
            output_id <- paste0(local_method_id, "_section_table_a")
            output[[output_id]] <- renderDT({
              DT::datatable(
                data,
                options = list(
                  pageLength = 10,
                  searching = FALSE,
                  lengthChange = FALSE
                ),
                rownames = FALSE,
                colnames = c("Section", "Distance from LEW (ft)", "Section Width (ft)")
              )
            })
          })

          # Calculate average time
          observeEvent(c(input[[paste0(local_method_id, "_start_time_a")]], input[[paste0(local_method_id, "_end_time_a")]]), {
            start_time <- input[[paste0(local_method_id, "_start_time_a")]]
            end_time <- input[[paste0(local_method_id, "_end_time_a")]]

            # Only calculate if both times are provided
            if(start_time != "" && end_time != "") {
              # Try to parse times
              tryCatch({
                start <- as.POSIXct(start_time, format = "%H:%M")
                end <- as.POSIXct(end_time, format = "%H:%M")

                # Calculate average time
                avg_time <- start + (difftime(end, start) / 2)
                avg_time_str <- format(avg_time, "%H:%M")

                # Update average time text
                output[[paste0(local_method_id, "_avg_time_a")]] <- renderText({
                  paste("Average Time:", avg_time_str)
                })
              }, error = function(e) {
                output[[paste0(local_method_id, "_avg_time_a")]] <- renderText({
                  "Average Time: (Invalid time format. Use HH:MM)"
                })
              })
            } else {
              output[[paste0(local_method_id, "_avg_time_a")]] <- renderText({
                "Average Time: (Enter both start and end times)"
              })
            }
          })
        })

        # For Set B (copy of Set A sections)
        local({
          local_method <- method
          local_method_id <- method_id

          # Create reactive for this specific method's num_verticals
          num_verticals_id_a <- paste0(local_method_id, "_num_verticals_a")
          num_verticals_id_b <- paste0(local_method_id, "_num_verticals_b")

          # Update num_verticals_b when num_verticals_a changes
          observeEvent(input[[num_verticals_id_a]], {
            updateNumericInput(session, num_verticals_id_b, value = input[[num_verticals_id_a]])
          })

          # Update section table when num verticals or edges change
          observeEvent(c(input[[num_verticals_id_b]], input$left_edge, input$right_edge), {
            req(input[[num_verticals_id_b]], input$left_edge, input$right_edge)

            # Calculate sections
            width <- channel_width()
            num_sections <- input[[num_verticals_id_b]]
            section_width <- width / num_sections

            data <- data.frame(
              Section = integer(),
              Distance_from_LEW = numeric(),
              Section_Width = numeric(),
              stringsAsFactors = FALSE
            )

            for(i in 1:num_sections) {
              section_distance <- input$left_edge + (i - 0.5) * section_width
              data <- rbind(data, data.frame(
                Section = i,
                Distance_from_LEW = round(section_distance, 2),
                Section_Width = round(section_width, 2)
              ))
            }

            # Render the table for this specific method
            output_id <- paste0(local_method_id, "_section_table_b")
            output[[output_id]] <- renderDT({
              DT::datatable(
                data,
                options = list(
                  pageLength = 10,
                  searching = FALSE,
                  lengthChange = FALSE
                ),
                rownames = FALSE,
                colnames = c("Section", "Distance from LEW (ft)", "Section Width (ft)")
              )
            })
          })

          # Calculate average time
          observeEvent(c(input[[paste0(local_method_id, "_start_time_b")]], input[[paste0(local_method_id, "_end_time_b")]]), {
            start_time <- input[[paste0(local_method_id, "_start_time_b")]]
            end_time <- input[[paste0(local_method_id, "_end_time_b")]]

            # Only calculate if both times are provided
            if(start_time != "" && end_time != "") {
              # Try to parse times
              tryCatch({
                start <- as.POSIXct(start_time, format = "%H:%M")
                end <- as.POSIXct(end_time, format = "%H:%M")

                # Calculate average time
                avg_time <- start + (difftime(end, start) / 2)
                avg_time_str <- format(avg_time, "%H:%M")

                # Update average time text
                output[[paste0(local_method_id, "_avg_time_b")]] <- renderText({
                  paste("Average Time:", avg_time_str)
                })
              }, error = function(e) {
                output[[paste0(local_method_id, "_avg_time_b")]] <- renderText({
                  "Average Time: (Invalid time format. Use HH:MM)"
                })
              })
            } else {
              output[[paste0(local_method_id, "_avg_time_b")]] <- renderText({
                "Average Time: (Enter both start and end times)"
              })
            }
          })
        })
        # For Set C (new)
        local({
          local_method <- method
          local_method_id <- method_id

          # Create reactive for this specific method's num_verticals
          num_verticals_id_a <- paste0(local_method_id, "_num_verticals_a")
          num_verticals_id_c <- paste0(local_method_id, "_num_verticals_c")

          # Update num_verticals_c when num_verticals_a changes
          observeEvent(input[[num_verticals_id_a]], {
            updateNumericInput(session, num_verticals_id_c, value = input[[num_verticals_id_a]])
          })

          # Update section table when num verticals or edges change
          observeEvent(c(input[[num_verticals_id_c]], input$left_edge, input$right_edge), {
            req(input[[num_verticals_id_c]], input$left_edge, input$right_edge)

            # Calculate sections
            width <- channel_width()
            num_sections <- input[[num_verticals_id_c]]
            section_width <- width / num_sections

            data <- data.frame(
              Section = integer(),
              Distance_from_LEW = numeric(),
              Section_Width = numeric(),
              stringsAsFactors = FALSE
            )

            for(i in 1:num_sections) {
              section_distance <- input$left_edge + (i - 0.5) * section_width
              data <- rbind(data, data.frame(
                Section = i,
                Distance_from_LEW = round(section_distance, 2),
                Section_Width = round(section_width, 2)
              ))
            }

            # Render the table for this specific method
            output_id <- paste0(local_method_id, "_section_table_c")
            output[[output_id]] <- renderDT({
              DT::datatable(
                data,
                options = list(
                  pageLength = 10,
                  searching = FALSE,
                  lengthChange = FALSE
                ),
                rownames = FALSE,
                colnames = c("Section", "Distance from LEW (ft)", "Section Width (ft)")
              )
            })
          })

          # Calculate average time
          observeEvent(c(input[[paste0(local_method_id, "_start_time_c")]], input[[paste0(local_method_id, "_end_time_c")]]), {
            start_time <- input[[paste0(local_method_id, "_start_time_c")]]
            end_time <- input[[paste0(local_method_id, "_end_time_c")]]

            # Only calculate if both times are provided
            if(start_time != "" && end_time != "") {
              # Try to parse times
              tryCatch({
                start <- as.POSIXct(start_time, format = "%H:%M")
                end <- as.POSIXct(end_time, format = "%H:%M")

                # Calculate average time
                avg_time <- start + (difftime(end, start) / 2)
                avg_time_str <- format(avg_time, "%H:%M")

                # Update average time text
                output[[paste0(local_method_id, "_avg_time_c")]] <- renderText({
                  paste("Average Time:", avg_time_str)
                })
              }, error = function(e) {
                output[[paste0(local_method_id, "_avg_time_c")]] <- renderText({
                  "Average Time: (Invalid time format. Use HH:MM)"
                })
              })
            } else {
              output[[paste0(local_method_id, "_avg_time_c")]] <- renderText({
                "Average Time: (Enter both start and end times)"
              })
            }
          })
        })
      } else if(method == "Multi-Vertical") {
        # For Multi-Vertical method
        local({
          local_method <- method
          local_method_id <- method_id

          # Create reactive for this specific method's num_verticals
          num_verticals_id <- paste0(local_method_id, "_num_verticals")

          # Update section table when num verticals or edges change
          observeEvent(c(input[[num_verticals_id]], input$left_edge, input$right_edge), {
            req(input[[num_verticals_id]], input$left_edge, input$right_edge)

            # Calculate sections
            width <- channel_width()
            num_sections <- input[[num_verticals_id]]
            section_width <- width / num_sections

            data <- data.frame(
              Section = integer(),
              Distance_from_LEW = numeric(),
              Section_Width = numeric(),
              stringsAsFactors = FALSE
            )

            for(i in 1:num_sections) {
              section_distance <- input$left_edge + i * section_width
              data <- rbind(data, data.frame(
                Section = i,
                Distance_from_LEW = round(section_distance, 2),
                Section_Width = NA  # User would define this
              ))
            }

            # Render the table for this specific method
            output_id <- paste0(local_method_id, "_section_table")
            output[[output_id]] <- renderDT({
              DT::datatable(
                data,
                options = list(
                  pageLength = 10,
                  searching = FALSE,
                  lengthChange = FALSE
                ),
                rownames = FALSE,
                colnames = c("Section", "Distance from LEW (ft)", "Section Width (ft)")
              )
            })
          })

          # Calculate average time
          observeEvent(c(input[[paste0(local_method_id, "_start_time")]], input[[paste0(local_method_id, "_end_time")]]), {
            start_time <- input[[paste0(local_method_id, "_start_time")]]
            end_time <- input[[paste0(local_method_id, "_end_time")]]

            # Only calculate if both times are provided
            if(start_time != "" && end_time != "") {
              # Try to parse times
              tryCatch({
                start <- as.POSIXct(start_time, format = "%H:%M")
                end <- as.POSIXct(end_time, format = "%H:%M")

                # Calculate average time
                avg_time <- start + (difftime(end, start) / 2)
                avg_time_str <- format(avg_time, "%H:%M")

                # Update average time text
                output[[paste0(local_method_id, "_avg_time")]] <- renderText({
                  paste("Average Time:", avg_time_str)
                })
              }, error = function(e) {
                output[[paste0(local_method_id, "_avg_time")]] <- renderText({
                  "Average Time: (Invalid time format. Use HH:MM)"
                })
              })
            } else {
              output[[paste0(local_method_id, "_avg_time")]] <- renderText({
                "Average Time: (Enter both start and end times)"
              })
            }
          })
        })
      } else if(method %in% c("Single Vertical", "Point", "Grab", "ISCO", "Box Single")) {
        # For single point methods
        local({
          local_method <- method
          local_method_id <- method_id

          # Calculate average time
          observeEvent(c(input[[paste0(local_method_id, "_start_time")]], input[[paste0(local_method_id, "_end_time")]]), {
            start_time <- input[[paste0(local_method_id, "_start_time")]]
            end_time <- input[[paste0(local_method_id, "_end_time")]]

            # Only calculate if both times are provided
            if(start_time != "" && end_time != "") {
              # Try to parse times
              tryCatch({
                start <- as.POSIXct(start_time, format = "%H:%M")
                end <- as.POSIXct(end_time, format = "%H:%M")

                # Calculate average time
                avg_time <- start + (difftime(end, start) / 2)
                avg_time_str <- format(avg_time, "%H:%M")

                # Update average time text
                output[[paste0(local_method_id, "_avg_time")]] <- renderText({
                  paste("Average Time:", avg_time_str)
                })
              }, error = function(e) {
                output[[paste0(local_method_id, "_avg_time")]] <- renderText({
                  "Average Time: (Invalid time format. Use HH:MM)"
                })
              })
            } else {
              output[[paste0(local_method_id, "_avg_time")]] <- renderText({
                "Average Time: (Enter both start and end times)"
              })
            }
          })
        })
      }
    }
  })

  # Display saved data
  output$data_output <- renderPrint({
    if(nrow(saved_data()) > 0) {
      saved_data()
    } else {
      "No data saved yet."
    }
  })

  # Gather all method data
  gather_method_data <- function() {
    req(input$selected_methods)

    method_data_list <- list()

    for(method in input$selected_methods) {
      method_id <- gsub(" ", "_", tolower(method))

# In the gather_method_data function, update the EWI methods section:
      if(method %in% c("EWI Iso", "EWI Non-Iso")) {
        # EWI methods with A, B, and C sets
        section_data_a <- NULL
        section_data_b <- NULL
        section_data_c <- NULL  # Add this for Set C

        # Get section data if available
        if(!is.null(input[[paste0(method_id, "_num_verticals_a")]])) {
          # Calculate sections for set A
          width <- channel_width()
          num_sections <- input[[paste0(method_id, "_num_verticals_a")]]
          section_width <- width / num_sections

          section_data_a <- data.frame(
            Section = integer(),
            Distance_from_LEW = numeric(),
            Section_Width = numeric(),
            stringsAsFactors = FALSE
          )

          for(i in 1:num_sections) {
            section_distance <- input$left_edge + (i - 0.5) * section_width
            section_data_a <- rbind(section_data_a, data.frame(
              Section = i,
              Distance_from_LEW = round(section_distance, 2),
              Section_Width = round(section_width, 2)
            ))
          }
        }

        if(!is.null(input[[paste0(method_id, "_num_verticals_b")]])) {
          # Calculate sections for set B
          width <- channel_width()
          num_sections <- input[[paste0(method_id, "_num_verticals_b")]]
          section_width <- width / num_sections

          section_data_b <- data.frame(
            Section = integer(),
            Distance_from_LEW = numeric(),
            Section_Width = numeric(),
            stringsAsFactors = FALSE
          )

          for(i in 1:num_sections) {
            section_distance <- input$left_edge + (i - 0.5) * section_width
            section_data_b <- rbind(section_data_b, data.frame(
              Section = i,
              Distance_from_LEW = round(section_distance, 2),
              Section_Width = round(section_width, 2)
            ))
          }
        }

        # Add section data for Set C
        if(!is.null(input[[paste0(method_id, "_num_verticals_c")]])) {
          # Calculate sections for set C
          width <- channel_width()
          num_sections <- input[[paste0(method_id, "_num_verticals_c")]]
          section_width <- width / num_sections

          section_data_c <- data.frame(
            Section = integer(),
            Distance_from_LEW = numeric(),
            Section_Width = numeric(),
            stringsAsFactors = FALSE
          )

          for(i in 1:num_sections) {
            section_distance <- input$left_edge + (i - 0.5) * section_width
            section_data_c <- rbind(section_data_c, data.frame(
              Section = i,
              Distance_from_LEW = round(section_distance, 2),
              Section_Width = round(section_width, 2)
            ))
          }
        }

        # Add data for this method
        method_data_list[[method]] <- list(
          Type = method,
          SetA = list(
            NumVerticals = input[[paste0(method_id, "_num_verticals_a")]],
            StartTime = input[[paste0(method_id, "_start_time_a")]],
            EndTime = input[[paste0(method_id, "_end_time_a")]],
            Sections = section_data_a
          ),
          SetB = list(
            NumVerticals = input[[paste0(method_id, "_num_verticals_b")]],
            StartTime = input[[paste0(method_id, "_start_time_b")]],
            EndTime = input[[paste0(method_id, "_end_time_b")]],
            Sections = section_data_b
          ),
          SetC = list(
            NumVerticals = input[[paste0(method_id, "_num_verticals_c")]],
            StartTime = input[[paste0(method_id, "_start_time_c")]],
            EndTime = input[[paste0(method_id, "_end_time_c")]],
            Sections = section_data_c
          )
        )
      } else if(method == "Multi-Vertical") {
        # Multi-Vertical method
        section_data <- NULL

        # Get section data if available
        if(!is.null(input[[paste0(method_id, "_num_verticals")]])) {
          # Calculate sections
          width <- channel_width()
          num_sections <- input[[paste0(method_id, "_num_verticals")]]
          section_width <- width / num_sections

          section_data <- data.frame(
            Section = integer(),
            Distance_from_LEW = numeric(),
            Section_Width = numeric(),
            stringsAsFactors = FALSE
          )

          for(i in 1:num_sections) {
            section_distance <- input$left_edge + i * section_width
            section_data <- rbind(section_data, data.frame(
              Section = i,
              Distance_from_LEW = round(section_distance, 2),
              Section_Width = NA
            ))
          }
        }

        # Add data for this method
        method_data_list[[method]] <- list(
          Type = method,
          NumVerticals = input[[paste0(method_id, "_num_verticals")]],
          StartTime = input[[paste0(method_id, "_start_time")]],
          EndTime = input[[paste0(method_id, "_end_time")]],
          Sections = section_data
        )
      } else if(method %in% c("Single Vertical", "Point")) {
        # Single location methods
        method_data_list[[method]] <- list(
          Type = method,
          StartTime = input[[paste0(method_id, "_start_time")]],
          EndTime = input[[paste0(method_id, "_end_time")]],
          Location = input[[paste0(method_id, "_location")]],
          Depth = input[[paste0(method_id, "_depth")]]
        )
      } else {
        # Other methods (ISCO, Grab, Box Single)
        method_data_list[[method]] <- list(
          Type = method,
          StartTime = input[[paste0(method_id, "_start_time")]],
          EndTime = input[[paste0(method_id, "_end_time")]],
          Notes = input[[paste0(method_id, "_notes")]]
        )
      }
    }

    return(method_data_list)
  }

  # Create record for saving or download
  create_record <- function() {
    # Get sampler type (include explanation if "Other" is selected)
    sampler_info <- input$sampler_type
    if(input$sampler_type == "Other" && !is.null(input$other_sampler) && input$other_sampler != "") {
      sampler_info <- paste0("Other: ", input$other_sampler)
    }

    # Get all method data
    methods_data <- gather_method_data()

    # Create a new data record
    record <- list(
      Station_Number = input$station_number,
      Date = format(input$date, "%Y-%m-%d"),
      Station_Name = input$station_name,
      Party = input$party,
      Water_Temperature = input$water_temp,
      Measurement_Type = input$measurement_type,
      Stage = input$stage,
      Location = paste0(input$location_value, " feet ", tolower(input$location_direction), " of gage"),
      Left_Edge = input$left_edge,
      Right_Edge = input$right_edge,
      Channel_Width = channel_width(),
      Sampling_Methods = methods_data,
      Sampler_Type = sampler_info,
      Timestamp = format(Sys.time(), "%Y-%m-%d %H:%M:%S")
    )

    return(record)
  }

  # Format filename with station number and date
  format_filename <- function() {
    # Get station name (use "Unknown" if empty)
    station_name <- ifelse(is.null(input$station_name) || input$station_name == "",
                          "Unknown",
                          gsub("[^a-zA-Z0-9]", "", input$station_name)) # Remove special characters

    # Get date in YYYYMMDD format
    date_str <- format(input$date, "%Y%m%d")

    # Create filename
    filename <- paste0("SS_", station_name, "_", date_str, ".xlsx")

    return(filename)
  }

  # Save button action
  observeEvent(input$save_btn, {
    # Create record
    record <- create_record()
    current_record(record)

    # Create simplified summary for display
    methods_summary <- paste(input$selected_methods, collapse = ", ")

    # Convert to data frame for saved_data
    record_df <- data.frame(
      Station_Number = record$Station_Number,
      Date = record$Date,
      Station_Name = record$Station_Name,
      Party = record$Party,
      Water_Temperature = record$Water_Temperature,
      Measurement_Type = record$Measurement_Type,
      Stage = record$Stage,
      Location = record$Location,
      Left_Edge = record$Left_Edge,
      Right_Edge = record$Right_Edge,
      Channel_Width = record$Channel_Width,
      Sampling_Methods = methods_summary,
      Sampler_Type = record$Sampler_Type,
      Timestamp = record$Timestamp,
      stringsAsFactors = FALSE
    )

    # Add to existing data
    saved_data(rbind(saved_data(), record_df))

    # Show confirmation message
    showModal(modalDialog(
      title = "Success",
      "Data has been saved successfully!",
      easyClose = TRUE,
      footer = modalButton("Close")
    ))
  })

  # Generate Excel Report
  output$download_excel <- downloadHandler(
    filename = function() {
      # Use the formatted filename with station number and date
      format_filename()
    },
    content = function(file) {
      # Create a new workbook
      wb <- openxlsx::createWorkbook()

      # Add a worksheet for the main data
      openxlsx::addWorksheet(wb, "Sediment Field Data")

      # Create title style with large font and bold
      title_style <- openxlsx::createStyle(
        fontSize = 14,
        fontColour = "#FFFFFF",
        halign = "center",
        fgFill = "#4682B4",
        textDecoration = "bold",
        border = "TopBottomLeftRight",
        borderColour = "#000000"
      )

      # Create section header style
      section_style <- openxlsx::createStyle(
        fontSize = 12,
        fontColour = "#FFFFFF",
        halign = "left",
        fgFill = "#4682B4",
        textDecoration = "bold",
        border = "bottom",
        borderColour = "#000000"
      )

      # Create field name style
      field_style <- openxlsx::createStyle(
        fontSize = 11,
        textDecoration = "bold",
        border = "bottom",
        borderColour = "#CCCCCC"
      )

      # Create value style
      value_style <- openxlsx::createStyle(
        fontSize = 11,
        border = "bottom",
        borderColour = "#CCCCCC"
      )

      # Create table header style
      table_header_style <- openxlsx::createStyle(
        fontSize = 11,
        fontColour = "#FFFFFF",
        halign = "center",
        fgFill = "#4682B4",
        textDecoration = "bold",
        border = "TopBottomLeftRight",
        borderColour = "#000000",
        wrapText = TRUE
      )

      # Create table data style
      table_data_style <- openxlsx::createStyle(
        fontSize = 11,
        border = "TopBottomLeftRight",
        borderColour = "#CCCCCC",
        halign = "center"
      )

      # Create alternating row style
      alt_row_style <- openxlsx::createStyle(
        fontSize = 11,
        border = "TopBottomLeftRight",
        borderColour = "#CCCCCC",
        fgFill = "#F2F2F2",
        halign = "center"
      )

      # Get the current record
      record <- create_record()

      # Title Row
      openxlsx::writeData(wb, "Sediment Field Data", "Suspended Sediment Field Data Sheet", startRow = 1, startCol = 1)
      openxlsx::mergeCells(wb, "Sediment Field Data", rows = 1, cols = 1:8)
      openxlsx::addStyle(wb, "Sediment Field Data", title_style, rows = 1, cols = 1:8)

      # Date Generated Row
      openxlsx::writeData(wb, "Sediment Field Data", paste("Generated on:", format(Sys.time(), "%Y-%m-%d %H:%M:%S")), startRow = 2, startCol = 1)
      openxlsx::mergeCells(wb, "Sediment Field Data", rows = 2, cols = 1:8)

      # Station Information
      current_row <- 4
      openxlsx::writeData(wb, "Sediment Field Data", "Station Information", startRow = current_row, startCol = 1)
      openxlsx::mergeCells(wb, "Sediment Field Data", rows = current_row, cols = 1:8)
      openxlsx::addStyle(wb, "Sediment Field Data", section_style, rows = current_row, cols = 1:8)

      current_row <- current_row + 1
      fields <- c("Station Number", "Date", "Station Name", "Party", "Water Temperature (°C)")
      values <- c(record$Station_Number, record$Date, record$Station_Name, record$Party, record$Water_Temperature)

      for(i in 1:length(fields)) {
        openxlsx::writeData(wb, "Sediment Field Data", fields[i], startRow = current_row, startCol = 1)
        openxlsx::writeData(wb, "Sediment Field Data", values[i], startRow = current_row, startCol = 2)
        openxlsx::addStyle(wb, "Sediment Field Data", field_style, rows = current_row, cols = 1)
        openxlsx::addStyle(wb, "Sediment Field Data", value_style, rows = current_row, cols = 2)
        current_row <- current_row + 1
      }

      # Measurement Information
      current_row <- current_row + 1
      openxlsx::writeData(wb, "Sediment Field Data", "Measurement Information", startRow = current_row, startCol = 1)
      openxlsx::mergeCells(wb, "Sediment Field Data", rows = current_row, cols = 1:8)
      openxlsx::addStyle(wb, "Sediment Field Data", section_style, rows = current_row, cols = 1:8)

      current_row <- current_row + 1
      fields <- c("Measurement Type", "Stage", "Location")
      values <- c(record$Measurement_Type, record$Stage, record$Location)

      for(i in 1:length(fields)) {
        openxlsx::writeData(wb, "Sediment Field Data", fields[i], startRow = current_row, startCol = 1)
        openxlsx::writeData(wb, "Sediment Field Data", values[i], startRow = current_row, startCol = 2)
        openxlsx::addStyle(wb, "Sediment Field Data", field_style, rows = current_row, cols = 1)
        openxlsx::addStyle(wb, "Sediment Field Data", value_style, rows = current_row, cols = 2)
        current_row <- current_row + 1
      }

      # Channel Information
      current_row <- current_row + 1
      openxlsx::writeData(wb, "Sediment Field Data", "Channel Information", startRow = current_row, startCol = 1)
      openxlsx::mergeCells(wb, "Sediment Field Data", rows = current_row, cols = 1:8)
      openxlsx::addStyle(wb, "Sediment Field Data", section_style, rows = current_row, cols = 1:8)

      current_row <- current_row + 1
      fields <- c("Left Edge of Water (ft)", "Right Edge of Water (ft)", "Channel Width (ft)")
      values <- c(record$Left_Edge, record$Right_Edge, record$Channel_Width)

      for(i in 1:length(fields)) {
        openxlsx::writeData(wb, "Sediment Field Data", fields[i], startRow = current_row, startCol = 1)
        openxlsx::writeData(wb, "Sediment Field Data", values[i], startRow = current_row, startCol = 2)
        openxlsx::addStyle(wb, "Sediment Field Data", field_style, rows = current_row, cols = 1)
        openxlsx::addStyle(wb, "Sediment Field Data", value_style, rows = current_row, cols = 2)
        current_row <- current_row + 1
      }

      # Equipment Information
      current_row <- current_row + 1
      openxlsx::writeData(wb, "Sediment Field Data", "Equipment Information", startRow = current_row, startCol = 1)
      openxlsx::mergeCells(wb, "Sediment Field Data", rows = current_row, cols = 1:8)
      openxlsx::addStyle(wb, "Sediment Field Data", section_style, rows = current_row, cols = 1:8)

      current_row <- current_row + 1
      openxlsx::writeData(wb, "Sediment Field Data", "Sampler Type", startRow = current_row, startCol = 1)
      openxlsx::writeData(wb, "Sediment Field Data", record$Sampler_Type, startRow = current_row, startCol = 2)
      openxlsx::addStyle(wb, "Sediment Field Data", field_style, rows = current_row, cols = 1)
      openxlsx::addStyle(wb, "Sediment Field Data", value_style, rows = current_row, cols = 2)
      current_row <- current_row + 1

      # Sampling Methods Information
      if(length(record$Sampling_Methods) > 0) {
        # For each sampling method
        for(method_name in names(record$Sampling_Methods)) {
          method_data <- record$Sampling_Methods[[method_name]]

          current_row <- current_row + 1
          openxlsx::writeData(wb, "Sediment Field Data", paste("Sampling Method:", method_name), startRow = current_row, startCol = 1)
          openxlsx::mergeCells(wb, "Sediment Field Data", rows = current_row, cols = 1:8)
          openxlsx::addStyle(wb, "Sediment Field Data", section_style, rows = current_row, cols = 1:8)

          if(method_name %in% c("EWI Iso", "EWI Non-Iso")) {
            # For EWI methods with A and B sets

            # Set A
            current_row <- current_row + 1
            openxlsx::writeData(wb, "Sediment Field Data", "Set A", startRow = current_row, startCol = 1)
            openxlsx::mergeCells(wb, "Sediment Field Data", rows = current_row, cols = 1:8)

            current_row <- current_row + 1
            fields <- c("Number of Verticals", "Start Time", "End Time")
            values <- c(
              method_data$SetA$NumVerticals,
              method_data$SetA$StartTime,
              method_data$SetA$EndTime
            )

            for(i in 1:length(fields)) {
              openxlsx::writeData(wb, "Sediment Field Data", fields[i], startRow = current_row, startCol = 1)
              openxlsx::writeData(wb, "Sediment Field Data", values[i], startRow = current_row, startCol = 2)
              openxlsx::addStyle(wb, "Sediment Field Data", field_style, rows = current_row, cols = 1)
              openxlsx::addStyle(wb, "Sediment Field Data", value_style, rows = current_row, cols = 2)
              current_row <- current_row + 1
            }

            # Table of sections for Set A
            if(!is.null(method_data$SetA$Sections) && nrow(method_data$SetA$Sections) > 0) {
              current_row <- current_row + 1
              openxlsx::writeData(wb, "Sediment Field Data", "Sampling Sections - Set A", startRow = current_row, startCol = 1)
              openxlsx::mergeCells(wb, "Sediment Field Data", rows = current_row, cols = 1:8)

              current_row <- current_row + 1
              table_headers <- c("Section", "Distance from LEW (ft)", "Section Width (ft)")
              openxlsx::writeData(wb, "Sediment Field Data", table_headers, startRow = current_row, startCol = 1, colNames = FALSE)
              for(j in 1:length(table_headers)) {
                openxlsx::addStyle(wb, "Sediment Field Data", table_header_style, rows = current_row, cols = j)
              }

              # Write section data with alternating row colors
              for(i in 1:nrow(method_data$SetA$Sections)) {
                current_row <- current_row + 1
                row_data <- c(
                  method_data$SetA$Sections$Section[i],
                  method_data$SetA$Sections$Distance_from_LEW[i],
                  method_data$SetA$Sections$Section_Width[i]
                )
                openxlsx::writeData(wb, "Sediment Field Data", row_data, startRow = current_row, startCol = 1, colNames = FALSE)

                # Apply alternating row styles
                if(i %% 2 == 0) {
                  for(j in 1:length(row_data)) {
                    openxlsx::addStyle(wb, "Sediment Field Data", alt_row_style, rows = current_row, cols = j)
                  }
                } else {
                  for(j in 1:length(row_data)) {
                    openxlsx::addStyle(wb, "Sediment Field Data", table_data_style, rows = current_row, cols = j)
                  }
                }
              }
            }

            # Set B
            current_row <- current_row + 2
            openxlsx::writeData(wb, "Sediment Field Data", "Set B", startRow = current_row, startCol = 1)
            openxlsx::mergeCells(wb, "Sediment Field Data", rows = current_row, cols = 1:8)

            current_row <- current_row + 1
            fields <- c("Number of Verticals", "Start Time", "End Time")
            values <- c(
              method_data$SetB$NumVerticals,
              method_data$SetB$StartTime,
              method_data$SetB$EndTime
            )

            for(i in 1:length(fields)) {
              openxlsx::writeData(wb, "Sediment Field Data", fields[i], startRow = current_row, startCol = 1)
              openxlsx::writeData(wb, "Sediment Field Data", values[i], startRow = current_row, startCol = 2)
              openxlsx::addStyle(wb, "Sediment Field Data", field_style, rows = current_row, cols = 1)
              openxlsx::addStyle(wb, "Sediment Field Data", value_style, rows = current_row, cols = 2)
              current_row <- current_row + 1
            }

            # Table of sections for Set B
            if(!is.null(method_data$SetB$Sections) && nrow(method_data$SetB$Sections) > 0) {
              current_row <- current_row + 1
              openxlsx::writeData(wb, "Sediment Field Data", "Sampling Sections - Set B", startRow = current_row, startCol = 1)
              openxlsx::mergeCells(wb, "Sediment Field Data", rows = current_row, cols = 1:8)

              current_row <- current_row + 1
              table_headers <- c("Section", "Distance from LEW (ft)", "Section Width (ft)")
              openxlsx::writeData(wb, "Sediment Field Data", table_headers, startRow = current_row, startCol = 1, colNames = FALSE)
              for(j in 1:length(table_headers)) {
                openxlsx::addStyle(wb, "Sediment Field Data", table_header_style, rows = current_row, cols = j)
              }

              # Write section data with alternating row colors
              for(i in 1:nrow(method_data$SetB$Sections)) {
                current_row <- current_row + 1
                row_data <- c(
                  method_data$SetB$Sections$Section[i],
                  method_data$SetB$Sections$Distance_from_LEW[i],
                  method_data$SetB$Sections$Section_Width[i]
                )
                openxlsx::writeData(wb, "Sediment Field Data", row_data, startRow = current_row, startCol = 1, colNames = FALSE)

                # Apply alternating row styles
                if(i %% 2 == 0) {
                  for(j in 1:length(row_data)) {
                    openxlsx::addStyle(wb, "Sediment Field Data", alt_row_style, rows = current_row, cols = j)
                  }
                } else {
                  for(j in 1:length(row_data)) {
                    openxlsx::addStyle(wb, "Sediment Field Data", table_data_style, rows = current_row, cols = j)
                  }
                }
              }
            }

          } else if(method_name == "Multi-Vertical") {
            # Multi-Vertical method
            current_row <- current_row + 1
            fields <- c("Number of Verticals", "Start Time", "End Time")
            values <- c(
              method_data$NumVerticals,
              method_data$StartTime,
              method_data$EndTime
            )

            for(i in 1:length(fields)) {
              openxlsx::writeData(wb, "Sediment Field Data", fields[i], startRow = current_row, startCol = 1)
              openxlsx::writeData(wb, "Sediment Field Data", values[i], startRow = current_row, startCol = 2)
              openxlsx::addStyle(wb, "Sediment Field Data", field_style, rows = current_row, cols = 1)
              openxlsx::addStyle(wb, "Sediment Field Data", value_style, rows = current_row, cols = 2)
              current_row <- current_row + 1
            }

            # Table of sections
            if(!is.null(method_data$Sections) && nrow(method_data$Sections) > 0) {
              current_row <- current_row + 1
              openxlsx::writeData(wb, "Sediment Field Data", "Sampling Sections", startRow = current_row, startCol = 1)
              openxlsx::mergeCells(wb, "Sediment Field Data", rows = current_row, cols = 1:8)

              current_row <- current_row + 1
              table_headers <- c("Section", "Distance from LEW (ft)", "Section Width (ft)")
              openxlsx::writeData(wb, "Sediment Field Data", table_headers, startRow = current_row, startCol = 1, colNames = FALSE)
              for(j in 1:length(table_headers)) {
                openxlsx::addStyle(wb, "Sediment Field Data", table_header_style, rows = current_row, cols = j)
              }

              # Write section data with alternating row colors
              for(i in 1:nrow(method_data$Sections)) {
                current_row <- current_row + 1
                row_data <- c(
                  method_data$Sections$Section[i],
                  method_data$Sections$Distance_from_LEW[i],
                  method_data$Sections$Section_Width[i]
                )
                openxlsx::writeData(wb, "Sediment Field Data", row_data, startRow = current_row, startCol = 1, colNames = FALSE)

                # Apply alternating row styles
                if(i %% 2 == 0) {
                  for(j in 1:length(row_data)) {
                    openxlsx::addStyle(wb, "Sediment Field Data", alt_row_style, rows = current_row, cols = j)
                  }
                } else {
                  for(j in 1:length(row_data)) {
                    openxlsx::addStyle(wb, "Sediment Field Data", table_data_style, rows = current_row, cols = j)
                  }
                }
              }
            }
            # Set C
            current_row <- current_row + 2
            openxlsx::writeData(wb, "Sediment Field Data", "Set C", startRow = current_row, startCol = 1)
            openxlsx::mergeCells(wb, "Sediment Field Data", rows = current_row, cols = 1:8)

            current_row <- current_row + 1
            fields <- c("Number of Verticals", "Start Time", "End Time")
            values <- c(
              method_data$SetC$NumVerticals,
              method_data$SetC$StartTime,
              method_data$SetC$EndTime
            )

            for(i in 1:length(fields)) {
              openxlsx::writeData(wb, "Sediment Field Data", fields[i], startRow = current_row, startCol = 1)
              openxlsx::writeData(wb, "Sediment Field Data", values[i], startRow = current_row, startCol = 2)
              openxlsx::addStyle(wb, "Sediment Field Data", field_style, rows = current_row, cols = 1)
              openxlsx::addStyle(wb, "Sediment Field Data", value_style, rows = current_row, cols = 2)
              current_row <- current_row + 1
            }

            # Table of sections for Set C
            if(!is.null(method_data$SetC$Sections) && nrow(method_data$SetC$Sections) > 0) {
              current_row <- current_row + 1
              openxlsx::writeData(wb, "Sediment Field Data", "Sampling Sections - Set C", startRow = current_row, startCol = 1)
              openxlsx::mergeCells(wb, "Sediment Field Data", rows = current_row, cols = 1:8)

              current_row <- current_row + 1
              table_headers <- c("Section", "Distance from LEW (ft)", "Section Width (ft)")
              openxlsx::writeData(wb, "Sediment Field Data", table_headers, startRow = current_row, startCol = 1, colNames = FALSE)
              for(j in 1:length(table_headers)) {
                openxlsx::addStyle(wb, "Sediment Field Data", table_header_style, rows = current_row, cols = j)
              }

              # Write section data with alternating row colors
              for(i in 1:nrow(method_data$SetC$Sections)) {
                current_row <- current_row + 1
                row_data <- c(
                  method_data$SetC$Sections$Section[i],
                  method_data$SetC$Sections$Distance_from_LEW[i],
                  method_data$SetC$Sections$Section_Width[i]
                )
                openxlsx::writeData(wb, "Sediment Field Data", row_data, startRow = current_row, startCol = 1, colNames = FALSE)

                # Apply alternating row styles
                if(i %% 2 == 0) {
                  for(j in 1:length(row_data)) {
                    openxlsx::addStyle(wb, "Sediment Field Data", alt_row_style, rows = current_row, cols = j)
                  }
                } else {
                  for(j in 1:length(row_data)) {
                    openxlsx::addStyle(wb, "Sediment Field Data", table_data_style, rows = current_row, cols = j)
                  }
                }
              }
            }

          } else if(method_name %in% c("Single Vertical", "Point")) {
            # Single location methods
            current_row <- current_row + 1
            fields <- c("Start Time", "End Time", "Location from LEW (ft)", "Depth (ft)")
            values <- c(
              method_data$StartTime,
              method_data$EndTime,
              method_data$Location,
              method_data$Depth
            )

            for(i in 1:length(fields)) {
              openxlsx::writeData(wb, "Sediment Field Data", fields[i], startRow = current_row, startCol = 1)
              openxlsx::writeData(wb, "Sediment Field Data", values[i], startRow = current_row, startCol = 2)
              openxlsx::addStyle(wb, "Sediment Field Data", field_style, rows = current_row, cols = 1)
              openxlsx::addStyle(wb, "Sediment Field Data", value_style, rows = current_row, cols = 2)
              current_row <- current_row + 1
            }

          } else {
            # Other methods (ISCO, Grab, Box Single)
            current_row <- current_row + 1
            fields <- c("Start Time", "End Time", "Notes")
            values <- c(
              method_data$StartTime,
              method_data$EndTime,
              method_data$Notes
            )

            for(i in 1:length(fields)) {
              openxlsx::writeData(wb, "Sediment Field Data", fields[i], startRow = current_row, startCol = 1)
              openxlsx::writeData(wb, "Sediment Field Data", values[i], startRow = current_row, startCol = 2)
              openxlsx::addStyle(wb, "Sediment Field Data", field_style, rows = current_row, cols = 1)
              openxlsx::addStyle(wb, "Sediment Field Data", value_style, rows = current_row, cols = 2)
              current_row <- current_row + 1
            }
          }
        }
      }

      # Set column widths
      openxlsx::setColWidths(wb, "Sediment Field Data", cols = 1, widths = 25)
      openxlsx::setColWidths(wb, "Sediment Field Data", cols = 2, widths = 30)
      openxlsx::setColWidths(wb, "Sediment Field Data", cols = 3, widths = 20)

      # Add footer
      current_row <- current_row + 3
      footer_text <- paste("Generated by Suspended Sediment Field Data Sheet App |", format(Sys.time(), "%Y-%m-%d %H:%M:%S"))
      openxlsx::writeData(wb, "Sediment Field Data", footer_text, startRow = current_row, startCol = 1)
      openxlsx::mergeCells(wb, "Sediment Field Data", rows = current_row, cols = 1:8)

      # Save the workbook
      openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
    }
  )

  # Clear form button action
  observeEvent(input$clear_btn, {
    # Reset all inputs (base values)
    updateTextInput(session, "station_number", value = "")
    updateDateInput(session, "date", value = Sys.Date())
    updateTextInput(session, "station_name", value = "")
    updateTextInput(session, "party", value = "")
    updateNumericInput(session, "water_temp", value = NA)
    updateSelectInput(session, "measurement_type", selected = measurement_types[1])
    updateRadioButtons(session, "stage", selected = stage_options[1])
    updateNumericInput(session, "location_value", value = 0)
    updateRadioButtons(session, "location_direction", selected = "Upstream")
    updateNumericInput(session, "left_edge", value = 0)
    updateNumericInput(session, "right_edge", value = 0)
    updateCheckboxGroupInput(session, "selected_methods", selected = "EWI Iso")
    updateSelectInput(session, "sampler_type", selected = sampler_types[1])
    updateTextAreaInput(session, "other_sampler", value = "")

    # Method-specific inputs will be recreated when selection changes
  })
}

# Add shinyjs dependency
if(!require("shinyjs", character.only = TRUE, quietly = TRUE)) {
  install.packages("shinyjs")
  library(shinyjs)
}

# Run the application with shinyjs enabled
shinyApp(ui = shinyUI(fluidPage(
  shinyjs::useShinyjs(),
  ui
)), server = server)

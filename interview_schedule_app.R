# Interview Scheduler - Interactive R Shiny Application
# Install required packages if not already installed
if (!require("shiny")) install.packages("shiny")
if (!require("DT")) install.packages("DT")
if (!require("readxl")) install.packages("readxl")
if (!require("dplyr")) install.packages("dplyr")
if (!require("lubridate")) install.packages("lubridate")
if (!require("shinythemes")) install.packages("shinythemes")
if (!require("shinyWidgets")) install.packages("shinyWidgets")
if (!require("tidyr")) install.packages("tidyr")
if (!require("ggplot2")) install.packages("ggplot2")

library(shiny)
library(DT)
library(readxl)
library(dplyr)
library(lubridate)
library(shinythemes)
library(shinyWidgets)
library(tidyr)
library(ggplot2)

# Load and prepare data
load_interview_data <- function() {
  # File path - update this to your local path
  file_path <- "Interview_matrix.xlsx"
  
  # Load ranking data
  ranking <- read_excel(file_path, sheet = "ranking")
  
  # Load program dates
  program_dates <- read_excel(file_path, sheet = "program")
  
  # Process program dates into a long format
  interview_options <- data.frame()
  
  for (i in 1:nrow(program_dates)) {
    program <- program_dates$Programs[i]
    rank <- ranking$`preference order`[ranking$Programs == program]
    
    if (length(rank) == 0) next
    
    # Get all date columns
    date_cols <- grep("^date", names(program_dates), value = TRUE)
    
    for (col in date_cols) {
      date_val <- program_dates[[col]][i]
      if (!is.na(date_val)) {
        interview_options <- rbind(interview_options, 
                                    data.frame(
                                      Program = program,
                                      Rank = rank,
                                      Date = as.Date(date_val),
                                      DateStr = format(as.Date(date_val), "%Y-%m-%d"),
                                      Weekday = weekdays(as.Date(date_val)),
                                      stringsAsFactors = FALSE
                                    ))
      }
    }
  }
  
  # Add priority level
  interview_options$Priority <- ifelse(interview_options$Rank <= 5, "TOP PRIORITY",
                                        ifelse(interview_options$Rank <= 10, "HIGH",
                                               ifelse(interview_options$Rank <= 20, "MEDIUM", "LOW")))
  
  # Add month for grouping
  interview_options$Month <- format(interview_options$Date, "%B %Y")
  
  # Calculate "centrality score" - how close to middle of interview season
  date_range <- range(interview_options$Date)
  mid_date <- mean(date_range)
  interview_options$DaysFromMiddle <- abs(as.numeric(interview_options$Date - mid_date))
  
  return(list(
    interview_options = interview_options,
    ranking = ranking
  ))
}

# Define UI
ui <- fluidPage(
  theme = shinytheme("flatly"),
  
  tags$head(
    tags$style(HTML("
      .top-priority { background-color: #d4edda !important; font-weight: bold; }
      .high-priority { background-color: #fff3cd !important; }
      .medium-priority { background-color: #f8f9fa !important; }
      .low-priority { background-color: #ffffff !important; }
      .selected-row { background-color: #b8daff !important; }
      .conflict-date { background-color: #f8d7da !important; }
      .centered { text-align: center; }
      .stats-box { 
        padding: 15px; 
        margin: 10px;
        border-radius: 5px;
        background-color: #f8f9fa;
        border: 1px solid #dee2e6;
      }
      .recommendation-box {
        padding: 15px;
        margin: 10px;
        border-radius: 5px;
        background-color: #e7f3ff;
        border: 1px solid #b8daff;
      }
    "))
  ),
  
  titlePanel(
    div(
      h2("📅 Interactive Interview Scheduler", style = "color: #2c3e50;"),
      h4("Optimize your residency interview schedule with mutual exclusion", style = "color: #7f8c8d;")
    )
  ),
  
  sidebarLayout(
    sidebarPanel(
      width = 3,
      
      h4("📊 Schedule Statistics"),
      div(class = "stats-box",
          uiOutput("stats")
      ),
      
      hr(),
      
      h4("🎯 Optimization Settings"),
      
      checkboxInput("prefer_middle", 
                    "Prioritize mid-season dates for top programs",
                    value = TRUE),
      
      checkboxInput("avoid_wedding",
                    "Avoid Nov 21-23 (Wedding)",
                    value = TRUE),
      
      hr(),
      
      h4("🔍 Filter Options"),
      
      pickerInput("priority_filter",
                  "Show Priority Levels:",
                  choices = c("TOP PRIORITY", "HIGH", "MEDIUM", "LOW"),
                  selected = c("TOP PRIORITY", "HIGH", "MEDIUM", "LOW"),
                  multiple = TRUE,
                  options = list(
                    `actions-box` = TRUE,
                    `selected-text-format` = "count > 2"
                  )),
      
      pickerInput("month_filter",
                  "Show Months:",
                  choices = NULL,  # Will be updated dynamically
                  selected = NULL,
                  multiple = TRUE,
                  options = list(
                    `actions-box` = TRUE,
                    `selected-text-format` = "count > 1"
                  )),
      
      hr(),
      
      actionButton("auto_optimize", 
                   "🤖 Auto-Optimize Schedule", 
                   class = "btn-primary btn-block"),
      
      br(),
      
      actionButton("clear_all", 
                   "🔄 Clear All Selections", 
                   class = "btn-warning btn-block"),
      
      br(),
      
      downloadButton("download_schedule", 
                     "💾 Download Schedule (CSV)", 
                     class = "btn-success btn-block"),
      
      hr(),
      
      h4("💼 Configuration Management"),
      
      downloadButton("save_config", 
                     "💾 Save Configuration", 
                     class = "btn-info btn-block"),
      
      br(),
      
      fileInput("load_config", 
                "📂 Load Configuration",
                accept = c(".rds"),
                buttonLabel = "Browse...",
                placeholder = "Select saved config file"),
      
      textInput("config_name",
                "Configuration Name:",
                value = paste0("schedule_", format(Sys.Date(), "%Y%m%d")),
                placeholder = "Enter config name")
    ),
    
    mainPanel(
      width = 9,
      
      tabsetPanel(
        id = "main_tabs",
        
        tabPanel("📅 Schedule Builder",
                 br(),
                 fluidRow(
                   column(12,
                          h4("Available Interview Slots"),
                          p("Click on a row to select/deselect an interview. 
                            Only one interview can be scheduled per day."),
                          DTOutput("available_interviews")
                   )
                 ),
                 br(),
                 fluidRow(
                   column(12,
                          h4("Your Selected Schedule"),
                          DTOutput("selected_schedule")
                   )
                 )
        ),
        
        tabPanel("📈 Schedule Analysis",
                 br(),
                 fluidRow(
                   column(12,
                          h4("Schedule Timeline"),
                          plotOutput("timeline_plot", height = "400px")
                   )
                 ),
                 br(),
                 fluidRow(
                   column(6,
                          h4("Programs Not Yet Scheduled"),
                          DTOutput("unscheduled_programs")
                   ),
                   column(6,
                          h4("Schedule Recommendations"),
                          div(class = "recommendation-box",
                              uiOutput("recommendations")
                          )
                   )
                 )
        ),
        
        tabPanel("📊 Comparison View",
                 br(),
                 fluidRow(
                   column(12,
                          h4("Date Availability Matrix"),
                          p("This matrix shows all programs and their available dates. 
                            Green = Selected, Gray = Available, Red = Conflict"),
                          DTOutput("availability_matrix")
                   )
                 )
        )
      )
    )
  )
)

# Define Server
server <- function(input, output, session) {
  
  # Initialize reactive values first
  values <- reactiveValues(
    selected_interviews = data.frame(),
    selected_dates = character(),
    selected_programs = character(),
    data_loaded = FALSE,
    interview_data = NULL,
    ranking_data = NULL
  )
  
  # Load data once at startup
  observe({
    if (!values$data_loaded) {
      result <- load_interview_data()
      values$interview_data <- result$interview_options
      values$ranking_data <- result$ranking
      values$data_loaded <- TRUE
    }
  })
  
  # Update month filter choices
  observe({
    req(values$data_loaded)
    interview_data <- values$interview_data
    months <- unique(interview_data$Month)
    months <- months[order(match(months, format(sort(unique(interview_data$Date)), "%B %Y")))]
    
    updatePickerInput(session, "month_filter",
                      choices = months,
                      selected = months)
  })
  
  # Filter interview options
  filtered_interviews <- reactive({
    req(values$data_loaded)
    interview_data <- values$interview_data
    
    # Check if filters are ready
    priority_filter <- if(is.null(input$priority_filter)) {
      c("TOP PRIORITY", "HIGH", "MEDIUM", "LOW")
    } else {
      input$priority_filter
    }
    
    month_filter <- if(is.null(input$month_filter)) {
      unique(interview_data$Month)
    } else {
      input$month_filter
    }
    
    # Apply filters
    filtered <- interview_data %>%
      filter(Priority %in% priority_filter,
             Month %in% month_filter)
    
    # Remove wedding dates if selected
    if (!is.null(input$avoid_wedding) && input$avoid_wedding) {
      wedding_dates <- as.Date(c("2025-11-21", "2025-11-22", "2025-11-23"))
      filtered <- filtered %>%
        filter(!(Date %in% wedding_dates))
    }
    
    # Sort by rank and date
    if (!is.null(input$prefer_middle) && input$prefer_middle) {
      filtered <- filtered %>%
        arrange(Rank, DaysFromMiddle, Date)
    } else {
      filtered <- filtered %>%
        arrange(Rank, Date)
    }
    
    # Add selection status
    filtered$Status <- ifelse(filtered$Program %in% values$selected_programs, "Program Already Scheduled",
                              ifelse(filtered$DateStr %in% values$selected_dates, "Date Conflict",
                                     "Available"))
    
    filtered
  })
  
  # Display available interviews
  output$available_interviews <- renderDT({
    df <- filtered_interviews() %>%
      select(Program, Rank, Date = DateStr, Weekday, Priority, Status)
    
    # Add a column to track if this specific combination is selected
    df$IsSelected <- paste(df$Program, df$Date) %in% 
                     paste(values$selected_programs, values$selected_dates)
    
    dt <- datatable(df,
              selection = 'single',
              options = list(
                pageLength = -1,  # Show all rows
                scrollY = "500px",
                scrollCollapse = TRUE,
                paging = FALSE,  # Disable pagination
                dom = 'ft',  # Just filter and table, no pagination info
                columnDefs = list(
                  list(className = 'dt-center', targets = '_all'),
                  list(visible = FALSE, targets = 6)  # Hide IsSelected column
                )
              ),
              rownames = FALSE) %>%
      formatStyle(
        'Priority',
        backgroundColor = styleEqual(
          c("TOP PRIORITY", "HIGH", "MEDIUM", "LOW"),
          c("#d4edda", "#fff3cd", "#f8f9fa", "#ffffff")
        )
      ) %>%
      formatStyle(
        'Status',
        color = styleEqual(
          c("Available", "Program Already Scheduled", "Date Conflict"),
          c("green", "orange", "red")
        ),
        fontWeight = styleEqual(
          c("Available", "Program Already Scheduled", "Date Conflict"),
          c("normal", "bold", "bold")
        )
      )
    
    # Add highlighting for selected rows
    dt %>% formatStyle(
      columns = 0:6,
      valueColumns = 'IsSelected',
      backgroundColor = styleEqual(TRUE, '#e6f2ff')
    )
  }, server = FALSE)
  
  # Handle row selection in available interviews
  observeEvent(input$available_interviews_rows_selected, {
    if (!is.null(input$available_interviews_rows_selected)) {
      selected_row <- filtered_interviews()[input$available_interviews_rows_selected, ]
      
      # Check if this specific interview is already selected (for toggle functionality)
      is_already_selected <- any(
        values$selected_programs == selected_row$Program & 
        values$selected_dates == selected_row$DateStr
      )
      
      if (is_already_selected) {
        # Remove from selected interviews (toggle off)
        values$selected_interviews <- values$selected_interviews %>%
          filter(!(Program == selected_row$Program & DateStr == selected_row$DateStr))
        
        # Rebuild the dates and programs arrays from the remaining selections
        if (nrow(values$selected_interviews) > 0) {
          values$selected_dates <- values$selected_interviews$DateStr
          values$selected_programs <- values$selected_interviews$Program
        } else {
          values$selected_dates <- character()
          values$selected_programs <- character()
        }
        
        showNotification(paste("Removed:", selected_row$Program, "on", selected_row$DateStr),
                         type = "message", duration = 2)
      } else if (selected_row$Status == "Available") {
        # Add to selected interviews
        values$selected_interviews <- rbind(values$selected_interviews, selected_row)
        values$selected_dates <- c(values$selected_dates, selected_row$DateStr)
        values$selected_programs <- c(values$selected_programs, selected_row$Program)
        
        showNotification(paste("Added:", selected_row$Program, "on", selected_row$DateStr),
                         type = "message", duration = 2)
      } else if (selected_row$Status == "Date Conflict") {
        showNotification(paste("Cannot select: Another interview already scheduled on", selected_row$DateStr),
                         type = "warning", duration = 3)
      } else {
        showNotification(paste("Cannot select: Program already has a scheduled date"),
                         type = "warning", duration = 3)
      }
    }
  })
  
  # Display selected schedule
  output$selected_schedule <- renderDT({
    if (nrow(values$selected_interviews) > 0) {
      df <- values$selected_interviews %>%
        arrange(Date) %>%
        select(Program, Rank, Date = DateStr, Weekday, Priority)
      
      datatable(df,
                selection = 'single',
                options = list(
                  pageLength = -1,  # Show all rows
                  scrollY = "300px",
                  scrollCollapse = TRUE,
                  paging = FALSE,  # Disable pagination
                  dom = 't'  # Just table, no pagination
                ),
                rownames = FALSE) %>%
        formatStyle(
          'Priority',
          backgroundColor = styleEqual(
            c("TOP PRIORITY", "HIGH", "MEDIUM", "LOW"),
            c("#d4edda", "#fff3cd", "#f8f9fa", "#ffffff")
          )
        )
    } else {
      datatable(data.frame(Message = "No interviews selected yet"),
                options = list(dom = 't'),
                rownames = FALSE)
    }
  }, server = FALSE)
  
  # Remove from selected schedule
  observeEvent(input$selected_schedule_rows_selected, {
    if (!is.null(input$selected_schedule_rows_selected) && nrow(values$selected_interviews) > 0) {
      sorted_selected <- values$selected_interviews %>% arrange(Date)
      row_to_remove <- sorted_selected[input$selected_schedule_rows_selected, ]
      
      # Remove from all tracking variables
      values$selected_interviews <- values$selected_interviews %>%
        filter(!(Program == row_to_remove$Program & DateStr == row_to_remove$DateStr))
      
      # Rebuild the dates and programs arrays from the remaining selections
      if (nrow(values$selected_interviews) > 0) {
        values$selected_dates <- values$selected_interviews$DateStr
        values$selected_programs <- values$selected_interviews$Program
      } else {
        values$selected_dates <- character()
        values$selected_programs <- character()
      }
      
      showNotification(paste("Removed:", row_to_remove$Program),
                       type = "message", duration = 2)
    }
  })
  
  # Statistics
  output$stats <- renderUI({
    total_selected <- nrow(values$selected_interviews)
    top5_selected <- sum(values$selected_interviews$Rank <= 5, na.rm = TRUE)
    top10_selected <- sum(values$selected_interviews$Rank <= 10, na.rm = TRUE)
    
    div(
      p(strong("Total Scheduled: "), total_selected),
      p(strong("Top 5 Programs: "), paste0(top5_selected, " / 5")),
      p(strong("Top 10 Programs: "), paste0(top10_selected, " / 10")),
      if (total_selected > 0) {
        p(strong("Date Range: "), 
          paste(min(values$selected_interviews$Date), "to", 
                max(values$selected_interviews$Date)))
      }
    )
  })
  
  # Auto-optimize function
  observeEvent(input$auto_optimize, {
    showNotification("Auto-optimizing schedule...", type = "message", duration = 2)
    
    # Clear current selections
    values$selected_interviews <- data.frame()
    values$selected_dates <- character()
    values$selected_programs <- character()
    
    # Get all interviews sorted by rank
    req(values$data_loaded)
    all_interviews <- values$interview_data
    
    # Remove wedding dates if needed
    if (!is.null(input$avoid_wedding) && input$avoid_wedding) {
      wedding_dates <- as.Date(c("2025-11-21", "2025-11-22", "2025-11-23"))
      all_interviews <- all_interviews %>%
        filter(!(Date %in% wedding_dates))
    }
    
    # Sort by rank and centrality if preferred
    if (!is.null(input$prefer_middle) && input$prefer_middle) {
      all_interviews <- all_interviews %>%
        arrange(Rank, DaysFromMiddle)
    } else {
      all_interviews <- all_interviews %>%
        arrange(Rank, Date)
    }
    
    # Select interviews - build the data frame properly
    selected_list <- list()
    selected_dates_temp <- character()
    selected_programs_temp <- character()
    
    for (i in 1:nrow(all_interviews)) {
      interview <- all_interviews[i, ]
      
      # Check if we can add this interview
      if (!(interview$Program %in% selected_programs_temp) && 
          !(interview$DateStr %in% selected_dates_temp)) {
        
        selected_list[[length(selected_list) + 1]] <- interview
        selected_dates_temp <- c(selected_dates_temp, interview$DateStr)
        selected_programs_temp <- c(selected_programs_temp, interview$Program)
      }
    }
    
    # Combine all selected interviews into a data frame
    if (length(selected_list) > 0) {
      values$selected_interviews <- do.call(rbind, selected_list)
      values$selected_dates <- selected_dates_temp
      values$selected_programs <- selected_programs_temp
    }
    
    showNotification(paste("Auto-optimization complete!", 
                           nrow(values$selected_interviews), "interviews scheduled"),
                     type = "message", duration = 3)
  })
  
  # Clear all selections
  observeEvent(input$clear_all, {
    values$selected_interviews <- data.frame()
    values$selected_dates <- character()
    values$selected_programs <- character()
    showNotification("All selections cleared", type = "message", duration = 2)
  })
  
  # Unscheduled programs
  output$unscheduled_programs <- renderDT({
    req(values$data_loaded)
    all_programs <- values$ranking_data
    scheduled <- unique(values$selected_programs)
    
    unscheduled <- all_programs %>%
      filter(!(Programs %in% scheduled)) %>%
      select(Program = Programs, Rank = `preference order`) %>%
      arrange(Rank)
    
    datatable(unscheduled,
              options = list(
                pageLength = -1,  # Show all rows
                scrollY = "300px",
                scrollCollapse = TRUE,
                paging = FALSE,  # Disable pagination
                dom = 't'  # Just table
              ),
              rownames = FALSE) %>%
      formatStyle(
        'Rank',
        backgroundColor = styleInterval(c(5, 10, 20),
                                         c("#d4edda", "#fff3cd", "#f8f9fa", "#ffffff"))
      )
  }, server = FALSE)
  
  # Recommendations
  output$recommendations <- renderUI({
    req(values$data_loaded)
    unscheduled <- values$ranking_data %>%
      filter(!(Programs %in% values$selected_programs))
    
    top10_unscheduled <- unscheduled %>%
      filter(`preference order` <= 10)
    
    recommendations <- list()
    
    if (nrow(top10_unscheduled) > 0) {
      recommendations <- append(recommendations, 
                                list(
                                  h5("⚠️ Top 10 Programs Not Scheduled:"),
                                  tags$ul(
                                    lapply(1:nrow(top10_unscheduled), function(i) {
                                      tags$li(paste(top10_unscheduled$Programs[i], 
                                                    "(Rank", top10_unscheduled$`preference order`[i], ")"))
                                    })
                                  )
                                ))
    }
    
    if (nrow(values$selected_interviews) > 0) {
      # Check for scheduling patterns
      dates <- as.Date(values$selected_interviews$Date)
      date_gaps <- diff(sort(dates))
      
      if (max(date_gaps) > 7) {
        recommendations <- append(recommendations,
                                  list(
                                    br(),
                                    p("📅 You have gaps of more than 7 days between some interviews. 
                                      Consider filling these gaps if possible.")
                                  ))
      }
      
      # Check top program positioning
      top5_dates <- values$selected_interviews %>%
        filter(Rank <= 5) %>%
        pull(Date)
      
      if (length(top5_dates) > 0) {
        all_dates <- values$selected_interviews$Date
        date_range <- range(all_dates)
        mid_date <- mean(date_range)
        
        avg_distance <- mean(abs(as.numeric(top5_dates - mid_date)))
        
        if (avg_distance > 10) {
          recommendations <- append(recommendations,
                                    list(
                                      br(),
                                      p("📍 Your top 5 programs are not well-centered in your schedule. 
                                        Consider adjusting dates to place them closer to mid-season.")
                                    ))
        }
      }
    }
    
    if (length(recommendations) == 0) {
      recommendations <- list(p("✅ Your schedule looks well-optimized!"))
    }
    
    do.call(tagList, recommendations)
  })
  
  # Timeline plot
  output$timeline_plot <- renderPlot({
    req(values$data_loaded)
    
    # Always use the full date range from all available interviews
    all_dates <- values$interview_data$Date
    date_range <- range(all_dates, na.rm = TRUE)
    
    # Create a sequence of all dates for grid marks
    all_date_seq <- seq.Date(from = date_range[1], to = date_range[2], by = "day")
    
    if (nrow(values$selected_interviews) > 0) {
      df <- values$selected_interviews %>%
        mutate(
          Date = as.Date(Date),
          PriorityNum = case_when(
            Priority == "TOP PRIORITY" ~ 1,
            Priority == "HIGH" ~ 2,
            Priority == "MEDIUM" ~ 3,
            TRUE ~ 4
          )
        ) %>%
        arrange(Date)
      
      # Create base plot with full date range
      p <- ggplot(df, aes(x = Date, y = 1, label = Program)) +
        geom_point(aes(color = Priority, size = 5 - PriorityNum), alpha = 0.7) +
        geom_text(angle = 45, hjust = 0, vjust = -1, size = 3) +
        scale_color_manual(values = c("TOP PRIORITY" = "#28a745", 
                                       "HIGH" = "#ffc107", 
                                       "MEDIUM" = "#6c757d",
                                       "LOW" = "#adb5bd")) +
        scale_size_continuous(range = c(3, 8), guide = "none")
    } else {
      # Create empty plot with same date range
      p <- ggplot(data.frame(Date = date_range, y = 1), aes(x = Date, y = y)) +
        annotate("text", x = mean(date_range), y = 1, 
                 label = "Select interviews to see them on timeline", 
                 size = 5, color = "gray40")
    }
    
    # Apply consistent scale and theme to all plots
    p + scale_x_date(limits = date_range,
                     breaks = seq.Date(from = date_range[1], to = date_range[2], by = "3 days"),
                     minor_breaks = all_date_seq,
                     date_labels = "%b %d",
                     expand = expansion(mult = 0.01)) +
        labs(title = "Interview Schedule Timeline",
             subtitle = paste("Interview Season:", format(date_range[1], "%B %d, %Y"), 
                            "to", format(date_range[2], "%B %d, %Y")),
             x = "Date",
             y = "",
             color = "Priority Level") +
        theme_minimal() +
        theme(
          axis.text.y = element_blank(),
          axis.ticks.y = element_blank(),
          panel.grid.major.y = element_blank(),
          panel.grid.minor.y = element_blank(),
          panel.grid.minor.x = element_line(color = "gray90", size = 0.25),
          panel.grid.major.x = element_line(color = "gray80", size = 0.5),
          axis.text.x = element_text(angle = 45, hjust = 1, size = 8),
          legend.position = "bottom",
          plot.subtitle = element_text(color = "gray50", size = 10)
        ) +
        ylim(0.5, 1.5)
  })
  
  # Date Availability Matrix for Comparison View
  output$availability_matrix <- renderDT({
    req(values$data_loaded)
    
    # Create a matrix of all programs and their available dates
    all_interviews <- values$interview_data
    
    # Get unique programs and dates
    programs <- unique(all_interviews$Program)
    dates <- sort(unique(all_interviews$Date))
    
    # Create matrix data frame
    matrix_data <- expand.grid(
      Program = programs,
      Date = dates,
      stringsAsFactors = FALSE
    )
    
    # Add date string for display
    matrix_data$DateStr <- format(matrix_data$Date, "%m/%d")
    
    # Check availability for each combination
    matrix_data$Status <- NA
    for (i in 1:nrow(matrix_data)) {
      prog <- matrix_data$Program[i]
      date <- matrix_data$Date[i]
      date_str <- format(date, "%Y-%m-%d")
      
      # Check if this combination exists in available interviews
      is_available <- any(all_interviews$Program == prog & all_interviews$Date == date)
      
      if (is_available) {
        # Check if it's selected
        is_selected <- any(values$selected_programs == prog & values$selected_dates == date_str)
        if (is_selected) {
          matrix_data$Status[i] <- "Selected"
        } else {
          # Check if there's a conflict
          date_conflict <- date_str %in% values$selected_dates
          program_conflict <- prog %in% values$selected_programs
          
          if (date_conflict && program_conflict) {
            matrix_data$Status[i] <- "Both Conflict"
          } else if (date_conflict) {
            matrix_data$Status[i] <- "Date Conflict"
          } else if (program_conflict) {
            matrix_data$Status[i] <- "Program Scheduled"
          } else {
            matrix_data$Status[i] <- "Available"
          }
        }
      } else {
        matrix_data$Status[i] <- "Not Offered"
      }
    }
    
    # Get program rankings
    rankings <- values$ranking_data
    program_ranks <- data.frame(
      Program = rankings$Programs,
      Rank = rankings$`preference order`
    )
    
    # Pivot to wide format for display
    matrix_wide <- matrix_data %>%
      select(Program, DateStr, Status) %>%
      tidyr::pivot_wider(
        names_from = DateStr,
        values_from = Status,
        values_fill = "Not Offered"
      )
    
    # Add rank column
    matrix_wide <- matrix_wide %>%
      left_join(program_ranks, by = "Program") %>%
      arrange(Rank) %>%
      select(Rank, Program, everything())
    
    # Create datatable with color coding
    dt <- datatable(
      matrix_wide,
      options = list(
        pageLength = -1,
        scrollY = "500px",
        scrollX = TRUE,
        scrollCollapse = TRUE,
        paging = FALSE,
        dom = 'ft',
        columnDefs = list(
          list(className = 'dt-center', targets = '_all'),
          list(width = '60px', targets = 0),
          list(width = '150px', targets = 1)
        )
      ),
      rownames = FALSE
    )
    
    # Apply color coding to date columns
    date_cols <- names(matrix_wide)[3:ncol(matrix_wide)]
    
    for (col in date_cols) {
      dt <- dt %>%
        formatStyle(
          col,
          backgroundColor = styleEqual(
            c("Selected", "Available", "Date Conflict", "Program Scheduled", "Both Conflict", "Not Offered"),
            c("#28a745", "#ffffff", "#ffc107", "#17a2b8", "#dc3545", "#e9ecef")
          ),
          color = styleEqual(
            c("Selected", "Available", "Date Conflict", "Program Scheduled", "Both Conflict", "Not Offered"),
            c("white", "black", "black", "white", "white", "#6c757d")
          ),
          fontWeight = styleEqual(
            c("Selected"),
            c("bold")
          )
        )
    }
    
    # Highlight top programs
    dt %>%
      formatStyle(
        'Rank',
        backgroundColor = styleInterval(
          c(5, 10, 20),
          c("#d4edda", "#fff3cd", "#f8f9fa", "#ffffff")
        )
      )
  }, server = FALSE)
  
  # Download handler for CSV export
  output$download_schedule <- downloadHandler(
    filename = function() {
      paste0("interview_schedule_", Sys.Date(), ".csv")
    },
    content = function(file) {
      if (nrow(values$selected_interviews) > 0) {
        schedule <- values$selected_interviews %>%
          arrange(Date) %>%
          select(Program, Rank, Date = DateStr, Weekday, Priority)
        write.csv(schedule, file, row.names = FALSE)
      } else {
        write.csv(data.frame(Message = "No interviews scheduled"), file, row.names = FALSE)
      }
    }
  )
  
  # Save configuration handler
  output$save_config <- downloadHandler(
    filename = function() {
      config_name <- if(is.null(input$config_name) || input$config_name == "") {
        paste0("schedule_config_", Sys.Date())
      } else {
        input$config_name
      }
      paste0(config_name, ".rds")
    },
    content = function(file) {
      # Create configuration object with all necessary data
      config <- list(
        selected_interviews = values$selected_interviews,
        selected_dates = values$selected_dates,
        selected_programs = values$selected_programs,
        config_name = input$config_name,
        saved_date = Sys.Date(),
        saved_time = Sys.time(),
        total_interviews = nrow(values$selected_interviews),
        settings = list(
          prefer_middle = input$prefer_middle,
          avoid_wedding = input$avoid_wedding,
          priority_filter = input$priority_filter,
          month_filter = input$month_filter
        )
      )
      
      # Save as RDS file
      saveRDS(config, file)
      
      showNotification(paste("Configuration saved:", input$config_name),
                       type = "message", duration = 3)
    }
  )
  
  # Load configuration handler
  observeEvent(input$load_config, {
    req(input$load_config)
    
    tryCatch({
      # Load the configuration file
      config <- readRDS(input$load_config$datapath)
      
      # Validate the configuration
      if (!all(c("selected_interviews", "selected_dates", "selected_programs") %in% names(config))) {
        stop("Invalid configuration file format")
      }
      
      # Clear current selections
      values$selected_interviews <- data.frame()
      values$selected_dates <- character()
      values$selected_programs <- character()
      
      # Load the saved selections
      if (!is.null(config$selected_interviews) && nrow(config$selected_interviews) > 0) {
        values$selected_interviews <- config$selected_interviews
        values$selected_dates <- config$selected_dates
        values$selected_programs <- config$selected_programs
      }
      
      # Update settings if they exist
      if (!is.null(config$settings)) {
        if (!is.null(config$settings$prefer_middle)) {
          updateCheckboxInput(session, "prefer_middle", value = config$settings$prefer_middle)
        }
        if (!is.null(config$settings$avoid_wedding)) {
          updateCheckboxInput(session, "avoid_wedding", value = config$settings$avoid_wedding)
        }
        if (!is.null(config$settings$priority_filter)) {
          updatePickerInput(session, "priority_filter", selected = config$settings$priority_filter)
        }
        if (!is.null(config$settings$month_filter)) {
          updatePickerInput(session, "month_filter", selected = config$settings$month_filter)
        }
      }
      
      # Update config name field
      if (!is.null(config$config_name)) {
        updateTextInput(session, "config_name", value = config$config_name)
      }
      
      # Show success message
      config_info <- paste0(
        "Configuration loaded: ", 
        ifelse(!is.null(config$config_name), config$config_name, "Unknown"),
        " (", config$total_interviews, " interviews)"
      )
      
      if (!is.null(config$saved_date)) {
        config_info <- paste0(config_info, "\nSaved on: ", config$saved_date)
      }
      
      showNotification(config_info, type = "message", duration = 4)
      
    }, error = function(e) {
      showNotification(paste("Error loading configuration:", e$message),
                       type = "error", duration = 5)
    })
  })
}

# Run the app
shinyApp(ui = ui, server = server)

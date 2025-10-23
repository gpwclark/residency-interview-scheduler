# Interactive Interview Scheduler with Mutual Exclusion
# Run this app with: shiny::runApp("interview_scheduler_app.R")

library(shiny)
library(shinydashboard)
library(DT)
library(readxl)
library(dplyr)
library(lubridate)
library(tidyr)
library(ggplot2)

# Configuration
file_path <- "Interview_matrix.xlsx"  # Update this path to your Excel file location

# Function to load and prepare data
load_interview_data <- function() {
  ranking <- read_excel(file_path, sheet = "ranking")
  program <- read_excel(file_path, sheet = "program")
  
  # Transform to long format
  program_long <- program %>%
    pivot_longer(cols = starts_with("date"), 
                 names_to = "date_option", 
                 values_to = "date") %>%
    filter(!is.na(date)) %>%
    mutate(date = as.Date(date))
  
  # Merge with rankings and add priority levels
  interview_data <- program_long %>%
    left_join(ranking %>% select(Programs, `preference order`), 
              by = "Programs") %>%
    arrange(`preference order`, date) %>%
    mutate(
      priority_level = case_when(
        `preference order` <= 5 ~ "TOP PRIORITY",
        `preference order` <= 10 ~ "HIGH",
        `preference order` <= 20 ~ "MEDIUM",
        TRUE ~ "LOW"
      )
    )
  
  return(interview_data)
}

# UI Definition
ui <- dashboardPage(
  dashboardHeader(title = "Interview Scheduler - Mutual Exclusion Mode"),
  
  dashboardSidebar(
    sidebarMenu(
      menuItem("Schedule Builder", tabName = "scheduler", icon = icon("calendar")),
      menuItem("Analysis", tabName = "analysis", icon = icon("chart-bar")),
      menuItem("Instructions", tabName = "help", icon = icon("question-circle"))
    ),
    br(),
    div(style = "padding: 10px;",
        h4("Controls"),
        checkboxGroupInput("priority_filter", "Show Priority Levels:",
                          choices = c("TOP PRIORITY", "HIGH", "MEDIUM", "LOW"),
                          selected = c("TOP PRIORITY", "HIGH")),
        br(),
        actionButton("optimize_middle", "Auto-Schedule Top 5", 
                    class = "btn-success", width = "100%"),
        br(), br(),
        actionButton("reset_schedule", "Reset All", 
                    class = "btn-warning", width = "100%"),
        br(), br(),
        downloadButton("download_schedule", "Export Schedule", 
                      class = "btn-primary", width = "100%")
    )
  ),
  
  dashboardBody(
    tags$head(
      tags$style(HTML("
        .content-wrapper, .right-side {
          background-color: #f4f4f4;
        }
        .top-priority { background-color: #d4edda !important; font-weight: bold; }
        .high-priority { background-color: #fff3cd !important; }
        .medium-priority { background-color: #d1ecf1 !important; }
        .low-priority { background-color: #f8f9fa !important; }
      "))
    ),
    
    tabItems(
      # Schedule Builder Tab
      tabItem(tabName = "scheduler",
        fluidRow(
          box(width = 12, title = "📅 Your Interview Schedule", 
              status = "primary", solidHeader = TRUE,
              DT::dataTableOutput("current_schedule")
          )
        ),
        fluidRow(
          box(width = 6, title = "🎯 Available Interview Slots", 
              status = "info", solidHeader = TRUE,
              p("Click any row to add it to your schedule"),
              DT::dataTableOutput("available_slots")
          ),
          box(width = 6, title = "📊 Program Status Overview", 
              status = "warning", solidHeader = TRUE,
              DT::dataTableOutput("program_status")
          )
        )
      ),
      
      # Analysis Tab
      tabItem(tabName = "analysis",
        fluidRow(
          valueBoxOutput("total_scheduled"),
          valueBoxOutput("top5_scheduled"),
          valueBoxOutput("top10_scheduled")
        ),
        fluidRow(
          box(width = 12, title = "Schedule Timeline", status = "primary",
              plotOutput("timeline_plot", height = "500px")
          )
        ),
        fluidRow(
          box(width = 6, title = "Priority Distribution", status = "success",
              plotOutput("priority_plot", height = "350px")
          ),
          box(width = 6, title = "Schedule Optimization Analysis", status = "info",
              verbatimTextOutput("optimization_analysis")
          )
        )
      ),
      
      # Help Tab
      tabItem(tabName = "help",
        box(width = 12, title = "How to Use This Scheduler", status = "info",
            h4("🎯 Key Features:"),
            p("This scheduler enforces mutual exclusion - only one interview per day, 
              and once a program is scheduled, it's removed from available slots."),
            br(),
            h4("📝 Instructions:"),
            tags$ol(
              tags$li("Click on rows in 'Available Interview Slots' to add them"),
              tags$li("The system automatically prevents date conflicts"),
              tags$li("Use 'Auto-Schedule Top 5' to optimize top programs for mid-November"),
              tags$li("Click the ❌ button in the schedule to remove an interview"),
              tags$li("Export your final schedule using the Download button")
            ),
            br(),
            h4("🎨 Color Coding:"),
            div(class = "top-priority", style = "padding: 5px; margin: 5px;", 
                "TOP PRIORITY (Ranks 1-5) - Schedule these in mid-November if possible"),
            div(class = "high-priority", style = "padding: 5px; margin: 5px;", 
                "HIGH (Ranks 6-10)"),
            div(class = "medium-priority", style = "padding: 5px; margin: 5px;", 
                "MEDIUM (Ranks 11-20)"),
            div(class = "low-priority", style = "padding: 5px; margin: 5px;", 
                "LOW (Ranks 21+)"),
            br(),
            h4("💡 Optimization Tips:"),
            tags$ul(
              tags$li("Top 5 programs should ideally be scheduled Nov 10-20"),
              tags$li("Avoid Nov 21-23 (Abby's Wedding)"),
              tags$li("Leave buffer days between interviews when possible"),
              tags$li("Contact Brigham/MGH (Rank #8) directly - no dates in system")
            )
        )
      )
    )
  )
)

# Server Logic
server <- function(input, output, session) {
  
  # Reactive values to store state
  vals <- reactiveValues(
    interview_data = NULL,
    scheduled = data.frame(),
    scheduled_dates = c(),
    scheduled_programs = c()
  )
  
  # Load data on startup
  observe({
    tryCatch({
      vals$interview_data <- load_interview_data()
      showNotification("Data loaded successfully!", type = "success")
    }, error = function(e) {
      showNotification(paste("Error loading data:", e$message), type = "error", duration = NULL)
    })
  })
  
  # Available slots (with mutual exclusion)
  available_slots <- reactive({
    req(vals$interview_data)
    
    vals$interview_data %>%
      filter(!date %in% vals$scheduled_dates) %>%
      filter(!Programs %in% vals$scheduled_programs) %>%
      filter(priority_level %in% input$priority_filter) %>%
      arrange(`preference order`, date) %>%
      select(
        Program = Programs, 
        Rank = `preference order`, 
        Date = date,
        Priority = priority_level
      ) %>%
      mutate(
        `Day of Week` = weekdays(Date),
        `Date Display` = format(Date, "%b %d, %Y")
      )
  })
  
  # Render available slots table
  output$available_slots <- DT::renderDataTable({
    slots_display <- available_slots() %>%
      select(Program, Rank, `Date Display`, `Day of Week`, Priority)
    
    DT::datatable(
      slots_display,
      selection = 'single',
      options = list(
        pageLength = 10,
        ordering = TRUE,
        order = list(list(1, 'asc'))
      ),
      rownames = FALSE
    ) %>%
      formatStyle(
        'Priority',
        backgroundColor = styleEqual(
          c("TOP PRIORITY", "HIGH", "MEDIUM", "LOW"),
          c("#d4edda", "#fff3cd", "#d1ecf1", "#f8f9fa")
        ),
        fontWeight = styleEqual("TOP PRIORITY", "bold")
      )
  })
  
  # Handle slot selection
  observeEvent(input$available_slots_rows_selected, {
    req(input$available_slots_rows_selected)
    
    selected <- available_slots()[input$available_slots_rows_selected, ]
    
    # Add to schedule
    vals$scheduled <- rbind(vals$scheduled, selected)
    vals$scheduled_dates <- c(vals$scheduled_dates, selected$Date)
    vals$scheduled_programs <- c(vals$scheduled_programs, selected$Program)
    
    showNotification(
      paste("✅ Added", selected$Program, "on", format(selected$Date, "%B %d")),
      type = "success"
    )
  })
  
  # Render current schedule
  output$current_schedule <- DT::renderDataTable({
    if (nrow(vals$scheduled) == 0) {
      data.frame(Message = "No interviews scheduled yet. Click available slots to add them.")
    } else {
      display_schedule <- vals$scheduled %>%
        arrange(Date) %>%
        mutate(
          Action = sprintf('<button class="btn btn-danger btn-sm" onclick="Shiny.setInputValue(\'remove_row\', %d, {priority: \'event\'})">❌</button>', row_number()),
          `Date Display` = format(Date, "%b %d, %Y"),
          `Day` = weekdays(Date)
        ) %>%
        select(Action, Program, Rank, `Date Display`, Day, Priority)
      
      DT::datatable(
        display_schedule,
        escape = FALSE,
        selection = 'none',
        options = list(pageLength = 15, dom = 't'),
        rownames = FALSE
      ) %>%
        formatStyle(
          'Priority',
          backgroundColor = styleEqual(
            c("TOP PRIORITY", "HIGH", "MEDIUM", "LOW"),
            c("#d4edda", "#fff3cd", "#d1ecf1", "#f8f9fa")
          ),
          fontWeight = styleEqual("TOP PRIORITY", "bold")
        )
    }
  })
  
  # Handle removal
  observeEvent(input$remove_row, {
    row_idx <- as.numeric(input$remove_row)
    if (row_idx <= nrow(vals$scheduled)) {
      sorted_schedule <- vals$scheduled %>% arrange(Date)
      removed <- sorted_schedule[row_idx, ]
      
      vals$scheduled_dates <- vals$scheduled_dates[vals$scheduled_dates != removed$Date]
      vals$scheduled_programs <- vals$scheduled_programs[vals$scheduled_programs != removed$Program]
      vals$scheduled <- vals$scheduled[!(vals$scheduled$Date == removed$Date & 
                                         vals$scheduled$Program == removed$Program), ]
      
      showNotification(paste("Removed", removed$Program), type = "warning")
    }
  })
  
  # Program status
  output$program_status <- DT::renderDataTable({
    req(vals$interview_data)
    
    status <- vals$interview_data %>%
      group_by(Programs, `preference order`) %>%
      summarise(.groups = "drop") %>%
      mutate(
        Status = ifelse(Programs %in% vals$scheduled_programs, "✅ Scheduled", "⏳ Available"),
        Priority = case_when(
          `preference order` <= 5 ~ "TOP",
          `preference order` <= 10 ~ "HIGH",
          TRUE ~ "OTHER"
        )
      ) %>%
      arrange(`preference order`) %>%
      select(Program = Programs, Rank = `preference order`, Status, Priority)
    
    DT::datatable(
      status,
      options = list(pageLength = 10, dom = 'ftp'),
      rownames = FALSE
    ) %>%
      formatStyle(
        'Status',
        color = styleEqual(c("✅ Scheduled", "⏳ Available"), c("green", "gray")),
        fontWeight = styleEqual("✅ Scheduled", "bold")
      )
  })
  
  # Auto-optimize top 5
  observeEvent(input$optimize_middle, {
    req(vals$interview_data)
    
    # Reset schedule
    vals$scheduled <- data.frame()
    vals$scheduled_dates <- c()
    vals$scheduled_programs <- c()
    
    # Get top 5 programs
    top5_programs <- vals$interview_data %>%
      filter(`preference order` <= 5) %>%
      pull(Programs) %>%
      unique()
    
    # Target mid-November dates
    target_start <- as.Date("2025-11-10")
    target_end <- as.Date("2025-11-20")
    
    # Schedule each program
    for (prog in top5_programs) {
      prog_data <- vals$interview_data %>%
        filter(Programs == prog, !date %in% vals$scheduled_dates)
      
      if (nrow(prog_data) == 0) next
      
      # Prefer dates in target range
      middle_dates <- prog_data %>% filter(date >= target_start, date <= target_end)
      
      selected <- if (nrow(middle_dates) > 0) middle_dates[1, ] else prog_data[1, ]
      
      # Add to schedule
      vals$scheduled <- rbind(vals$scheduled, data.frame(
        Program = selected$Programs,
        Rank = selected$`preference order`,
        Date = selected$date,
        Priority = selected$priority_level
      ))
      vals$scheduled_dates <- c(vals$scheduled_dates, selected$date)
      vals$scheduled_programs <- c(vals$scheduled_programs, selected$Programs)
    }
    
    showNotification("✨ Top 5 programs optimized for mid-November!", type = "success")
  })
  
  # Reset schedule
  observeEvent(input$reset_schedule, {
    vals$scheduled <- data.frame()
    vals$scheduled_dates <- c()
    vals$scheduled_programs <- c()
    showNotification("Schedule cleared", type = "info")
  })
  
  # Download handler
  output$download_schedule <- downloadHandler(
    filename = function() {
      paste0("interview_schedule_", Sys.Date(), ".csv")
    },
    content = function(file) {
      if (nrow(vals$scheduled) > 0) {
        export <- vals$scheduled %>%
          arrange(Date) %>%
          mutate(Date = format(Date, "%Y-%m-%d"), Day = weekdays(Date)) %>%
          select(Date, Day, Program, Rank, Priority)
        write.csv(export, file, row.names = FALSE)
      }
    }
  )
  
  # Value boxes
  output$total_scheduled <- renderValueBox({
    valueBox(
      nrow(vals$scheduled),
      "Total Scheduled",
      icon = icon("calendar"),
      color = "blue"
    )
  })
  
  output$top5_scheduled <- renderValueBox({
    count <- sum(vals$scheduled$Rank <= 5, na.rm = TRUE)
    valueBox(
      paste0(count, "/5"),
      "Top 5 Programs",
      icon = icon("star"),
      color = if(count == 5) "green" else "yellow"
    )
  })
  
  output$top10_scheduled <- renderValueBox({
    count <- sum(vals$scheduled$Rank <= 10, na.rm = TRUE)
    valueBox(
      paste0(count, "/10"),
      "Top 10 Programs",
      icon = icon("trophy"),
      color = if(count >= 9) "green" else if(count >= 7) "yellow" else "red"
    )
  })
  
  # Timeline plot
  output$timeline_plot <- renderPlot({
    if (nrow(vals$scheduled) == 0) {
      plot(1, type = "n", axes = FALSE, xlab = "", ylab = "")
      text(1, 1, "No interviews scheduled yet", cex = 2, col = "gray")
    } else {
      plot_data <- vals$scheduled %>%
        mutate(
          Color = case_when(
            Rank <= 5 ~ "darkgreen",
            Rank <= 10 ~ "orange",
            Rank <= 20 ~ "blue",
            TRUE ~ "gray"
          ),
          Label = paste0("#", Rank)
        )
      
      ggplot(plot_data, aes(x = Date, y = reorder(Program, -Rank))) +
        geom_point(aes(color = Color), size = 8) +
        geom_text(aes(label = Label), color = "white", size = 3) +
        scale_color_identity() +
        scale_x_date(date_breaks = "1 week", date_labels = "%b %d") +
        labs(x = "Interview Date", y = "Program (ordered by rank)", 
             title = "Interview Schedule Timeline") +
        theme_minimal() +
        theme(
          axis.text.x = element_text(angle = 45, hjust = 1),
          panel.grid.minor = element_blank(),
          text = element_text(size = 12)
        ) +
        geom_vline(xintercept = as.Date(c("2025-11-21", "2025-11-22", "2025-11-23")),
                   linetype = "dashed", color = "red", alpha = 0.3) +
        annotate("text", x = as.Date("2025-11-22"), y = 0.5, 
                 label = "Wedding", color = "red", angle = 90)
    }
  })
  
  # Priority distribution plot
  output$priority_plot <- renderPlot({
    if (nrow(vals$scheduled) == 0) {
      plot(1, type = "n", axes = FALSE, xlab = "", ylab = "")
      text(1, 1, "No data", cex = 1.5, col = "gray")
    } else {
      priority_summary <- vals$scheduled %>%
        group_by(Priority) %>%
        summarise(Count = n(), .groups = "drop") %>%
        mutate(Priority = factor(Priority, 
                                 levels = c("TOP PRIORITY", "HIGH", "MEDIUM", "LOW")))
      
      ggplot(priority_summary, aes(x = Priority, y = Count, fill = Priority)) +
        geom_bar(stat = "identity") +
        scale_fill_manual(values = c(
          "TOP PRIORITY" = "darkgreen",
          "HIGH" = "orange",
          "MEDIUM" = "lightblue",
          "LOW" = "gray"
        )) +
        geom_text(aes(label = Count), vjust = -0.5, size = 5) +
        labs(title = "Interviews by Priority Level", x = "", y = "Count") +
        theme_minimal() +
        theme(legend.position = "none", text = element_text(size = 12))
    }
  })
  
  # Optimization analysis
  output$optimization_analysis <- renderPrint({
    if (nrow(vals$scheduled) == 0) {
      cat("No interviews scheduled.\n")
    } else {
      cat("=== SCHEDULE ANALYSIS ===\n\n")
      
      date_range <- range(vals$scheduled$Date)
      cat("Date Range:", format(date_range[1], "%b %d"), "-", 
          format(date_range[2], "%b %d, %Y"), "\n")
      cat("Duration:", diff(date_range), "days\n\n")
      
      # Top 5 in optimal range
      top5 <- vals$scheduled %>% filter(Rank <= 5)
      if (nrow(top5) > 0) {
        optimal <- sum(top5$Date >= as.Date("2025-11-10") & 
                      top5$Date <= as.Date("2025-11-20"))
        cat("Top 5 in mid-Nov:", optimal, "/", nrow(top5), "\n\n")
      }
      
      # Missing top 10
      missing <- setdiff(1:10, vals$scheduled$Rank)
      if (length(missing) > 0) {
        cat("Missing Top 10 Ranks:", paste(missing, collapse = ", "), "\n")
        if (8 %in% missing) {
          cat("⚠️ Contact Brigham/MGH directly!\n")
        }
      } else {
        cat("✅ All Top 10 scheduled!\n")
      }
    }
  })
}

# Run the application
shinyApp(ui = ui, server = server)

library(shiny)    #For dashboard
library(dplyr)    #For data manipulation
library(ggplot2)  #For enhanced visualization
library(readxl)   #For loading excel data set

# Define file path (update if necessary)
file_path <- "D:/tulgaa/Hult/MBAN/Spring/R/Air France Case Spreadsheet Supplement.xls"
sheet_name <- "DoubleClick"

# Load the dataset once
my_data <- read_excel(file_path, sheet = sheet_name) %>%
  as.data.frame()

# UI: Defines the layout
ui <- fluidPage(
  titlePanel("AirFrance Analysis Dashboard"),
  
  sidebarLayout(
    sidebarPanel(
      selectInput("selectlabel",
                  "Select a Publisher",
                  choices = unique(my_data$`Publisher Name`),
                  multiple = TRUE,
                  selected = unique(my_data$`Publisher Name`)[1]),  # choose the 1st publisher on the file by default
      
      sliderInput("cpc_filter",
                  "Filter CPC (Cost Per Click)",
                  min = min(my_data$`Avg. Cost per Click`, na.rm = TRUE),
                  max = max(my_data$`Avg. Cost per Click`, na.rm = TRUE),
                  value = c(min(my_data$`Avg. Cost per Click`, na.rm = TRUE),
                            max(my_data$`Avg. Cost per Click`, na.rm = TRUE))),
      
      sliderInput("revenue_filter",
                  "Filter Total Revenue",
                  min = min(my_data$Amount, na.rm = TRUE),
                  max = max(my_data$Amount, na.rm = TRUE),
                  value = c(min(my_data$Amount, na.rm = TRUE),
                            max(my_data$Amount, na.rm = TRUE)))
    ),
    
    mainPanel(
      tabsetPanel(
        tabPanel("Summary Statistics", tableOutput("summary_table")),
        tabPanel("Revenue Analysis",
                 tableOutput("revenue_publisher"),
                 tableOutput("revenue_match"),
                 tableOutput("revenue_campaign")),
        tabPanel("CPC Distribution", plotOutput("cpc_plot"))
      )
    )
  )
)

# Server: Defines backend logic
server <- function(input, output, session) {
  # Filtered Data based on CPC & Revenue selection
  filtered_data <- reactive({
    req(my_data)
    
    my_data %>%
      filter(`Publisher Name` %in% input$selectlabel,
             `Avg. Cost per Click` >= input$cpc_filter[1],
             `Avg. Cost per Click` <= input$cpc_filter[2],
             Amount >= input$revenue_filter[1],
             Amount <= input$revenue_filter[2])
  })
  
  # Compute summary statistics based on selection
  summary_data <- reactive({
    req(filtered_data())
    
    filtered_data() %>%
      group_by(`Publisher Name`) %>%
      summarise(
        Total_Clicks = sum(Clicks, na.rm = TRUE),
        Total_Click_Charges = sum(`Click Charges`, na.rm = TRUE),
        Total_Impressions = sum(Impressions, na.rm = TRUE),
        Total_Net_Revenue = sum(Amount, na.rm = TRUE) - Total_Click_Charges,
        Avg_Cost_Per_Click = Total_Click_Charges / Total_Clicks,
        Total_Bookings = sum(`Total Volume of Bookings`, na.rm = TRUE),
        Revenue_Per_Booking = sum(Amount, na.rm = TRUE) / Total_Bookings,
        ROA = Total_Net_Revenue / Total_Click_Charges * 100,
        Conversion_Rate_Impressions = Total_Bookings / Total_Impressions * 100,
        Conversion_Rate_Clicks = Total_Bookings / Total_Clicks * 100,
        Cost_Per_Booking = Total_Click_Charges / Total_Bookings
      ) %>%
      arrange(`Publisher Name`)
  })
  
  # Render summary table
  output$summary_table <- renderTable({
    req(summary_data())
    summary_data()
  })
  
  # Render conversion histogram
  output$conversion_plot <- renderPlot({
    req(filtered_data())
    
    # Create binary conversion variable
    filtered_data()$Conversion_Booking <- ifelse(filtered_data()$`Total Volume of Bookings` > 0, 1, 0)
    
    ggplot(filtered_data(), aes(x = Conversion_Booking)) +
      geom_histogram(bins = 30, fill = "blue", alpha = 0.7) +
      facet_wrap(~`Publisher Name`) +
      labs(title = "Distribution of Conversion",
           x = "Conversion = 1",
           y = "Count") +
      theme_minimal()
  })
  
  # Revenue Analysis Tables
  output$revenue_publisher <- renderTable({
    filtered_data() %>%
      group_by(`Publisher Name`) %>%
      summarise(Total_Revenue = sum(Amount, na.rm = TRUE)) %>%
      arrange(desc(Total_Revenue))
  })
  
  output$revenue_match <- renderTable({
    filtered_data() %>%
      group_by(`Match Type`) %>%
      summarise(Total_Revenue = sum(Amount, na.rm = TRUE)) %>%
      arrange(desc(Total_Revenue))
  })
  
  output$revenue_campaign <- renderTable({
    filtered_data() %>%
      group_by(Campaign) %>%
      summarise(Total_Revenue = sum(Amount, na.rm = TRUE)) %>%
      arrange(desc(Total_Revenue))
  })
  
  # CPC Distribution Plot
  output$cpc_plot <- renderPlot({
    req(filtered_data())
    
    ggplot(filtered_data(), aes(x = `Avg. Cost per Click`)) +
      geom_histogram(bins = 30, fill = "blue", alpha = 0.7) +
      labs(title = "Distribution of Cost Per Click",
           x = "Cost Per Click",
           y = "Count") +
      theme_minimal()
  })
}

# Run the app
shinyApp(ui, server)
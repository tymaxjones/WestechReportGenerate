
# You can learn more about package authoring with RStudio at:
#
#   http://r-pkgs.had.co.nz/
#
# Some useful keyboard shortcuts for package authoring:
#
#   Install Package:           'Ctrl + Shift + B'
#   Check Package:             'Ctrl + Shift + E'
#   Test Package:              'Ctrl + Shift + T'

#' Generate Report Function
#'
#' @param
#'
#' @return Will return a Word Document on report generate
#' @export
#'
#' @examples
#' GenereateReport()
GenerateReport <- function(test_input) {

  months <- c("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

  #Shiny Widget CSS
  CSS <- '

                              .bg-grey {
                                background-color: #d6d6d6;
                              }
                              .bg-dark-grey {
                                background-color: #696868;
                                color: white;
                              }
                              .bg-blue {
                                background-color: #34A6E7;
                                color: white;
                              }
                              hr{
                                height: 1px;
                                background-color: #707070;
                                border: none;
                              }
                                .irs-grid-pol.small {height: 0px;}

                              .sidenav {
                              height: 100%;
                              width: 100px;
                              position: fixed;
                              z-index: 1;
                              top: 0;
                              left: 0;
                              background-color: #F5F5F5;
                              overflow-x: hidden;
                              padding-top: 20px;
                              }


                            .sidenav a {
                            padding: 6px 8px 6px 16px;
                            text-decoration: none;
                            font-size: 18px;
                            color: #818181;
                            display: block;
                            }

                            .sidenav a:hover {
                            color: #a1a1a1;
                            }

                            .selectize-dropdown {
                            bottom: 100% !important; top: auto !important;
                            font-size: 14px;
                            height: 75px;
                            width: 100%;
                            }

                            #run{background-color:#34A6E7}

'



  #User Interface
  ui <- bootstrapPage(

    options(shiny.maxRequestSize = 30*1024^2),

    #bring in the CSS
    tags$head(tags$style(HTML(CSS))),

    tags$div(class = "container-fluid",
             #Create Navbar
             tags$div(class = "sidenav",
                      br(),
                      br(),
                      br(),
                      br(),
                      br(),
                      br(),
                      br(),
                      tags$a(href = "#in",`data-toggle` = "tab",
                             "Input"),
                      br(),
                      br(),
                      br(),
                      tags$a(href = "#out",`data-toggle` = "tab",
                             "Output")
             ),

             #Create each tab adjust the margins for the sidebar
             div(class = "tab-content", style = "margin-left: 100px;",

                 #Input Tab
                 tags$div(class = "tab-pane active", id = "in",

                          #Header
                          div(class = "row text-center",
                              div(class = "col-xl-12 bg-blue",
                                  "Input Dataframes"
                              )
                          ),
                          div(class = "row text-center",
                              div(class = "col-xl-12",
                                  column(width = 12, align = "center",
                                         br(),
                                         fileInput("file1", "Upload Health and Safety Report",
                                                   multiple = TRUE,
                                                   accept = c("text/csv",
                                                              "text/comma-separated-values,text/plain",
                                                              ".csv")),
                                         fileInput("file2", "Upload Trial Balance Sheet",
                                                   multiple = TRUE,
                                                   accept = c("text/csv",
                                                              "text/comma-separated-values,text/plain",
                                                              ".csv")),
                                         fileInput("file3", "Upload Scorecard Sheet",
                                                   multiple = TRUE,
                                                   accept = c("text/csv",
                                                              "text/comma-separated-values,text/plain",
                                                              ".csv")),
                                         fileInput("file4", "Upload Report Forecast",
                                                   multiple = TRUE,
                                                   accept = c("text/csv",
                                                              "text/comma-separated-values,text/plain",
                                                              ".csv"))

                                  )
                              )
                          )
                 ),

                 #output Tab
                 tags$div(class = "tab-pane", id = "out",
                          div(class = "row text-center",
                              div(class = "col-xl-12",
                                  column(width = 12, align = "center",
                                         br(),
                                         br(),
                                         br(),
                                         br(),
                                         br(),
                                         br(),
                                         br(),
                                         br(),
                                         br(),
                                         actionButton("run","Create Report"),
                                         br(),
                                         br(),
                                         selectizeInput("current_month",
                                                     "Report Month",
                                                     choices = months,
                                                     selected = months[as.numeric(substr(as.character(Sys.Date()), 6 ,7)) - 1]
                                         )





                                  )
                              )
                          )



                 )

             )
    )
  )

  #Defining server logic
  server <- function(input, output, session) {

    ############################################################################
    ###############     SERVER LOGIC       #####################################
    ############################################################################



    observeEvent(input$run,{

      #Read in our dataframes
      df1 <- read_excel(input$file1$datapath, sheet =  "TOT Co. Rolling Data")
      df2 <- read_excel(input$file2$datapath, sheet =  "Inc by entity")
      df3a <- read_excel(input$file3$datapath, sheet = "Main")
      df3b <- read_excel(input$file3$datapath, sheet = "PM & Estimating")
      df3c <- read_excel(input$file3$datapath, sheet = "Quality & Efficiency")
      df3d <- read_excel(input$file3$datapath, sheet = "Backlog")
      df4 <- read_excel(input$file4$datapath)

      # Convert data frame into tibble
      df1 <- as_tibble(df1)

      # Extract meaningful data
      # Column X
      X <- 24
      # Remove top rows & select designated rows
      df1 <- df1[c(13,14,18),c(X)]
      colnames(df1) <- '2021 Actual'

      # Round the big numbers
      df1$'2021 Actual' <- round(as.numeric(df1$'2021 Actual'),2)
      df1$`2021 Actual` <- as.character(df1$`2021 Actual`)

      # Convert data frame into tibble
      df2 <- as_tibble(df2)

      # Extract meaningful data
      # Column W
      W <- 23
      # Select important data
      df2 <- df2[c(47,23,9,7,33,9,13),c(W)]
      # Rename Columns
      colnames(df2) <- c('2021 Actual')
      df2 <- df2 %>% add_column('Key' = c('1','2','3','4','5','6','7'),.before = 1)


      # Strip the commas
      df2$'2021 Actual' <- as.numeric(gsub(",","",df2$'2021 Actual'))

      df2[3,2] <- df2[2,2] / df2[3,2]
      df2[7,2] <- df2[7,2] / df2[6,2]

      df2 <- df2 %>%
        mutate(
          '2021 Actual' <- case_when(
            Key %in% c(1,2,4,5,6) ~ round(df2$'2021 Actual'/1000000,2),
            Key %in% c(3,7) ~ df2$'2021 Actual'
          )
        )
      colnames(df2) <- c('Key','2021 Actual','new_column')
      df2 <- df2[-c(1,2)]
      colnames(df2) <- c('2021 Actual')
      # Round the big numbers
      df2$'2021 Actual' <- round(df2$'2021 Actual',2)
      df2[3,1] <- df2[3,1]*100
      df2[7,1] <- df2[7,1]*100
      df2$`2021 Actual` <- as.character(df2$`2021 Actual`)
      df2[1,1] <- paste('$',df2[1,1])
      df2[2,1] <- paste('$',df2[2,1])
      df2[3,1] <- paste(df2[3,1],'%')
      df2[4,1] <- paste('$',df2[4,1])
      df2[5,1] <- paste('$',df2[5,1])
      df2[6,1] <- paste('$',df2[6,1])
      df2[7,1] <- paste(df2[7,1],'%')


      # Convert data frame into tibble
      df3a <- as_tibble(df3a)

      # Extract meaningful data
      # Column D
      D <- 3
      # Remove top rows & select designated rows
      df3a <- df3a[c(16),c(D)]
      colnames(df3a) <- '2021 Actual'
      #df3a %>%
      #add_column("Performance Indicators"="0" ) %>%
      #add_column("Dataset"="A" )
      # Round the big numbers
      df3a$'2021 Actual' <- round(as.numeric(df3a$'2021 Actual'),2)
      df3a$'2021 Actual' <- as.character(df3a$'2021 Actual')
      df3a[1,1] <- paste(df3a[1,1],'%')


      # Convert data frame into tibble
      df3b <- as_tibble(df3b)

      # Extract meaningful data
      # Column D
      D <- 3
      # Remove top rows & select designated rows
      df3b <- df3b[c(9,7,8),c(D)]
      colnames(df3b) <- '2021 Actual'
      # Round the big numbers
      df3b$'2021 Actual' <- round(as.numeric(df3b$'2021 Actual'),2)
      df3b
      df3b[1,1] <- df3b[1,1]*100
      df3b[2,1] <- df3b[2,1]*100
      df3b$'2021 Actual' <- as.character(df3b$'2021 Actual')
      df3b[1,1] <- paste(df3b[1,1],'%')
      df3b[2,1] <- paste(df3b[2,1],'%')


      # Convert data frame into tibble
      df3c <- as_tibble(df3c)

      # Extract meaningful data
      # Column C
      C <- 2
      # Remove top rows & select designated rows
      df3c <- df3c[c(6),c(C)]
      colnames(df3c) <- '2021 Actual'
      # Round the big numbers
      df3c$'2021 Actual' <- round(as.numeric(df3c$'2021 Actual'),2)
      df3c[1,1] <- df3c[1,1]*100
      df3c$'2021 Actual' <- as.character(df3c$'2021 Actual')
      df3c[1,1] <- paste(df3c[1,1],'%')


      # Convert data frame into tibble
      df3d <- as_tibble(df3d)

      # Extract meaningful data
      # Column C
      C <- 2
      # Column D
      D <- 3
      # Remove top rows & select designated rows
      df3d <- df3d[c(6),c(C,D)]
      colnames(df3d) <- c('Current','Year_End')
      df3d$'Current' <- as.numeric(df3d$'Current')
      df3d$'Year_End' <- as.numeric(df3d$'Year_End')
      df3d[1,1] <- df3d[1,1] - df3d[2,1]
      df3d <- df3d[-c(1)]
      colnames(df3d) <- c('2021 Actual')
      # Round the big numbers
      df3d$'2021 Actual' <- round(as.numeric(df3d$'2021 Actual'/1000000),2)
      df3d$'2021 Actual' <- as.character(df3d$'2021 Actual')


      # Convert data frame into tibble
      df4 <- as_tibble(df4)

      # Extract meaningful data
      # January
      Jan <- 3
      # February
      Feb <- 4
      # March
      Mar <- 5
      # April
      Apr <- 6
      # May
      May <- 7
      # June
      Jun <- 8
      # July
      Jul <- 9
      # August
      Aug <- 10
      # September
      Sep <- 11
      # October
      Oct <- 12
      # November
      Nov <- 13
      # December
      Dec <- 14
      # Remove top rows & select designated rows
      df4 <- df4[c(129),c(Mar)]
      colnames(df4) <- '2021 Actual'
      # Round the big numbers
      df4$'2021 Actual' <- round(as.numeric(df4$'2021 Actual')/1000000,2)
      df4$'2021 Actual' <- as.character(df4$'2021 Actual')


      df = bind_rows(df1,df2)
      df = bind_rows(df,df3b)
      df = bind_rows(df,df3a,)
      df = bind_rows(df,df3c)
      df = bind_rows(df,df3d)
      #df = bind_rows(df,df4)

      df <- df %>% add_row('2021 Actual' = '0',.before = 7)
      df <- df %>% add_row('2021 Actual' = '0',.before = 9)
      df <- df %>% add_row('2021 Actual' = '0',.before = 13)
      df <- df %>% add_row('2021 Actual' = '0',.before = 14)
      df <- df %>% add_row('2021 Actual' = '0',.before = 15)
      df <- df %>% add_row('2021 Actual' = '0',.before = 16)
      df <- df %>% add_row('2021 Actual' = '0',.before = 17)
      df <- df %>% add_row('2021 Actual' = '0',.before = 18)
      df <- df %>% add_row('2021 Actual' = '0',.before = 19)
      df <- df %>% add_row('2021 Actual' = '0%',.before = 20)
      df <- df %>% add_row('2021 Actual' = '$0',.before = 21)
      df <- df %>% add_column('Key' = c('1','2','3','4','5','6','7','8','A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','Q','P','R','9'),.before = 1)

      df <- df[order(df$Key),]

      df <- df[-c(1)]


      # Create Report
      df <- df %>%

        # Format Columns
        add_column("Performance Indicators"="0", .before = "2021 Actual" ) %>%
        add_column("Code"="H1", .before = "Performance Indicators") %>%
        add_column(`Time Base` = "YTD", .after = "Performance Indicators") %>%
        add_column(`2021 Forecast` = "YTD", .before = "2021 Actual") %>%
        add_column(`Status vs Budget` = "Color", .after = "2021 Actual") %>%
        add_column(`Direction of Trend vs Budget` = "Image", .after = "Status vs Budget") %>%

        # Format Rows
        add_row(`Performance Indicators` = "Health and Safety", .before = 1)
      df[2,1] <- "H1"
      df[2,2] <- "Total Recordable Injury Rate"
      df[2,3] <- "YTD"
      df[2,4] <- "< 1.15"
      df[2,7] <- "▲"

      df[3,1] <- "H2"
      df[3,2] <- "Days Away / Restricted Time Rate"
      df[3,3] <- "YTD"
      df[3,4] <- "< 1.15"
      df[3,7] <- "▼"

      df[4,1] <- "H3"
      df[4,2] <- "Number of Slam Dunks"
      df[4,3] <- "YTD"
      df[4,4] <- "> 40"
      df[4,7] <- "►"

      df <- df %>%
        add_row(`Performance Indicators` = "Financial Success", .before = 5)
      df[6,1] <- "F1a"
      df[6,2] <- "EBITDA ($Millions)"
      df[6,3] <- "Month"
      df[6,4] <- "$1.46"
      df[6,7] <- "▲"

      df[7,1] <- "F1b"
      df[7,2] <- "EBITDA ($Millions)"
      df[7,3] <- "YTD"
      df[7,4] <- "$2.97"
      df[7,7] <- "▲"

      df[8,1] <- "F1c"
      df[8,2] <- "EBITDA % of Sales"
      df[8,3] <- "YTD"
      df[8,4] <- "7.6%"
      df[8,7] <- "▲"

      df[9,1] <- "F2"
      df[9,2] <- "EBITE ($Millions)"
      df[9,3] <- "YTD"
      df[9,4] <- "0"
      df[9,7] <- "▲"

      df[10,1] <- "F3a"
      df[10,2] <- "Bookings ($Millions)"
      df[10,3] <- "YTD"
      df[10,4] <- "$61.9"
      df[10,7] <- "▲"

      df[11,1] <- "F4"
      df[11,2] <- "Backlog ($Millions)"
      df[11,3] <- "Period End"
      df[11,4] <- "0"
      df[11,7] <- "▲"

      df[12,1] <- "F5"
      df[12,2] <- "Booked GM %"
      df[12,3] <- "YTD"
      df[12,4] <- "0"
      df[12,7] <- "▲"

      df[13,1] <- "F6a"
      df[13,2] <- "Revenue ($Millions)"
      df[13,3] <- "Month"
      df[13,4] <- "$19.7"
      df[13,7] <- "▲"

      df[14,1] <- "F6b"
      df[14,2] <- "Revenue ($Millions)"
      df[14,3] <- "YTD"
      df[14,4] <- "$39"
      df[14,7] <- "▲"

      df[15,1] <- "F7"
      df[15,2] <- "Gross Margin %"
      df[15,3] <- "YTD"
      df[15,4] <- "31%"
      df[15,7] <- "▲"

      df[16,1] <- "F8a"
      df[16,2] <- "Cap Ex ($Millions)"
      df[16,3] <- "Month"
      df[16,4] <- "0"
      df[16,7] <- "▲"

      df[17,1] <- "F8b"
      df[17,2] <- "Cap Ex ($Millions)"
      df[17,3] <- "YTD"
      df[17,4] <- "0"
      df[17,7] <- "▲"

      df[18,1] <- "F9"
      df[18,2] <- "Net Bank Debt ($Millions)"
      df[18,3] <- "Period End"
      df[18,4] <- "0"
      df[18,7] <- "▲"

      df[19,1] <- "F10"
      df[19,2] <- "RONA (NI/(FA+NWC)"
      df[19,3] <- "Period End"
      df[19,4] <- "0"
      df[19,7] <- "▲"

      df[20,1] <- "F11"
      df[20,2] <- "Rental Unit Utilization"
      df[20,3] <- "Period End"
      df[20,4] <- "0"
      df[20,7] <- "▲"

      df <- df %>%
        add_row(`Performance Indicators` = "People Services", .before = 21)
      df[22,1] <- "P1a"
      df[22,2] <- "Headcount USA"
      df[22,3] <- "Period End"
      df[22,4] <- "575"
      df[22,7] <- "▲"

      df[23,1] <- "P1b"
      df[23,2] <- "Headcount Total"
      df[23,3] <- "Period End"
      df[23,4] <- "656"
      df[23,7] <- "▲"

      df[24,1] <- "P2"
      df[24,2] <- "Employee Turnover Rate"
      df[24,3] <- "Rolling 12 Month"
      df[24,4] <- "12%"
      df[24,7] <- "▲"

      df[25,1] <- "P3"
      df[25,2] <- "GM$ / Employee Total ($’000)"
      df[25,3] <- "Rolling 12 Month"
      df[25,4] <- "$72"
      df[25,7] <- "▲"

      df <- df %>%
        add_row(`Performance Indicators` = "Customer", .before = 26)
      df[27,1] <- "C1"
      df[27,2] <- "On Time Delivery %"
      df[27,3] <- "YTD"
      df[27,4] <- "97%"
      df[27,7] <- "▲"

      df <- df %>%
        add_row(`Performance Indicators` = "Winning Team", .before = 28)
      df[29,1] <- "OE1"
      df[29,2] <- "On Time Submittal %"
      df[29,3] <- "YTD"
      df[29,4] <- "97%"
      df[29,7] <- "▲"

      df[30,1] <- "OE2"
      df[30,2] <- "On Time Release to Purchasing Median Weeks to Fab"
      df[30,3] <- "YTD"
      df[30,4] <- "97%"
      df[30,7] <- "▲"

      df[31,1] <- "OE3"
      df[31,2] <- "Median Weeks to Fab"
      df[31,3] <- "YTD"
      df[31,4] <- "16"
      df[31,7] <- "▲"

      df[32,1] <- "OE4"
      df[32,2] <- "Warranty %"
      df[32,3] <- "18 Month Rolling"
      df[32,4] <- "2.0%"
      df[32,7] <- "▲"


      # Color Codes
      grey = "#c9c9c9"
      green = "#5ceb54"
      red = "#ff4a4a"
      yellow = "#fcff47"

      # Creat Flextable
      ft <- flextable(df)

      # Edit Flextable
      ft <- ft %>%
        hrule(rule = "exact",part = "header") %>%
        height(height = .37, part = "header") %>%
        align(align = "center", part = "header") %>%
        hrule(rule = "exact", part = "body") %>%
        height(height = .23, part = "body") %>%


        # Adjust Column Widths
        width(j = 1, width = .4) %>%
        width(j = 2, width = 2) %>%
        width(j = 3, width = 1) %>%
        width(j = 4, width = .8) %>%
        width(j = 5, width = .8) %>%
        width(j = 7, width = 1.2) %>%

        # Create cell Borders
        border_inner(border = fp_border(color = "#c9c9c9", width = 1.5)) %>%
        border_outer(border = fp_border(color = "#c9c9c9", width = 1.5)) %>%

        # Merge Header Rows
        merge_h(i = 1) %>%
        merge_h(i = 5) %>%
        merge_h(i = 21) %>%
        merge_h(i = 26) %>%
        merge_h(i = 28) %>%

        # Color Actual Column Grey
        bg(j = 5, bg = "#c9c9c9", part = "body") %>%

        # Color Header Rows Grey
        bg(i = 1, bg = "#c9c9c9", part = "body") %>%
        bg(i = 5, bg = "#c9c9c9", part = "body") %>%
        bg(i = 21, bg = "#c9c9c9", part = "body") %>%
        bg(i = 26, bg = "#c9c9c9", part = "body") %>%
        bg(i = 28, bg = "#c9c9c9", part = "body") %>%

        ############################################################################
      # Add Title
      set_caption("WesTech Monthly Level 1 Report - Apr 2021") %>%
        ############################################################################

      # Edit Table Font
      font(font = "Calibri (Body)", part = "all") %>%
        fontsize(size = 8, part = "all") %>%
        bold(bold = TRUE, part = "all")

      # Add Conditional Formatting
      ft <- bg(ft, i = ~ `2021 Actual`  < `2021 Forecast`, j = c(6), bg=red)
      ft <- bg(ft, i = ~ `2021 Actual`  == `2021 Forecast`, j = c(6), bg=yellow)
      ft <- bg(ft, i = ~ `2021 Actual`  > `2021 Forecast`, j = c(6), bg=green)

      # Print table to Word Document
      print(ft,"docx")


    })

  }

  runGadget(ui, server)
}


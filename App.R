# Load libraries ##########################################################

library(shiny)
library(data.table)
library(openxlsx)
library(DT)
library(faosws)
library(faoswsUtil)
library(SwsApiClient)
library(writexl)
library(lubridate)

# Connection to SWS #######################################################

if(CheckDebug()){
  
  library(faoswsModules)
  SETT <- ReadSettings("sws.yml")
  
  # SetClientFiles(SETT[["certdir"]])
  GetTestEnvironment(baseUrl = SETT[["server"]], token = SETT[["token"]])

  # # Initialize client
  # initializeClient()
}

# Get the list of codelists ###############################################

# all_codelists <- getAllCodelists()

# Import utils functions ##################################################

source("utils.R")

# -------------------------------------------------------------------------
# UI
# -------------------------------------------------------------------------

ui <- fluidPage(
  
  tags$head(
    tags$style(HTML("

      /* ── Color palette & typography ───────────────────────────── */
      @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

      :root {
        --bg:        #b6bfc1;
        --surface:   #97acc8;
        --border:    #051230;
        --accent:    #051230;
        --accent2:   #7c3aed;
        --text:      #051230;
        --text-muted:#051230;
        --success:   #10b981;
        --radius:    8px;
      }

      body {
        background: var(--bg);
        color: var(--text);
        font-family: 'IBM Plex Sans', sans-serif;
        font-size: 14px;
        margin: 0;
        padding: 0;
      }

      /* ── Header ─────────────────────────────────────────── */
      .app-header {
        background: #b6bfc1;
        border-bottom: 0px solid var(--border);
        padding: 28px 40px 10px;
        display: flex;
        align-items: center;
        justify-content: left;
        gap: 16px;
        flex-direction: row;
      }
      .app-title {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 22px;
        font-weight: 600;
        color: var(--accent);
        letter-spacing: -0.5px;
        margin: 0 0 4px;
      }
      .app-subtitle {
        color: var(--text-muted);
        font-size: 13px;
        margin: 0;
      }

      /* ── Main body ─────────────────────────────────── */
      .main-body {
        padding: 32px 40px;
        max-width: 1100px;
      }

      /* ── Control section ───────────────────────────────── */
      .control-card {
        background: var(--surface);
        border: 0px solid var(--border);
        border-radius: var(--radius);
        padding: 24px 28px;
        margin-bottom: 28px;
        display: flex;
        align-items: flex-end;
        gap: 24px;
        flex-wrap: wrap;
        box-shadow: 0 4px 24px rgba(0, 0, 0, 0.4);
      }
      .ctrl-group {
        display: flex;
        flex-direction: column;
        gap: 6px;
      }
      .ctrl-label {
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        color: var(--text-muted);
      }

      /* ── Select custom ───────────────────────────────────── */
      .shiny-input-container {
        margin-bottom: 0 !important;
      }
      select.form-control {
        height: 40px;
        background: #051230 !important;
        color: var(--text) !important;
        border: 1px solid var(--border) !important;
        border-radius: var(--radius) !important;
        font-family: 'IBM Plex Mono', monospace !important;
        font-size: 13px !important;
        padding: 10px 14px !important;
        min-width: 220px;
        appearance: none;
        background-image: url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%238892a4' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E\") !important;
        background-repeat: no-repeat !important;
        background-position: right 14px center !important;
        cursor: pointer;
        transition: border-color .2s;
      }
      select.form-control:focus {
        border-color: var(--accent) !important;
        outline: none !important;
        box-shadow: 0 0 0 3px rgba(79,142,247,.15) !important;
      }

      /* ── Download button ───────────────────────────── */
      .btn-download {
        height: 40px;
        background: #051230 !important;
        color: #b6bfc1 !important;
        border: none !important;
        border-radius: var(--radius) !important;
        font-family: 'IBM Plex Sans', sans-serif !important;
        font-weight: 600 !important;
        font-size: 13px !important;
        padding: 11px 22px !important;
        cursor: pointer;
        transition: opacity .2s, transform .15s;
        display: flex;
        align-items: center;
        gap: 8px;
        white-space: nowrap;
      }
      .btn-download:hover  { opacity: .88; transform: translateY(-1px); }
      .btn-download:active { transform: translateY(0); opacity: 1; }

      /* ── Info lines badge ────────────────────────────────── */
      .info-badge {
        height: 40px;
        font-family: 'IBM Plex Mono', monospace;
        font-size: 12px;
        color: #b6bfc1;
        background: #051230;
        border: 0px solid #051230;
        border-radius: 20px;
        padding: 5px 14px;
        white-space: nowrap;
        display: flex;
        align-items: center;
        justify-content: center;
      }

      /* ── Table section ───────────────────────────────────── */
      .table-card {
        background: var(--surface);
        border: 0px solid var(--border);
        border-radius: var(--radius);
        overflow: hidden;
        box-shadow: 0 4px 24px rgba(0, 0, 0, 0.4);
      }
      .table-card-header {
        background:#051230;
        padding: 16px 24px;
        border-bottom: 1px solid var(--border);
        display: flex;
        justify-content: space-between;
        align-items: center;
      }
      .table-card-title {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 13px;
        font-weight: 600;
        color: #b6bfc1;
      }

      /* ── DT overrides ────────────────────────────────────── */
      .dataTables_wrapper { padding: 0 !important; }
      table.dataTable {
        background: transparent !important;
        color: var(--text) !important;
        border: none !important;
        font-family: 'IBM Plex Sans', sans-serif;
        font-size: 13px;
      }
      table.dataTable thead th {
        background: #97acc8 !important;
        color: #051230 !important;
        font-size: 11px !important;
        font-weight: 600 !important;
        text-transform: uppercase;
        letter-spacing: .6px;
        border-bottom: 1px solid var(--border) !important;
        border-top: none !important;
      }
      table.dataTable tbody tr {
        background: transparent !important;
        border-bottom: 1px solid var(--border) !important;
      }
      table.dataTable tbody tr:hover td { background: rgba(79,142,247,.05) !important; }
      table.dataTable tbody td {
        padding: 11px 16px !important;
        border: none !important;
        max-height: 50px;
        overflow: hidden;
        white-space: nowrap;
        text-overflow: ellipsis;
      }
      
      .dataTables_info, .dataTables_length, .dataTables_filter, .dataTables_paginate {
        color: var(--text-muted) !important;
        font-size: 12px !important;
        padding: 12px 24px !important;
      }
      .dataTables_filter input, .dataTables_length select {
        background: #b6bfc1 !important;
        color: var(--text) !important;
        border: 1px solid var(--border) !important;
        border-radius: 6px !important;
        padding: 4px 10px !important;
        font-size: 12px !important;
      }
      .paginate_button {
        color: var(--text-muted) !important;
        border-radius: 6px !important;
        border: none !important;
      }
      .paginate_button.current, .paginate_button:hover {
        background: rgba(79,142,247,.15) !important;
        color: var(--accent) !important;
        border: none !important;
      }
    "))
  ),
  
  # Header
  div(class = "app-header",
      div(
        p(class = "app-title",    "codelist2Excel"),
        p(class = "app-subtitle", "Select a codelist, visualize it and export it into Excel.")
      )
  ),
  
  # Body
  div(class = "main-body",
      
      # ── Control section ──────────────────────────
      div(class = "control-card",
          
          div(class = "ctrl-group",
              div(class = "ctrl-label", "Select your codelist"),
              selectInput(
                  inputId  = "selected_codelist",
                  label    = NULL,
                  choices  = c(" " = ""),
                  selected = "",
                  width    = "240px"
                )
              ),
          
          div(class = "ctrl-group",
              div(class = "ctrl-label", "Export"),
              downloadButton(
                outputId = "btn_excel",
                label    = "Download as Excel",
                class    = "btn-download",
                icon     = icon("file-excel")
              )
          ),
          
          uiOutput("info_badge")
      ),
      
      # ── Table section ───────────────────────────────
      div(class = "table-card",
          div(class = "table-card-header",
              div(class = "table-card-title", textOutput("table_title", inline = TRUE))
          ),
          DTOutput("table")
      )
  )
)

# -------------------------------------------------------------------------
# Server
# -------------------------------------------------------------------------

server <- function(input, output, session) {

  codelists_r <- reactiveVal(NULL)

  user <- reactiveVal(NULL)

  observeEvent(TRUE, {
    tryCatch({
      initialiseClient(session = session, sws_endpoint = "https://sws.fao.org")
      user(getCurrentUser())
      
      codelists <- getAllCodelists()
      updateSelectInput(session, "selected_codelist",
                        choices = c(" " = "", codelists$id))
      
    }, error = function(e) {
      showNotification(paste0("Initialization failed: ", e$message), type = "error")
    })
  }, once = TRUE, ignoreNULL = FALSE)
  
  # Reactive data in function of the dropdown
  selected_dt <- reactive({
    req(input$selected_codelist)
    get_codelist_info(input$selected_codelist)
  })
  
  # # Get parent-children data
  # pc_data <- reactive({
  #   req(input$selected_codelist)
  #   dt <- selected_dt()
  #   dt[, .(children = unlist(strsplit(as.character(children), ", "))), by = code]
  # })
  
  # Badge with number of rows/columns
  output$info_badge <- renderUI({
    req(input$selected_codelist)
    dt <- selected_dt()
    div(class = "ctrl-group",
        div(class = "ctrl-label", "\u00a0"),
        div(class = "info-badge",
            sprintf("%d lines · %d columns", nrow(dt), ncol(dt))
        )
    )
  })
  
  # Title of the chart
  output$table_title <- renderText({
    req(input$selected_codelist)
    input$selected_codelist
  })
  
  # Excel output
  output$table <- renderDT({
    dt <- selected_dt()
    
    idx_cols <- as.vector(which(sapply(dt, is.character)) - 1)
    
    col_defs <- if (length(idx_cols) > 0) {
      lapply(idx_cols, function(i) list(
        targets = i,
        render  = JS("function(data, type, row, meta) {
      if (type === 'display' && data && data.length > 50) {
        return '<span title=\"' + data + '\">' + data.substr(0, 50) + '...</span>';
      }
      return data;
    }")
      ))
    } else {
      list()
    }
    
    datatable(
      dt,
      options = list(
        pageLength = 10,
        dom        = "frtip",
        scrollX    = TRUE,
        columnDefs = col_defs
      ),
      rownames = FALSE,
      class    = "display"
    )
  })
  
  # Download codelist as Excel
  output$btn_excel <- downloadHandler(
    filename = function() {
      paste0(input$selected_codelist, "_", format(Sys.Date(), "%Y%m%d"), ".xlsx")
    },
    content = function(file) {
      write_xlsx(
        list(
          "Codelist"        = selected_dt()
          # "Parent-Children" = pc_data()
        ),
        path = file
      )
    }
  )
  
}

# -------------------------------------------------------------------------
# Launching
# -------------------------------------------------------------------------

app <- shinyApp(ui = ui, server = server)

tryCatch({
  runApp(app, host = '0.0.0.0', port = 8080)
}, error = function(e) {
  log_msg(paste("ERROR: Failed to start app -", e$message))
  stop(e)
})
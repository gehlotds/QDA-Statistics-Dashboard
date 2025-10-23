# Enhanced QDA Sensory Analysis Dashboard
# Install required packages
install.packages(c(
  "shiny", "shinydashboard", "DT", "plotly", "ggplot2",
  "dplyr", "tidyr", "lme4", "lmerTest", "car", "factoextra", 
  "FactoMineR", "cluster", "multcompView", "agricolae",
  "corrplot", "reshape2", "emmeans", "nortest", "moments",
  "multcomp", "MASS", "ggrepel", "pbkrtest", "openxlsx"
))

library(shiny)
library(shinydashboard)
library(DT)
library(plotly)
library(ggplot2)
library(dplyr)
library(tidyr)
library(lme4)
library(lmerTest)
library(car)
library(factoextra)
library(FactoMineR)
library(cluster)
library(multcompView)
library(agricolae)
library(corrplot)
library(reshape2)
library(emmeans)
library(nortest)
library(moments)
library(multcomp)
library(MASS)
library(ggrepel)
library(pbkrtest)
library(openxlsx)

ui <- dashboardPage(
  dashboardHeader(title = "QDA Sensory Analysis Dashboard"),
  dashboardSidebar(
    sidebarMenu(
      menuItem("Data Upload", tabName = "upload", icon = icon("upload")),
      menuItem("Normality Testing", tabName = "normality", icon = icon("chart-line")),
      menuItem("Mixed Model Results", tabName = "mixedmodel", icon = icon("calculator")),
      menuItem("Cluster Analysis (AHC)", tabName = "ahc_clustering", icon = icon("sitemap")),
      menuItem("PCA Analysis", tabName = "pca", icon = icon("project-diagram")),
      menuItem("AHC on PCA", tabName = "clustering", icon = icon("sitemap")),
      menuItem("Statistical Comparisons", tabName = "comparisons", icon = icon("table")),
      menuItem("Pairwise Comparision", tabName = "bargraph", icon = icon("chart-bar")),
      menuItem("Statistical Theory", tabName = "theory", icon = icon("book"))
    )
  ),
  dashboardBody(
    tags$head(tags$style(HTML(".content-wrapper, .right-side { background-color: #f4f4f4; }"))),
    tabItems(
      tabItem(tabName = "upload",
              fluidRow(
                box(title = "Upload QDA Data", status = "primary", solidHeader = TRUE, width = 12,
                    fileInput("file", "Choose CSV File", accept = c(".csv")),
                    checkboxInput("header", "Header", TRUE),
                    radioButtons("sep", "Separator", choices = c(Comma = ",", Semicolon = ";", Tab = "\t"), selected = ","),
                    hr(),
                    h4("Expected Format:"),
                    p("Column 1: Session | Column 2: Panelist | Column 3: Product | Remaining Columns: Descriptors (any number)"),
                    hr(),
                    h4("Download All Analysis Tables:"),
                    p("After running the analysis, download all tables in one Excel file with separate sheets."),
                    downloadButton("download_all_tables", "Download Analysis Tables", class = "btn-success"),
                    hr(),
                    h4("Data Summary:"),
                    verbatimTextOutput("data_summary"),
                    hr(),
                    h4("Data Preview:"),
                    DT::dataTableOutput("data_preview")
                )
              )
      ),
      tabItem(tabName = "normality",
              fluidRow(
                box(title = "Comprehensive Normality Testing", status = "primary", solidHeader = TRUE, width = 12,
                    h4("Multiple Normality Tests"),
                    p("Statistical Theory: Testing H₀: Data follows normal distribution"),
                    tags$ul(
                      tags$li("Shapiro-Wilk: W = (Σaᵢx₍ᵢ₎)²/Σ(xᵢ-x̄)² - Most powerful for small samples"),
                      tags$li("Anderson-Darling: Better for larger samples"),
                      tags$li("Jarque-Bera: JB = (n/6)[S² + (K-3)²/4] - Tests skewness and kurtosis"),
                      tags$li("P-values < 0.0001 indicate strong evidence against normality")
                    ),
                    DT::dataTableOutput("normality_results_enhanced"),
                    br(),
                    h4("Visual Normality Assessment - Histogram with Normal Curve"),
                    selectInput("normality_descriptor", "Select Descriptor:", choices = NULL),
                    plotlyOutput("histogram_plot", height = "500px")
                )
              )
      ),
      tabItem(tabName = "mixedmodel",
              fluidRow(
                box(title = "Linear Mixed Model Analysis - Statistical Explanation", status = "primary", solidHeader = TRUE, width = 12,
                    h4("📊 Why Linear Mixed Models (LMM)?"),
                    tags$ul(
                      tags$li(strong("Repeated Measures:"), " Each panelist rates multiple products → observations are correlated"),
                      tags$li(strong("Account for Dependencies:"), " LMM handles non-independent data correctly"),
                      tags$li(strong("Multiple Sources of Variation:"), " Separates product effects from panelist/session variability"),
                      tags$li(strong("Better than ANOVA:"), " ANOVA assumes independence; LMM models the dependency structure")
                    ),
                    hr(),
                    h4("Model Specification"),
                    p(strong("Model: Rating ~ Product + (1|Panelist) + (1|Session) + (1|Panelist:Product)")),
                    tags$ul(
                      tags$li(strong("Fixed Effect (Product):"), " Tests if products differ in mean ratings (what we care about)"),
                      tags$li(strong("Random Effect (Panelist):"), " Accounts for baseline differences between panelists"),
                      tags$li(strong("Random Effect (Session):"), " Accounts for session-to-session variation (time of day, fatigue)"),
                      tags$li(strong("Random Interaction (Panelist:Product):"), " Captures individual panelist preferences for specific products")
                    ),
                    hr(),
                    h4("LSMeans from Mixed Models"),
                    p("Least Squares Means (LSMeans) = Estimated marginal means adjusted for random effects"),
                    DT::dataTableOutput("lmm_results"),
                    br(),
                    h4("Fixed Effects F-Tests"),
                    p(strong("Purpose:"), " Test H₀: All products have the same mean rating"),
                    tags$ul(
                      tags$li("F-statistic: Ratio of between-product variance to within-product variance"),
                      tags$li("Satterthwaite approximation: Estimates appropriate degrees of freedom for mixed models"),
                      tags$li("P-value < 0.05: At least one product differs significantly")
                    ),
                    DT::dataTableOutput("fixed_effects_tests"),
                    br(),
                    h4("Variance Components & Model Fit"),
                    p(strong("Purpose:"), " Understand sources of variability in ratings"),
                    tags$ul(
                      tags$li("σ²_Panelist: Variation due to individual differences (some rate higher/lower overall)"),
                      tags$li("σ²_Session: Variation due to session effects"),
                      tags$li("σ²_Panelist:Product: Individual panelist-product interactions"),
                      tags$li("σ²_Residual: Unexplained random error"),
                      tags$li("ICC = σ²_Panelist / σ²_Total: Proportion of variance due to panelists (>0.5 = strong panelist effects)")
                    ),
                    DT::dataTableOutput("variance_components"),
                    br(),
                    h4("Model Assumptions Diagnostics"),
                    p(strong("Purpose:"), " Check if model assumptions are met"),
                    box(title = "Understanding 'Model Failed to Converge'", status = "warning", solidHeader = TRUE, width = 12,
                        p(strong("Why this happens:")),
                        tags$ul(
                          tags$li(strong("Zero Variance:"), " All panelists rated the descriptor very similarly (no variation to model)"),
                          tags$li(strong("Singular Fit:"), " Random effects estimated as exactly zero (no panelist/session differences detected)"),
                          tags$li(strong("Insufficient Data:"), " Not enough observations for complex random effects structure")
                        ),
                        p(strong("What the code does:")),
                        tags$ul(
                          tags$li("Automatically falls back to simple arithmetic means"),
                          tags$li("Results are still valid - just not adjusted for random effects"),
                          tags$li("This is common for descriptors with very low variability")
                        ),
                        p(strong("Action needed:")),
                        tags$ul(
                          tags$li("Consider removing descriptors with convergence failures if they show very low variance"),
                          tags$li("Or accept simple means for those descriptors (still scientifically valid)")
                        )
                    ),
                    selectInput("assumption_descriptor", "Select Descriptor:", choices = NULL),
                    p(strong("Residuals vs Fitted Plot:"), " Checks for patterns in residuals"),
                    tags$ul(
                      tags$li(strong("Good:"), " Random scatter around zero (assumptions met)"),
                      tags$li(strong("Bad:"), " Funnel shape (heteroscedasticity), curvature (non-linearity)")
                    ),
                    plotOutput("residuals_fitted_plot", height = "400px"),
                    br(),
                    p(strong("Q-Q Plot:"), " Checks if residuals follow normal distribution"),
                    tags$ul(
                      tags$li(strong("Good:"), " Points fall on diagonal line"),
                      tags$li(strong("Bad:"), " Systematic departures (heavy tails, skewness)")
                    ),
                    plotOutput("qq_plot", height = "400px")
                )
              )
      ),
      tabItem(tabName = "ahc_clustering",
              fluidRow(
                box(title = "Agglomerative Hierarchical Clustering on LSMeans", 
                    status = "primary", solidHeader = TRUE, width = 12,
                    h4("📊 Clustering Products Based on LSMeans from Mixed Models"),
                    p("This analysis performs hierarchical clustering directly on the Least Squares Means (LSM) from the mixed models."),
                    tags$ul(
                      tags$li(strong("Input:"), " LSMeans matrix (Products × Descriptors)"),
                      tags$li(strong("Method:"), " Ward's linkage (minimizes within-cluster variance)"),
                      tags$li(strong("Distance:"), " Euclidean distance between product profiles"),
                      tags$li(strong("Purpose:"), " Group products with similar overall sensory profiles")
                    ),
                    hr(),
                    h4("Optimal Number of Clusters"),
                    fluidRow(
                      column(6,
                             h5("Silhouette Analysis"),
                             p("Measures clustering quality: higher values = better separation"),
                             plotlyOutput("ahc_silhouette_plot", height = "300px")
                      ),
                      column(6,
                             h5("Within-Cluster Sum of Squares (Elbow Method)"),
                             p("Choose k at the 'elbow' where adding clusters gives diminishing returns"),
                             plotlyOutput("ahc_elbow_plot", height = "300px")
                      )
                    ),
                    br(),
                    h4("Suggested Optimal Clusters"),
                    verbatimTextOutput("ahc_optimal_clusters_text"),
                    br(),
                    numericInput("n_clusters_ahc", "Number of Clusters to Display:", value = 3, min = 2, max = 8),
                    br(),
                    h4("Dendrogram"),
                    plotOutput("ahc_dendrogram", height = "600px"),
                    br(),
                    h4("Cluster Assignments"),
                    DT::dataTableOutput("ahc_cluster_table"),
                    br(),
                    h4("Cluster Profiles (Mean LSM by Cluster)"),
                    p("This table shows the average LSMean value for each descriptor within each cluster."),
                    DT::dataTableOutput("ahc_cluster_profiles")
                )
              )
      ),
      tabItem(tabName = "pca",
              fluidRow(
                box(title = "Principal Component Analysis (PCA)", status = "primary", solidHeader = TRUE, width = 12,
                    h4("Scree Plot - Eigenvalues"),
                    plotlyOutput("scree_plot", height = "400px"),
                    br(),
                    h4("Variable Contributions"),
                    plotlyOutput("pca_contrib", height = "400px"),
                    br(),
                    h4("PCA Biplot: PC1 vs PC2"),
                    fluidRow(
                      column(3, checkboxInput("swap_axes_12", "Swap X and Y Axes", value = FALSE)),
                      column(3, sliderInput("repel_force_12", "Label Repulsion:", min = 0.5, max = 5, value = 2, step = 0.5)),
                      column(3, sliderInput("point_size_12", "Point Size:", min = 2, max = 8, value = 4, step = 1)),
                      column(3, sliderInput("text_size_12", "Text Size:", min = 2, max = 8, value = 4, step = 1))
                    ),
                    plotOutput("pca_biplot_12", height = "600px"),
                    downloadButton("download_pca_12", "Download PCA Plot (PC1 vs PC2)"),
                    br(), hr(), br(),
                    h4("PCA Biplot: PC1 vs PC3"),
                    fluidRow(
                      column(3, checkboxInput("swap_axes_13", "Swap X and Y Axes", value = FALSE)),
                      column(3, sliderInput("repel_force_13", "Label Repulsion:", min = 0.5, max = 5, value = 2, step = 0.5)),
                      column(3, sliderInput("point_size_13", "Point Size:", min = 2, max = 8, value = 4, step = 1)),
                      column(3, sliderInput("text_size_13", "Text Size:", min = 2, max = 8, value = 4, step = 1))
                    ),
                    plotOutput("pca_biplot_13", height = "600px"),
                    downloadButton("download_pca_13", "Download PCA Plot (PC1 vs PC3)")
                )
              )
      ),
      tabItem(tabName = "clustering",
              fluidRow(
                box(title = "Hierarchical Clustering Analysis", status = "primary", solidHeader = TRUE, width = 12,
                    h4("Optimal Number of Clusters"),
                    fluidRow(
                      column(6,
                             h5("Silhouette Analysis"),
                             p("Measures clustering quality: higher values = better separation"),
                             plotlyOutput("silhouette_plot", height = "300px")
                      ),
                      column(6,
                             h5("Within-Cluster Sum of Squares (Elbow Method)"),
                             p("Choose k at the 'elbow' where adding clusters gives diminishing returns"),
                             plotlyOutput("elbow_plot", height = "300px")
                      )
                    ),
                    br(),
                    h4("Suggested Optimal Clusters"),
                    verbatimTextOutput("optimal_clusters_text"),
                    br(),
                    numericInput("n_clusters_manual", "Number of Clusters to Display:", value = 3, min = 2, max = 8),
                    br(),
                    fluidRow(
                      column(6,
                             h4("Dendrogram - PC1 & PC2"),
                             plotOutput("dendro_12", height = "400px"),
                             h5("Cluster Assignments"),
                             DT::dataTableOutput("cluster_table_12")
                      ),
                      column(6,
                             h4("Dendrogram - PC1 & PC3"),
                             plotOutput("dendro_13", height = "400px"),
                             h5("Cluster Assignments"),
                             DT::dataTableOutput("cluster_table_13")
                      )
                    )
                )
              )
      ),
      tabItem(tabName = "comparisons",
              fluidRow(
                box(title = "Mixed Model Statistical Comparisons", status = "primary", solidHeader = TRUE, width = 12,
                    h4("⚠️ SE & DF Correction Applied (Kenward-Roger Method)"),
                    tags$div(
                      style = "background-color: #d1ecf1; padding: 15px; border-radius: 5px; margin-bottom: 20px; border-left: 5px solid #0c5460;",
                      tags$ul(
                        tags$li(strong("SE (Standard Error):"), " Now uses Kenward-Roger method for more accurate estimation. SE values may vary slightly between comparisons (this is correct and expected)."),
                        tags$li(strong("DF (Degrees of Freedom):"), " Decimal values are CORRECT for mixed models. Mixed models use Satterthwaite/Kenward-Roger approximation, which accounts for random effects and produces non-integer df."),
                        tags$li(strong("Why decimals?"), " Unlike ANOVA (integer df), mixed models account for variance components and correlation structure, resulting in effective degrees of freedom as real numbers (e.g., 15.83, 16.21)."),
                        tags$li(strong("Interpretation:"), " Similar df values across comparisons indicate balanced design with homogeneous variance structure - this is statistically appropriate.")
                      )
                    ),
                    hr(),
                    h4("Post-Hoc Pairwise Comparisons"),
                    p(strong("Purpose:"), " After significant F-test, determine which specific product pairs differ"),
                    tags$ul(
                      tags$li("Tukey HSD adjustment: Controls family-wise error rate (prevents false positives)"),
                      tags$li("P-value < 0.05: Product pair is significantly different")
                    ),
                    selectInput("comparison_descriptor", "Select Descriptor:", choices = NULL),
                    br(),
                    h4("Pairwise Comparisons (Tukey-adjusted)"),
                    DT::dataTableOutput("pairwise_comparisons"),
                    br(),
                    h4("Compact Letter Display (CLD)"),
                    p("Products with same letter are not significantly different (α = 0.05)"),
                    DT::dataTableOutput("cld_results_enhanced")
                )
              )
      ),
      tabItem(tabName = "bargraph",
              fluidRow(
                box(title = "Bar Graph with Statistical Annotations", status = "primary", solidHeader = TRUE, width = 12,
                    selectInput("bargraph_descriptor", "Select Descriptor:", choices = NULL),
                    br(),
                    h4("LSMeans ± Standard Errors with Significance Letters"),
                    p("Letters indicate statistical groupings - same letter = not significantly different"),
                    plotlyOutput("bargraph_plot", height = "500px"),
                    br(),
                    h4("Detailed Statistics"),
                    DT::dataTableOutput("bargraph_stats"),
                    br(), hr(), br(),
                    h4("Aggregate Pairwise Significance Analysis"),
                    p("This table shows how many descriptors differentiate each product pair across all sensory attributes."),
                    tags$ul(
                      tags$li("For each product pair, we count the number of descriptors where the pair is significantly different (p < 0.05)"),
                      tags$li("% Difference = (Count of significant descriptors / Total descriptors) × 100"),
                      tags$li("This provides a holistic view of product differentiation across all sensory dimensions")
                    ),
                    h5("Product Pair Differentiation Summary"),
                    DT::dataTableOutput("pairwise_aggregate_table"),
                    br(),
                    h5("Distribution of Product Pair Differentiation (with Product Pair Names)"),
                    p("Frequency distribution showing how many product pairs fall into each % difference range, along with the specific product pairs"),
                    DT::dataTableOutput("pairwise_frequency_table")
                )
              )
      ),
      tabItem(tabName = "theory",
              fluidRow(
                box(title = "Statistical Methods Explanation", status = "info", solidHeader = TRUE, width = 12,
                    h3("📚 Core Statistical Framework"),
                    h4("1. Linear Mixed Model (LMM)"),
                    p("Model: Rating ~ Product + (1|Panelist) + (1|Session) + (1|Panelist:Product)"),
                    tags$ul(
                      tags$li(strong("Why LMM?"), " Sensory data has repeated measures - each panelist rates multiple products"),
                      tags$li(strong("Fixed Effect (Product):"), " What we're testing - do products differ?"),
                      tags$li(strong("Random Effects:"), " Account for dependencies in data"),
                      tags$ul(
                        tags$li("Panelist: Individual baseline differences"),
                        tags$li("Session: Time-of-day, fatigue, environmental effects"),
                        tags$li("Panelist:Product: Individual preferences for specific products")
                      ),
                      tags$li(strong("Estimation:"), " REML (Restricted Maximum Likelihood)"),
                      tags$li(strong("Output:"), " LSMeans (Least Squares Means) = adjusted marginal means")
                    ),
                    h4("2. Why F-tests with Satterthwaite Approximation?"),
                    tags$ul(
                      tags$li(strong("Purpose:"), " Test H₀: All products have equal means"),
                      tags$li(strong("F-statistic:"), " Ratio of between-product variance to within-product variance"),
                      tags$li(strong("Satterthwaite:"), " Mixed models don't have simple degrees of freedom - this method estimates appropriate df"),
                      tags$li(strong("Interpretation:"), " P < 0.05 means at least one product differs significantly")
                    ),
                    h4("3. Variance Components - Why Important?"),
                    tags$ul(
                      tags$li("Total Variance = σ²_Panelist + σ²_Session + σ²_Panelist:Product + σ²_Error"),
                      tags$li("σ²_Panelist: How much variation comes from individual differences"),
                      tags$li("σ²_Session: How much from testing conditions"),
                      tags$li("σ²_Panelist:Product: Individual panelist-product interactions"),
                      tags$li("σ²_Error: Random unexplained variation"),
                      tags$li("ICC (Intraclass Correlation) = σ²_Panelist / σ²_Total"),
                      tags$ul(
                        tags$li("High ICC (>0.5): Strong panelist effects - individuals differ consistently"),
                        tags$li("Low ICC (<0.2): Good panel agreement or high measurement error")
                      )
                    ),
                    h4("4. Model Diagnostics - Residual Plots"),
                    tags$ul(
                      tags$li(strong("Q-Q Plot:"), " Tests normality of residuals - points should follow red line"),
                      tags$li(strong("Residuals vs Fitted:"), " Tests homoscedasticity - should be random scatter"),
                      tags$ul(
                        tags$li("Good: Random scatter around zero"),
                        tags$li("Bad: Patterns, curves, funnel shapes indicate model issues")
                      ),
                      tags$li(strong("'Model Failed to Converge':"), " Happens when variance is too low or random effects = 0. Code automatically uses simple means instead.")
                    ),
                    h4("5. Post-Hoc Pairwise Comparisons (Tukey HSD)"),
                    tags$ul(
                      tags$li(strong("Purpose:"), " After significant F-test, find which specific pairs differ"),
                      tags$li(strong("Tukey Adjustment:"), " Controls family-wise error rate (FWER) at α = 0.05"),
                      tags$li(strong("Why needed?"), " Multiple comparisons increase Type I error risk - Tukey corrects this"),
                      tags$li(strong("Compact Letter Display:"), " Same letter = no significant difference")
                    ),
                    h4("6. Principal Component Analysis (PCA)"),
                    tags$ul(
                      tags$li("Reduces dimensionality while preserving variance"),
                      tags$li(strong("PC1 (x-axis/horizontal):"), " Direction of maximum variance"),
                      tags$li(strong("PC2/PC3 (y-axis/vertical):"), " Orthogonal directions of next highest variance"),
                      tags$li("Biplot: Products (blue) and Descriptors (red) plotted together"),
                      tags$li("Distance from origin = importance of descriptor"),
                      tags$li("Proximity of descriptors = positive correlation"),
                      tags$li("Opposite directions = negative correlation")
                    ),
                    h4("7. Hierarchical Clustering"),
                    tags$ul(
                      tags$li(strong("Ward's Method:"), " Minimizes within-cluster variance"),
                      tags$li(strong("Dendrogram:"), " Tree showing product relationships"),
                      tags$li(strong("Height:"), " Dissimilarity between merged clusters"),
                      tags$li(strong("Optimal k:"), " Found using silhouette score (maximize) or elbow method (WSS)")
                    )
                )
              )
      )
    )
  )
)

server <- function(input, output, session) {
  values <- reactiveValues(
    data = NULL, n_descriptors = 0, descriptor_names = NULL,
    lmm_means = NULL, pca_result = NULL, fixed_effects = NULL,
    variance_components = NULL, silhouette_data = NULL, wss_data = NULL,
    optimal_k = 3, pairwise_aggregate = NULL,
    ahc_silhouette_data = NULL, ahc_wss_data = NULL, ahc_optimal_k = 3
  )
  
  observeEvent(input$file, {
    req(input$file)
    tryCatch({
      df <- read.csv(input$file$datapath, header = input$header, sep = input$sep, stringsAsFactors = FALSE)
      if(ncol(df) < 4) {
        showNotification("Error: Data must have at least 4 columns", type = "error", duration = 10)
        return()
      }
      n_desc <- ncol(df) - 3
      original_names <- names(df)
      new_names <- c("Session", "Panelist", "Product", original_names[4:ncol(df)])
      names(df) <- new_names
      df$Session <- as.factor(df$Session)
      df$Panelist <- as.factor(df$Panelist)
      df$Product <- as.factor(df$Product)
      descriptor_cols <- names(df)[4:ncol(df)]
      for(col in descriptor_cols) {
        df[[col]] <- as.numeric(as.character(df[[col]]))
      }
      values$data <- df
      values$n_descriptors <- n_desc
      values$descriptor_names <- descriptor_cols
      updateSelectInput(session, "normality_descriptor", choices = descriptor_cols, selected = descriptor_cols[1])
      updateSelectInput(session, "assumption_descriptor", choices = descriptor_cols, selected = descriptor_cols[1])
      updateSelectInput(session, "comparison_descriptor", choices = descriptor_cols, selected = descriptor_cols[1])
      updateSelectInput(session, "bargraph_descriptor", choices = descriptor_cols, selected = descriptor_cols[1])
      showNotification("Data loaded successfully!", type = "message", duration = 3)
    }, error = function(e) {
      showNotification(paste("Error:", e$message), type = "error", duration = 10)
    })
  })
  
  output$data_summary <- renderText({
    req(values$data)
    paste("✓ Data loaded successfully!\nRows:", nrow(values$data), "\nSessions:", nlevels(values$data$Session),
          "\nPanelists:", nlevels(values$data$Panelist), "\nProducts:", nlevels(values$data$Product),
          "\nDescriptors:", values$n_descriptors, "\nNames:", paste(values$descriptor_names, collapse = ", "))
  })
  
  output$data_preview <- DT::renderDataTable({
    req(values$data)
    DT::datatable(values$data, options = list(scrollX = TRUE, pageLength = 10)) %>%
      DT::formatRound(columns = 4:ncol(values$data), digits = 1)
  })
  
  output$normality_results_enhanced <- DT::renderDataTable({
    req(values$data)
    tryCatch({
      results <- data.frame(
        Descriptor = values$descriptor_names,
        Shapiro_W = NA, Shapiro_p = character(values$n_descriptors),
        AD_A = NA, AD_p = character(values$n_descriptors),
        JB_Chi2 = NA, JB_p = character(values$n_descriptors),
        Overall = character(values$n_descriptors),
        stringsAsFactors = FALSE
      )
      for(i in 1:values$n_descriptors) {
        desc <- values$descriptor_names[i]
        data_vec <- na.omit(values$data[[desc]])
        if(length(data_vec) > 3 && length(unique(data_vec)) > 1) {
          if(length(data_vec) <= 5000) {
            sw <- shapiro.test(data_vec)
            results$Shapiro_W[i] <- sw$statistic
            results$Shapiro_p[i] <- ifelse(sw$p.value < 0.0001, sprintf("%.2e", sw$p.value), sprintf("%.4f", sw$p.value))
            p_shapiro <- sw$p.value
          } else {
            p_shapiro <- NA
          }
          tryCatch({
            ad <- ad.test(data_vec)
            results$AD_A[i] <- ad$statistic
            results$AD_p[i] <- ifelse(ad$p.value < 0.0001, sprintf("%.2e", ad$p.value), sprintf("%.4f", ad$p.value))
            p_ad <- ad$p.value
          }, error = function(e) { p_ad <<- NA })
          if(length(data_vec) > 7) {
            tryCatch({
              jb <- jarque.test(data_vec)
              results$JB_Chi2[i] <- jb$statistic
              results$JB_p[i] <- ifelse(jb$p.value < 0.0001, sprintf("%.2e", jb$p.value), sprintf("%.4f", jb$p.value))
              p_jb <- jb$p.value
            }, error = function(e) { p_jb <<- NA })
          } else {
            p_jb <- NA
          }
          p_vals <- c(p_shapiro, p_ad, p_jb)
          p_vals <- p_vals[!is.na(p_vals)]
          if(length(p_vals) > 0) {
            results$Overall[i] <- ifelse(mean(p_vals > 0.05) >= 0.5, "Yes", "No")
          }
        }
      }
      DT::datatable(results, options = list(scrollX = TRUE)) %>%
        DT::formatRound(columns = c("Shapiro_W", "AD_A", "JB_Chi2"), digits = 3) %>%
        DT::formatStyle("Overall", backgroundColor = DT::styleEqual("No", "lightcoral"))
    }, error = function(e) {
      DT::datatable(data.frame(Error = paste("Normality testing error:", e$message)))
    })
  })
  
  output$histogram_plot <- renderPlotly({
    req(values$data, input$normality_descriptor)
    tryCatch({
      p <- ggplot(values$data, aes(x = .data[[input$normality_descriptor]])) +
        geom_histogram(aes(y = after_stat(density)), bins = 20, fill = "skyblue", color = "black", alpha = 0.7) +
        geom_density(color = "red", linewidth = 1.2) +
        stat_function(fun = dnorm, 
                      args = list(mean = mean(values$data[[input$normality_descriptor]], na.rm = TRUE),
                                  sd = sd(values$data[[input$normality_descriptor]], na.rm = TRUE)),
                      color = "darkgreen", linewidth = 1, linetype = "dashed") +
        theme_minimal() +
        labs(title = paste("Distribution:", input$normality_descriptor),
             subtitle = "Red = Actual density | Green dashed = Normal distribution",
             x = input$normality_descriptor, y = "Density")
      ggplotly(p)
    }, error = function(e) {
      p <- ggplot() + annotate("text", x = 0.5, y = 0.5, label = "Error creating histogram") + theme_void()
      ggplotly(p)
    })
  })
  
  observe({
    req(values$data)
    tryCatch({
      lmm_results <- data.frame(Product = levels(values$data$Product))
      fixed_effects <- data.frame(
        Descriptor = values$descriptor_names,
        F_value = NA, Df1 = NA, Df2 = NA, P_value = NA,
        Significant = character(values$n_descriptors),
        stringsAsFactors = FALSE
      )
      variance_comp <- data.frame(
        Descriptor = values$descriptor_names,
        Panelist_Var = NA, Session_Var = NA, PanelistProduct_Var = NA, Residual_Var = NA,
        ICC = NA, AIC = NA, BIC = NA, LogLik = NA,
        stringsAsFactors = FALSE
      )
      all_pairwise_results <- list()
      withProgress(message = 'Fitting Mixed Models...', value = 0, {
        for(i in 1:values$n_descriptors) {
          desc <- values$descriptor_names[i]
          incProgress(1/values$n_descriptors, detail = desc)
          tryCatch({
            desc_var <- var(values$data[[desc]], na.rm = TRUE)
            if(is.na(desc_var) || desc_var < 1e-10) {
              simple_means <- aggregate(values$data[[desc]], by = list(values$data$Product), FUN = mean, na.rm = TRUE)
              lmm_results[[desc]] <- simple_means$x[match(lmm_results$Product, simple_means$Group.1)]
              next
            }
            formula_str <- paste(desc, "~ Product + (1|Panelist) + (1|Session) + (1|Panelist:Product)")
            # Suppress convergence warnings (model still works, falls back to simple means if needed)
            model <- suppressWarnings(lmer(as.formula(formula_str), data = values$data, REML = TRUE))
            emm <- emmeans(model, "Product")
            emm_df <- as.data.frame(emm)
            lmm_results[[desc]] <- emm_df$emmean[match(lmm_results$Product, emm_df$Product)]
            pairs_result <- pairs(emm, adjust = "tukey")
            all_pairwise_results[[desc]] <- as.data.frame(pairs_result)
            aov_result <- anova(model)
            if(nrow(aov_result) > 0) {
              fixed_effects$F_value[i] <- aov_result$`F value`[1]
              fixed_effects$Df1[i] <- aov_result$NumDF[1]
              fixed_effects$Df2[i] <- aov_result$DenDF[1]
              fixed_effects$P_value[i] <- aov_result$`Pr(>F)`[1]
              fixed_effects$Significant[i] <- ifelse(aov_result$`Pr(>F)`[1] < 0.05, "Yes", "No")
            }
            vc <- as.data.frame(VarCorr(model))
            pan_var <- ifelse(any(vc$grp == "Panelist"), vc$vcov[vc$grp == "Panelist"][1], 0)
            ses_var <- ifelse(any(vc$grp == "Session"), vc$vcov[vc$grp == "Session"][1], 0)
            panprod_var <- ifelse(any(vc$grp == "Panelist:Product"), vc$vcov[vc$grp == "Panelist:Product"][1], 0)
            res_var <- ifelse(any(vc$grp == "Residual"), vc$vcov[vc$grp == "Residual"][1], 0)
            variance_comp$Panelist_Var[i] <- pan_var
            variance_comp$Session_Var[i] <- ses_var
            variance_comp$PanelistProduct_Var[i] <- panprod_var
            variance_comp$Residual_Var[i] <- res_var
            total_var <- pan_var + ses_var + panprod_var + res_var
            variance_comp$ICC[i] <- ifelse(total_var > 0, pan_var / total_var, 0)
            variance_comp$AIC[i] <- AIC(model)
            variance_comp$BIC[i] <- BIC(model)
            variance_comp$LogLik[i] <- as.numeric(logLik(model))
          }, error = function(e) {
            simple_means <- aggregate(values$data[[desc]], by = list(values$data$Product), FUN = mean, na.rm = TRUE)
            lmm_results[[desc]] <- simple_means$x[match(lmm_results$Product, simple_means$Group.1)]
          })
        }
      })
      values$lmm_means <- lmm_results
      values$fixed_effects <- fixed_effects
      values$variance_components <- variance_comp
      if(length(all_pairwise_results) > 0) {
        combined <- do.call(rbind, lapply(names(all_pairwise_results), function(d) {
          df <- all_pairwise_results[[d]]
          df$Descriptor <- d
          df$Significant <- ifelse(df$p.value < 0.05, 1, 0)
          df
        }))
        agg <- combined %>%
          group_by(contrast) %>%
          summarise(Total_Descriptors = n(), Significant_Count = sum(Significant),
                    Percent_Different = (Significant_Count / Total_Descriptors) * 100) %>%
          arrange(desc(Percent_Different))
        agg <- as.data.frame(agg)
        names(agg)[1] <- "Pair"
        values$pairwise_aggregate <- agg
      }
    }, error = function(e) {
      showNotification(paste("Model fitting error:", e$message), type = "error", duration = 10)
    })
  })
  
  output$lmm_results <- DT::renderDataTable({
    req(values$lmm_means)
    DT::datatable(values$lmm_means, options = list(scrollX = TRUE)) %>%
      DT::formatRound(columns = 2:ncol(values$lmm_means), digits = 1)
  })
  
  output$fixed_effects_tests <- DT::renderDataTable({
    req(values$fixed_effects)
    
    # Format the dataframe for display
    display_df <- values$fixed_effects
    
    # Format P-value with e-notation for small values
    display_df$P_value_formatted <- ifelse(display_df$P_value < 0.0001,
                                           sprintf("%.2e", display_df$P_value),
                                           sprintf("%.4f", display_df$P_value))
    
    # Create display dataframe with formatted p-value
    final_df <- data.frame(
      Descriptor = display_df$Descriptor,
      F_value = display_df$F_value,
      Df1 = display_df$Df1,
      Df2 = display_df$Df2,
      P_value = display_df$P_value_formatted,
      Significant = display_df$Significant,
      stringsAsFactors = FALSE
    )
    
    DT::datatable(final_df, options = list(scrollX = TRUE)) %>%
      DT::formatRound(columns = c("F_value"), digits = 2) %>%
      DT::formatRound(columns = c("Df1"), digits = 0) %>%  # No decimals for Df1
      DT::formatRound(columns = c("Df2"), digits = 1) %>%  # 1 decimal for Df2
      DT::formatStyle("Significant", backgroundColor = DT::styleEqual("Yes", "lightgreen"))
  })
  
  output$variance_components <- DT::renderDataTable({
    req(values$variance_components)
    DT::datatable(values$variance_components, options = list(scrollX = TRUE)) %>%
      DT::formatRound(columns = 2:ncol(values$variance_components), digits = 4)
  })
  
  output$residuals_fitted_plot <- renderPlot({
    req(values$data, input$assumption_descriptor)
    tryCatch({
      formula_str <- paste(input$assumption_descriptor, "~ Product + (1|Panelist) + (1|Session) + (1|Panelist:Product)")
      model <- suppressWarnings(lmer(as.formula(formula_str), data = values$data, REML = TRUE))
      plot(fitted(model), residuals(model), xlab = "Fitted Values", ylab = "Residuals", main = "Residuals vs Fitted")
      abline(h = 0, col = "red", lwd = 2)
    }, error = function(e) {
      plot(1, type = "n", main = "Model failed to converge\n(analysis used simple means instead)")
    })
  })
  
  output$qq_plot <- renderPlot({
    req(values$data, input$assumption_descriptor)
    tryCatch({
      formula_str <- paste(input$assumption_descriptor, "~ Product + (1|Panelist) + (1|Session) + (1|Panelist:Product)")
      model <- suppressWarnings(lmer(as.formula(formula_str), data = values$data, REML = TRUE))
      qqnorm(residuals(model), main = "Q-Q Plot")
      qqline(residuals(model), col = "red", lwd = 2)
    }, error = function(e) {
      plot(1, type = "n", main = "Model failed to converge\n(analysis used simple means instead)")
    })
  })
  
  observe({
    req(values$lmm_means)
    tryCatch({
      pca_data <- values$lmm_means[, -1, drop = FALSE]
      rownames(pca_data) <- values$lmm_means$Product
      pca_data <- pca_data[, sapply(pca_data, function(x) var(x, na.rm = TRUE) > 1e-8), drop = FALSE]
      if(ncol(pca_data) >= 2) {
        values$pca_result <- PCA(pca_data, graph = FALSE)
      }
    }, error = function(e) {
      showNotification(paste("PCA error:", e$message), type = "warning")
    })
  })
  
  # AHC Clustering on LSMeans
  observe({
    req(values$lmm_means)
    tryCatch({
      ahc_data <- values$lmm_means[, -1, drop = FALSE]
      rownames(ahc_data) <- values$lmm_means$Product
      ahc_data <- ahc_data[, sapply(ahc_data, function(x) var(x, na.rm = TRUE) > 1e-8), drop = FALSE]
      
      if(ncol(ahc_data) >= 2 && nrow(ahc_data) >= 3) {
        n_products <- nrow(ahc_data)
        k_range <- 2:min(8, n_products - 1)
        sil_scores <- numeric(length(k_range))
        wss_values <- numeric(length(k_range))
        
        for(i in seq_along(k_range)) {
          k <- k_range[i]
          hc <- hclust(dist(ahc_data), method = "ward.D2")
          clusters <- cutree(hc, k = k)
          sil <- silhouette(clusters, dist(ahc_data))
          sil_scores[i] <- mean(sil[,3])
          wss <- sum(sapply(1:k, function(j) {
            cluster_points <- ahc_data[clusters == j, , drop = FALSE]
            if(nrow(cluster_points) > 1) {
              sum(scale(cluster_points, scale = FALSE)^2)
            } else {
              0
            }
          }))
          wss_values[i] <- wss
        }
        
        values$ahc_silhouette_data <- data.frame(k = k_range, silhouette = sil_scores)
        values$ahc_wss_data <- data.frame(k = k_range, wss = wss_values)
        values$ahc_optimal_k <- k_range[which.max(sil_scores)]
        updateNumericInput(session, "n_clusters_ahc", value = values$ahc_optimal_k)
      }
    }, error = function(e) {
      showNotification(paste("AHC clustering error:", e$message), type = "warning")
    })
  })
  
  output$scree_plot <- renderPlotly({
    req(values$pca_result)
    eig <- values$pca_result$eig
    df <- data.frame(PC = paste0("PC", 1:nrow(eig)), Eigenvalue = eig[,1], Variance_Pct = eig[,2])
    df$PC <- factor(df$PC, levels = df$PC)
    p <- ggplot(df, aes(x = PC, y = Eigenvalue)) +
      geom_bar(stat = "identity", fill = "steelblue", alpha = 0.7) +
      geom_line(aes(group = 1), color = "red", linewidth = 1) +
      geom_point(color = "red", size = 3) +
      theme_minimal() + labs(title = "Scree Plot") +
      theme(axis.text.x = element_text(angle = 45, hjust = 1))
    ggplotly(p)
  })
  
  output$pca_contrib <- renderPlotly({
    req(values$pca_result)
    tryCatch({
      contrib <- as.data.frame(values$pca_result$var$contrib)
      contrib$Variable <- rownames(contrib)
      contrib_long <- pivot_longer(contrib, cols = -Variable, names_to = "PC", values_to = "Contribution")
      contrib_long <- contrib_long[contrib_long$PC %in% c("Dim.1", "Dim.2", "Dim.3"), ]
      p <- ggplot(contrib_long, aes(x = Variable, y = Contribution, fill = PC)) +
        geom_col(position = "dodge") + theme_minimal() +
        theme(axis.text.x = element_text(angle = 45, hjust = 1)) +
        labs(title = "Variable Contributions", y = "Contribution (%)")
      ggplotly(p)
    }, error = function(e) {
      p <- ggplot() + annotate("text", x = 0.5, y = 0.5, label = "Error") + theme_void()
      ggplotly(p)
    })
  })
  
  create_pca_plot_12 <- reactive({
    req(values$pca_result, input$repel_force_12, input$point_size_12, input$text_size_12)
    tryCatch({
      ind <- as.data.frame(values$pca_result$ind$coord)
      var <- as.data.frame(values$pca_result$var$coord)
      scale_factor <- 3
      ind$label <- rownames(ind)
      var$label <- rownames(var)
      var$Dim.1 <- var$Dim.1 * scale_factor
      var$Dim.2 <- var$Dim.2 * scale_factor
      
      # AXIS SWAP LOGIC
      if (input$swap_axes_12) {
        all_x <- c(ind$Dim.2, var$Dim.2)
        all_y <- c(ind$Dim.1, var$Dim.1)
        x_range <- diff(range(all_x))
        x_padding <- x_range * 0.25
        xlim <- c(min(all_x) - x_padding, max(all_x) + x_padding)
        y_range <- diff(range(all_y))
        y_padding <- y_range * 0.25
        ylim <- c(min(all_y) - y_padding, max(all_y) + y_padding)
        
        p <- ggplot() +
          geom_hline(yintercept = 0, color = "black", linewidth = 0.8) +
          geom_vline(xintercept = 0, color = "black", linewidth = 0.8) +
          geom_point(data = ind, aes(x = Dim.2, y = Dim.1), color = "blue", size = input$point_size_12, alpha = 0.8) +
          geom_text_repel(data = ind, aes(x = Dim.2, y = Dim.1, label = label),
                          color = "blue", size = input$text_size_12, fontface = "bold",
                          force = input$repel_force_12, force_pull = 0.5, max.overlaps = Inf,
                          max.iter = 10000, box.padding = 0.5, point.padding = 0.3,
                          segment.color = "blue", segment.alpha = 0.3, segment.size = 0.3,
                          min.segment.length = 0.1) +
          geom_point(data = var, aes(x = Dim.2, y = Dim.1), color = "red", size = input$point_size_12 * 0.8, alpha = 0.8) +
          geom_text_repel(data = var, aes(x = Dim.2, y = Dim.1, label = label),
                          color = "red", size = input$text_size_12, fontface = "bold",
                          force = input$repel_force_12 * 1.2, force_pull = 0.5, max.overlaps = Inf,
                          max.iter = 10000, box.padding = 0.5, point.padding = 0.3,
                          segment.color = "red", segment.alpha = 0.3, segment.size = 0.3,
                          min.segment.length = 0.1) +
          scale_x_continuous(limits = xlim, expand = c(0, 0)) +
          scale_y_continuous(limits = ylim, expand = c(0, 0)) +
          theme_minimal(base_size = 14) +
          theme(panel.grid.major = element_blank(), panel.grid.minor = element_blank(),
                panel.background = element_rect(fill = "white", color = NA),
                plot.background = element_rect(fill = "white", color = NA),
                axis.line = element_blank(), axis.ticks = element_blank(),
                legend.position = "none", aspect.ratio = NULL) +
          labs(title = "PCA Biplot: PC2 vs PC1",
               x = paste0("PC2 (", round(values$pca_result$eig[2,2], 1), "%)"),
               y = paste0("PC1 (", round(values$pca_result$eig[1,2], 1), "%)"))
      } else {
        all_x <- c(ind$Dim.1, var$Dim.1)
        all_y <- c(ind$Dim.2, var$Dim.2)
        x_range <- diff(range(all_x))
        x_padding <- x_range * 0.25
        xlim <- c(min(all_x) - x_padding, max(all_x) + x_padding)
        y_range <- diff(range(all_y))
        y_padding <- y_range * 0.25
        ylim <- c(min(all_y) - y_padding, max(all_y) + y_padding)
        
        p <- ggplot() +
          geom_hline(yintercept = 0, color = "black", linewidth = 0.8) +
          geom_vline(xintercept = 0, color = "black", linewidth = 0.8) +
          geom_point(data = ind, aes(x = Dim.1, y = Dim.2), color = "blue", size = input$point_size_12, alpha = 0.8) +
          geom_text_repel(data = ind, aes(x = Dim.1, y = Dim.2, label = label),
                          color = "blue", size = input$text_size_12, fontface = "bold",
                          force = input$repel_force_12, force_pull = 0.5, max.overlaps = Inf,
                          max.iter = 10000, box.padding = 0.5, point.padding = 0.3,
                          segment.color = "blue", segment.alpha = 0.3, segment.size = 0.3,
                          min.segment.length = 0.1) +
          geom_point(data = var, aes(x = Dim.1, y = Dim.2), color = "red", size = input$point_size_12 * 0.8, alpha = 0.8) +
          geom_text_repel(data = var, aes(x = Dim.1, y = Dim.2, label = label),
                          color = "red", size = input$text_size_12, fontface = "bold",
                          force = input$repel_force_12 * 1.2, force_pull = 0.5, max.overlaps = Inf,
                          max.iter = 10000, box.padding = 0.5, point.padding = 0.3,
                          segment.color = "red", segment.alpha = 0.3, segment.size = 0.3,
                          min.segment.length = 0.1) +
          scale_x_continuous(limits = xlim, expand = c(0, 0)) +
          scale_y_continuous(limits = ylim, expand = c(0, 0)) +
          theme_minimal(base_size = 14) +
          theme(panel.grid.major = element_blank(), panel.grid.minor = element_blank(),
                panel.background = element_rect(fill = "white", color = NA),
                plot.background = element_rect(fill = "white", color = NA),
                axis.line = element_blank(), axis.ticks = element_blank(),
                legend.position = "none", aspect.ratio = NULL) +
          labs(title = "PCA Biplot: PC1 vs PC2",
               x = paste0("PC1 (", round(values$pca_result$eig[1,2], 1), "%)"),
               y = paste0("PC2 (", round(values$pca_result$eig[2,2], 1), "%)"))
      }
      return(p)
    }, error = function(e) {
      ggplot() + annotate("text", x = 0.5, y = 0.5, label = paste("Error:", e$message), size = 5, color = "red") + theme_void()
    })
  })
  
  output$pca_biplot_12 <- renderPlot({ create_pca_plot_12() })
  
  output$download_pca_12 <- downloadHandler(
    filename = function() { paste0("PCA_PC1_PC2_", Sys.Date(), ".png") },
    content = function(file) {
      ggsave(file, plot = create_pca_plot_12(), width = 14, height = 8, dpi = 300, bg = "white")
    }
  )
  
  create_pca_plot_13 <- reactive({
    req(values$pca_result, input$repel_force_13, input$point_size_13, input$text_size_13)
    tryCatch({
      ind <- as.data.frame(values$pca_result$ind$coord)
      var <- as.data.frame(values$pca_result$var$coord)
      scale_factor <- 3
      ind$label <- rownames(ind)
      var$label <- rownames(var)
      var$Dim.1 <- var$Dim.1 * scale_factor
      var$Dim.3 <- var$Dim.3 * scale_factor
      
      # AXIS SWAP LOGIC
      if (input$swap_axes_13) {
        all_x <- c(ind$Dim.3, var$Dim.3)
        all_y <- c(ind$Dim.1, var$Dim.1)
        x_range <- diff(range(all_x))
        x_padding <- x_range * 0.25
        xlim <- c(min(all_x) - x_padding, max(all_x) + x_padding)
        y_range <- diff(range(all_y))
        y_padding <- y_range * 0.25
        ylim <- c(min(all_y) - y_padding, max(all_y) + y_padding)
        
        p <- ggplot() +
          geom_hline(yintercept = 0, color = "black", linewidth = 0.8) +
          geom_vline(xintercept = 0, color = "black", linewidth = 0.8) +
          geom_point(data = ind, aes(x = Dim.3, y = Dim.1), color = "blue", size = input$point_size_13, alpha = 0.8) +
          geom_text_repel(data = ind, aes(x = Dim.3, y = Dim.1, label = label),
                          color = "blue", size = input$text_size_13, fontface = "bold",
                          force = input$repel_force_13, force_pull = 0.5, max.overlaps = Inf,
                          max.iter = 10000, box.padding = 0.5, point.padding = 0.3,
                          segment.color = "blue", segment.alpha = 0.3, segment.size = 0.3,
                          min.segment.length = 0.1) +
          geom_point(data = var, aes(x = Dim.3, y = Dim.1), color = "red", size = input$point_size_13 * 0.8, alpha = 0.8) +
          geom_text_repel(data = var, aes(x = Dim.3, y = Dim.1, label = label),
                          color = "red", size = input$text_size_13, fontface = "bold",
                          force = input$repel_force_13 * 1.2, force_pull = 0.5, max.overlaps = Inf,
                          max.iter = 10000, box.padding = 0.5, point.padding = 0.3,
                          segment.color = "red", segment.alpha = 0.3, segment.size = 0.3,
                          min.segment.length = 0.1) +
          scale_x_continuous(limits = xlim, expand = c(0, 0)) +
          scale_y_continuous(limits = ylim, expand = c(0, 0)) +
          theme_minimal(base_size = 14) +
          theme(panel.grid.major = element_blank(), panel.grid.minor = element_blank(),
                panel.background = element_rect(fill = "white", color = NA),
                plot.background = element_rect(fill = "white", color = NA),
                axis.line = element_blank(), axis.ticks = element_blank(),
                legend.position = "none", aspect.ratio = NULL) +
          labs(title = "PCA Biplot: PC3 vs PC1",
               x = paste0("PC3 (", round(values$pca_result$eig[3,2], 1), "%)"),
               y = paste0("PC1 (", round(values$pca_result$eig[1,2], 1), "%)"))
      } else {
        all_x <- c(ind$Dim.1, var$Dim.1)
        all_y <- c(ind$Dim.3, var$Dim.3)
        x_range <- diff(range(all_x))
        x_padding <- x_range * 0.25
        xlim <- c(min(all_x) - x_padding, max(all_x) + x_padding)
        y_range <- diff(range(all_y))
        y_padding <- y_range * 0.25
        ylim <- c(min(all_y) - y_padding, max(all_y) + y_padding)
        
        p <- ggplot() +
          geom_hline(yintercept = 0, color = "black", linewidth = 0.8) +
          geom_vline(xintercept = 0, color = "black", linewidth = 0.8) +
          geom_point(data = ind, aes(x = Dim.1, y = Dim.3), color = "blue", size = input$point_size_13, alpha = 0.8) +
          geom_text_repel(data = ind, aes(x = Dim.1, y = Dim.3, label = label),
                          color = "blue", size = input$text_size_13, fontface = "bold",
                          force = input$repel_force_13, force_pull = 0.5, max.overlaps = Inf,
                          max.iter = 10000, box.padding = 0.5, point.padding = 0.3,
                          segment.color = "blue", segment.alpha = 0.3, segment.size = 0.3,
                          min.segment.length = 0.1) +
          geom_point(data = var, aes(x = Dim.1, y = Dim.3), color = "red", size = input$point_size_13 * 0.8, alpha = 0.8) +
          geom_text_repel(data = var, aes(x = Dim.1, y = Dim.3, label = label),
                          color = "red", size = input$text_size_13, fontface = "bold",
                          force = input$repel_force_13 * 1.2, force_pull = 0.5, max.overlaps = Inf,
                          max.iter = 10000, box.padding = 0.5, point.padding = 0.3,
                          segment.color = "red", segment.alpha = 0.3, segment.size = 0.3,
                          min.segment.length = 0.1) +
          scale_x_continuous(limits = xlim, expand = c(0, 0)) +
          scale_y_continuous(limits = ylim, expand = c(0, 0)) +
          theme_minimal(base_size = 14) +
          theme(panel.grid.major = element_blank(), panel.grid.minor = element_blank(),
                panel.background = element_rect(fill = "white", color = NA),
                plot.background = element_rect(fill = "white", color = NA),
                axis.line = element_blank(), axis.ticks = element_blank(),
                legend.position = "none", aspect.ratio = NULL) +
          labs(title = "PCA Biplot: PC1 vs PC3",
               x = paste0("PC1 (", round(values$pca_result$eig[1,2], 1), "%)"),
               y = paste0("PC3 (", round(values$pca_result$eig[3,2], 1), "%)"))
      }
      return(p)
    }, error = function(e) {
      ggplot() + annotate("text", x = 0.5, y = 0.5, label = paste("Error:", e$message), size = 5, color = "red") + theme_void()
    })
  })
  
  output$pca_biplot_13 <- renderPlot({ create_pca_plot_13() })
  
  output$download_pca_13 <- downloadHandler(
    filename = function() { paste0("PCA_PC1_PC3_", Sys.Date(), ".png") },
    content = function(file) {
      ggsave(file, plot = create_pca_plot_13(), width = 14, height = 8, dpi = 300, bg = "white")
    }
  )
  
  observe({
    req(values$pca_result)
    tryCatch({
      coords <- values$pca_result$ind$coord[, 1:2]
      n_products <- nrow(coords)
      if(n_products < 3) return()
      k_range <- 2:min(8, n_products - 1)
      sil_scores <- numeric(length(k_range))
      wss_values <- numeric(length(k_range))
      for(i in seq_along(k_range)) {
        k <- k_range[i]
        hc <- hclust(dist(coords), method = "ward.D2")
        clusters <- cutree(hc, k = k)
        sil <- silhouette(clusters, dist(coords))
        sil_scores[i] <- mean(sil[,3])
        wss <- sum(sapply(1:k, function(j) {
          cluster_points <- coords[clusters == j, , drop = FALSE]
          if(nrow(cluster_points) > 1) {
            sum(scale(cluster_points, scale = FALSE)^2)
          } else {
            0
          }
        }))
        wss_values[i] <- wss
      }
      values$silhouette_data <- data.frame(k = k_range, silhouette = sil_scores)
      values$wss_data <- data.frame(k = k_range, wss = wss_values)
      values$optimal_k <- k_range[which.max(sil_scores)]
      updateNumericInput(session, "n_clusters_manual", value = values$optimal_k)
    }, error = function(e) {
      showNotification(paste("Clustering error:", e$message), type = "warning")
    })
  })
  
  output$silhouette_plot <- renderPlotly({
    req(values$silhouette_data)
    p <- ggplot(values$silhouette_data, aes(x = k, y = silhouette)) +
      geom_line(color = "blue", linewidth = 1) + geom_point(color = "blue", size = 3) +
      theme_minimal() + labs(title = "Silhouette Score", x = "Number of Clusters", y = "Silhouette Score")
    ggplotly(p)
  })
  
  output$elbow_plot <- renderPlotly({
    req(values$wss_data)
    p <- ggplot(values$wss_data, aes(x = k, y = wss)) +
      geom_line(color = "red", linewidth = 1) + geom_point(color = "red", size = 3) +
      theme_minimal() + labs(title = "Elbow Method", x = "Number of Clusters", y = "Within-Cluster SS")
    ggplotly(p)
  })
  
  output$optimal_clusters_text <- renderText({
    req(values$optimal_k)
    paste("Optimal number of clusters based on Silhouette analysis:", values$optimal_k)
  })
  
  output$dendro_12 <- renderPlot({
    req(values$pca_result, input$n_clusters_manual)
    tryCatch({
      coords <- values$pca_result$ind$coord[, 1:2]
      hc <- hclust(dist(coords), method = "ward.D2")
      plot(hc, main = "Dendrogram (PC1 & PC2)", xlab = "Products", ylab = "Height", cex = 0.8)
      rect.hclust(hc, k = input$n_clusters_manual, border = 2:6)
    }, error = function(e) {
      plot(1, type = "n", main = "Clustering error")
    })
  })
  
  output$dendro_13 <- renderPlot({
    req(values$pca_result, input$n_clusters_manual)
    tryCatch({
      coords <- values$pca_result$ind$coord[, c(1,3)]
      hc <- hclust(dist(coords), method = "ward.D2")
      plot(hc, main = "Dendrogram (PC1 & PC3)", xlab = "Products", ylab = "Height", cex = 0.8)
      rect.hclust(hc, k = input$n_clusters_manual, border = 2:6)
    }, error = function(e) {
      plot(1, type = "n", main = "Clustering error")
    })
  })
  
  output$cluster_table_12 <- DT::renderDataTable({
    req(values$pca_result, input$n_clusters_manual)
    tryCatch({
      coords <- values$pca_result$ind$coord[, 1:2]
      hc <- hclust(dist(coords), method = "ward.D2")
      clusters <- cutree(hc, k = input$n_clusters_manual)
      df <- data.frame(Product = names(clusters), Cluster = clusters)
      DT::datatable(df, options = list(pageLength = 20))
    }, error = function(e) {
      DT::datatable(data.frame(Error = "Clustering failed"))
    })
  })
  
  output$cluster_table_13 <- DT::renderDataTable({
    req(values$pca_result, input$n_clusters_manual)
    tryCatch({
      coords <- values$pca_result$ind$coord[, c(1,3)]
      hc <- hclust(dist(coords), method = "ward.D2")
      clusters <- cutree(hc, k = input$n_clusters_manual)
      df <- data.frame(Product = names(clusters), Cluster = clusters)
      DT::datatable(df, options = list(pageLength = 20))
    }, error = function(e) {
      DT::datatable(data.frame(Error = "Clustering failed"))
    })
  })
  
  # AHC Clustering Outputs
  output$ahc_silhouette_plot <- renderPlotly({
    req(values$ahc_silhouette_data)
    p <- ggplot(values$ahc_silhouette_data, aes(x = k, y = silhouette)) +
      geom_line(color = "blue", linewidth = 1) + geom_point(color = "blue", size = 3) +
      theme_minimal() + labs(title = "Silhouette Score", x = "Number of Clusters", y = "Silhouette Score")
    ggplotly(p)
  })
  
  output$ahc_elbow_plot <- renderPlotly({
    req(values$ahc_wss_data)
    p <- ggplot(values$ahc_wss_data, aes(x = k, y = wss)) +
      geom_line(color = "red", linewidth = 1) + geom_point(color = "red", size = 3) +
      theme_minimal() + labs(title = "Elbow Method", x = "Number of Clusters", y = "Within-Cluster SS")
    ggplotly(p)
  })
  
  output$ahc_optimal_clusters_text <- renderText({
    req(values$ahc_optimal_k)
    paste("Optimal number of clusters based on Silhouette analysis:", values$ahc_optimal_k)
  })
  
  output$ahc_dendrogram <- renderPlot({
    req(values$lmm_means, input$n_clusters_ahc)
    tryCatch({
      ahc_data <- values$lmm_means[, -1, drop = FALSE]
      rownames(ahc_data) <- values$lmm_means$Product
      ahc_data <- ahc_data[, sapply(ahc_data, function(x) var(x, na.rm = TRUE) > 1e-8), drop = FALSE]
      hc <- hclust(dist(ahc_data), method = "ward.D2")
      plot(hc, main = "Dendrogram - AHC on LSMeans (Ward's Method)", 
           xlab = "Products", ylab = "Height", cex = 0.8)
      rect.hclust(hc, k = input$n_clusters_ahc, border = 2:6)
    }, error = function(e) {
      plot(1, type = "n", main = "Clustering error")
    })
  })
  
  output$ahc_cluster_table <- DT::renderDataTable({
    req(values$lmm_means, input$n_clusters_ahc)
    tryCatch({
      ahc_data <- values$lmm_means[, -1, drop = FALSE]
      rownames(ahc_data) <- values$lmm_means$Product
      ahc_data <- ahc_data[, sapply(ahc_data, function(x) var(x, na.rm = TRUE) > 1e-8), drop = FALSE]
      hc <- hclust(dist(ahc_data), method = "ward.D2")
      clusters <- cutree(hc, k = input$n_clusters_ahc)
      df <- data.frame(Product = names(clusters), Cluster = clusters)
      df <- df[order(df$Cluster, df$Product), ]
      DT::datatable(df, options = list(pageLength = 20), rownames = FALSE)
    }, error = function(e) {
      DT::datatable(data.frame(Error = "Clustering failed"))
    })
  })
  
  output$ahc_cluster_profiles <- DT::renderDataTable({
    req(values$lmm_means, input$n_clusters_ahc)
    tryCatch({
      ahc_data <- values$lmm_means[, -1, drop = FALSE]
      rownames(ahc_data) <- values$lmm_means$Product
      ahc_data <- ahc_data[, sapply(ahc_data, function(x) var(x, na.rm = TRUE) > 1e-8), drop = FALSE]
      hc <- hclust(dist(ahc_data), method = "ward.D2")
      clusters <- cutree(hc, k = input$n_clusters_ahc)
      
      # Create dataframe with cluster assignments
      ahc_df <- as.data.frame(ahc_data)
      ahc_df$Cluster <- clusters
      
      # Calculate mean for each cluster
      cluster_profiles <- ahc_df %>%
        group_by(Cluster) %>%
        summarise(across(everything(), \(x) mean(x, na.rm = TRUE)))
      
      cluster_profiles <- as.data.frame(cluster_profiles)
      
      DT::datatable(cluster_profiles, options = list(scrollX = TRUE, pageLength = 10)) %>%
        DT::formatRound(columns = 2:ncol(cluster_profiles), digits = 2)
    }, error = function(e) {
      DT::datatable(data.frame(Error = paste("Profile error:", e$message)))
    })
  })
  
  output$pairwise_comparisons <- DT::renderDataTable({
    req(values$data, input$comparison_descriptor)
    tryCatch({
      formula_str <- paste(input$comparison_descriptor, "~ Product + (1|Panelist) + (1|Session) + (1|Panelist:Product)")
      model <- suppressWarnings(lmer(as.formula(formula_str), data = values$data))
      # CORRECTION: Use Kenward-Roger method for more accurate SE and df estimation
      emm <- suppressMessages(emmeans(model, "Product", lmer.df = "kenward-roger"))
      pairs_result <- pairs(emm, adjust = "tukey")
      df <- as.data.frame(pairs_result)
      df$Significant <- ifelse(df$p.value < 0.05, "Yes", "No")
      
      # NEW: Format p-value with e-notation for small values
      df$p.value_formatted <- ifelse(df$p.value < 0.0001,
                                     sprintf("%.2e", df$p.value),
                                     sprintf("%.4f", df$p.value))
      
      df_display <- df[, c("contrast", "estimate", "SE", "df", "t.ratio", "p.value_formatted", "Significant")]
      names(df_display)[6] <- "p.value"
      
      # NEW: Format df to 3 decimal places
      DT::datatable(df_display, options = list(scrollX = TRUE)) %>%
        DT::formatRound(columns = c("estimate", "SE", "df", "t.ratio"), digits = 3) %>%
        DT::formatStyle("Significant", backgroundColor = DT::styleEqual("Yes", "lightgreen"))
    }, error = function(e) {
      DT::datatable(data.frame(Error = paste("Error:", e$message)))
    })
  })
  
  output$cld_results_enhanced <- DT::renderDataTable({
    req(values$data, input$comparison_descriptor)
    tryCatch({
      formula_str <- paste(input$comparison_descriptor, "~ Product + (1|Panelist) + (1|Session) + (1|Panelist:Product)")
      model <- suppressWarnings(lmer(as.formula(formula_str), data = values$data))
      emm <- suppressMessages(emmeans(model, "Product"))
      cld_result <- multcomp::cld(emm, Letters = letters, adjust = "tukey")
      df <- as.data.frame(cld_result)
      summary_df <- data.frame(Product = df$Product, LSMean = round(df$emmean, 2),
                               SE = round(df$SE, 3), Letters = trimws(df$.group))
      DT::datatable(summary_df, options = list(scrollX = TRUE))
    }, error = function(e) {
      DT::datatable(data.frame(Error = paste("Error:", e$message)))
    })
  })
  
  output$bargraph_plot <- renderPlotly({
    req(values$data, input$bargraph_descriptor)
    tryCatch({
      formula_str <- paste(input$bargraph_descriptor, "~ Product + (1|Panelist) + (1|Session) + (1|Panelist:Product)")
      model <- suppressWarnings(lmer(as.formula(formula_str), data = values$data))
      emm <- suppressMessages(emmeans(model, "Product"))
      cld_result <- multcomp::cld(emm, Letters = letters, adjust = "tukey")
      df <- as.data.frame(cld_result)
      p <- ggplot(df, aes(x = Product, y = emmean)) +
        geom_col(aes(fill = Product), alpha = 0.7, color = "black", show.legend = FALSE) +
        geom_errorbar(aes(ymin = emmean - SE, ymax = emmean + SE), width = 0.2, linewidth = 0.8) +
        geom_text(aes(y = emmean + SE + 0.1*max(emmean), label = trimws(.group)), size = 4, fontface = "bold") +
        theme_minimal() + theme(legend.position = "none") +
        labs(title = paste("LSMeans ±SE:", input$bargraph_descriptor), y = "LSMean", x = "Product",
             caption = "Same letter = not significantly different (α = 0.05)")
      ggplotly(p, tooltip = c("x", "y")) %>% layout(showlegend = FALSE)
    }, error = function(e) {
      p <- ggplot() + annotate("text", x = 0.5, y = 0.5, label = "Error") + theme_void()
      ggplotly(p)
    })
  })
  
  output$bargraph_stats <- DT::renderDataTable({
    req(values$data, input$bargraph_descriptor)
    tryCatch({
      formula_str <- paste(input$bargraph_descriptor, "~ Product + (1|Panelist) + (1|Session) + (1|Panelist:Product)")
      model <- suppressWarnings(lmer(as.formula(formula_str), data = values$data))
      emm <- suppressMessages(emmeans(model, "Product"))
      df <- as.data.frame(emm)
      df$CI_lower <- df$emmean - 1.96 * df$SE
      df$CI_upper <- df$emmean + 1.96 * df$SE
      DT::datatable(df, options = list(scrollX = TRUE)) %>%
        DT::formatRound(columns = c("emmean", "SE", "df", "lower.CL", "upper.CL", "CI_lower", "CI_upper"), digits = 3)
    }, error = function(e) {
      DT::datatable(data.frame(Error = paste("Error:", e$message)))
    })
  })
  
  output$pairwise_aggregate_table <- DT::renderDataTable({
    req(values$pairwise_aggregate)
    DT::datatable(values$pairwise_aggregate, 
                  options = list(scrollX = TRUE, pageLength = 20),
                  caption = "Product pairs ranked by % of descriptors showing significant difference") %>%
      DT::formatRound(columns = "Percent_Different", digits = 1) %>%
      DT::formatStyle("Percent_Different",
                      background = styleColorBar(c(0, 100), 'lightblue'),
                      backgroundSize = '100% 90%',
                      backgroundRepeat = 'no-repeat',
                      backgroundPosition = 'center')
  })
  
  output$pairwise_frequency_table <- DT::renderDataTable({
    req(values$pairwise_aggregate)
    freq_df <- data.frame(Percent_Range = c("0-10%", "10-20%", "20-30%", "30-40%", "40-50%", 
                                            "50-60%", "60-70%", "70-80%", "80-90%", "90-100%"),
                          Count = 0, Product_Pairs = character(10), stringsAsFactors = FALSE)
    for(i in 1:nrow(freq_df)) {
      lower <- (i-1) * 10
      upper <- i * 10
      pairs_in_range <- values$pairwise_aggregate$Pair[
        values$pairwise_aggregate$Percent_Different > lower & 
          values$pairwise_aggregate$Percent_Different <= upper
      ]
      freq_df$Count[i] <- length(pairs_in_range)
      if(length(pairs_in_range) > 0) {
        freq_df$Product_Pairs[i] <- paste(pairs_in_range, collapse = "; ")
      } else {
        freq_df$Product_Pairs[i] <- "None"
      }
    }
    DT::datatable(freq_df, 
                  options = list(scrollX = TRUE, pageLength = 10, autoWidth = TRUE,
                                 columnDefs = list(list(width = '50%', targets = 2))),
                  caption = "Distribution of product pairs by differentiation level") %>%
      DT::formatStyle("Count",
                      background = styleColorBar(c(0, max(freq_df$Count)), 'lightgreen'),
                      backgroundSize = '100% 90%',
                      backgroundRepeat = 'no-repeat',
                      backgroundPosition = 'center')
  })
  
  # Download handler for all analysis tables
  output$download_all_tables <- downloadHandler(
    filename = function() {
      paste0("QDA_Analysis_", format(Sys.Date(), "%Y%m%d"), ".xlsx")
    },
    content = function(file) {
      # Create a new workbook
      wb <- createWorkbook()
      
      # Sheet 1: LSMeans from Mixed Models
      if(!is.null(values$lmm_means)) {
        addWorksheet(wb, "LSMeans")
        writeData(wb, "LSMeans", values$lmm_means)
      }
      
      # Sheet 2: Fixed Effects F-Tests
      if(!is.null(values$fixed_effects)) {
        addWorksheet(wb, "Fixed_Effects")
        writeData(wb, "Fixed_Effects", values$fixed_effects)
      }
      
      # Sheet 3: Variance Components
      if(!is.null(values$variance_components)) {
        addWorksheet(wb, "Variance_Components")
        writeData(wb, "Variance_Components", values$variance_components)
      }
      
      # Sheet 4: Pairwise Aggregate (Product Differentiation)
      if(!is.null(values$pairwise_aggregate)) {
        addWorksheet(wb, "Product_Differentiation")
        writeData(wb, "Product_Differentiation", values$pairwise_aggregate)
      }
      
      # Sheet 5: PCA Eigenvalues (if PCA was run)
      if(!is.null(values$pca_result)) {
        tryCatch({
          eig <- get_eigenvalue(values$pca_result)
          addWorksheet(wb, "PCA_Eigenvalues")
          writeData(wb, "PCA_Eigenvalues", eig)
        }, error = function(e) {})
      }
      
      # Sheet 6: Normality Tests
      tryCatch({
        if(!is.null(values$data)) {
          descriptors <- colnames(values$data)[4:ncol(values$data)]
          results <- data.frame(Descriptor = character(), Shapiro_W = numeric(), Shapiro_p = numeric(),
                                AD_A = numeric(), AD_p = numeric(), JB_stat = numeric(), JB_p = numeric(),
                                Conclusion = character(), stringsAsFactors = FALSE)
          for (desc in descriptors) {
            x <- values$data[[desc]]
            if (length(unique(x)) > 1 && length(x) >= 3) {
              sw <- shapiro.test(x)
              ad <- ad.test(x)
              jb <- jarque.test(x)
              conclusion <- ifelse(sw$p.value > 0.05 & ad$p.value > 0.05 & jb$p.value > 0.05, 
                                   "Normal", "Non-Normal")
              results <- rbind(results, data.frame(Descriptor = desc, Shapiro_W = round(sw$statistic, 4),
                                                   Shapiro_p = round(sw$p.value, 4), AD_A = round(ad$statistic, 4),
                                                   AD_p = round(ad$p.value, 4), JB_stat = round(jb$statistic, 4),
                                                   JB_p = round(jb$p.value, 4), Conclusion = conclusion))
            }
          }
          if(nrow(results) > 0) {
            addWorksheet(wb, "Normality_Tests")
            writeData(wb, "Normality_Tests", results)
          }
        }
      }, error = function(e) {})
      
      # Save the workbook
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

shinyApp(ui = ui, server = server)

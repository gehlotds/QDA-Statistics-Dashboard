# QDA Sensory Analysis Dashboard

A comprehensive R Shiny dashboard for conducting Quantitative Descriptive Analysis (QDA) in sensory science. This tool implements Linear Mixed Models with Kenward-Roger approximation and provides detailed statistical analysis for sensory profiling studies with interactive visualizations and exportable results.

## Features

- **Interactive Data Upload**: CSV upload with automatic validation and data preview
- **Linear Mixed Models**: Implements Kenward-Roger method for accurate small-sample inference
- **Comprehensive Analysis**: LSMeans, F-tests, variance components, and Tukey post-hoc comparisons
- **Real-time Visualization**: Interactive plots, dendrograms, biplots, and bar graphs using Plotly
- **Multivariate Analysis**: PCA, hierarchical clustering (AHC), and K-means clustering
- **One-Click Excel Export**: Download all analysis tables in a single workbook with multiple sheets
- **Normality Testing**: Multiple tests (Shapiro-Wilk, Anderson-Darling, Jarque-Bera) with visual diagnostics
- **Built-in Statistical Theory**: Comprehensive documentation of all methods and formulas

## Installation

### Prerequisites

Ensure you have R (version 4.0+ recommended) installed on your system.

### Required R Packages

```r
# Install required packages
install.packages(c(
  "shiny", "shinydashboard", "DT", "plotly", "ggplot2",
  "dplyr", "tidyr", "lme4", "lmerTest", "car", "factoextra", 
  "FactoMineR", "cluster", "multcompView", "agricolae",
  "corrplot", "reshape2", "emmeans", "nortest", "moments",
  "multcomp", "MASS", "ggrepel", "pbkrtest", "openxlsx"
))
```

### Running the Application

1. Clone this repository:
```bash
git clone https://github.com/yourusername/qda-sensory-dashboard.git
cd qda-sensory-dashboard
```

2. Open R/RStudio and run:
```r
# Load required libraries
source("app.R")  # This will launch the Shiny app
```

3. The application will open in your default web browser.

## Usage

### Basic Workflow

1. **Data Upload**: Upload your QDA data as a CSV file
   - Required format: Session | Panelist | Product | Descriptor1 | Descriptor2 | ...
   - Accepts comma, semicolon, or tab-separated files
   - Preview and validate data structure before analysis

2. **Check Normality**: Review normality tests for all descriptors
   - Multiple statistical tests with visual diagnostics
   - Identifies descriptors that may need transformation

3. **Run Analysis**: Automatic fitting of Linear Mixed Models
   - LSMeans calculated for all product-descriptor combinations
   - Fixed effects F-tests for product differences
   - Variance components decomposition

4. **Interpret Results**: View comprehensive results including:
   - Product differentiation analysis
   - Pairwise comparisons with significance letters
   - Interactive visualizations
   - Model diagnostics

### Data Format

The application expects CSV data with the following structure:

```
Session,Panelist,Product,Sweet,Sour,Bitter,Fruity,Floral,Spicy
1,P1,ProductA,5.2,3.8,2.1,6.5,4.3,3.1
1,P1,ProductB,6.4,4.2,1.8,5.9,3.7,2.5
1,P1,ProductC,4.1,5.3,3.2,4.8,2.9,4.1
1,P2,ProductA,4.8,3.5,2.3,6.8,4.5,3.3
```

## Technical Details

### Linear Mixed Model Implementation

- **Model**: `Rating ~ Product + (1|Panelist) + (1|Session) + (1|Panelist:Product)`
- **Estimation**: Restricted Maximum Likelihood (REML)
- **Degrees of Freedom**: Kenward-Roger approximation for small samples
- **Post-hoc**: Tukey HSD with family-wise error control

### Key Calculations

- **LSMeans**: `emmeans(model, "Product", lmer.df = "kenward-roger")`
- **F-test**: `F = MS_between / MS_within ~ F(df1, df2)`
- **Sensitivity (da)**: `da = intercept / √((1 + slope²) / 2)`
- **AUC**: Trapezoidal integration of ROC curve

### Validation

The dashboard implements methods consistent with:
- ISO 11132:2021 (Sensory analysis standards)
- Kenward & Roger (1997) small-sample inference
- Standard sensory evaluation practices

## Output Interpretation

### Primary Metrics

- **LSMeans**: Product mean ratings adjusted for random effects
- **F-statistic**: Tests H₀: All products have the same mean rating
- **p-value**: Statistical significance (p < 0.05 indicates product differences)
- **Variance Components**: Decomposition of variability sources
- **ICC**: Intraclass correlation (panelist consistency measure)

### Model Fit Statistics

- **Variance Components**: Panelist, Session, Interaction, Residual effects
- **p-values**: Fixed effects significance (p < 0.05 indicates good discrimination)
- **AIC/BIC**: Information criteria for model comparison
- **Convergence**: Optimization success indicator

## Dependencies

```r
shiny         # Web application framework
shinydashboard # Dashboard UI components
DT            # Interactive data tables
plotly        # Interactive plotting
ggplot2       # Static plotting
dplyr         # Data manipulation
tidyr         # Data tidying
lme4          # Linear mixed models
lmerTest      # Mixed model tests
emmeans       # Estimated marginal means
pbkrtest      # Kenward-Roger method
openxlsx      # Excel export
FactoMineR    # PCA analysis
factoextra    # PCA visualization
cluster       # Clustering algorithms
multcompView  # Compact letter display
nortest       # Normality tests
```

## File Structure

```
qda-sensory-dashboard/
│
├── app.R                 # Main Shiny application
├── README.md            # This file
├── example_data.csv     # Sample QDA data
└── LICENSE              # License file (if applicable)
```

## Contributing

Contributions are welcome! Please feel free to submit issues, feature requests, or pull requests.

### Development Guidelines

- Follow existing code style and commenting patterns
- Test with example data before submitting
- Ensure all new features include appropriate documentation
- Validate statistical implementations against published methods

## Citation

```
Kenward, M.G. & Roger, J.H. (1997). Small sample inference for fixed effects 
from restricted maximum likelihood. Biometrics, 53(3), 983-997.
```

## Acknowledgments

- Statistical methods based on ISO 11132:2021 standards
- Kenward-Roger implementation via pbkrtest package
- PCA and clustering methods from FactoMineR
- Implements statistical methods from sensory evaluation literature

## Contact

gehlotds1995@gmail.com

---

## Troubleshooting

### Common Issues

1. **Installation Problems**: Ensure all required packages are installed
2. **Data Upload Fails**: Check that first 3 columns are Session, Panelist, Product
3. **Model Convergence Warnings**: This is normal for descriptors with low variance (dashboard handles automatically)
4. **Download Not Working**: Ensure analysis has completed and tables are populated

### Getting Help

- Check the Statistical Theory section in the app
- Review the built-in help text and tooltips
- Submit issues via GitHub Issues tab

## For more information, please see the built-in Statistical Theory tab

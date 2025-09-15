# Sophrosyne
A comprehensive bias detection and fairness analysis tool built entirely within Google Apps Script. Automatically detects algorithmic bias in datasets, generates detailed reports with visualizations, and provides actionable recommendations for bias mitigation.
![imagealt](https://github.com/Mahabu-Subhani/Sophrosyne/blob/d22161fbc5e34bfc3715638eecaeb6b277cc6804/Sophrosyne%20AI%20-%20Dashboard.jpeg)
# Features

Core Bias Detection

Disparate Impact Analysis: Automated calculation of 80% rule compliance
Statistical Parity: Measures outcome differences across protected groups
Equal Opportunity: Evaluates fairness in positive prediction rates
Protected Attribute Detection: Automatically identifies gender, race, age, ethnicity, and other sensitive attributes
Target Variable Recognition: Detects prediction, score, outcome, and classification columns

![imagealt](https://github.com/Mahabu-Subhani/Sophrosyne/blob/5a50db89b69cb8a6644bc77537792d957728e656/Dashboard%20-%20Sophrosyne%20AI%20.jpeg)


Advanced Analytics

Intersectional Bias Analysis: Multi-dimensional bias detection across attribute combinations
Statistical Significance Testing: Chi-square, t-tests, and Kolmogorov-Smirnov tests
Temporal Bias Tracking: Monitor bias evolution over time periods
Feature Importance: Identify features most correlated with protected attributes
Individual Fairness: Assess similar treatment for similar individuals

![imagealt](https://github.com/Mahabu-Subhani/Sophrosyne/blob/fb128198bd8f457bd620793e718e7635cd9e6588/Report.jpeg)

Reporting & Visualization

Comprehensive Reports: Auto-generated bias analysis with executive summaries
Interactive Charts: Group distribution and bias severity visualizations
Risk Assessment: Color-coded bias flags and severity indicators
Historical Tracking: Maintain analysis history with trend identification
Actionable Recommendations: Specific steps for bias mitigation

![imagealt](https://github.com/Mahabu-Subhani/Sophrosyne/blob/a716536e670795b937505e0f885a639520e8c88c/Report%20History.jpeg)


# Bias Thresholds
``` javascript
BIAS_THRESHOLDS: {
  DISPARATE_IMPACT: 0.8,    // 80% rule compliance
  STATISTICAL_PARITY: 0.1,  // 10% difference threshold  
  EQUAL_OPPORTUNITY: 0.1    // 10% opportunity difference
}
```
# Custom Configuration
Modify the CONFIG object to customize:
```javascript
const CONFIG = {
  REPORT_SHEET_NAME: 'Custom_Report_Name',
  BIAS_THRESHOLDS: {
    DISPARATE_IMPACT: 0.75  // Stricter threshold
  }
};
```
# Core Functions
runBiasAnalysis()
Executes complete bias analysis pipeline

Detects protected attributes and target variables
Calculates fairness metrics
Generates comprehensive report
Updates analysis history

calculateBiasMetrics(data, columnAnalysis)
Performs statistical bias calculations
```javascript
Returns: {
  disparateImpact: number,
  statisticalParity: number, 
  equalOpportunity: number,
  biasSeverity: number
}
generateAIInsights(results)
Creates interpretive analysis of bias findings
javascriptReturns: Array<{
  type: string,
  message: string,
  confidence: number
}>
```
# Utility Functions
```
analyzeColumns(data)
Automatically detects column types and protected attributes
createBiasReport(biasResults, aiInsights, columnAnalysis)
Generates formatted bias analysis report with visualizations
```
# Reporting Issues
Please include:
```
Google Apps Script version,
Dataset size and structure,
Error messages or unexpected behavior,
Steps to reproduce the issue.
```
# Limitations
Technical Constraints
```javascript
Google Apps Script 6-minute execution limit
Maximum 50,000 rows for optimal performance
Limited to Google Sheets environment
No real-time processing capabilities
```
Statistical Limitations
```
1. Simplified individual fairness implementation
2. Basic counterfactual fairness approximation
3. Assumes binary classification for some metrics
4. Limited to demographic parity definitions
```
# License
MIT License - see LICENSE file for details

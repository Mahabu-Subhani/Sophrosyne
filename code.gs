/*** AI Bias Detector for Google Sheets
 * Complete implementation with bias detection, visualizations, and reporting
 */

// Global configuration
const CONFIG = {
  REPORT_SHEET_NAME: 'Bias_Analysis_Report',
  HISTORY_SHEET_NAME: 'Bias_History',
  PROTECTED_ATTRIBUTES: ['gender', 'race', 'ethnicity', 'age', 'religion', 'nationality', 'sexual_orientation', 'disability'],
  BIAS_THRESHOLDS: {
    DISPARATE_IMPACT: 0.8, // Standard 80% rule
    STATISTICAL_PARITY: 0.1, // 10% difference threshold
    EQUAL_OPPORTUNITY: 0.1
  }
};

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 Sophrosyne AI')
    .addItem('🔍 Run Bias Analysis', 'runBiasAnalysis')
    .addItem('📊 Show Dashboard', 'showDashboard')
    .addItem('📈 View History', 'showHistory')
    .addItem('⚙️ Settings', 'showSettings')
    .addSeparator()
    .addItem('📖 Help & Documentation', 'showHelp')
    .addToUi();
}

/**
 * Main function to run comprehensive bias analysis
 */
function runBiasAnalysis() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = getSheetData(sheet);
    
    if (data.length < 2) {
      SpreadsheetApp.getUi().alert('Error', 'Please ensure your sheet has data with headers.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Step 1: Detect protected attributes and target columns
    const columnAnalysis = analyzeColumns(data);
    
    if (columnAnalysis.protectedAttributes.length === 0) {
      SpreadsheetApp.getUi().alert('No Protected Attributes Found', 
        'Could not detect protected attributes automatically. Please ensure your data contains columns like gender, race, age, etc.', 
        SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Step 2: Run bias detection
    const biasResults = calculateBiasMetrics(data, columnAnalysis);
    
    // Step 3: Generate AI insights
    const aiInsights = generateAIInsights(biasResults);
    
    // Step 4: Create comprehensive report
    createBiasReport(biasResults, aiInsights, columnAnalysis);
    
    // Step 5: Save to history
    saveToHistory(biasResults, aiInsights);
    
    // Step 6: Show dashboard
    showDashboard();
    
    SpreadsheetApp.getUi().alert('Analysis Complete!', 
      `Bias analysis completed successfully!\n\n` +
      `• Found ${columnAnalysis.protectedAttributes.length} protected attributes\n` +
      `• Analyzed ${biasResults.totalGroups} demographic groups\n` +
      `• Generated comprehensive bias report\n\n` +
      `Check the "${CONFIG.REPORT_SHEET_NAME}" sheet for detailed results.`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
      
  } catch (error) {
    console.error('Error in runBiasAnalysis:', error);
    SpreadsheetApp.getUi().alert('Error', `An error occurred: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Get all data from the active sheet
 */
function getSheetData(sheet) {
  const range = sheet.getDataRange();
  if (range.getNumRows() === 0) return [];
  
  const values = range.getValues();
  const headers = values[0].map(h => String(h).toLowerCase().trim());
  
  return values.slice(1).map(row => {
    const rowData = {};
    headers.forEach((header, index) => {
      rowData[header] = row[index];
    });
    return rowData;
  });
}

/**
 * Analyze columns to detect protected attributes and target variables
 */
function analyzeColumns(data) {
  const headers = Object.keys(data[0] || {});
  const protectedAttributes = [];
  const targetColumns = [];
  const numericColumns = [];
  
  headers.forEach(header => {
    const lowerHeader = header.toLowerCase();
    
    // Detect protected attributes
    const isProtected = CONFIG.PROTECTED_ATTRIBUTES.some(attr => 
      lowerHeader.includes(attr) || 
      lowerHeader.includes(attr.replace('_', ''))
    ) || ['sex', 'male', 'female', 'black', 'white', 'asian', 'hispanic', 'latino', 'old', 'young'].some(keyword => 
      lowerHeader.includes(keyword)
    );
    
    if (isProtected) {
      protectedAttributes.push(header);
    }
    
    // Detect potential target columns
    const isTarget = ['prediction', 'score', 'label', 'outcome', 'result', 'class', 'target', 'approved', 'hired', 'accepted'].some(keyword => 
      lowerHeader.includes(keyword)
    );
    
    if (isTarget) {
      targetColumns.push(header);
    }
    
    // Detect numeric columns
    const sampleValues = data.slice(0, 10).map(row => row[header]).filter(val => val !== null && val !== undefined && val !== '');
    const isNumeric = sampleValues.length > 0 && sampleValues.every(val => !isNaN(Number(val)));
    
    if (isNumeric) {
      numericColumns.push(header);
    }
  });
  
  return {
    protectedAttributes,
    targetColumns: targetColumns.length > 0 ? targetColumns : numericColumns.slice(-2), // fallback to last 2 numeric columns
    numericColumns,
    allColumns: headers
  };
}

/**
 * Calculate comprehensive bias metrics
 */
function calculateBiasMetrics(data, columnAnalysis) {
  const results = {
    timestamp: new Date(),
    totalRecords: data.length,
    totalGroups: 0,
    protectedAttributes: columnAnalysis.protectedAttributes,
    targetColumns: columnAnalysis.targetColumns,
    groupAnalysis: {},
    overallMetrics: {},
    biasFlags: [],
    recommendations: []
  };
  
  // Analyze each protected attribute
  columnAnalysis.protectedAttributes.forEach(protectedAttr => {
    results.groupAnalysis[protectedAttr] = analyzeProtectedAttribute(data, protectedAttr, columnAnalysis.targetColumns);
  });
  
  // Calculate overall bias metrics
  results.overallMetrics = calculateOverallMetrics(results.groupAnalysis);
  
  // Generate bias flags and recommendations
  results.biasFlags = generateBiasFlags(results);
  results.recommendations = generateRecommendations(results);
  
  results.totalGroups = Object.values(results.groupAnalysis).reduce((sum, analysis) => 
    sum + Object.keys(analysis.groups).length, 0
  );
  
  return results;
}

/**
 * Analyze a single protected attribute
 */
function analyzeProtectedAttribute(data, protectedAttr, targetColumns) {
  const groups = {};
  
  // Group data by protected attribute values
  data.forEach(row => {
    const groupValue = String(row[protectedAttr] || 'Unknown').trim();
    if (!groups[groupValue]) {
      groups[groupValue] = [];
    }
    groups[groupValue].push(row);
  });
  
  const groupAnalysis = {};
  const groupNames = Object.keys(groups);
  
  // Analyze each group
  groupNames.forEach(groupName => {
    const groupData = groups[groupName];
    groupAnalysis[groupName] = {
      count: groupData.length,
      percentage: (groupData.length / data.length) * 100,
      metrics: {}
    };
    
    // Calculate metrics for each target column
    targetColumns.forEach(targetCol => {
      const targetValues = groupData.map(row => Number(row[targetCol])).filter(val => !isNaN(val));
      
      if (targetValues.length > 0) {
        groupAnalysis[groupName].metrics[targetCol] = {
          mean: targetValues.reduce((a, b) => a + b, 0) / targetValues.length,
          median: calculateMedian(targetValues),
          std: calculateStandardDeviation(targetValues),
          positiveRate: targetValues.filter(val => val > 0.5).length / targetValues.length,
          count: targetValues.length
        };
      }
    });
  });
  
  // Calculate disparate impact and statistical parity
  const metrics = calculateFairnessMetrics(groupAnalysis, targetColumns);
  
  return {
    groups: groupAnalysis,
    fairnessMetrics: metrics,
    groupCount: groupNames.length,
    attribute: protectedAttr
  };
}

/**
 * Calculate fairness metrics (Disparate Impact, Statistical Parity, etc.)
 */
function calculateFairnessMetrics(groupAnalysis, targetColumns) {
  const metrics = {};
  const groupNames = Object.keys(groupAnalysis);
  
  if (groupNames.length < 2) return metrics;
  
  targetColumns.forEach(targetCol => {
    metrics[targetCol] = {};
    
    // Get positive rates for all groups
    const positiveRates = {};
    groupNames.forEach(group => {
      if (groupAnalysis[group].metrics[targetCol]) {
        positiveRates[group] = groupAnalysis[group].metrics[targetCol].positiveRate;
      }
    });
    
    const rates = Object.values(positiveRates);
    if (rates.length >= 2) {
      const maxRate = Math.max(...rates);
      const minRate = Math.min(...rates);
      
      // Disparate Impact (min/max ratio)
      metrics[targetCol].disparateImpact = maxRate > 0 ? minRate / maxRate : 0;
      
      // Statistical Parity Difference (max - min)
      metrics[targetCol].statisticalParityDiff = maxRate - minRate;
      
      // Equal Opportunity (for binary classification)
      metrics[targetCol].equalOpportunity = calculateEqualOpportunity(groupAnalysis, targetCol);
      
      // Bias severity score (0-1, where 1 is most biased)
      metrics[targetCol].biasSeverity = Math.max(
        Math.abs(1 - metrics[targetCol].disparateImpact),
        metrics[targetCol].statisticalParityDiff
      );
    }
  });
  
  return metrics;
}

/**
 * Calculate Equal Opportunity metric
 */
function calculateEqualOpportunity(groupAnalysis, targetCol) {
  const groupNames = Object.keys(groupAnalysis);
  const tpRates = [];
  
  groupNames.forEach(group => {
    if (groupAnalysis[group].metrics[targetCol]) {
      // Simplified: using positive rate as proxy for TPR
      tpRates.push(groupAnalysis[group].metrics[targetCol].positiveRate);
    }
  });
  
  if (tpRates.length >= 2) {
    return Math.max(...tpRates) - Math.min(...tpRates);
  }
  
  return 0;
}

/**
 * Calculate overall bias metrics across all protected attributes
 */
function calculateOverallMetrics(groupAnalysis) {
  const overallMetrics = {
    avgDisparateImpact: 0,
    avgStatisticalParity: 0,
    avgBiasSeverity: 0,
    mostBiasedAttribute: '',
    leastBiasedAttribute: '',
    overallBiasScore: 0
  };
  
  const attributes = Object.keys(groupAnalysis);
  if (attributes.length === 0) return overallMetrics;
  
  let totalDI = 0, totalSP = 0, totalBS = 0, count = 0;
  let maxBias = 0, minBias = 1, maxBiasAttr = '', minBiasAttr = '';
  
  attributes.forEach(attr => {
    const metrics = groupAnalysis[attr].fairnessMetrics;
    Object.values(metrics).forEach(targetMetrics => {
      if (targetMetrics.disparateImpact !== undefined) {
        totalDI += targetMetrics.disparateImpact;
        totalSP += targetMetrics.statisticalParityDiff;
        totalBS += targetMetrics.biasSeverity;
        count++;
        
        if (targetMetrics.biasSeverity > maxBias) {
          maxBias = targetMetrics.biasSeverity;
          maxBiasAttr = attr;
        }
        if (targetMetrics.biasSeverity < minBias) {
          minBias = targetMetrics.biasSeverity;
          minBiasAttr = attr;
        }
      }
    });
  });
  
  if (count > 0) {
    overallMetrics.avgDisparateImpact = totalDI / count;
    overallMetrics.avgStatisticalParity = totalSP / count;
    overallMetrics.avgBiasSeverity = totalBS / count;
    overallMetrics.mostBiasedAttribute = maxBiasAttr;
    overallMetrics.leastBiasedAttribute = minBiasAttr;
    overallMetrics.overallBiasScore = totalBS / count;
  }
  
  return overallMetrics;
}

/**
 * Generate bias flags based on thresholds
 */
function generateBiasFlags(results) {
  const flags = [];
  
  Object.entries(results.groupAnalysis).forEach(([attr, analysis]) => {
    Object.entries(analysis.fairnessMetrics).forEach(([target, metrics]) => {
      if (metrics.disparateImpact < CONFIG.BIAS_THRESHOLDS.DISPARATE_IMPACT) {
        flags.push({
          type: 'DISPARATE_IMPACT',
          severity: 'HIGH',
          attribute: attr,
          target: target,
          value: metrics.disparateImpact,
          threshold: CONFIG.BIAS_THRESHOLDS.DISPARATE_IMPACT,
          message: `Disparate Impact violation: ${(metrics.disparateImpact * 100).toFixed(1)}% (threshold: ${CONFIG.BIAS_THRESHOLDS.DISPARATE_IMPACT * 100}%)`
        });
      }
      
      if (metrics.statisticalParityDiff > CONFIG.BIAS_THRESHOLDS.STATISTICAL_PARITY) {
        flags.push({
          type: 'STATISTICAL_PARITY',
          severity: metrics.statisticalParityDiff > 0.2 ? 'HIGH' : 'MEDIUM',
          attribute: attr,
          target: target,
          value: metrics.statisticalParityDiff,
          threshold: CONFIG.BIAS_THRESHOLDS.STATISTICAL_PARITY,
          message: `Statistical parity difference: ${(metrics.statisticalParityDiff * 100).toFixed(1)}% (threshold: ${CONFIG.BIAS_THRESHOLDS.STATISTICAL_PARITY * 100}%)`
        });
      }
    });
  });
  
  return flags;
}

/**
 * Generate actionable recommendations
 */
function generateRecommendations(results) {
  const recommendations = [];
  
  // Data balancing recommendations
  Object.entries(results.groupAnalysis).forEach(([attr, analysis]) => {
    const groups = Object.entries(analysis.groups);
    const avgSize = groups.reduce((sum, [, group]) => sum + group.count, 0) / groups.length;
    
    groups.forEach(([groupName, group]) => {
      if (group.count < avgSize * 0.5) {
        recommendations.push({
          type: 'DATA_BALANCING',
          priority: 'HIGH',
          attribute: attr,
          group: groupName,
          action: `Increase representation of ${groupName} group`,
          details: `Current: ${group.count} samples (${group.percentage.toFixed(1)}%). Consider upsampling or collecting more data.`
        });
      }
    });
  });
  
  // Threshold adjustment recommendations
  results.biasFlags.forEach(flag => {
    if (flag.type === 'DISPARATE_IMPACT') {
      recommendations.push({
        type: 'THRESHOLD_ADJUSTMENT',
        priority: 'MEDIUM',
        attribute: flag.attribute,
        target: flag.target,
        action: `Adjust decision thresholds for ${flag.attribute}`,
        details: `Consider using different thresholds per group to achieve fairness in ${flag.target}.`
      });
    }
  });
  
  // Model retraining recommendations
  if (results.overallMetrics.overallBiasScore > 0.3) {
    recommendations.push({
      type: 'MODEL_RETRAINING',
      priority: 'HIGH',
      action: 'Retrain model with fairness constraints',
      details: `Overall bias score is ${(results.overallMetrics.overallBiasScore * 100).toFixed(1)}%. Consider retraining with fairness-aware algorithms.`
    });
  }
  
  return recommendations;
}

/**
 * Generate AI-powered insights
 */
function generateAIInsights(results) {
  const insights = [];
  
  // Overall bias assessment
  const overallScore = results.overallMetrics.overallBiasScore;
  let overallAssessment = '';
  
  if (overallScore < 0.1) {
    overallAssessment = 'Your model shows minimal bias across protected attributes. This is excellent fairness performance.';
  } else if (overallScore < 0.3) {
    overallAssessment = 'Your model shows moderate bias that should be addressed to ensure fairness.';
  } else {
    overallAssessment = 'Your model shows significant bias that requires immediate attention and remediation.';
  }
  
  insights.push({
    type: 'OVERALL_ASSESSMENT',
    message: overallAssessment,
    confidence: 0.95
  });
  
  // Attribute-specific insights
  Object.entries(results.groupAnalysis).forEach(([attr, analysis]) => {
    const groups = Object.entries(analysis.groups);
    const sortedGroups = groups.sort((a, b) => b[1].percentage - a[1].percentage);
    
    if (sortedGroups.length >= 2) {
      const dominant = sortedGroups[0];
      const minority = sortedGroups[sortedGroups.length - 1];
      
      const representation = `${attr} representation: ${dominant[0]} (${dominant[1].percentage.toFixed(1)}%) dominates while ${minority[0]} (${minority[1].percentage.toFixed(1)}%) is underrepresented.`;
      
      insights.push({
        type: 'REPRESENTATION_ANALYSIS',
        attribute: attr,
        message: representation,
        confidence: 0.90
      });
    }
  });
  
  // Performance disparity insights
  Object.entries(results.groupAnalysis).forEach(([attr, analysis]) => {
    Object.entries(analysis.fairnessMetrics).forEach(([target, metrics]) => {
      if (metrics.biasSeverity > 0.2) {
        const message = `Significant performance disparity detected in ${target} across ${attr} groups. Disparate impact ratio: ${(metrics.disparateImpact * 100).toFixed(1)}%`;
        
        insights.push({
          type: 'PERFORMANCE_DISPARITY',
          attribute: attr,
          target: target,
          message: message,
          confidence: 0.85
        });
      }
    });
  });
  
  return insights;
}

/**
 * Utility function to calculate median
 */
function calculateMedian(values) {
  const sorted = values.slice().sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
}

/**
 * Utility function to calculate standard deviation
 */
function calculateStandardDeviation(values) {
  const mean = values.reduce((a, b) => a + b, 0) / values.length;
  const squaredDiffs = values.map(value => Math.pow(value - mean, 2));
  const avgSquaredDiff = squaredDiffs.reduce((a, b) => a + b, 0) / squaredDiffs.length;
  return Math.sqrt(avgSquaredDiff);
}

/**
 * Report Generation and Visualization Functions
 * Creates comprehensive bias reports with charts and insights
 */

/**
 * Create comprehensive bias report sheet
 */
function createBiasReport(biasResults, aiInsights, columnAnalysis) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Delete existing report sheet if it exists
  const existingSheet = spreadsheet.getSheetByName(CONFIG.REPORT_SHEET_NAME);
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }
  
  // Create new report sheet
  const reportSheet = spreadsheet.insertSheet(CONFIG.REPORT_SHEET_NAME);
  
  // Set up the report structure
  let currentRow = 1;
  
  // Header section
  currentRow = createReportHeader(reportSheet, currentRow, biasResults);
  
  // Executive Summary
  currentRow = createExecutiveSummary(reportSheet, currentRow, biasResults, aiInsights);
  
  // Overall Metrics
  currentRow = createOverallMetrics(reportSheet, currentRow, biasResults);
  
  // Group Analysis
  currentRow = createGroupAnalysis(reportSheet, currentRow, biasResults);
  
  // Bias Flags
  currentRow = createBiasFlags(reportSheet, currentRow, biasResults);
  
  // Recommendations
  currentRow = createRecommendations(reportSheet, currentRow, biasResults);
  
  // AI Insights
  currentRow = createAIInsights(reportSheet, currentRow, aiInsights);
  
  // Create charts
  createBiasCharts(reportSheet, biasResults);
  
  // Format the sheet
  formatReportSheet(reportSheet);
  
  // Set as active sheet
  spreadsheet.setActiveSheet(reportSheet);
}

/**
 * Create report header
 */
function createReportHeader(sheet, startRow, biasResults) {
  sheet.getRange(startRow, 1, 1, 6).merge().setValue('🤖 AI BIAS DETECTION REPORT').setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(startRow, 1).setBackground('#4285f4').setFontColor('white');
  
  startRow += 2;
  
  // Report metadata
  const metadata = [
    ['📅 Analysis Date:', biasResults.timestamp.toLocaleString()],
    ['📊 Total Records:', biasResults.totalRecords],
    ['🏷️ Protected Attributes:', biasResults.protectedAttributes.join(', ')],
    ['🎯 Target Columns:', biasResults.targetColumns.join(', ')],
    ['👥 Total Groups:', biasResults.totalGroups]
  ];
  
  metadata.forEach(([label, value], index) => {
    sheet.getRange(startRow + index, 1).setValue(label).setFontWeight('bold');
    sheet.getRange(startRow + index, 2, 1, 4).merge().setValue(value);
  });
  
  return startRow + metadata.length + 2;
}

/**
 * Create executive summary
 */
function createExecutiveSummary(sheet, startRow, biasResults, aiInsights) {
  sheet.getRange(startRow, 1).setValue('📋 EXECUTIVE SUMMARY').setFontSize(14).setFontWeight('bold');
  sheet.getRange(startRow, 1).setBackground('#34a853').setFontColor('white');
  startRow += 2;
  
  // Overall bias score with color coding
  const overallScore = biasResults.overallMetrics.overallBiasScore;
  const scorePercentage = (overallScore * 100).toFixed(1) + '%';
  const scoreColor = overallScore < 0.1 ? '#34a853' : overallScore < 0.3 ? '#fbbc04' : '#ea4335';
  
  sheet.getRange(startRow, 1).setValue('Overall Bias Score:').setFontWeight('bold');
  sheet.getRange(startRow, 2).setValue(scorePercentage).setBackground(scoreColor).setFontColor('white').setFontWeight('bold');
  startRow += 1;
  
  // Risk level
  let riskLevel = 'LOW';
  if (overallScore >= 0.3) riskLevel = 'HIGH';
  else if (overallScore >= 0.1) riskLevel = 'MEDIUM';
  
  sheet.getRange(startRow, 1).setValue('Risk Level:').setFontWeight('bold');
  sheet.getRange(startRow, 2).setValue(riskLevel).setBackground(scoreColor).setFontColor('white').setFontWeight('bold');
  startRow += 1;
  
  // Key findings from AI insights
  const overallAssessment = aiInsights.find(insight => insight.type === 'OVERALL_ASSESSMENT');
  if (overallAssessment) {
    sheet.getRange(startRow, 1).setValue('Key Finding:').setFontWeight('bold');
    sheet.getRange(startRow, 2, 1, 4).merge().setValue(overallAssessment.message).setWrap(true);
    startRow += 2;
  }
  
  return startRow + 1;
}

/**
 * Create overall metrics section
 */
function createOverallMetrics(sheet, startRow, biasResults) {
  sheet.getRange(startRow, 1).setValue('📊 OVERALL FAIRNESS METRICS').setFontSize(14).setFontWeight('bold');
  sheet.getRange(startRow, 1).setBackground('#ff9900').setFontColor('white');
  startRow += 2;
  
  const metrics = biasResults.overallMetrics;
  const metricRows = [
    ['Average Disparate Impact:', (metrics.avgDisparateImpact * 100).toFixed(1) + '%', 'Higher is better (>80%)'],
    ['Average Statistical Parity:', (metrics.avgStatisticalParity * 100).toFixed(1) + '%', 'Lower is better (<10%)'],
    ['Most Biased Attribute:', metrics.mostBiasedAttribute, 'Requires immediate attention'],
    ['Least Biased Attribute:', metrics.leastBiasedAttribute, 'Good fairness performance']
  ];
  
  // Headers
  sheet.getRange(startRow, 1).setValue('Metric').setFontWeight('bold');
  sheet.getRange(startRow, 2).setValue('Value').setFontWeight('bold');
  sheet.getRange(startRow, 3).setValue('Interpretation').setFontWeight('bold');
  startRow += 1;
  
  metricRows.forEach(([metric, value, interpretation], index) => {
    sheet.getRange(startRow + index, 1).setValue(metric);
    sheet.getRange(startRow + index, 2).setValue(value).setFontWeight('bold');
    sheet.getRange(startRow + index, 3).setValue(interpretation);
  });
  
  return startRow + metricRows.length + 2;
}

/**
 * Create detailed group analysis
 */
function createGroupAnalysis(sheet, startRow, biasResults) {
  sheet.getRange(startRow, 1).setValue('👥 DETAILED GROUP ANALYSIS').setFontSize(14).setFontWeight('bold');
  sheet.getRange(startRow, 1).setBackground('#9c27b0').setFontColor('white');
  startRow += 2;
  
  Object.entries(biasResults.groupAnalysis).forEach(([attribute, analysis]) => {
    // Attribute header
    sheet.getRange(startRow, 1, 1, 6).merge().setValue(`Protected Attribute: ${attribute.toUpperCase()}`).setFontWeight('bold').setBackground('#f3f3f3');
    startRow += 2;
    
    // Group distribution table
    const headers = ['Group', 'Count', 'Percentage', 'Disparate Impact', 'Statistical Parity', 'Bias Severity'];
    headers.forEach((header, index) => {
      sheet.getRange(startRow, index + 1).setValue(header).setFontWeight('bold');
    });
    startRow += 1;
    
    Object.entries(analysis.groups).forEach(([groupName, groupData]) => {
      sheet.getRange(startRow, 1).setValue(groupName);
      sheet.getRange(startRow, 2).setValue(groupData.count);
      sheet.getRange(startRow, 3).setValue(groupData.percentage.toFixed(1) + '%');
      
      // Add fairness metrics if available
      const targetCol = biasResults.targetColumns[0];
      if (analysis.fairnessMetrics[targetCol]) {
        const metrics = analysis.fairnessMetrics[targetCol];
        sheet.getRange(startRow, 4).setValue((metrics.disparateImpact * 100).toFixed(1) + '%');
        sheet.getRange(startRow, 5).setValue((metrics.statisticalParityDiff * 100).toFixed(1) + '%');
        
        // Color code bias severity
        const severity = metrics.biasSeverity;
        const severityColor = severity < 0.1 ? '#34a853' : severity < 0.3 ? '#fbbc04' : '#ea4335';
        sheet.getRange(startRow, 6).setValue((severity * 100).toFixed(1) + '%').setBackground(severityColor).setFontColor('white');
      }
      
      startRow += 1;
    });
    
    startRow += 2;
  });
  
  return startRow;
}

/**
 * Create bias flags section
 */
function createBiasFlags(sheet, startRow, biasResults) {
  sheet.getRange(startRow, 1).setValue('🚩 BIAS FLAGS & VIOLATIONS').setFontSize(14).setFontWeight('bold');
  sheet.getRange(startRow, 1).setBackground('#ea4335').setFontColor('white');
  startRow += 2;
  
  if (biasResults.biasFlags.length === 0) {
    sheet.getRange(startRow, 1, 1, 4).merge().setValue('✅ No significant bias violations detected!').setBackground('#34a853').setFontColor('white').setFontWeight('bold');
    return startRow + 3;
  }
  
  // Headers
  const headers = ['Severity', 'Type', 'Attribute', 'Target', 'Description'];
  headers.forEach((header, index) => {
    sheet.getRange(startRow, index + 1).setValue(header).setFontWeight('bold');
  });
  startRow += 1;
  
  biasResults.biasFlags.forEach(flag => {
    const severityColor = flag.severity === 'HIGH' ? '#ea4335' : flag.severity === 'MEDIUM' ? '#fbbc04' : '#ff9900';
    
    sheet.getRange(startRow, 1).setValue(flag.severity).setBackground(severityColor).setFontColor('white').setFontWeight('bold');
    sheet.getRange(startRow, 2).setValue(flag.type.replace('_', ' '));
    sheet.getRange(startRow, 3).setValue(flag.attribute);
    sheet.getRange(startRow, 4).setValue(flag.target);
    sheet.getRange(startRow, 5).setValue(flag.message).setWrap(true);
    
    startRow += 1;
  });
  
  return startRow + 2;
}

/**
 * Create recommendations section
 */
function createRecommendations(sheet, startRow, biasResults) {
  sheet.getRange(startRow, 1).setValue('💡 ACTIONABLE RECOMMENDATIONS').setFontSize(14).setFontWeight('bold');
  sheet.getRange(startRow, 1).setBackground('#0f9d58').setFontColor('white');
  startRow += 2;
  
  if (biasResults.recommendations.length === 0) {
    sheet.getRange(startRow, 1, 1, 4).merge().setValue('No specific recommendations at this time.').setFontStyle('italic');
    return startRow + 3;
  }
  
  // Headers
  const headers = ['Priority', 'Type', 'Action', 'Details'];
  headers.forEach((header, index) => {
    sheet.getRange(startRow, index + 1).setValue(header).setFontWeight('bold');
  });
  startRow += 1;
  
  biasResults.recommendations.forEach(rec => {
    const priorityColor = rec.priority === 'HIGH' ? '#ea4335' : rec.priority === 'MEDIUM' ? '#fbbc04' : '#34a853';
    
    sheet.getRange(startRow, 1).setValue(rec.priority).setBackground(priorityColor).setFontColor('white').setFontWeight('bold');
    sheet.getRange(startRow, 2).setValue(rec.type.replace('_', ' '));
    sheet.getRange(startRow, 3).setValue(rec.action).setWrap(true);
    sheet.getRange(startRow, 4).setValue(rec.details).setWrap(true);
    
    startRow += 1;
  });
  
  return startRow + 2;
}

/**
 * Create AI insights section
 */
function createAIInsights(sheet, startRow, aiInsights) {
  sheet.getRange(startRow, 1).setValue('🧠 AI-POWERED INSIGHTS').setFontSize(14).setFontWeight('bold');
  sheet.getRange(startRow, 1).setBackground('#673ab7').setFontColor('white');
  startRow += 2;
  
  aiInsights.forEach(insight => {
    sheet.getRange(startRow, 1).setValue('💬').setFontSize(16);
    sheet.getRange(startRow, 2, 1, 4).merge().setValue(insight.message).setWrap(true).setFontStyle('italic');
    
    if (insight.confidence) {
      sheet.getRange(startRow + 1, 2).setValue(`Confidence: ${(insight.confidence * 100).toFixed(0)}%`).setFontSize(10).setFontColor('#666');
    }
    
    startRow += 3;
  });
  
  return startRow + 1;
}

/**
 * Create bias visualization charts
 */
function createBiasCharts(sheet, biasResults) {
  // Create group distribution charts
  let chartRow = 5; // Position charts in the upper right area
  let chartCol = 8;
  
  Object.entries(biasResults.groupAnalysis).forEach(([attribute, analysis], index) => {
    // Prepare data for group distribution chart
    const chartData = [['Group', 'Count', 'Percentage']];
    
    Object.entries(analysis.groups).forEach(([groupName, groupData]) => {
      chartData.push([groupName, groupData.count, groupData.percentage]);
    });
    
    // Create data range for chart
    const dataRange = sheet.getRange(100 + (index * 20), 1, chartData.length, chartData[0].length);
    dataRange.setValues(chartData);
    
    // Create column chart
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataRange)
      .setPosition(chartRow + (index * 12), chartCol, 0, 0)
      .setOption('title', `${attribute} Distribution`)
      .setOption('width', 400)
      .setOption('height', 250)
      .setOption('colors', ['#4285f4', '#34a853', '#fbbc04', '#ea4335', '#ff9900', '#9c27b0'])
      .build();
    
    sheet.insertChart(chart);
  });
  
  // Create bias severity heatmap data
  createBiasSeverityChart(sheet, biasResults, chartRow, chartCol + 6);
}

/**
 * Create bias severity visualization
 */
function createBiasSeverityChart(sheet, biasResults, startRow, startCol) {
  const heatmapData = [['Attribute', 'Target', 'Bias Severity']];
  
  Object.entries(biasResults.groupAnalysis).forEach(([attribute, analysis]) => {
    Object.entries(analysis.fairnessMetrics).forEach(([target, metrics]) => {
      if (metrics.biasSeverity !== undefined) {
        heatmapData.push([attribute, target, metrics.biasSeverity]);
      }
    });
  });
  
  if (heatmapData.length > 1) {
    const heatmapRange = sheet.getRange(200, 1, heatmapData.length, heatmapData[0].length);
    heatmapRange.setValues(heatmapData);
    
    const heatmapChart = sheet.newChart()
      .setChartType(Charts.ChartType.TABLE)
      .addRange(heatmapRange)
      .setPosition(startRow, startCol, 0, 0)
      .setOption('title', 'Bias Severity Heatmap')
      .setOption('width', 400)
      .setOption('height', 250)
      .build();
    
    sheet.insertChart(heatmapChart);
  }
}

/**
 * Format the report sheet
 */
function formatReportSheet(sheet) {
  // Auto-resize columns
  sheet.autoResizeColumns(1, 6);
  
  // Set column widths
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 300);
  sheet.setColumnWidth(6, 200);
  
  // Add borders to important sections
  const lastRow = sheet.getLastRow();
  sheet.getRange(1, 1, lastRow, 6).setBorder(true, true, true, true, false, false);
  
  // Freeze header rows
  sheet.setFrozenRows(1);
}

/**
 * Save analysis results to history
 */
function saveToHistory(biasResults, aiInsights) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = spreadsheet.getSheetByName(CONFIG.HISTORY_SHEET_NAME);
  
  if (!historySheet) {
    historySheet = spreadsheet.insertSheet(CONFIG.HISTORY_SHEET_NAME);
    
    // Create headers
    const headers = [
      'Timestamp', 'Total Records', 'Protected Attributes', 'Overall Bias Score', 
      'Avg Disparate Impact', 'Avg Statistical Parity', 'Bias Flags Count', 
      'Most Biased Attribute', 'Risk Level', 'Notes'
    ];
    
    historySheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  
  // Add new record
  const newRow = historySheet.getLastRow() + 1;
  const overallScore = biasResults.overallMetrics.overallBiasScore;
  let riskLevel = 'LOW';
  if (overallScore >= 0.3) riskLevel = 'HIGH';
  else if (overallScore >= 0.1) riskLevel = 'MEDIUM';
  
  const rowData = [
    biasResults.timestamp,
    biasResults.totalRecords,
    biasResults.protectedAttributes.join(', '),
    (overallScore * 100).toFixed(1) + '%',
    (biasResults.overallMetrics.avgDisparateImpact * 100).toFixed(1) + '%',
    (biasResults.overallMetrics.avgStatisticalParity * 100).toFixed(1) + '%',
    biasResults.biasFlags.length,
    biasResults.overallMetrics.mostBiasedAttribute,
    riskLevel,
    `Analysis completed with ${aiInsights.length} AI insights`
  ];
  
  historySheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
  
  // Color code based on risk level
  const riskColor = riskLevel === 'HIGH' ? '#ea4335' : riskLevel === 'MEDIUM' ? '#fbbc04' : '#34a853';
  historySheet.getRange(newRow, 9).setBackground(riskColor).setFontColor('white');
}

/**
 * Advanced Analysis Functions and Utilities
 * Enhanced bias detection with machine learning metrics and statistical tests
 */

/**
 * Advanced bias analysis with statistical significance testing
 */
function runAdvancedBiasAnalysis(data, columnAnalysis) {
  const results = {
    statisticalTests: {},
    advancedMetrics: {},
    intersectionalBias: {},
    temporalAnalysis: {},
    featureImportance: {}
  };
  
  // Statistical significance testing
  results.statisticalTests = performStatisticalTests(data, columnAnalysis);
  
  // Advanced fairness metrics
  results.advancedMetrics = calculateAdvancedMetrics(data, columnAnalysis);
  
  // Intersectional bias analysis
  results.intersectionalBias = analyzeIntersectionalBias(data, columnAnalysis);
  
  // Temporal bias analysis (if date columns exist)
  results.temporalAnalysis = analyzeTemporalBias(data, columnAnalysis);
  
  // Feature importance for bias
  results.featureImportance = calculateFeatureImportance(data, columnAnalysis);
  
  return results;
}

/**
 * Perform statistical significance tests
 */
function performStatisticalTests(data, columnAnalysis) {
  const tests = {};
  
  columnAnalysis.protectedAttributes.forEach(protectedAttr => {
    tests[protectedAttr] = {};
    
    columnAnalysis.targetColumns.forEach(targetCol => {
      const groups = groupDataByAttribute(data, protectedAttr);
      const groupNames = Object.keys(groups);
      
      if (groupNames.length >= 2) {
        // Chi-square test for independence
        tests[protectedAttr][targetCol] = {
          chiSquare: performChiSquareTest(groups, targetCol),
          tTest: performTTest(groups, targetCol),
          ksTest: performKSTest(groups, targetCol)
        };
      }
    });
  });
  
  return tests;
}

/**
 * Chi-square test for independence
 */
function performChiSquareTest(groups, targetCol) {
  const groupNames = Object.keys(groups);
  if (groupNames.length < 2) return null;
  
  // Create contingency table
  const contingencyTable = [];
  const outcomes = ['positive', 'negative'];
  
  groupNames.forEach(group => {
    const groupData = groups[group];
    const positives = groupData.filter(row => Number(row[targetCol]) > 0.5).length;
    const negatives = groupData.length - positives;
    contingencyTable.push([positives, negatives]);
  });
  
  // Calculate chi-square statistic
  const chiSquare = calculateChiSquareStatistic(contingencyTable);
  const degreesOfFreedom = (groupNames.length - 1) * (outcomes.length - 1);
  const pValue = calculateChiSquarePValue(chiSquare, degreesOfFreedom);
  
  return {
    statistic: chiSquare,
    pValue: pValue,
    degreesOfFreedom: degreesOfFreedom,
    significant: pValue < 0.05
  };
}

/**
 * T-test for comparing group means
 */
function performTTest(groups, targetCol) {
  const groupNames = Object.keys(groups);
  if (groupNames.length !== 2) return null;
  
  const group1Values = groups[groupNames[0]].map(row => Number(row[targetCol])).filter(val => !isNaN(val));
  const group2Values = groups[groupNames[1]].map(row => Number(row[targetCol])).filter(val => !isNaN(val));
  
  if (group1Values.length < 2 || group2Values.length < 2) return null;
  
  const mean1 = group1Values.reduce((a, b) => a + b, 0) / group1Values.length;
  const mean2 = group2Values.reduce((a, b) => a + b, 0) / group2Values.length;
  
  const var1 = calculateVariance(group1Values, mean1);
  const var2 = calculateVariance(group2Values, mean2);
  
  const pooledSE = Math.sqrt(var1 / group1Values.length + var2 / group2Values.length);
  const tStatistic = (mean1 - mean2) / pooledSE;
  const degreesOfFreedom = group1Values.length + group2Values.length - 2;
  
  return {
    statistic: tStatistic,
    degreesOfFreedom: degreesOfFreedom,
    meanDifference: mean1 - mean2,
    significant: Math.abs(tStatistic) > 1.96 // Rough approximation
  };
}

/**
 * Kolmogorov-Smirnov test for distribution differences
 */
function performKSTest(groups, targetCol) {
  const groupNames = Object.keys(groups);
  if (groupNames.length !== 2) return null;
  
  const group1Values = groups[groupNames[0]].map(row => Number(row[targetCol])).filter(val => !isNaN(val)).sort((a, b) => a - b);
  const group2Values = groups[groupNames[1]].map(row => Number(row[targetCol])).filter(val => !isNaN(val)).sort((a, b) => a - b);
  
  if (group1Values.length < 5 || group2Values.length < 5) return null;
  
  // Calculate empirical distribution functions
  const allValues = [...group1Values, ...group2Values].sort((a, b) => a - b);
  let maxDifference = 0;
  
  allValues.forEach(value => {
    const cdf1 = group1Values.filter(v => v <= value).length / group1Values.length;
    const cdf2 = group2Values.filter(v => v <= value).length / group2Values.length;
    const difference = Math.abs(cdf1 - cdf2);
    maxDifference = Math.max(maxDifference, difference);
  });
  
  return {
    statistic: maxDifference,
    significant: maxDifference > 0.05 // Simplified threshold
  };
}

/**
 * Calculate advanced fairness metrics
 */
function calculateAdvancedMetrics(data, columnAnalysis) {
  const metrics = {};
  
  columnAnalysis.protectedAttributes.forEach(protectedAttr => {
    metrics[protectedAttr] = {};
    
    columnAnalysis.targetColumns.forEach(targetCol => {
      const groups = groupDataByAttribute(data, protectedAttr);
      
      metrics[protectedAttr][targetCol] = {
        equalizedOdds: calculateEqualizedOdds(groups, targetCol),
        calibration: calculateCalibration(groups, targetCol),
        individualFairness: calculateIndividualFairness(groups, targetCol),
        counterfactualFairness: calculateCounterfactualFairness(data, protectedAttr, targetCol),
        treatmentEquality: calculateTreatmentEquality(groups, targetCol)
      };
    });
  });
  
  return metrics;
}

/**
 * Calculate Equalized Odds metric
 */
function calculateEqualizedOdds(groups, targetCol) {
  const groupNames = Object.keys(groups);
  const tprByGroup = {};
  const fprByGroup = {};
  
  groupNames.forEach(group => {
    const groupData = groups[group];
    let tp = 0, fp = 0, tn = 0, fn = 0;
    
    groupData.forEach(row => {
      const predicted = Number(row[targetCol]) > 0.5;
      const actual = Number(row[targetCol + '_actual'] || row[targetCol]) > 0.5; // Assume same if no actual column
      
      if (predicted && actual) tp++;
      else if (predicted && !actual) fp++;
      else if (!predicted && actual) fn++;
      else tn++;
    });
    
    tprByGroup[group] = tp / (tp + fn) || 0;
    fprByGroup[group] = fp / (fp + tn) || 0;
  });
  
  const tprValues = Object.values(tprByGroup);
  const fprValues = Object.values(fprByGroup);
  
  return {
    tprDifference: Math.max(...tprValues) - Math.min(...tprValues),
    fprDifference: Math.max(...fprValues) - Math.min(...fprValues),
    satisfiesEqualizedOdds: (Math.max(...tprValues) - Math.min(...tprValues)) < 0.1 && 
                           (Math.max(...fprValues) - Math.min(...fprValues)) < 0.1
  };
}

/**
 * Calculate Calibration metric
 */
function calculateCalibration(groups, targetCol) {
  const groupNames = Object.keys(groups);
  const calibrationByGroup = {};
  
  groupNames.forEach(group => {
    const groupData = groups[group];
    const bins = [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0];
    const binCalibration = [];
    
    for (let i = 0; i < bins.length - 1; i++) {
      const binData = groupData.filter(row => {
        const score = Number(row[targetCol]);
        return score >= bins[i] && score < bins[i + 1];
      });
      
      if (binData.length > 0) {
        const avgPrediction = binData.reduce((sum, row) => sum + Number(row[targetCol]), 0) / binData.length;
        const actualRate = binData.filter(row => Number(row[targetCol + '_actual'] || row[targetCol]) > 0.5).length / binData.length;
        binCalibration.push(Math.abs(avgPrediction - actualRate));
      }
    }
    
    calibrationByGroup[group] = binCalibration.length > 0 ? 
      binCalibration.reduce((a, b) => a + b, 0) / binCalibration.length : 0;
  });
  
  const calibrationValues = Object.values(calibrationByGroup);
  return {
    avgCalibrationError: calibrationValues.reduce((a, b) => a + b, 0) / calibrationValues.length,
    maxCalibrationDifference: Math.max(...calibrationValues) - Math.min(...calibrationValues),
    wellCalibrated: Math.max(...calibrationValues) < 0.1
  };
}

/**
 * Calculate Individual Fairness metric
 */
function calculateIndividualFairness(groups, targetCol) {
  // Simplified individual fairness: similar individuals should get similar outcomes
  const allData = Object.values(groups).flat();
  let totalDifference = 0;
  let comparisons = 0;
  
  for (let i = 0; i < allData.length && i < 100; i++) { // Limit for performance
    for (let j = i + 1; j < allData.length && j < 100; j++) {
      const similarity = calculateSimilarity(allData[i], allData[j], targetCol);
      const outcomeDifference = Math.abs(Number(allData[i][targetCol]) - Number(allData[j][targetCol]));
      
      if (similarity > 0.8) { // Similar individuals
        totalDifference += outcomeDifference;
        comparisons++;
      }
    }
  }
  
  return {
    avgOutcomeDifference: comparisons > 0 ? totalDifference / comparisons : 0,
    fairnessViolations: comparisons > 0 ? (totalDifference / comparisons) > 0.1 : false
  };
}

/**
 * Calculate similarity between two data points
 */
function calculateSimilarity(row1, row2, excludeCol) {
  const keys = Object.keys(row1).filter(key => key !== excludeCol);
  let matches = 0;
  let total = 0;
  
  keys.forEach(key => {
    if (typeof row1[key] === 'number' && typeof row2[key] === 'number') {
      const diff = Math.abs(row1[key] - row2[key]);
      const maxVal = Math.max(Math.abs(row1[key]), Math.abs(row2[key]), 1);
      matches += 1 - (diff / maxVal);
      total++;
    } else if (row1[key] === row2[key]) {
      matches++;
      total++;
    } else if (total < keys.length) {
      total++;
    }
  });
  
  return total > 0 ? matches / total : 0;
}

/**
 * Calculate Counterfactual Fairness
 */
function calculateCounterfactualFairness(data, protectedAttr, targetCol) {
  // Simplified counterfactual fairness: what if protected attribute was different?
  const groups = groupDataByAttribute(data, protectedAttr);
  const groupNames = Object.keys(groups);
  
  if (groupNames.length < 2) return { score: 0, violations: 0 };
  
  let violations = 0;
  let total = 0;
  
  // Compare similar individuals from different groups
  groupNames.forEach((group1, i) => {
    groupNames.slice(i + 1).forEach(group2 => {
      const group1Data = groups[group1].slice(0, 50); // Limit for performance
      const group2Data = groups[group2].slice(0, 50);
      
      group1Data.forEach(row1 => {
        // Find most similar individual in other group
        let maxSimilarity = 0;
        let mostSimilar = null;
        
        group2Data.forEach(row2 => {
          const similarity = calculateSimilarity(row1, row2, protectedAttr);
          if (similarity > maxSimilarity) {
            maxSimilarity = similarity;
            mostSimilar = row2;
          }
        });
        
        if (mostSimilar && maxSimilarity > 0.7) {
          const outcomeDiff = Math.abs(Number(row1[targetCol]) - Number(mostSimilar[targetCol]));
          if (outcomeDiff > 0.1) violations++;
          total++;
        }
      });
    });
  });
  
  return {
    score: total > 0 ? 1 - (violations / total) : 1,
    violations: violations,
    totalComparisons: total
  };
}

/**
 * Calculate Treatment Equality
 */
function calculateTreatmentEquality(groups, targetCol) {
  const groupNames = Object.keys(groups);
  const errorRatios = {};
  
  groupNames.forEach(group => {
    const groupData = groups[group];
    let fp = 0, fn = 0;
    
    groupData.forEach(row => {
      const predicted = Number(row[targetCol]) > 0.5;
      const actual = Number(row[targetCol + '_actual'] || row[targetCol]) > 0.5;
      
      if (predicted && !actual) fp++;
      else if (!predicted && actual) fn++;
    });
    
    errorRatios[group] = fn > 0 ? fp / fn : (fp > 0 ? Infinity : 1);
  });
  
  const ratios = Object.values(errorRatios).filter(r => r !== Infinity);
  
  return {
    ratioRange: ratios.length > 0 ? Math.max(...ratios) - Math.min(...ratios) : 0,
    satisfiesTreatmentEquality: ratios.length > 0 ? (Math.max(...ratios) - Math.min(...ratios)) < 0.1 : true
  };
}

/**
 * Analyze intersectional bias
 */
function analyzeIntersectionalBias(data, columnAnalysis) {
  const intersectionalResults = {};
  
  // Analyze combinations of protected attributes
  if (columnAnalysis.protectedAttributes.length >= 2) {
    const combinations = generateCombinations(columnAnalysis.protectedAttributes, 2);
    
    combinations.forEach(combo => {
      const [attr1, attr2] = combo;
      const intersectionKey = `${attr1}_x_${attr2}`;
      
      // Create intersectional groups
      const intersectionalGroups = {};
      
      data.forEach(row => {
        const value1 = String(row[attr1] || 'Unknown');
        const value2 = String(row[attr2] || 'Unknown');
        const intersectionValue = `${value1}_${value2}`;
        
        if (!intersectionalGroups[intersectionValue]) {
          intersectionalGroups[intersectionValue] = [];
        }
        intersectionalGroups[intersectionValue].push(row);
      });
      
      // Calculate bias metrics for intersectional groups
      columnAnalysis.targetColumns.forEach(targetCol => {
        const intersectionalMetrics = calculateFairnessMetrics(
          Object.fromEntries(
            Object.entries(intersectionalGroups).map(([key, groupData]) => [
              key,
              {
                count: groupData.length,
                percentage: (groupData.length / data.length) * 100,
                metrics: {
                  [targetCol]: {
                    positiveRate: groupData.filter(row => Number(row[targetCol]) > 0.5).length / groupData.length
                  }
                }
              }
            ])
          ),
          [targetCol]
        );
        
        intersectionalResults[intersectionKey] = {
          attributes: combo,
          groups: intersectionalGroups,
          metrics: intersectionalMetrics,
          groupCount: Object.keys(intersectionalGroups).length
        };
      });
    });
  }
  
  return intersectionalResults;
}

/**
 * Generate combinations of array elements
 */
function generateCombinations(arr, size) {
  if (size > arr.length) return [];
  if (size === 1) return arr.map(item => [item]);
  
  const combinations = [];
  for (let i = 0; i <= arr.length - size; i++) {
    const smallerCombos = generateCombinations(arr.slice(i + 1), size - 1);
    smallerCombos.forEach(combo => {
      combinations.push([arr[i], ...combo]);
    });
  }
  
  return combinations;
}

/**
 * Analyze temporal bias patterns
 */
function analyzeTemporalBias(data, columnAnalysis) {
  // Look for date/time columns
  const headers = Object.keys(data[0] || {});
  const dateColumns = headers.filter(header => {
    const sampleValues = data.slice(0, 5).map(row => row[header]);
    return sampleValues.some(val => val instanceof Date || (typeof val === 'string' && /\d{4}[-/]\d{2}[-/]\d{2}/.test(val)));
  });
  
  if (dateColumns.length === 0) {
    return { message: 'No date columns found for temporal analysis' };
  }
  
  const temporalResults = {};
  
  dateColumns.forEach(dateCol => {
    // Group data by time periods
    const timeGroups = {};
    
    data.forEach(row => {
      const dateValue = row[dateCol];
      let period;
      
      if (dateValue instanceof Date) {
        period = `${dateValue.getFullYear()}-${String(dateValue.getMonth() + 1).padStart(2, '0')}`;
      } else if (typeof dateValue === 'string') {
        const match = dateValue.match(/(\d{4})[-/](\d{2})/);
        if (match) {
          period = `${match[1]}-${match[2]}`;
        }
      }
      
      if (period) {
        if (!timeGroups[period]) timeGroups[period] = [];
        timeGroups[period].push(row);
      }
    });
    
    // Analyze bias trends over time
    const timePeriods = Object.keys(timeGroups).sort();
    const biasOverTime = [];
    
    timePeriods.forEach(period => {
      const periodData = timeGroups[period];
      const periodAnalysis = analyzeColumns(periodData);
      
      if (periodAnalysis.protectedAttributes.length > 0) {
        const periodBias = calculateBiasMetrics(periodData, periodAnalysis);
        biasOverTime.push({
          period: period,
          biasScore: periodBias.overallMetrics.overallBiasScore,
          disparateImpact: periodBias.overallMetrics.avgDisparateImpact,
          recordCount: periodData.length
        });
      }
    });
    
    temporalResults[dateCol] = {
      periods: timePeriods.length,
      biasOverTime: biasOverTime,
      trend: calculateBiasTrend(biasOverTime)
    };
  });
  
  return temporalResults;
}

/**
 * Calculate bias trend (improving, worsening, stable)
 */
function calculateBiasTrend(biasOverTime) {
  if (biasOverTime.length < 2) return 'insufficient_data';
  
  const scores = biasOverTime.map(item => item.biasScore);
  const firstHalf = scores.slice(0, Math.floor(scores.length / 2));
  const secondHalf = scores.slice(Math.ceil(scores.length / 2));
  
  const firstAvg = firstHalf.reduce((a, b) => a + b, 0) / firstHalf.length;
  const secondAvg = secondHalf.reduce((a, b) => a + b, 0) / secondHalf.length;
  
  const improvement = firstAvg - secondAvg;
  
  if (improvement > 0.05) return 'improving';
  else if (improvement < -0.05) return 'worsening';
  else return 'stable';
}

/**
 * Calculate feature importance for bias
 */
function calculateFeatureImportance(data, columnAnalysis) {
  const importance = {};
  
  // Analyze correlation between features and bias
  const allColumns = Object.keys(data[0] || {});
  const nonProtectedFeatures = allColumns.filter(col => 
    !columnAnalysis.protectedAttributes.includes(col) && 
    !columnAnalysis.targetColumns.includes(col)
  );
  
  columnAnalysis.protectedAttributes.forEach(protectedAttr => {
    importance[protectedAttr] = {};
    
    nonProtectedFeatures.forEach(feature => {
      // Calculate correlation between feature and protected attribute
      const correlation = calculateCorrelation(data, feature, protectedAttr);
      
      importance[protectedAttr][feature] = {
        correlation: correlation,
        biasRisk: Math.abs(correlation) > 0.3 ? 'HIGH' : Math.abs(correlation) > 0.1 ? 'MEDIUM' : 'LOW',
        recommendation: Math.abs(correlation) > 0.3 ? 
          `Feature ${feature} shows high correlation with ${protectedAttr}. Consider feature engineering or removal.` :
          `Feature ${feature} shows low bias risk.`
      };
    });
  });
  
  return importance;
}

/**
 * Calculate correlation between two variables
 */
function calculateCorrelation(data, var1, var2) {
  const pairs = data.map(row => ({
    x: Number(row[var1]) || 0,
    y: Number(row[var2]) || 0
  })).filter(pair => !isNaN(pair.x) && !isNaN(pair.y));
  
  if (pairs.length < 2) return 0;
  
  const meanX = pairs.reduce((sum, pair) => sum + pair.x, 0) / pairs.length;
  const meanY = pairs.reduce((sum, pair) => sum + pair.y, 0) / pairs.length;
  
  let numerator = 0;
  let denomX = 0;
  let denomY = 0;
  
  pairs.forEach(pair => {
    const diffX = pair.x - meanX;
    const diffY = pair.y - meanY;
    numerator += diffX * diffY;
    denomX += diffX * diffX;
    denomY += diffY * diffY;
  });
  
  const denominator = Math.sqrt(denomX * denomY);
  return denominator === 0 ? 0 : numerator / denominator;
}

/**
 * Utility functions for statistical calculations
 */

function groupDataByAttribute(data, attribute) {
  const groups = {};
  data.forEach(row => {
    const groupValue = String(row[attribute] || 'Unknown').trim();
    if (!groups[groupValue]) groups[groupValue] = [];
    groups[groupValue].push(row);
  });
  return groups;
}

function calculateVariance(values, mean) {
  if (values.length < 2) return 0;
  const squaredDiffs = values.map(value => Math.pow(value - mean, 2));
  return squaredDiffs.reduce((a, b) => a + b, 0) / (values.length - 1);
}

function calculateChiSquareStatistic(contingencyTable) {
  if (contingencyTable.length < 2 || contingencyTable[0].length < 2) return 0;
  
  const rows = contingencyTable.length;
  const cols = contingencyTable[0].length;
  
  // Calculate row and column totals
  const rowTotals = contingencyTable.map(row => row.reduce((a, b) => a + b, 0));
  const colTotals = [];
  for (let j = 0; j < cols; j++) {
    colTotals[j] = contingencyTable.reduce((sum, row) => sum + row[j], 0);
  }
  const grandTotal = rowTotals.reduce((a, b) => a + b, 0);
  
  // Calculate chi-square statistic
  let chiSquare = 0;
  for (let i = 0; i < rows; i++) {
    for (let j = 0; j < cols; j++) {
      const observed = contingencyTable[i][j];
      const expected = (rowTotals[i] * colTotals[j]) / grandTotal;
      if (expected > 0) {
        chiSquare += Math.pow(observed - expected, 2) / expected;
      }
    }
  }
  
  return chiSquare;
}

function calculateChiSquarePValue(chiSquare, degreesOfFreedom) {
  // Simplified p-value calculation (approximation)
  // In a real implementation, you would use a proper chi-square distribution function
  if (degreesOfFreedom === 1) {
    if (chiSquare > 10.83) return 0.001;
    else if (chiSquare > 6.63) return 0.01;
    else if (chiSquare > 3.84) return 0.05;
    else if (chiSquare > 2.71) return 0.1;
    else return 0.5;
  }
  
  // Very rough approximation for other degrees of freedom
  const criticalValue = 3.84 + (degreesOfFreedom - 1) * 2;
  return chiSquare > criticalValue ? 0.05 : 0.5;
}

/**
 * Enhanced bias report generation with advanced metrics
 */
function createAdvancedBiasReport(biasResults, advancedResults, aiInsights, columnAnalysis) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create advanced report sheet
  const advancedReportName = CONFIG.REPORT_SHEET_NAME + '_Advanced';
  let advancedSheet = spreadsheet.getSheetByName(advancedReportName);
  
  if (advancedSheet) {
    spreadsheet.deleteSheet(advancedSheet);
  }
  
  advancedSheet = spreadsheet.insertSheet(advancedReportName);
  
  let currentRow = 1;
  
  // Header
  advancedSheet.getRange(currentRow, 1, 1, 8).merge().setValue('🧠 ADVANCED AI BIAS ANALYSIS REPORT')
    .setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#673ab7').setFontColor('white');
  currentRow += 3;
  
  // Statistical Tests Section
  if (advancedResults.statisticalTests) {
    currentRow = createStatisticalTestsSection(advancedSheet, currentRow, advancedResults.statisticalTests);
  }
  
  // Advanced Metrics Section
  if (advancedResults.advancedMetrics) {
    currentRow = createAdvancedMetricsSection(advancedSheet, currentRow, advancedResults.advancedMetrics);
  }
  
  // Intersectional Bias Section
  if (advancedResults.intersectionalBias) {
    currentRow = createIntersectionalBiasSection(advancedSheet, currentRow, advancedResults.intersectionalBias);
  }
  
  // Temporal Analysis Section
  if (advancedResults.temporalAnalysis) {
    currentRow = createTemporalAnalysisSection(advancedSheet, currentRow, advancedResults.temporalAnalysis);
  }
  
  // Feature Importance Section
  if (advancedResults.featureImportance) {
    currentRow = createFeatureImportanceSection(advancedSheet, currentRow, advancedResults.featureImportance);
  }
  
  // Format the advanced report
  formatReportSheet(advancedSheet);
  
  return advancedSheet;
}

/**
 * Create statistical tests section in advanced report
 */
function createStatisticalTestsSection(sheet, startRow, statisticalTests) {
  sheet.getRange(startRow, 1).setValue('📊 STATISTICAL SIGNIFICANCE TESTS')
    .setFontSize(14).setFontWeight('bold').setBackground('#2196f3').setFontColor('white');
  startRow += 2;
  
  // Headers
  const headers = ['Protected Attribute', 'Target', 'Test', 'Statistic', 'P-Value', 'Significant', 'Interpretation'];
  headers.forEach((header, index) => {
    sheet.getRange(startRow, index + 1).setValue(header).setFontWeight('bold');
  });
  startRow += 1;
  
  Object.entries(statisticalTests).forEach(([attr, targets]) => {
    Object.entries(targets).forEach(([target, tests]) => {
      Object.entries(tests).forEach(([testName, result]) => {
        if (result) {
          const significant = result.significant || result.pValue < 0.05;
          const interpretation = significant ? 
            'Statistically significant difference detected' : 
            'No statistically significant difference';
          
          sheet.getRange(startRow, 1).setValue(attr);
          sheet.getRange(startRow, 2).setValue(target);
          sheet.getRange(startRow, 3).setValue(testName.toUpperCase());
          sheet.getRange(startRow, 4).setValue(result.statistic?.toFixed(4) || 'N/A');
          sheet.getRange(startRow, 5).setValue(result.pValue?.toFixed(4) || 'N/A');
          sheet.getRange(startRow, 6).setValue(significant ? 'YES' : 'NO')
            .setBackground(significant ? '#ffcdd2' : '#c8e6c9')
            .setFontWeight('bold');
          sheet.getRange(startRow, 7).setValue(interpretation).setWrap(true);
          
          startRow += 1;
        }
      });
    });
  });
  
  return startRow + 2;
}

/**
 * Create advanced metrics section
 */
function createAdvancedMetricsSection(sheet, startRow, advancedMetrics) {
  sheet.getRange(startRow, 1).setValue('⚖️ ADVANCED FAIRNESS METRICS')
    .setFontSize(14).setFontWeight('bold').setBackground('#4caf50').setFontColor('white');
  startRow += 2;
  
  Object.entries(advancedMetrics).forEach(([attr, targets]) => {
    sheet.getRange(startRow, 1, 1, 6).merge().setValue(`Protected Attribute: ${attr.toUpperCase()}`)
      .setFontWeight('bold').setBackground('#f5f5f5');
    startRow += 2;
    
    Object.entries(targets).forEach(([target, metrics]) => {
      // Equalized Odds
      if (metrics.equalizedOdds) {
        sheet.getRange(startRow, 1).setValue('Equalized Odds:').setFontWeight('bold');
        sheet.getRange(startRow, 2).setValue(`TPR Diff: ${(metrics.equalizedOdds.tprDifference * 100).toFixed(2)}%`);
        sheet.getRange(startRow, 3).setValue(`FPR Diff: ${(metrics.equalizedOdds.fprDifference * 100).toFixed(2)}%`);
        sheet.getRange(startRow, 4).setValue(metrics.equalizedOdds.satisfiesEqualizedOdds ? 'PASS' : 'FAIL')
          .setBackground(metrics.equalizedOdds.satisfiesEqualizedOdds ? '#c8e6c9' : '#ffcdd2');
        startRow += 1;
      }
      
      // Calibration
      if (metrics.calibration) {
        sheet.getRange(startRow, 1).setValue('Calibration:').setFontWeight('bold');
        sheet.getRange(startRow, 2).setValue(`Avg Error: ${(metrics.calibration.avgCalibrationError * 100).toFixed(2)}%`);
        sheet.getRange(startRow, 3).setValue(`Max Diff: ${(metrics.calibration.maxCalibrationDifference * 100).toFixed(2)}%`);
        sheet.getRange(startRow, 4).setValue(metrics.calibration.wellCalibrated ? 'PASS' : 'FAIL')
          .setBackground(metrics.calibration.wellCalibrated ? '#c8e6c9' : '#ffcdd2');
        startRow += 1;
      }
    });
    
    startRow += 1;
  });
  
  return startRow + 2;
}

/**
 * Create intersectional bias section
 */
function createIntersectionalBiasSection(sheet, startRow, intersectionalBias) {
  sheet.getRange(startRow, 1).setValue('🔀 INTERSECTIONAL BIAS ANALYSIS')
    .setFontSize(14).setFontWeight('bold').setBackground('#ff9800').setFontColor('white');
  startRow += 2;
  
  if (Object.keys(intersectionalBias).length === 0) {
    sheet.getRange(startRow, 1, 1, 4).merge().setValue('No intersectional analysis performed (requires 2+ protected attributes)')
      .setFontStyle('italic');
    return startRow + 3;
  }
  
  Object.entries(intersectionalBias).forEach(([intersection, analysis]) => {
    sheet.getRange(startRow, 1, 1, 6).merge()
      .setValue(`Intersection: ${analysis.attributes.join(' × ').toUpperCase()}`)
      .setFontWeight('bold').setBackground('#fff3e0');
    startRow += 1;
    
    sheet.getRange(startRow, 1).setValue('Total Groups:').setFontWeight('bold');
    sheet.getRange(startRow, 2).setValue(analysis.groupCount);
    startRow += 1;
    
    // Show most and least represented intersectional groups
    const groupSizes = Object.entries(analysis.groups).map(([name, data]) => ({
      name: name,
      size: data.length
    })).sort((a, b) => b.size - a.size);
    
    if (groupSizes.length > 0) {
      sheet.getRange(startRow, 1).setValue('Largest Group:').setFontWeight('bold');
      sheet.getRange(startRow, 2, 1, 2).merge().setValue(`${groupSizes[0].name} (${groupSizes[0].size} records)`);
      startRow += 1;
      
      sheet.getRange(startRow, 1).setValue('Smallest Group:').setFontWeight('bold');
      sheet.getRange(startRow, 2, 1, 2).merge()
        .setValue(`${groupSizes[groupSizes.length - 1].name} (${groupSizes[groupSizes.length - 1].size} records)`);
      startRow += 1;
    }
    
    startRow += 2;
  });
  
  return startRow + 1;
}

/**
 * Create temporal analysis section
 */
function createTemporalAnalysisSection(sheet, startRow, temporalAnalysis) {
  sheet.getRange(startRow, 1).setValue('📈 TEMPORAL BIAS ANALYSIS')
    .setFontSize(14).setFontWeight('bold').setBackground('#9c27b0').setFontColor('white');
  startRow += 2;
  
  if (temporalAnalysis.message) {
    sheet.getRange(startRow, 1, 1, 4).merge().setValue(temporalAnalysis.message).setFontStyle('italic');
    return startRow + 3;
  }
  
  Object.entries(temporalAnalysis).forEach(([dateCol, analysis]) => {
    sheet.getRange(startRow, 1).setValue(`Date Column: ${dateCol}`).setFontWeight('bold');
    sheet.getRange(startRow, 2).setValue(`Periods Analyzed: ${analysis.periods}`);
    startRow += 1;
    
    sheet.getRange(startRow, 1).setValue('Bias Trend:').setFontWeight('bold');
    const trendColor = analysis.trend === 'improving' ? '#4caf50' : 
                      analysis.trend === 'worsening' ? '#f44336' : '#ff9800';
    sheet.getRange(startRow, 2).setValue(analysis.trend.toUpperCase())
      .setBackground(trendColor).setFontColor('white').setFontWeight('bold');
    startRow += 2;
  });
  
  return startRow + 1;
}

/**
 * Create feature importance section
 */
function createFeatureImportanceSection(sheet, startRow, featureImportance) {
  sheet.getRange(startRow, 1).setValue('🎯 FEATURE IMPORTANCE FOR BIAS')
    .setFontSize(14).setFontWeight('bold').setBackground('#607d8b').setFontColor('white');
  startRow += 2;
  
  Object.entries(featureImportance).forEach(([attr, features]) => {
    sheet.getRange(startRow, 1, 1, 6).merge()
      .setValue(`Features correlated with ${attr.toUpperCase()}`)
      .setFontWeight('bold').setBackground('#eceff1');
    startRow += 2;
    
    // Headers
    const headers = ['Feature', 'Correlation', 'Bias Risk', 'Recommendation'];
    headers.forEach((header, index) => {
      sheet.getRange(startRow, index + 1).setValue(header).setFontWeight('bold');
    });
    startRow += 1;
    
    Object.entries(features).forEach(([feature, importance]) => {
      sheet.getRange(startRow, 1).setValue(feature);
      sheet.getRange(startRow, 2).setValue(importance.correlation.toFixed(3));
      
      const riskColor = importance.biasRisk === 'HIGH' ? '#ffcdd2' : 
                        importance.biasRisk === 'MEDIUM' ? '#fff3e0' : '#c8e6c9';
      sheet.getRange(startRow, 3).setValue(importance.biasRisk)
        .setBackground(riskColor).setFontWeight('bold');
      
      sheet.getRange(startRow, 4).setValue(importance.recommendation).setWrap(true);
      startRow += 1;
    });
    
    startRow += 2;
  });
  
  return startRow;
}
function showDashboard() {
  const html = createDashboardHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setTitle('🤖 AI Bias Detector Dashboard')
    .setWidth(400);
    
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

/**
 * Dashboard and User Interface Functions
 * Creates interactive dashboard and sidebar interfaces
 */

/**
 * Show interactive dashboard sidebar
 */


/**
 * Create dashboard HTML
 */
function createDashboardHTML() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      font-family: 'Google Sans', Arial, sans-serif;
      margin: 0;
      padding: 20px;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: #333;
      min-height: 100vh;
    }
    
    .dashboard-container {
      background: white;
      border-radius: 12px;
      padding: 24px;
      box-shadow: 0 8px 32px rgba(0,0,0,0.1);
      backdrop-filter: blur(10px);
      border: 1px solid rgba(255,255,255,0.2);
    }
    
    .header {
      text-align: center;
      margin-bottom: 30px;
    }
    
    .header h1 {
      margin: 0;
      font-size: 24px;
      font-weight: 600;
      background: linear-gradient(45deg, #4285f4, #34a853);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      background-clip: text;
    }
    
    .header p {
      margin: 8px 0 0 0;
      color: #666;
      font-size: 14px;
    }
    
    .action-buttons {
      display: flex;
      flex-direction: column;
      gap: 12px;
      margin-bottom: 30px;
    }
    
    .btn {
      padding: 14px 20px;
      border: none;
      border-radius: 8px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.3s ease;
      text-align: left;
      display: flex;
      align-items: center;
      gap: 10px;
    }
    
    .btn-primary {
      background: linear-gradient(45deg, #4285f4, #1976d2);
      color: white;
    }
    
    .btn-primary:hover {
      transform: translateY(-2px);
      box-shadow: 0 4px 12px rgba(66, 133, 244, 0.3);
    }
    
    .btn-secondary {
      background: #f8f9fa;
      color: #333;
      border: 1px solid #e8eaed;
    }
    
    .btn-secondary:hover {
      background: #e8f0fe;
      border-color: #4285f4;
    }
    
    .stats-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 16px;
      margin-bottom: 30px;
    }
    
    .stat-card {
      background: #f8f9fa;
      padding: 16px;
      border-radius: 8px;
      text-align: center;
      border-left: 4px solid #4285f4;
    }
    
    .stat-value {
      font-size: 24px;
      font-weight: 700;
      margin-bottom: 4px;
      color: #1976d2;
    }
    
    .stat-label {
      font-size: 12px;
      color: #666;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    
    .recent-analysis {
      background: #f8f9fa;
      padding: 20px;
      border-radius: 8px;
      margin-bottom: 20px;
    }
    
    .recent-analysis h3 {
      margin: 0 0 16px 0;
      font-size: 16px;
      color: #333;
    }
    
    .analysis-item {
      padding: 12px;
      background: white;
      border-radius: 6px;
      margin-bottom: 8px;
      border-left: 3px solid #34a853;
    }
    
    .analysis-date {
      font-size: 12px;
      color: #666;
      margin-bottom: 4px;
    }
    
    .analysis-result {
      font-weight: 600;
      color: #333;
    }
    
    .quick-actions {
      margin-top: 20px;
    }
    
    .quick-actions h3 {
      margin: 0 0 16px 0;
      font-size: 16px;
      color: #333;
    }
    
    .action-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 12px;
    }
    
    .action-card {
      padding: 16px;
      background: white;
      border-radius: 8px;
      border: 1px solid #e8eaed;
      cursor: pointer;
      transition: all 0.3s ease;
      text-align: center;
    }
    
    .action-card:hover {
      border-color: #4285f4;
      transform: translateY(-2px);
      box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    
    .action-icon {
      font-size: 24px;
      margin-bottom: 8px;
    }
    
    .action-title {
      font-weight: 600;
      font-size: 12px;
      color: #333;
    }
    
    .loading {
      text-align: center;
      padding: 20px;
    }
    
    .spinner {
      border: 3px solid #f3f3f3;
      border-top: 3px solid #4285f4;
      border-radius: 50%;
      width: 30px;
      height: 30px;
      animation: spin 1s linear infinite;
      margin: 0 auto 16px;
    }
    
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
    
    .status-indicator {
      display: inline-block;
      width: 8px;
      height: 8px;
      border-radius: 50%;
      margin-right: 8px;
    }
    
    .status-good { background: #34a853; }
    .status-warning { background: #fbbc04; }
    .status-error { background: #ea4335; }
    
    .tooltip {
      position: relative;
      display: inline-block;
      cursor: help;
    }
    
    .tooltip .tooltiptext {
      visibility: hidden;
      width: 200px;
      background-color: #333;
      color: #fff;
      text-align: center;
      border-radius: 6px;
      padding: 8px;
      position: absolute;
      z-index: 1;
      bottom: 125%;
      left: 50%;
      margin-left: -100px;
      font-size: 12px;
    }
    
    .tooltip:hover .tooltiptext {
      visibility: visible;
    }
  </style>
</head>
<body>
  <div class="dashboard-container">
    <div class="header">
      <h1>🤖 AI Bias Detector</h1>
      <p>Automated fairness analysis for your data</p>
    </div>
    
    <div class="action-buttons">
      <button class="btn btn-primary" onclick="runAnalysis()">
        <span>🔍</span>
        Run Full Bias Analysis
      </button>
      <button class="btn btn-secondary" onclick="quickScan()">
        <span>⚡</span>
        Quick Bias Scan
      </button>
    </div>
    
    <div id="stats-section" style="display: none;">
      <div class="stats-grid">
        <div class="stat-card">
          <div class="stat-value" id="bias-score">--</div>
          <div class="stat-label">Overall Bias Score</div>
        </div>
        <div class="stat-card">
          <div class="stat-value" id="risk-level">--</div>
          <div class="stat-label">Risk Level</div>
        </div>
      </div>
    </div>
    
    <div class="recent-analysis">
      <h3>📊 Last Analysis</h3>
      <div id="recent-results">
        <p style="color: #666; font-style: italic;">No analysis performed yet</p>
      </div>
    </div>
    
    <div class="quick-actions">
      <h3>⚡ Quick Actions</h3>
      <div class="action-grid">
        <div class="action-card" onclick="viewReport()">
          <div class="action-icon">📋</div>
          <div class="action-title">View Report</div>
        </div>
        <div class="action-card" onclick="viewHistory()">
          <div class="action-icon">📈</div>
          <div class="action-title">View History</div>
        </div>
        <div class="action-card" onclick="exportReport()">
          <div class="action-icon">📤</div>
          <div class="action-title">Export Report</div>
        </div>
        <div class="action-card" onclick="showSettings()">
          <div class="action-icon">⚙️</div>
          <div class="action-title">Settings</div>
        </div>
      </div>
    </div>
    
    <div id="loading" class="loading" style="display: none;">
      <div class="spinner"></div>
      <p>Analyzing data for bias patterns...</p>
    </div>
  </div>
  
  <script>
    function runAnalysis() {
      showLoading();
      google.script.run
        .withSuccessHandler(onAnalysisComplete)
        .withFailureHandler(onAnalysisError)
        .runBiasAnalysis();
    }
    
    function quickScan() {
      showLoading();
      google.script.run
        .withSuccessHandler(onQuickScanComplete)
        .withFailureHandler(onAnalysisError)
        .runQuickBiasScan();
    }
    
    function viewReport() {
      google.script.run.openReportSheet();
    }
    
    function viewHistory() {
      google.script.run.showHistory();
    }
    
    function exportReport() {
      google.script.run.exportBiasReport();
    }
    
    function showSettings() {
      google.script.run.showSettings();
    }
    
    function showLoading() {
      document.getElementById('loading').style.display = 'block';
    }
    
    function hideLoading() {
      document.getElementById('loading').style.display = 'none';
    }
    
    function onAnalysisComplete(result) {
      hideLoading();
      updateDashboard(result);
      showNotification('Analysis completed successfully!', 'success');
    }
    
    function onQuickScanComplete(result) {
      hideLoading();
      updateQuickStats(result);
      showNotification('Quick scan completed!', 'success');
    }
    
    function onAnalysisError(error) {
      hideLoading();
      showNotification('Error: ' + error.message, 'error');
    }
    
    function updateDashboard(results) {
      document.getElementById('stats-section').style.display = 'block';
      document.getElementById('bias-score').textContent = (results.overallBiasScore * 100).toFixed(1) + '%';
      document.getElementById('risk-level').textContent = results.riskLevel;
      
      const recentResults = document.getElementById('recent-results');
      recentResults.innerHTML = \`
        <div class="analysis-item">
          <div class="analysis-date">\${new Date().toLocaleString()}</div>
          <div class="analysis-result">
            <span class="status-indicator status-\${results.riskLevel.toLowerCase()}"></span>
            Analysis completed - \${results.riskLevel} risk level detected
          </div>
        </div>
      \`;
    }
    
    function updateQuickStats(results) {
      const recentResults = document.getElementById('recent-results');
      recentResults.innerHTML = \`
        <div class="analysis-item">
          <div class="analysis-date">\${new Date().toLocaleString()}</div>
          <div class="analysis-result">
            <span class="status-indicator status-warning"></span>
            Quick scan - \${results.flagCount} potential issues found
          </div>
        </div>
      \`;
    }
    
    function showNotification(message, type) {
      // Create notification element
      const notification = document.createElement('div');
      notification.style.cssText = \`
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 16px 20px;
        border-radius: 8px;
        color: white;
        font-weight: 600;
        z-index: 1000;
        transform: translateX(400px);
        transition: transform 0.3s ease;
        background: \${type === 'success' ? '#34a853' : '#ea4335'};
      \`;
      notification.textContent = message;
      
      document.body.appendChild(notification);
      
      // Show notification
      setTimeout(() => {
        notification.style.transform = 'translateX(0)';
      }, 100);
      
      // Hide notification after 3 seconds
      setTimeout(() => {
        notification.style.transform = 'translateX(400px)';
        setTimeout(() => {
          document.body.removeChild(notification);
        }, 300);
      }, 3000);
    }
    
    // Load initial data
    google.script.run
      .withSuccessHandler(updateInitialData)
      .getLastAnalysisResults();
      
    function updateInitialData(data) {
      if (data) {
        updateDashboard(data);
      }
    }
  </script>
</body>
</html>
  `;
}

/**
 * Show history analysis
 */
function showHistory() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = spreadsheet.getSheetByName(CONFIG.HISTORY_SHEET_NAME);
  
  if (historySheet) {
    spreadsheet.setActiveSheet(historySheet);
  } else {
    SpreadsheetApp.getUi().alert('No History Found', 'No analysis history available yet. Run a bias analysis first.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Show settings dialog
 */
function showSettings() {
  const html = createSettingsHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(600);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '⚙️ Bias Detector Settings');
}

/**
 * Create settings HTML
 */
function createSettingsHTML() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { 
      font-family: 'Google Sans', Arial, sans-serif; 
      padding: 20px; 
      background: #f8f9fa;
    }
    
    .settings-container {
      background: white;
      padding: 24px;
      border-radius: 8px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    .setting-group {
      margin-bottom: 24px;
      padding-bottom: 20px;
      border-bottom: 1px solid #e8eaed;
    }
    
    .setting-group:last-child {
      border-bottom: none;
    }
    
    .setting-title {
      font-size: 16px;
      font-weight: 600;
      margin-bottom: 8px;
      color: #333;
    }
    
    .setting-description {
      font-size: 14px;
      color: #666;
      margin-bottom: 16px;
    }
    
    .form-group {
      margin-bottom: 16px;
    }
    
    label {
      display: block;
      font-weight: 500;
      margin-bottom: 4px;
      color: #333;
    }
    
    input[type="number"], select, input[type="text"] {
      width: 100%;
      padding: 8px 12px;
      border: 1px solid #dadce0;
      border-radius: 4px;
      font-size: 14px;
    }
    
    input[type="checkbox"] {
      margin-right: 8px;
    }
    
    .btn {
      padding: 10px 20px;
      border: none;
      border-radius: 4px;
      font-weight: 500;
      cursor: pointer;
      margin-right: 12px;
    }
    
    .btn-primary {
      background: #4285f4;
      color: white;
    }
    
    .btn-secondary {
      background: #f8f9fa;
      color: #333;
      border: 1px solid #dadce0;
    }
    
    .advanced-settings {
      background: #f8f9fa;
      padding: 16px;
      border-radius: 6px;
      margin-top: 16px;
    }
  </style>
</head>
<body>
  <div class="settings-container">
    <div class="setting-group">
      <div class="setting-title">🎯 Bias Detection Thresholds</div>
      <div class="setting-description">Adjust the sensitivity of bias detection</div>
      
      <div class="form-group">
        <label for="disparateImpact">Disparate Impact Threshold (0.8 = 80% rule)</label>
        <input type="number" id="disparateImpact" min="0.1" max="1.0" step="0.1" value="0.8">
      </div>
      
      <div class="form-group">
        <label for="statisticalParity">Statistical Parity Threshold</label>
        <input type="number" id="statisticalParity" min="0.01" max="0.5" step="0.01" value="0.1">
      </div>
    </div>
    
    <div class="setting-group">
      <div class="setting-title">🏷️ Protected Attributes</div>
      <div class="setting-description">Select which attributes to consider as protected</div>
      
      <div class="form-group">
        <label><input type="checkbox" checked> Gender</label>
        <label><input type="checkbox" checked> Race/Ethnicity</label>
        <label><input type="checkbox" checked> Age</label>
        <label><input type="checkbox" checked> Religion</label>
        <label><input type="checkbox" checked> Nationality</label>
        <label><input type="checkbox"> Sexual Orientation</label>
        <label><input type="checkbox"> Disability Status</label>
      </div>
    </div>
    
    <div class="setting-group">
      <div class="setting-title">📊 Reporting Options</div>
      <div class="setting-description">Customize report generation and visualization</div>
      
      <div class="form-group">
        <label><input type="checkbox" checked> Generate visualizations</label>
        <label><input type="checkbox" checked> Include AI insights</label>
        <label><input type="checkbox" checked> Save to history</label>
        <label><input type="checkbox"> Auto-export to PDF</label>
      </div>
      
      <div class="form-group">
        <label for="reportFormat">Default Report Format</label>
        <select id="reportFormat">
          <option value="detailed">Detailed Report</option>
          <option value="summary">Executive Summary</option>
          <option value="technical">Technical Report</option>
        </select>
      </div>
    </div>
    
    <div class="setting-group">
      <div class="setting-title">🤖 AI Analysis Settings</div>
      <div class="setting-description">Configure AI-powered insights and recommendations</div>
      
      <div class="form-group">
        <label for="insightLevel">Insight Detail Level</label>
        <select id="insightLevel">
          <option value="basic">Basic</option>
          <option value="detailed" selected>Detailed</option>
          <option value="comprehensive">Comprehensive</option>
        </select>
      </div>
      
      <div class="form-group">
        <label><input type="checkbox" checked> Generate recommendations</label>
        <label><input type="checkbox" checked> Highlight biased data points</label>
        <label><input type="checkbox"> Advanced statistical analysis</label>
      </div>
    </div>
    
    <div class="advanced-settings">
      <div class="setting-title">🔧 Advanced Settings</div>
      
      <div class="form-group">
        <label for="minGroupSize">Minimum Group Size for Analysis</label>
        <input type="number" id="minGroupSize" min="10" max="1000" value="30">
      </div>
      
      <div class="form-group">
        <label for="confidenceLevel">Confidence Level for Statistical Tests</label>
        <select id="confidenceLevel">
          <option value="0.90">90%</option>
          <option value="0.95" selected>95%</option>
          <option value="0.99">99%</option>
        </select>
      </div>
    </div>
    
    <div style="text-align: right; margin-top: 24px;">
      <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
      <button class="btn btn-primary" onclick="saveSettings()">Save Settings</button>
    </div>
  </div>
  
  <script>
    function saveSettings() {
      const settings = {
        disparateImpact: parseFloat(document.getElementById('disparateImpact').value),
        statisticalParity: parseFloat(document.getElementById('statisticalParity').value),
        reportFormat: document.getElementById('reportFormat').value,
        insightLevel: document.getElementById('insightLevel').value,
        minGroupSize: parseInt(document.getElementById('minGroupSize').value),
        confidenceLevel: parseFloat(document.getElementById('confidenceLevel').value)
      };
      
      google.script.run
        .withSuccessHandler(() => {
          alert('Settings saved successfully!');
          google.script.host.close();
        })
        .withFailureHandler((error) => {
          alert('Error saving settings: ' + error.message);
        })
        .saveSettings(settings);
    }
  </script>
</body>
</html>
  `;
}

/**
 * Run quick bias scan (lighter analysis)
 */
function runQuickBiasScan() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = getSheetData(sheet);
    
    if (data.length < 2) {
      return { error: 'Insufficient data' };
    }
    
    const columnAnalysis = analyzeColumns(data);
    
    if (columnAnalysis.protectedAttributes.length === 0) {
      return { error: 'No protected attributes found' };
    }
    
    // Quick analysis - just check basic metrics
    let flagCount = 0;
    
    columnAnalysis.protectedAttributes.forEach(attr => {
      const groups = {};
      data.forEach(row => {
        const groupValue = String(row[attr] || 'Unknown').trim();
        if (!groups[groupValue]) groups[groupValue] = 0;
        groups[groupValue]++;
      });
      
      const groupSizes = Object.values(groups);
      const maxSize = Math.max(...groupSizes);
      const minSize = Math.min(...groupSizes);
      
      // Simple imbalance check
      if (maxSize > minSize * 3) {
        flagCount++;
      }
    });
    
    return {
      flagCount: flagCount,
      attributeCount: columnAnalysis.protectedAttributes.length,
      riskLevel: flagCount > 2 ? 'HIGH' : flagCount > 0 ? 'MEDIUM' : 'LOW'
    };
    
  } catch (error) {
    return { error: error.message };
  }
}

/**
 * Open bias report sheet
 */
function openReportSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = spreadsheet.getSheetByName(CONFIG.REPORT_SHEET_NAME);
  
  if (reportSheet) {
    spreadsheet.setActiveSheet(reportSheet);
  } else {
    SpreadsheetApp.getUi().alert('No Report Found', 'No bias analysis report found. Please run an analysis first.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Export bias report
 */
function exportBiasReport() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const reportSheet = spreadsheet.getSheetByName(CONFIG.REPORT_SHEET_NAME);
    
    if (!reportSheet) {
      SpreadsheetApp.getUi().alert('No Report Found', 'No bias analysis report found. Please run an analysis first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Create a temporary spreadsheet for export
    const exportSpreadsheet = SpreadsheetApp.create(`Bias_Report_Export_${new Date().toISOString().split('T')[0]}`);
    const exportSheet = exportSpreadsheet.getActiveSheet();
    
    // Copy data from report sheet
    const sourceRange = reportSheet.getDataRange();
    const sourceValues = sourceRange.getValues();
    const sourceFormats = sourceRange.getBackgrounds();
    
    exportSheet.getRange(1, 1, sourceValues.length, sourceValues[0].length).setValues(sourceValues);
    exportSheet.getRange(1, 1, sourceFormats.length, sourceFormats[0].length).setBackgrounds(sourceFormats);
    
    // Format the export sheet
    exportSheet.setName('Bias_Analysis_Report');
    exportSheet.autoResizeColumns(1, sourceValues[0].length);
    
    const exportUrl = exportSpreadsheet.getUrl();
    
    // Show export options dialog
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'Export Complete',
      `Report exported successfully!\n\nExport URL: ${exportUrl}\n\nWould you like to open the exported report?`,
      ui.ButtonSet.YES_NO
    );
    
    if (result === ui.Button.YES) {
      // Note: Cannot directly open URL in Google Apps Script, but we can copy it to clipboard
      SpreadsheetApp.getUi().alert('Export URL', exportUrl, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Export Error', `Failed to export report: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Save user settings
 */
function saveSettings(settings) {
  try {
    const properties = PropertiesService.getDocumentProperties();
    properties.setProperties({
      'biasDetector.disparateImpact': settings.disparateImpact.toString(),
      'biasDetector.statisticalParity': settings.statisticalParity.toString(),
      'biasDetector.reportFormat': settings.reportFormat,
      'biasDetector.insightLevel': settings.insightLevel,
      'biasDetector.minGroupSize': settings.minGroupSize.toString(),
      'biasDetector.confidenceLevel': settings.confidenceLevel.toString()
    });
    
    return { success: true };
  } catch (error) {
    throw new Error(`Failed to save settings: ${error.message}`);
  }
}

/**
 * Get user settings
 */
function getSettings() {
  const properties = PropertiesService.getDocumentProperties();
  const settings = properties.getProperties();
  
  return {
    disparateImpact: parseFloat(settings['biasDetector.disparateImpact'] || '0.8'),
    statisticalParity: parseFloat(settings['biasDetector.statisticalParity'] || '0.1'),
    reportFormat: settings['biasDetector.reportFormat'] || 'detailed',
    insightLevel: settings['biasDetector.insightLevel'] || 'detailed',
    minGroupSize: parseInt(settings['biasDetector.minGroupSize'] || '30'),
    confidenceLevel: parseFloat(settings['biasDetector.confidenceLevel'] || '0.95')
  };
}

/**
 * Get last analysis results for dashboard
 */
function getLastAnalysisResults() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = spreadsheet.getSheetByName(CONFIG.HISTORY_SHEET_NAME);
    
    if (!historySheet || historySheet.getLastRow() < 2) {
      return null;
    }
    
    const lastRow = historySheet.getLastRow();
    const data = historySheet.getRange(lastRow, 1, 1, 10).getValues()[0];
    
    return {
      timestamp: data[0],
      totalRecords: data[1],
      overallBiasScore: parseFloat(data[3].replace('%', '')) / 100,
      riskLevel: data[8]
    };
    
  } catch (error) {
    console.error('Error getting last analysis results:', error);
    return null;
  }
}

/**
 * Show help documentation
 */
function showHelp() {
  const html = createHelpHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setTitle('📖 AI Bias Detector Help')
    .setWidth(600)
    .setHeight(700);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Help & Documentation');
}

/**
 * Create help HTML
 */
function createHelpHTML() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { 
      font-family: 'Google Sans', Arial, sans-serif; 
      padding: 20px; 
      line-height: 1.6;
      background: #f8f9fa;
    }
    
    .help-container {
      background: white;
      padding: 24px;
      border-radius: 8px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
      max-width: 100%;
    }
    
    h1 { color: #1976d2; margin-bottom: 24px; }
    h2 { color: #333; border-bottom: 2px solid #e8eaed; padding-bottom: 8px; }
    h3 { color: #555; }
    
    .section { margin-bottom: 32px; }
    
    .feature-list {
      background: #f8f9fa;
      padding: 16px;
      border-radius: 6px;
      margin: 16px 0;
    }
    
    .metric-explanation {
      background: #e8f0fe;
      padding: 16px;
      border-left: 4px solid #4285f4;
      margin: 12px 0;
    }
    
    .tip {
      background: #e8f5e8;
      padding: 12px;
      border-left: 4px solid #34a853;
      margin: 12px 0;
    }
    
    code {
      background: #f1f3f4;
      padding: 2px 6px;
      border-radius: 3px;
      font-family: 'Courier New', monospace;
    }
    
    .btn-close {
      background: #4285f4;
      color: white;
      border: none;
      padding: 10px 20px;
      border-radius: 4px;
      cursor: pointer;
      float: right;
      margin-top: 20px;
    }
  </style>
</head>
<body>
  <div class="help-container">
    <h1>🤖 AI Bias Detector Documentation</h1>
    
    <div class="section">
      <h2>🎯 What is Bias Detection?</h2>
      <p>AI Bias Detection helps identify unfair treatment of different groups in your data. It automatically analyzes your dataset to find disparities in outcomes across protected attributes like gender, race, age, etc.</p>
      
      <div class="feature-list">
        <h3>Key Features:</h3>
        <ul>
          <li>🔍 <strong>Automatic Detection:</strong> Finds protected attributes in your data</li>
          <li>📊 <strong>Comprehensive Metrics:</strong> Calculates disparate impact, statistical parity, and more</li>
          <li>🧠 <strong>AI Insights:</strong> Natural language explanations of bias patterns</li>
          <li>📈 <strong>Visualizations:</strong> Charts and graphs to illustrate bias</li>
          <li>💡 <strong>Recommendations:</strong> Actionable steps to reduce bias</li>
          <li>📝 <strong>Professional Reports:</strong> Detailed analysis reports</li>
        </ul>
      </div>
    </div>
    
    <div class="section">
      <h2>📏 Understanding Bias Metrics</h2>
      
      <div class="metric-explanation">
        <h3>Disparate Impact (80% Rule)</h3>
        <p>Measures the ratio of positive outcomes between groups. Values below 80% indicate potential bias.</p>
        <p><strong>Example:</strong> If 90% of men get loan approvals but only 60% of women do, the disparate impact is 67% (60/90), indicating bias.</p>
      </div>
      
      <div class="metric-explanation">
        <h3>Statistical Parity</h3>
        <p>Measures the difference in positive outcome rates between groups. Lower differences indicate better fairness.</p>
        <p><strong>Example:</strong> If the approval rate difference between groups is 30%, this suggests significant bias.</p>
      </div>
      
      <div class="metric-explanation">
        <h3>Equal Opportunity</h3>
        <p>Ensures that qualified individuals from all groups have equal chances of positive outcomes.</p>
      </div>
    </div>
    
    <div class="section">
      <h2>🚀 How to Use</h2>
      
      <h3>Step 1: Prepare Your Data</h3>
      <ul>
        <li>Ensure your spreadsheet has headers in the first row</li>
        <li>Include columns with protected attributes (gender, race, age, etc.)</li>
        <li>Have target/outcome columns (predictions, scores, approvals, etc.)</li>
        <li>Clean your data (remove empty rows, fix inconsistent values)</li>
      </ul>
      
      <div class="tip">
        <strong>💡 Tip:</strong> Protected attributes should have consistent values (e.g., "Male/Female" not "M/F/Male/Female")
      </div>
      
      <h3>Step 2: Run Analysis</h3>
      <ul>
        <li>Go to <code>🤖 AI Bias Detector</code> → <code>🔍 Run Bias Analysis</code></li>
        <li>The tool will automatically detect protected attributes</li>
        <li>Analysis results appear in a new "Bias_Analysis_Report" sheet</li>
        <li>View the dashboard for quick insights</li>
      </ul>
      
      <h3>Step 3: Interpret Results</h3>
      <ul>
        <li><strong>Green indicators:</strong> Low bias risk</li>
        <li><strong>Yellow indicators:</strong> Moderate bias - monitor closely</li>
        <li><strong>Red indicators:</strong> High bias - immediate action needed</li>
      </ul>
    </div>
    
    <div class="section">
      <h2>⚠️ Common Issues & Solutions</h2>
      
      <h3>"No Protected Attributes Found"</h3>
      <p>The tool couldn't automatically detect protected attributes. Ensure column names include keywords like:</p>
      <ul>
        <li>gender, sex, male, female</li>
        <li>race, ethnicity, black, white, asian, hispanic</li>
        <li>age, old, young</li>
        <li>religion, nationality</li>
      </ul>
      
      <h3>"Insufficient Data"</h3>
      <p>You need at least 30 records per group for reliable analysis. Consider:</p>
      <ul>
        <li>Collecting more data</li>
        <li>Combining similar groups</li>
        <li>Using the Quick Scan for preliminary insights</li>
      </ul>
    </div>
    
    <div class="section">
      <h2>🔧 Best Practices</h2>
      <ul>
        <li><strong>Regular Monitoring:</strong> Run bias checks regularly, especially after model updates</li>
        <li><strong>Multiple Metrics:</strong> Don't rely on a single metric - consider all fairness measures</li>
        <li><strong>Context Matters:</strong> Understand your domain-specific fairness requirements</li>
        <li><strong>Documentation:</strong> Keep records of bias analyses for compliance</li>
        <li><strong>Action-Oriented:</strong> Use recommendations to actually improve fairness</li>
      </ul>
    </div>
    
    <button class="btn-close" onclick="google.script.host.close()">Close</button>
  </div>
</body>
</html>
  `;
}

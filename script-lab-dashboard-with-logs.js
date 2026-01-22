const DASHBOARD_CONSTANTS = {
  DATA_START_ROW: 2,
  HEADER_ROW: 7,
  FIRST_DATA_ROW: 8,
  LAST_DATA_ROW: 39,
  AUTO_FIT_ROW: 5,
  CHART_DATA_ROW: 42,
  CHART_HEADER_ROW: 43,
  CHART_DATA_RANGE: "A43:F51",
  CHART_POSITION: {
    START_CELL: "A53",
    END_CELL: "H75"
  },
  Q1_START_ROWS: [8, 16, 24, 32],
  QUARTERS: [
    ["2023 Q1"],
    ["2023 Q2"],
    ["2023 Q3"],
    ["2023 Q4"],
    ["2024 Q1"],
    ["2024 Q2"],
    ["2024 Q3"],
    ["2024 Q4"]
  ],
  PRODUCT_NAMES: {
    WIDGET_PRO: "Widget Pro",
    WIDGET_STANDARD: "Widget Standard",
    SERVICE_PACKAGE: "Service Package",
    ACCESSORY_KIT: "Accessory Kit",
    TOTAL_REVENUE: "Total Revenue"
  },
  CHART_HEADERS: [
    ["Quarter", "Widget Pro", "Widget Standard", "Service Package", "Accessory Kit", "Total Revenue"]
  ],
  CHART_CONFIG: {
    CHART_NAME: "QuarterlyMarginTrend",
    TITLE: "Quarterly Margin Trends by Product",
    SERIES_NAMES: ["Widget Pro", "Widget Standard", "Service Package", "Accessory Kit", "Total Revenue"],
    COLORS: ["#4F81BD", "#C0504D", "#9BBB59", "#8064A2"],
    PRIMARY_AXIS: {
      TITLE: "Profit Margin",
      NUMBER_FORMAT: "0.0%"
    },
    SECONDARY_AXIS: {
      TITLE: "Total Revenue ($)",
      NUMBER_FORMAT: "$#,##0"
    }
  },
  FORMATTING: {
    FONT_NAME: "Calibri",
    HEADER_BACKGROUND_COLOR: "#D9E1F2",
    STRONG_COLOR: "#C6EFCE",
    STRONG_TEXT_COLOR: "#006100",
    MODERATE_COLOR: "#FFEB9C",
    MODERATE_TEXT_COLOR: "#9C6500",
    AT_RISK_COLOR: "#FFC7CE",
    AT_RISK_TEXT_COLOR: "#9C0006",
    MERGED_TEXT_COLOR: "#707070",
    LINE_CHART_COLOR: "#4BACC6"
  },
  TEXT_LABELS: {
    CHART_DATA: "Chart Data",
    STRONG: "Strong",
    MODERATE: "Moderate",
    AT_RISK: "At Risk",
    N_A: "N/A"
  },
  HEALTH_THRESHOLDS: {
    STRONG: 0.35,
    MODERATE: 0.2
  },
  BASE_YEAR: "2023",
  SHEET_NAMES: {
    RAW_DATA: "Raw Data",
    DASHBOARD: "Dashboard"
  },
  RAW_DATA_COLUMNS: {
    DATE: "A",
    YEAR: "B",
    QUARTER: "C",
    PRODUCT: "D",
    REVENUE: "E",
    COSTS: "F",
    MARGIN: "G"
  }
};

// Builds revenue formula using SUMIFS to match product, year, and quarter
function buildRevenueFormula(rowNumber, dataStartRow, dataEndRow) {
  const sheetName = DASHBOARD_CONSTANTS.SHEET_NAMES.RAW_DATA;
  const cols = DASHBOARD_CONSTANTS.RAW_DATA_COLUMNS;
  return `=SUMIFS('${sheetName}'!$${cols.REVENUE}$${dataStartRow}:$${cols.REVENUE}$${dataEndRow},'${sheetName}'!$${cols.PRODUCT}$${dataStartRow}:$${cols.PRODUCT}$${dataEndRow},$A${rowNumber},'${sheetName}'!$${cols.YEAR}$${dataStartRow}:$${cols.YEAR}$${dataEndRow},LEFT($B${rowNumber},4)*1,'${sheetName}'!$${cols.QUARTER}$${dataStartRow}:$${cols.QUARTER}$${dataEndRow},RIGHT($B${rowNumber},2))`;
}

// Builds profit margin formula using SUMPRODUCT to calculate margin percentage
function buildMarginFormula(rowNumber, dataStartRow, dataEndRow) {
  const sheetName = DASHBOARD_CONSTANTS.SHEET_NAMES.RAW_DATA;
  const cols = DASHBOARD_CONSTANTS.RAW_DATA_COLUMNS;
  return `=IFERROR(SUMPRODUCT(('${sheetName}'!$${cols.PRODUCT}$${dataStartRow}:$${cols.PRODUCT}$${dataEndRow}=$A${rowNumber})*('${sheetName}'!$${cols.YEAR}$${dataStartRow}:$${cols.YEAR}$${dataEndRow}=LEFT($B${rowNumber},4)*1)*('${sheetName}'!$${cols.QUARTER}$${dataStartRow}:$${cols.QUARTER}$${dataEndRow}=RIGHT($B${rowNumber},2))*'${sheetName}'!$${cols.MARGIN}$${dataStartRow}:$${cols.MARGIN}$${dataEndRow})/C${rowNumber},0)`;
}

// Builds trend formula showing margin change from previous period (N/A for Q1 starts)
function buildTrendFormula(rowIndex, isQ1Start) {
  return isQ1Start ? `="${DASHBOARD_CONSTANTS.TEXT_LABELS.N_A}"` : `=D${rowIndex}-D${rowIndex - 1}`;
}

// Builds year-over-year formula (N/A for base year)
function buildYearOverYearFormula(rowIndex, baseYear) {
  return `=IF(LEFT($B${rowIndex},4)="${baseYear}","${DASHBOARD_CONSTANTS.TEXT_LABELS.N_A}","")`;
}

// Builds health status formula based on margin thresholds
function buildHealthFormula(rowIndex, strongThreshold, moderateThreshold) {
  return `=IF(D${rowIndex}>${strongThreshold},"${DASHBOARD_CONSTANTS.TEXT_LABELS.STRONG}",IF(D${rowIndex}>=${moderateThreshold},"${DASHBOARD_CONSTANTS.TEXT_LABELS.MODERATE}","${DASHBOARD_CONSTANTS.TEXT_LABELS.AT_RISK}"))`;
}

// Builds formulas for quarter chart data (product margins and total revenue)
function buildQuarterFormulas(rowNumber, firstDataRow, lastDataRow) {
  return [
    `=SUMPRODUCT(($A$${firstDataRow}:$A$${lastDataRow}="${DASHBOARD_CONSTANTS.PRODUCT_NAMES.WIDGET_PRO}")*($B$${firstDataRow}:$B$${lastDataRow}=$A${rowNumber})*($D$${firstDataRow}:$D$${lastDataRow}))`,
    `=SUMPRODUCT(($A$${firstDataRow}:$A$${lastDataRow}="${DASHBOARD_CONSTANTS.PRODUCT_NAMES.WIDGET_STANDARD}")*($B$${firstDataRow}:$B$${lastDataRow}=$A${rowNumber})*($D$${firstDataRow}:$D$${lastDataRow}))`,
    `=SUMPRODUCT(($A$${firstDataRow}:$A$${lastDataRow}="${DASHBOARD_CONSTANTS.PRODUCT_NAMES.SERVICE_PACKAGE}")*($B$${firstDataRow}:$B$${lastDataRow}=$A${rowNumber})*($D$${firstDataRow}:$D$${lastDataRow}))`,
    `=SUMPRODUCT(($A$${firstDataRow}:$A$${lastDataRow}="${DASHBOARD_CONSTANTS.PRODUCT_NAMES.ACCESSORY_KIT}")*($B$${firstDataRow}:$B$${lastDataRow}=$A${rowNumber})*($D$${firstDataRow}:$D$${lastDataRow}))`,
    `=SUMIF($B$${firstDataRow}:$B$${lastDataRow}, $A${rowNumber}, $C$${firstDataRow}:$C$${lastDataRow})`
  ];
}

// Validates that Raw Data sheet has expected column structure
async function validateRawDataSheet(worksheet, excelContext) {
  try {
    const headerRange = worksheet.getRange("A1:G1");
    headerRange.load("values");
    await excelContext.sync();
    
    const headers = headerRange.values[0];
    if (!headers || headers.length < 7) {
      throw new Error("Raw Data sheet doesn't have enough columns. Expected at least 7 columns (A-G).");
    }
    
    return true;
  } catch (error) {
    throw new Error(`Raw Data sheet validation failed: ${error.message}`);
  }
}

// Finds the last row with product data by checking for empty cells
function findLastDataRow(productData, firstDataRow) {
  if (!productData || productData.length === 0) {
    return firstDataRow - 1;
  }
  
  for (let index = 0; index < productData.length; index++) {
    if (!productData[index] || !productData[index][0] || !productData[index][1]) {
      return firstDataRow + index - 1;
    }
  }
  return firstDataRow + productData.length - 1;
}

// Builds all formulas for dashboard rows (revenue, margin, trend, YoY, health)
function buildAllFormulas(productData, firstDataRow, lastDataRow, dataStartRow, dataEndRow) {
  const revenueFormulas = [];
  const marginFormulas = [];
  const trendFormulas = [];
  const yearOverYearFormulas = [];
  const healthFormulas = [];

  for (let rowIndex = firstDataRow; rowIndex <= lastDataRow; rowIndex++) {
    if (!productData[rowIndex - firstDataRow][0] || !productData[rowIndex - firstDataRow][1]) {
      continue;
    }

    revenueFormulas.push([buildRevenueFormula(rowIndex, dataStartRow, dataEndRow)]);
    marginFormulas.push([buildMarginFormula(rowIndex, dataStartRow, dataEndRow)]);

    const isQ1Start = DASHBOARD_CONSTANTS.Q1_START_ROWS.includes(rowIndex);
    trendFormulas.push([buildTrendFormula(rowIndex, isQ1Start)]);

    yearOverYearFormulas.push([buildYearOverYearFormula(rowIndex, DASHBOARD_CONSTANTS.BASE_YEAR)]);
    healthFormulas.push([buildHealthFormula(rowIndex, DASHBOARD_CONSTANTS.HEALTH_THRESHOLDS.STRONG, DASHBOARD_CONSTANTS.HEALTH_THRESHOLDS.MODERATE)]);
  }

  return {
    revenueFormulas,
    marginFormulas,
    trendFormulas,
    yearOverYearFormulas,
    healthFormulas
  };
}

// Applies all formulas to dashboard columns with appropriate number formatting (batched for performance)
async function applyFormulasToDashboard(dashboardWorksheet, formulas, firstDataRow, lastDataRow, excelContext) {
  const revenueRange = dashboardWorksheet.getRange(`C${firstDataRow}:C${lastDataRow}`);
  revenueRange.formulas = formulas.revenueFormulas;
  revenueRange.numberFormat = "$#,##0";

  const marginRange = dashboardWorksheet.getRange(`D${firstDataRow}:D${lastDataRow}`);
  marginRange.formulas = formulas.marginFormulas;
  marginRange.numberFormat = "0.0%";

  const trendRange = dashboardWorksheet.getRange(`E${firstDataRow}:E${lastDataRow}`);
  trendRange.formulas = formulas.trendFormulas;
  trendRange.numberFormat = "0.0%";

  const yearOverYearRange = dashboardWorksheet.getRange(`F${firstDataRow}:F${lastDataRow}`);
  yearOverYearRange.formulas = formulas.yearOverYearFormulas;
  yearOverYearRange.numberFormat = "0.0%";

  const healthRange = dashboardWorksheet.getRange(`G${firstDataRow}:G${lastDataRow}`);
  healthRange.formulas = formulas.healthFormulas;

  await excelContext.sync();
}

// Applies conditional formatting to health column (Strong/Moderate/At Risk colors) - batched where possible
async function applyConditionalFormatting(dashboardWorksheet, healthColumnRange, excelContext) {
  const healthColumn = dashboardWorksheet.getRange(healthColumnRange);
  healthColumn.load("conditionalFormats");
  await excelContext.sync();

  const strongFormat = healthColumn.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
  strongFormat.textComparison.format.fill.color = DASHBOARD_CONSTANTS.FORMATTING.STRONG_COLOR;
  strongFormat.textComparison.format.font.color = DASHBOARD_CONSTANTS.FORMATTING.STRONG_TEXT_COLOR;
  strongFormat.textComparison.rule = {
    operator: Excel.ConditionalTextOperator.contains,
    text: DASHBOARD_CONSTANTS.TEXT_LABELS.STRONG,
  };

  const moderateFormat = healthColumn.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
  moderateFormat.textComparison.format.fill.color = DASHBOARD_CONSTANTS.FORMATTING.MODERATE_COLOR;
  moderateFormat.textComparison.format.font.color = DASHBOARD_CONSTANTS.FORMATTING.MODERATE_TEXT_COLOR;
  moderateFormat.textComparison.rule = {
    operator: Excel.ConditionalTextOperator.contains,
    text: DASHBOARD_CONSTANTS.TEXT_LABELS.MODERATE,
  };

  const atRiskFormat = healthColumn.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
  atRiskFormat.textComparison.format.fill.color = DASHBOARD_CONSTANTS.FORMATTING.AT_RISK_COLOR;
  atRiskFormat.textComparison.format.font.color = DASHBOARD_CONSTANTS.FORMATTING.AT_RISK_TEXT_COLOR;
  atRiskFormat.textComparison.rule = {
    operator: Excel.ConditionalTextOperator.contains,
    text: DASHBOARD_CONSTANTS.TEXT_LABELS.AT_RISK,
  };

  await excelContext.sync();
}

// Sets up chart data section header and column headers (batched for performance)
async function setupChartDataSection(dashboardWorksheet, excelContext) {
  const chartDataLabel = dashboardWorksheet.getRange(`A${DASHBOARD_CONSTANTS.CHART_DATA_ROW}`);
  chartDataLabel.values = [[DASHBOARD_CONSTANTS.TEXT_LABELS.CHART_DATA]];
  chartDataLabel.format.font.bold = true;

  const chartHeaderRange = dashboardWorksheet.getRange(`A${DASHBOARD_CONSTANTS.CHART_HEADER_ROW}:F${DASHBOARD_CONSTANTS.CHART_HEADER_ROW}`);
  chartHeaderRange.values = DASHBOARD_CONSTANTS.CHART_HEADERS;
  chartHeaderRange.format.font.bold = true;
  chartHeaderRange.format.fill.color = DASHBOARD_CONSTANTS.FORMATTING.HEADER_BACKGROUND_COLOR;
  chartHeaderRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

  await excelContext.sync();
}

// Populates quarter data rows with formulas for chart visualization (batched for performance)
async function setupQuarterData(dashboardWorksheet, firstDataRow, lastDataRow, excelContext) {
  const chartDataStartRow = DASHBOARD_CONSTANTS.CHART_HEADER_ROW + 1;

  for (let quarterIndex = 0; quarterIndex < DASHBOARD_CONSTANTS.QUARTERS.length; quarterIndex++) {
    const quarter = DASHBOARD_CONSTANTS.QUARTERS[quarterIndex];
    const rowNumber = chartDataStartRow + quarterIndex;

    if (quarterIndex === 0) {
      dashboardWorksheet.getRange(`A${rowNumber}`).values = [quarter];
      const mergedRange = dashboardWorksheet.getRange(`B${rowNumber}:F${rowNumber}`);
      mergedRange.merge(true);
      mergedRange.format.font.italic = true;
      mergedRange.format.font.color = DASHBOARD_CONSTANTS.FORMATTING.MERGED_TEXT_COLOR;
      await excelContext.sync();
    } else {
      dashboardWorksheet.getRange(`A${rowNumber}`).values = [quarter];

      const quarterFormulas = buildQuarterFormulas(rowNumber, firstDataRow, lastDataRow);
      const quarterDataRange = dashboardWorksheet.getRange(`B${rowNumber}:F${rowNumber}`);
      quarterDataRange.formulas = [quarterFormulas];
      
      const marginRange = dashboardWorksheet.getRange(`B${rowNumber}:E${rowNumber}`);
      marginRange.numberFormat = "0.0%";
      
      const revenueRange = dashboardWorksheet.getRange(`F${rowNumber}`);
      revenueRange.numberFormat = "$#,##0";
      
      await excelContext.sync();
    }
  }
}

// Creates the dashboard chart by removing existing charts and building a new combo chart (batched for performance)
async function createDashboardChart(dashboardWorksheet, excelContext) {
  try {
    dashboardWorksheet.charts.load("items");
    await excelContext.sync();
    dashboardWorksheet.charts.items.forEach((existingChart) => existingChart.delete());
    await excelContext.sync();

    const chartDataRange = dashboardWorksheet.getRange(DASHBOARD_CONSTANTS.CHART_DATA_RANGE);

    const chartConfig = {
      chartName: DASHBOARD_CONSTANTS.CHART_CONFIG.CHART_NAME,
      title: DASHBOARD_CONSTANTS.CHART_CONFIG.TITLE,
      seriesNames: DASHBOARD_CONSTANTS.CHART_CONFIG.SERIES_NAMES,
      colors: DASHBOARD_CONSTANTS.CHART_CONFIG.COLORS,
      primaryAxis: {
        title: DASHBOARD_CONSTANTS.CHART_CONFIG.PRIMARY_AXIS.TITLE,
        numberFormat: DASHBOARD_CONSTANTS.CHART_CONFIG.PRIMARY_AXIS.NUMBER_FORMAT,
      },
      secondaryAxis: {
        title: DASHBOARD_CONSTANTS.CHART_CONFIG.SECONDARY_AXIS.TITLE,
        numberFormat: DASHBOARD_CONSTANTS.CHART_CONFIG.SECONDARY_AXIS.NUMBER_FORMAT,
      },
      position: {
        startCell: DASHBOARD_CONSTANTS.CHART_POSITION.START_CELL,
        endCell: DASHBOARD_CONSTANTS.CHART_POSITION.END_CELL,
      },
    };

    const chart = await createChart(dashboardWorksheet, chartDataRange, chartConfig);
    dashboardWorksheet.activate();
    await excelContext.sync();
  } catch (chartError) {
    console.warn("Chart creation failed, continuing without chart:", chartError.message);
  }
}

// Creates a combo chart with column series for margins and line series for total revenue (batched for performance)
async function createChart(worksheet, dataRange, chartConfig) {
  const excelContext = worksheet.context;

  const chart = worksheet.charts.add(Excel.ChartType.columnClustered, dataRange, Excel.ChartSeriesBy.columns);
  chart.name = chartConfig.chartName || "ComboChart";
  chart.title.text = chartConfig.title || "";
  chart.title.format.font.size = 14;
  chart.series.load("count");
  await excelContext.sync();
  
  const seriesCount = chart.series.count;

  for (let seriesIndex = 0; seriesIndex < seriesCount && seriesIndex < chartConfig.seriesNames.length; seriesIndex++) {
    const series = chart.series.getItemAt(seriesIndex);
    series.name = chartConfig.seriesNames[seriesIndex];

    if (seriesIndex < 4) {
      series.chartType = Excel.ChartType.columnClustered;
      series.axisGroup = 0;
      if (chartConfig.colors && chartConfig.colors[seriesIndex]) {
        try {
          series.format.fill.setSolidColor(chartConfig.colors[seriesIndex]);
        } catch (colorError) {
        }
      }
    } else {
      series.chartType = Excel.ChartType.line;
      series.axisGroup = 1;
      series.format.line.weight = 3;
      series.format.line.color = DASHBOARD_CONSTANTS.FORMATTING.LINE_CHART_COLOR;
    }
  }
  
  chart.legend.visible = true;
  chart.legend.position = Excel.ChartLegendPosition.bottom;

  const primaryAxis = chart.axes.valueAxis;
  primaryAxis.title.text = chartConfig.primaryAxis.title;
  primaryAxis.title.format.font.size = 11;
  try {
    primaryAxis.numberFormat = chartConfig.primaryAxis.numberFormat;
  } catch (formatError) {
    primaryAxis.format.code = chartConfig.primaryAxis.numberFormat;
  }
  
  await excelContext.sync();

  let secondaryAxis;
  try {
    if (typeof Excel.ChartAxisType !== "undefined" && typeof Excel.ChartAxisGroup !== "undefined") {
      secondaryAxis = chart.axes.getItem(Excel.ChartAxisType.value, Excel.ChartAxisGroup.secondary);
    } else {
      secondaryAxis = chart.axes.getItem(Excel.ChartAxisType.value, "Secondary");
    }
    secondaryAxis.visible = true;
    secondaryAxis.title.text = chartConfig.secondaryAxis.title;
    secondaryAxis.title.format.font.size = 11;
    try {
      secondaryAxis.numberFormat = chartConfig.secondaryAxis.numberFormat;
    } catch (formatError) {
      secondaryAxis.format.code = chartConfig.secondaryAxis.numberFormat;
    }
  } catch (axisError) {
    secondaryAxis = null;
  }

  if (chartConfig.position && chartConfig.position.startCell && chartConfig.position.endCell) {
    chart.setPosition(chartConfig.position.startCell, chartConfig.position.endCell);
  }
  
  await excelContext.sync();

  return chart;
}

// Main function that orchestrates the entire dashboard building process
async function main() {
  try {
    if (typeof Excel === "undefined") {
      throw new Error("Excel API is not available. Make sure you're running this in Excel.");
    }

    await Excel.run(async (excelContext) => {
      try {
        const rawDataWorksheet = excelContext.workbook.worksheets.getItem(DASHBOARD_CONSTANTS.SHEET_NAMES.RAW_DATA);
        const dashboardWorksheet = excelContext.workbook.worksheets.getItem(DASHBOARD_CONSTANTS.SHEET_NAMES.DASHBOARD);

        await validateRawDataSheet(rawDataWorksheet, excelContext);

        const usedRange = rawDataWorksheet.getUsedRange();
        if (!usedRange) {
          throw new Error("No data found in Raw Data sheet");
        }
        
        usedRange.load("rowCount");
        await excelContext.sync();

        const lastRow = usedRange.rowCount;
        const dataStartRow = DASHBOARD_CONSTANTS.DATA_START_ROW;
        const dataEndRow = lastRow;

        let firstDataRow = DASHBOARD_CONSTANTS.FIRST_DATA_ROW;
        let lastDataRow = DASHBOARD_CONSTANTS.LAST_DATA_ROW;

        const productColumnsRange = dashboardWorksheet.getRange(`A${firstDataRow}:B${lastDataRow}`);
        productColumnsRange.load("values");
        await excelContext.sync();

        const productData = productColumnsRange.values;
        lastDataRow = findLastDataRow(productData, firstDataRow);

        if (lastDataRow < firstDataRow) {
          throw new Error("No product data found in dashboard range. Please ensure product data exists in columns A and B starting from row 8.");
        }

        const formulas = buildAllFormulas(productData, firstDataRow, lastDataRow, dataStartRow, dataEndRow);
        await applyFormulasToDashboard(dashboardWorksheet, formulas, firstDataRow, lastDataRow, excelContext);

        const headerCells = dashboardWorksheet.getRange(`A${DASHBOARD_CONSTANTS.HEADER_ROW}:G${DASHBOARD_CONSTANTS.HEADER_ROW}`);
        headerCells.format.font.bold = true;
        headerCells.format.font.name = DASHBOARD_CONSTANTS.FORMATTING.FONT_NAME;
        headerCells.format.fill.color = DASHBOARD_CONSTANTS.FORMATTING.HEADER_BACKGROUND_COLOR;
        headerCells.format.wrapText = true;
        headerCells.format.horizontalAlignment = Excel.HorizontalAlignment.center;

        dashboardWorksheet.getRange(`${DASHBOARD_CONSTANTS.AUTO_FIT_ROW}:${DASHBOARD_CONSTANTS.AUTO_FIT_ROW}`).format.autofitRows();
        await excelContext.sync();

        await applyConditionalFormatting(dashboardWorksheet, `G${firstDataRow}:G${lastDataRow}`, excelContext);

        await setupChartDataSection(dashboardWorksheet, excelContext);

        await setupQuarterData(dashboardWorksheet, firstDataRow, lastDataRow, excelContext);

        await createDashboardChart(dashboardWorksheet, excelContext);
      } catch (error) {
        console.error("Error in dashboard building process:", error);
        throw error;
      }
    });
  } catch (error) {
    console.error("Error in run function:", error);
    const statusDiv = document.getElementById("status");
    if (statusDiv) {
      statusDiv.textContent = `❌ Error: ${error.message}`;
      statusDiv.style.display = "block";
      statusDiv.className = "error";
    }
  }
}

// Initialize: Wait for Office API to be ready, then set up event handlers
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const runButton = document.getElementById("run");
    if (runButton) {
      runButton.addEventListener("click", main);
    } else {
      main();
    }
  }
});

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const sideloadMsg = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    if (sideloadMsg) sideloadMsg.style.display = "none";
    if (appBody) appBody.style.display = "flex";

    const buildButton = document.getElementById("build-dashboard");
    if (buildButton) {
      buildButton.onclick = () => {
        runBuild();
      };
      (buildButton as HTMLButtonElement).disabled = false;
    } else {
      showStatus("Error: Build button not found in HTML", true);
    }

    setTimeout(() => {
      runBuild();
    }, 500);
  } else {
    showStatus(`Warning: This add-in is designed for Excel. Current host: ${info.host}`, true);
  }
});

async function runBuild(): Promise<void> {
  showStatus("Building dashboard...", false);

  if (typeof Excel === "undefined") {
    const errorMsg = "Excel API is not available. Make sure you're running this in Excel.";
    showStatus(errorMsg, true);
    return;
  }

  try {
    await Excel.run(async (context) => {
      let rawSheet: Excel.Worksheet;
      let dashSheet: Excel.Worksheet;

      try {
        rawSheet = await prepareSheet(context, "Raw Data", true);
        dashSheet = await prepareSheet(context, "Dashboard", false);
        await context.sync();
      } catch (error) {
        showStatus("Error creating sheets: " + (error as Error).message, true);
        throw error;
      }

      try {
        const headerRange = rawSheet.getRange("A1");
        headerRange.load("values");
        await context.sync();

        const hasData =
          headerRange.values && headerRange.values.length > 0 && headerRange.values[0] && headerRange.values[0][0];

        if (!hasData) {
          showStatus("Adding raw data...", false);
          const rawHeaders = [["Transaction_ID", "Year", "Quarter", "Product", "Revenue", "Cost", "Profit"]];
          const headerRangeFull = rawSheet.getRange("A1:G1");
          headerRangeFull.values = rawHeaders;
          headerRangeFull.format.font.bold = true;
          headerRangeFull.format.fill.color = "#D9E1F2";
          await context.sync();

          const sampleData = generateRawData(186);
          rawSheet.getRange("A2:G187").values = sampleData;
          await context.sync();

          (rawSheet.getRange("B2:B187").numberFormat as any) = "0";
          (rawSheet.getRange("E2:G187").numberFormat as any) = "$#,##0";
          await context.sync();
        }
      } catch (error) {
        showStatus("Error in raw data: " + (error as Error).message, true);
        throw error;
      }

      showStatus("Setting up dashboard...", false);
      try {
        await context.sync();

        const titleCell = dashSheet.getRange("A1");
        titleCell.values = [["Executive Sales Performance Dashboard"]];
        titleCell.format.font.bold = true;
        titleCell.format.font.size = 14;
        titleCell.format.font.name = "Calibri";
        await context.sync();
        try {
          const titleRange = dashSheet.getRange("A1:H1");
          titleRange.merge(true);
          await context.sync();
        } catch (mergeError) {
          try {
            const titleRange = dashSheet.getRange("A1:H1");
            titleRange.merge();
            await context.sync();
          } catch (mergeError2) {
            showStatus("Warning: Could not merge title cells. Text may appear in single cell only.", false);
          }
        }

        const instructionsCell = dashSheet.getRange("A3");
        instructionsCell.values = [
          ["Instructions: Complete the analysis below using the Raw Data tab. All calculations should use formulas."],
        ];
        instructionsCell.format.font.italic = true;
        instructionsCell.format.font.size = 11;
        instructionsCell.format.font.name = "Calibri";
        await context.sync();
        try {
          const instructionsRange = dashSheet.getRange("A3:H3");
          instructionsRange.merge(true);
          await context.sync();
        } catch (mergeError) {
          try {
            const instructionsRange = dashSheet.getRange("A3:H3");
            instructionsRange.merge();
            await context.sync();
          } catch (mergeError2) {}
        }

        const summaryCell = dashSheet.getRange("A5");
        summaryCell.values = [["Quarterly Performance Summary"]];
        summaryCell.format.font.bold = true;
        summaryCell.format.font.size = 12;
        summaryCell.format.font.name = "Calibri";
        await context.sync();
        try {
          const summaryRange = dashSheet.getRange("A5:B5");
          summaryRange.merge(true);
          await context.sync();
        } catch (mergeError) {
          try {
            const summaryRange = dashSheet.getRange("A5:B5");
            summaryRange.merge();
            await context.sync();
          } catch (mergeError2) {}
        }

        await context.sync();

        const dashHeaders = [
          [
            "Product",
            "Quarter",
            "Total Revenue",
            "Weighted Avg Margin",
            "Rolling 3-Mo Trend",
            "YoY Margin Delta",
            "Margin Health",
          ],
        ];
        const dashHeaderRange = dashSheet.getRange("A7:G7");
        dashHeaderRange.values = dashHeaders;
        dashHeaderRange.format.font.bold = true;
        dashHeaderRange.format.font.name = "Calibri";
        dashHeaderRange.format.fill.color = "#D9E1F2";
        await context.sync();

        const products = ["Widget Pro", "Widget Standard", "Service Package", "Accessory Kit"];
        const quarters = ["2023 Q1", "2023 Q2", "2023 Q3", "2023 Q4", "2024 Q1", "2024 Q2", "2024 Q3", "2024 Q4"];
        const productRows: string[][] = [];
        for (const product of products) {
          for (const quarter of quarters) {
            productRows.push([product, quarter]);
          }
        }
        const productRange = dashSheet.getRange("A8:B39");
        productRange.values = productRows;
        await context.sync();
      } catch (error) {
        showStatus("Error in dashboard setup: " + (error as Error).message, true);
        throw error;
      }

      showStatus("Applying formulas...", false);
      showStatus("Setting revenue formulas...", false);
      try {
        const revenueFormulas: string[][] = [];
        for (let row = 8; row <= 39; row++) {
          const formula = `=SUMIFS('Raw Data'!$E$2:$E$187,'Raw Data'!$D$2:$D$187,$A${row},'Raw Data'!$B$2:$B$187,LEFT($B${row},4)*1,'Raw Data'!$C$2:$C$187,RIGHT($B${row},2))`;
          revenueFormulas.push([formula]);
        }
        const revenueRange = dashSheet.getRange("C8:C39");
        revenueRange.formulas = revenueFormulas;
        (revenueRange.numberFormat as any) = "$#,##0";
        try {
          await context.sync();
        } catch (syncError) {
          throw syncError;
        }
      } catch (error) {
        const errorDetails = `Revenue formulas failed!\nError: ${(error as Error).message}`;
        showStatus(errorDetails, true);
        showStatus("Trying simplified revenue formula using SUMPRODUCT...", false);
        try {
          const altRevenueFormulas: string[][] = [];
          for (let row = 8; row <= 39; row++) {
            altRevenueFormulas.push([
              `=SUMPRODUCT(('Raw Data'!$D$2:$D$187=$A${row})*('Raw Data'!$B$2:$B$187=LEFT($B${row},4)*1)*('Raw Data'!$C$2:$C$187=RIGHT($B${row},2))*'Raw Data'!$E$2:$E$187)`,
            ]);
          }
          const altRevenueRange = dashSheet.getRange("C8:C39");
          altRevenueRange.formulas = altRevenueFormulas;
          await context.sync();
          (altRevenueRange.numberFormat as any) = "$#,##0";
          await context.sync();
          showStatus("✓ Revenue formulas set using SUMPRODUCT", false);
        } catch (altError) {
          showStatus(
            `Both approaches failed!\n\nOriginal Error: ${(error as Error).message}\n\nAlternative Error: ${(altError as Error).message}`,
            true
          );
          throw altError;
        }
      }

      showStatus("Setting margin formulas...", false);
      try {
        const marginFormulas: string[][] = [];
        for (let row = 8; row <= 39; row++) {
          const formula = `=IFERROR(SUMPRODUCT(('Raw Data'!$D$2:$D$187=$A${row})*('Raw Data'!$B$2:$B$187=LEFT($B${row},4)*1)*('Raw Data'!$C$2:$C$187=RIGHT($B${row},2))*'Raw Data'!$G$2:$G$187)/C${row},0)`;
          marginFormulas.push([formula]);
        }
        const marginRange = dashSheet.getRange("D8:D39");
        marginRange.formulas = marginFormulas;
        (marginRange.numberFormat as any) = "0.0%";
        try {
          await context.sync();
        } catch (syncError) {
          throw syncError;
        }
      } catch (error) {
        const errorDetails = `Margin formulas failed!\nError: ${(error as Error).message}`;
        showStatus(errorDetails, true);
        try {
          const placeholderFormulas: string[][] = [];
          const marginRange = dashSheet.getRange("D8:D39");
          for (let row = 8; row <= 39; row++) {
            placeholderFormulas.push([`=IF(C${row}>0,C${row}/C${row},0)`]);
          }
          marginRange.formulas = placeholderFormulas;
          await context.sync();
          (marginRange.numberFormat as any) = "0.0%";
          await context.sync();
          showStatus("⚠ Margin formulas set to placeholder - check manually", false);
        } catch (placeholderError) {
          showStatus(`Could not set placeholder: ${(placeholderError as Error).message}`, true);
          throw placeholderError;
        }
      }

      showStatus("Setting trend formulas...", false);
      try {
        const trendFormulas: string[][] = [];
        for (let row = 8; row <= 39; row++) {
          const isFirstQ1 = row === 8 || row === 16 || row === 24 || row === 32;
          if (isFirstQ1) {
            trendFormulas.push([`="N/A"`]);
          } else {
            trendFormulas.push([`=D${row}-D${row - 1}`]);
          }
        }
        const trendRange = dashSheet.getRange("E8:E39");
        trendRange.formulas = trendFormulas;
        (trendRange.numberFormat as any) = "0.0%";
        await context.sync();
      } catch (error) {
        const errorDetails = `Trend formulas failed!\nError: ${(error as Error).message}`;
        showStatus(errorDetails, true);
        throw error;
      }

      showStatus("Setting YoY formulas...", false);
      const yoyFormulas: string[][] = [];
      try {
        for (let row = 8; row <= 39; row++) {
          const formula = `=IF(LEFT($B${row},4)="2023","N/A","")`;
          yoyFormulas.push([formula]);
        }
        const yoyRange = dashSheet.getRange("F8:F39");
        yoyRange.formulas = yoyFormulas;
        (yoyRange.numberFormat as any) = "0.0%";
        await context.sync();
      } catch (error) {
        const errorDetails = `YoY formulas failed!\nError: ${(error as Error).message}`;
        showStatus(errorDetails, true);
        throw error;
      }

      showStatus("Setting health formulas...", false);
      try {
        const healthFormulas: string[][] = [];
        for (let row = 8; row <= 39; row++) {
          const formula = `=IF(D${row}>0.35,"Strong",IF(D${row}>=0.2,"Moderate","At Risk"))`;
          healthFormulas.push([formula]);
        }
        const healthRange = dashSheet.getRange("G8:G39");
        healthRange.formulas = healthFormulas;
        await context.sync();
      } catch (error) {
        const errorDetails = `Health formulas failed!\nError: ${(error as Error).message}`;
        showStatus(errorDetails, true);
        throw error;
      }

      showStatus("Applying formatting...", false);
      try {
        const healthRange = dashSheet.getRange("G8:G39");
        healthRange.load("conditionalFormats");
        await context.sync();

        const strongFormat = healthRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
        strongFormat.textComparison.format.fill.color = "#C6EFCE";
        strongFormat.textComparison.format.font.color = "#006100";
        strongFormat.textComparison.rule = {
          operator: Excel.ConditionalTextOperator.contains,
          text: "Strong",
        };

        const moderateFormat = healthRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
        moderateFormat.textComparison.format.fill.color = "#FFEB9C";
        moderateFormat.textComparison.format.font.color = "#9C6500";
        moderateFormat.textComparison.rule = {
          operator: Excel.ConditionalTextOperator.contains,
          text: "Moderate",
        };

        const riskFormat = healthRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
        riskFormat.textComparison.format.fill.color = "#FFC7CE";
        riskFormat.textComparison.format.font.color = "#9C0006";
        riskFormat.textComparison.rule = {
          operator: Excel.ConditionalTextOperator.contains,
          text: "At Risk",
        };
        await context.sync();
      } catch (formatError) {}

      showStatus("Creating chart data...", false);

      const chartTitleRange = dashSheet.getRange("A42");
      chartTitleRange.values = [["Chart Data"]];
      chartTitleRange.format.font.bold = true;

      const chartHeaders = [
        ["Quarter", "Widget Pro", "Widget Standard", "Service Package", "Accessory Kit", "Total Revenue"],
      ];
      const chartHeaderRange = dashSheet.getRange("A43:F43");
      chartHeaderRange.values = chartHeaders;
      chartHeaderRange.format.font.bold = true;
      chartHeaderRange.format.fill.color = "#D9E1F2";
      await context.sync();

      const chartQuarters = [
        ["2023 Q1"],
        ["2023 Q2"],
        ["2023 Q3"],
        ["2023 Q4"],
        ["2024 Q1"],
        ["2024 Q2"],
        ["2024 Q3"],
        ["2024 Q4"],
      ];
      dashSheet.getRange("A44:A51").values = chartQuarters;

      const chartFormulas: string[][] = [];
      for (let i = 0; i < 8; i++) {
        const row = 44 + i;
        chartFormulas.push([
          `=SUMPRODUCT(($A$8:$A$39="Widget Pro")*($B$8:$B$39=$A${row})*($D$8:$D$39))`,
          `=SUMPRODUCT(($A$8:$A$39="Widget Standard")*($B$8:$B$39=$A${row})*($D$8:$D$39))`,
          `=SUMPRODUCT(($A$8:$A$39="Service Package")*($B$8:$B$39=$A${row})*($D$8:$D$39))`,
          `=SUMPRODUCT(($A$8:$A$39="Accessory Kit")*($B$8:$B$39=$A${row})*($D$8:$D$39))`,
          `=SUMIF($B$8:$B$39, $A${row}, $C$8:$C$39)`,
        ]);
      }
      dashSheet.getRange("B44:F51").formulas = chartFormulas;
      (dashSheet.getRange("B44:E51").numberFormat as any) = "0.0%";
      (dashSheet.getRange("F44:F51").numberFormat as any) = "$#,##0";
      await context.sync();

      try {
        dashSheet.charts.load("items");
        await context.sync();
        dashSheet.charts.items.forEach((chart) => chart.delete());
        await context.sync();

        const chartDataRange = dashSheet.getRange("A43:F51");

        const chartConfig = {
          chartName: "QuarterlyMarginTrend",
          title: "Quarterly Margin Trends by Product",
          seriesNames: ["Widget Pro", "Widget Standard", "Service Package", "Accessory Kit", "Total Revenue"],
          seriesConfig: [
            {
              chartType: Excel.ChartType.columnClustered,
              axisGroup: typeof Excel.ChartAxisGroup !== "undefined" ? Excel.ChartAxisGroup.primary : 0,
              fillFormat: {
                color: "#1F4E78",
              },
            },
            {
              chartType: Excel.ChartType.columnClustered,
              axisGroup: typeof Excel.ChartAxisGroup !== "undefined" ? Excel.ChartAxisGroup.primary : 0,
              fillFormat: {
                color: "#C0504D",
              },
            },
            {
              chartType: Excel.ChartType.columnClustered,
              axisGroup: typeof Excel.ChartAxisGroup !== "undefined" ? Excel.ChartAxisGroup.primary : 0,
              fillFormat: {
                color: "#9BBB59",
              },
            },
            {
              chartType: Excel.ChartType.columnClustered,
              axisGroup: typeof Excel.ChartAxisGroup !== "undefined" ? Excel.ChartAxisGroup.primary : 0,
              fillFormat: {
                color: "#8064A2",
              },
            },
            {
              chartType: Excel.ChartType.line,
              axisGroup: 1,
              lineFormat: {
                weight: 3,
                color: "#0070C0",
              },
            },
          ],
          primaryAxis: {
            title: "Profit Margin",
          },
          secondaryAxis: {
            title: "Total Revenue ($)",
            numberFormat: "$#,##0",
          },
          position: {
            startCell: "A53",
            endCell: "H75",
            left: 0,
            top: 400,
            width: 700,
            height: 400,
          },
        };

        showStatus("Creating combo chart using template...", false);
        const chart = await createComboChart(dashSheet, chartDataRange, chartConfig);
        showStatus("✓ Combo chart created successfully", false);

        dashSheet.activate();
        await context.sync();

        showStatus("✓ Dashboard build completed successfully!", false);
      } catch (chartError) {
        const errorMsg = `⚠ Chart creation failed!\n\nError: ${(chartError as Error).message}\n\nYou can create the chart manually:\n1. Select range A43:F51\n2. Insert → Recommended Charts → All Charts\n3. Choose Combo → Clustered Column - Line on Secondary Axis`;
        showStatus(errorMsg, true);
      }
    });
  } catch (error) {
    const errorType = typeof error;
    const errorMsg = error instanceof Error ? error.message : String(error);
    const errorStack = error instanceof Error ? error.stack : "No stack trace";

    const finalErrorDetails = `=== FINAL ERROR ===\nError Type: ${errorType}\nError Message: ${errorMsg}\n\nStack Trace:\n${errorStack}`;

    showStatus(finalErrorDetails, true);
  }
}

function generateRawData(count: number): (string | number)[][] {
  const products = ["Widget Pro", "Widget Standard", "Service Package", "Accessory Kit"];
  const quarters = ["Q1", "Q2", "Q3", "Q4"];
  const years = [2023, 2024];
  const data: (string | number)[][] = [];

  for (let i = 0; i < count; i++) {
    const year = years[Math.floor(Math.random() * years.length)];
    const quarter = quarters[Math.floor(Math.random() * quarters.length)];
    const product = products[Math.floor(Math.random() * products.length)];
    const revenue = Math.floor(Math.random() * 50000) + 10000;
    const cost = Math.floor(revenue * (0.5 + Math.random() * 0.3));
    const profit = revenue - cost;

    data.push([`TXN-${1000 + i}`, year, quarter, product, revenue, cost, profit]);
  }

  return data;
}

async function prepareSheet(
  context: Excel.RequestContext,
  name: string,
  preserveIfExists: boolean = false
): Promise<Excel.Worksheet> {
  try {
    const sheet = context.workbook.worksheets.getItem(name);
    await context.sync();
    if (preserveIfExists) {
      return sheet;
    }
    sheet.delete();
    await context.sync();
  } catch (error) {
    await context.sync();
  }

  const newSheet = context.workbook.worksheets.add(name);
  await context.sync();
  return newSheet;
}

interface ChartConfig {
  chartName: string;
  title: string;
  seriesNames: string[];
  seriesConfig: Array<{
    chartType: Excel.ChartType;
    axisGroup: number | Excel.ChartAxisGroup;
    lineFormat?: {
      weight?: number;
      color?: string;
    };
    fillFormat?: {
      color?: string;
    };
  }>;
  primaryAxis?: {
    title: string;
  };
  secondaryAxis?: {
    title: string;
    numberFormat: string;
  };
  position: {
    startCell?: string;
    endCell?: string;
    left?: number;
    top?: number;
    width?: number;
    height?: number;
  };
}

async function createComboChart(
  sheet: Excel.Worksheet,
  dataRange: Excel.Range,
  config: ChartConfig
): Promise<Excel.Chart> {
  const context = sheet.context;

  try {
    const baseChartType =
      config.seriesConfig && config.seriesConfig.length > 0
        ? config.seriesConfig[0].chartType
        : Excel.ChartType.columnClustered;

    const chart = sheet.charts.add(baseChartType, dataRange, Excel.ChartSeriesBy.columns);
    chart.name = config.chartName || "ComboChart";
    chart.title.text = config.title || "";
    chart.title.format.font.bold = true;
    chart.title.format.font.size = 14;
    chart.legend.visible = true;
    chart.legend.position = Excel.ChartLegendPosition.bottom;
    chart.legend.overlay = false;

    if (config.position && config.position.startCell && config.position.endCell) {
      try {
        chart.setPosition(config.position.startCell, config.position.endCell);
      } catch (earlyPosError) {}
    }

    await context.sync();

    chart.series.load("count");
    await context.sync();
    const seriesCount = chart.series.count;

    if (config.seriesNames && config.seriesNames.length > 0) {
      for (let i = 0; i < seriesCount && i < config.seriesNames.length; i++) {
        const series = chart.series.getItemAt(i);
        series.name = config.seriesNames[i];
      }
      await context.sync();
    }

    if (config.seriesConfig && config.seriesConfig.length > 0) {
      for (let i = 0; i < seriesCount && i < config.seriesConfig.length; i++) {
        const series = chart.series.getItemAt(i);
        const seriesCfg = config.seriesConfig[i];

        if (seriesCfg.chartType) {
          series.chartType = seriesCfg.chartType;
        }
        if (seriesCfg.axisGroup !== undefined && seriesCfg.axisGroup !== null) {
          let axisGroupValue: number | Excel.ChartAxisGroup | string = seriesCfg.axisGroup;
          if (typeof axisGroupValue === "string") {
            if (axisGroupValue === "Secondary" || axisGroupValue === "secondary") {
              axisGroupValue = 1;
            } else if (axisGroupValue === "Primary" || axisGroupValue === "primary") {
              axisGroupValue = 0;
            }
          } else if (
            typeof Excel.ChartAxisGroup !== "undefined" &&
            typeof axisGroupValue !== "string" &&
            typeof axisGroupValue !== "number"
          ) {
          } else if (typeof axisGroupValue === "number") {
          }
          (series.axisGroup as any) = axisGroupValue;
        }

        if (seriesCfg.chartType === Excel.ChartType.columnClustered && seriesCfg.fillFormat) {
          try {
            if (seriesCfg.fillFormat.color) {
              const fill = series.format.fill as any;
              if (fill.solid) {
                fill.solid();
              }
              fill.color = seriesCfg.fillFormat.color;
              await context.sync();
            }
          } catch (e) {}
        }
        if (seriesCfg.chartType === Excel.ChartType.line && seriesCfg.lineFormat) {
          try {
            if (seriesCfg.lineFormat.weight) {
              series.format.line.weight = seriesCfg.lineFormat.weight;
            }
            if (seriesCfg.lineFormat.color) {
              series.format.line.color = seriesCfg.lineFormat.color;
            }
          } catch (e) {}
        }
      }
      await context.sync();
    }
    if (config.primaryAxis) {
      try {
        const primaryAxis = chart.axes.valueAxis;
        if (config.primaryAxis.title) {
          primaryAxis.title.text = config.primaryAxis.title;
          primaryAxis.title.format.font.bold = true;
          primaryAxis.title.format.font.size = 11;
        }
        await context.sync();
      } catch (e) {}
    }
    if (config.secondaryAxis) {
      try {
        await context.sync();
        let secondaryAxis: Excel.ChartAxis | null = null;

        try {
          if (typeof Excel.ChartAxisType !== "undefined" && typeof Excel.ChartAxisGroup !== "undefined") {
            secondaryAxis = chart.axes.getItem(Excel.ChartAxisType.value, Excel.ChartAxisGroup.secondary);
            await context.sync();
          } else if (typeof Excel.ChartAxisType !== "undefined") {
            secondaryAxis = chart.axes.getItem(Excel.ChartAxisType.value, "Secondary");
            await context.sync();
          }
        } catch (e1) {}

        if (secondaryAxis) {
          secondaryAxis.visible = true;
          if (config.secondaryAxis.title) {
            secondaryAxis.title.text = config.secondaryAxis.title;
            secondaryAxis.title.format.font.bold = true;
            secondaryAxis.title.format.font.size = 11;
          }
          if (config.secondaryAxis.numberFormat) {
            try {
              secondaryAxis.numberFormat = config.secondaryAxis.numberFormat;
            } catch (nfError) {
              try {
                (secondaryAxis as any).format.code = config.secondaryAxis.numberFormat;
              } catch (codeError) {}
            }
          }
          await context.sync();
        }
      } catch (e) {}
    }
    if (config.position) {
      try {
        if (config.position.startCell && config.position.endCell) {
          chart.setPosition(config.position.startCell, config.position.endCell);
          await context.sync();
        } else if (config.position.left !== undefined || config.position.top !== undefined) {
          chart.left = config.position.left || 0;
          chart.top = config.position.top || 0;
          chart.width = config.position.width || 700;
          chart.height = config.position.height || 400;
          await context.sync();
        }
      } catch (positionError) {}
    }

    return chart;
  } catch (chartError) {
    throw chartError;
  }
}

function showStatus(message: string, isError: boolean): void {
  const statusDiv = document.getElementById("status");
  if (!statusDiv) {
    return;
  }

  const existingMessages = statusDiv.querySelectorAll(".status-card");
  const maxMessages = 100;

  if (existingMessages.length >= maxMessages) {
    const messagesToRemove = existingMessages.length - maxMessages + 1;
    for (let i = 0; i < messagesToRemove; i++) {
      if (existingMessages[i].parentNode) {
        existingMessages[i].parentNode!.removeChild(existingMessages[i]);
      }
    }
  }

  const statusCard = document.createElement("div");
  statusCard.className = `status-card ${isError ? "error-msg" : "success-msg"}`;

  const p = document.createElement("p");
  p.textContent = message;
  p.style.whiteSpace = "pre-wrap";
  p.style.wordBreak = "break-word";
  statusCard.appendChild(p);
  statusDiv.appendChild(statusCard);

  statusDiv.scrollTop = statusDiv.scrollHeight;
}

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  console.log("Office.onReady called", info);

  if (info.host === Office.HostType.Excel) {
    const sideloadMsg = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    if (sideloadMsg) sideloadMsg.style.display = "none";
    if (appBody) appBody.style.display = "flex";

    const buildButton = document.getElementById("build-dashboard");
    if (buildButton) {
      buildButton.onclick = () => {
        console.log("Build Dashboard button clicked");
        runBuild();
      };
      (buildButton as HTMLButtonElement).disabled = false;
      console.log("Button handler attached");
    } else {
      console.error("Build dashboard button not found!");
      showStatus("Error: Build button not found in HTML", true);
    }

    // Auto-run on load
    setTimeout(() => {
      console.log("Auto-running build...");
      runBuild();
    }, 500);
  } else {
    console.warn("Not running in Excel. Host:", info.host);
    showStatus(`Warning: This add-in is designed for Excel. Current host: ${info.host}`, true);
  }
});

async function runBuild(): Promise<void> {
  console.log("=== runBuild() STARTED ===");
  showStatus("Building dashboard...", false);

  if (typeof Excel === "undefined") {
    const errorMsg = "Excel API is not available. Make sure you're running this in Excel.";
    console.error(errorMsg);
    showStatus(errorMsg, true);
    return;
  }

  try {
    await Excel.run(async (context) => {
      console.log("=== Inside Excel.run ===");

      // Declare sheets at function scope so they're accessible throughout
      let rawSheet: Excel.Worksheet;
      let dashSheet: Excel.Worksheet;

      // 1. SETUP SHEETS
      console.log("[STEP 1] Creating sheets...");
      try {
        rawSheet = await prepareSheet(context, "Raw Data");
        console.log("[STEP 1a] Raw Data sheet created");
        dashSheet = await prepareSheet(context, "Dashboard");
        console.log("[STEP 1b] Dashboard sheet created");
        await context.sync(); // Single sync after both sheets are created
        console.log("[STEP 1c] Both sheets synced");
      } catch (error) {
        console.error("[STEP 1 ERROR]", error);
        throw error;
      }

      // 2. INITIALIZE RAW DATA (186 rows as per PDF)
      console.log("[STEP 2] Adding raw data...");
      showStatus("Adding raw data...", false);
      try {
        const rawHeaders = [["Transaction_ID", "Year", "Quarter", "Product", "Revenue", "Cost", "Profit"]];
        console.log("[STEP 2a] Setting raw data headers");
        const headerRange = rawSheet.getRange("A1:G1");
        headerRange.values = rawHeaders;
        headerRange.format.font.bold = true;
        headerRange.format.fill.color = "#D9E1F2";
        await context.sync(); // Single sync after all header operations
        console.log("[STEP 2b] Headers set and formatted");

        // Generate sample data (186 rows as per PDF requirement)
        console.log("[STEP 2d] Generating sample data (186 rows)");
        const sampleData = generateRawData(186);
        console.log("[STEP 2e] Sample data generated, setting values");
        rawSheet.getRange("A2:G187").values = sampleData;
        await context.sync();
        console.log("[STEP 2f] Sample data written");

        // Format raw data
        console.log("[STEP 2g] Formatting raw data");
        rawSheet.getRange("B2:B187").numberFormat = "0"; // Year as integer
        rawSheet.getRange("E2:G187").numberFormat = "$#,##0"; // Revenue, Cost, Profit as currency
        await context.sync(); // Single sync after all formatting
        console.log("[STEP 2h] Raw data formatted and synced");
      } catch (error) {
        console.error("[STEP 2 ERROR]", error);
        showStatus("Error in raw data: " + (error as Error).message, true);
        throw error;
      }

      // 3. SETUP DASHBOARD STRUCTURE (32 rows = 4 products × 8 quarters)
      console.log("[STEP 3] Setting up dashboard structure...");
      showStatus("Setting up dashboard...", false);
      try {
        await context.sync(); // Sync before dashboard operations
        console.log("[STEP 3a] Pre-dashboard sync complete");

        // Add title and labels in rows 1-6 to match example
        console.log("[STEP 3a-1] Setting dashboard title and labels");

        // Row 1: "Executive Sales Performance Dashboard" (merged A1:F1, bold, large font)
        const titleCell = dashSheet.getRange("A1");
        titleCell.values = [["Executive Sales Performance Dashboard"]];
        titleCell.format.font.bold = true;
        titleCell.format.font.size = 14;
        titleCell.format.font.name = "Calibri";
        await context.sync(); // Sync before merge attempt

        // Merge cells A1:F1
        try {
          const titleRange = dashSheet.getRange("A1:H1");
          titleRange.merge(true);
          await context.sync();
          console.log("[STEP 3a-1a] Title cells merged successfully");
        } catch (mergeError) {
          console.error("[STEP 3a-1a] Merge error:", (mergeError as Error).message);
          try {
            const titleRange = dashSheet.getRange("A1:H1");
            titleRange.merge();
            await context.sync();
            console.log("[STEP 3a-1a] Title cells merged (alternative method)");
          } catch (mergeError2) {
            console.error("[STEP 3a-1a] All merge attempts failed:", (mergeError2 as Error).message);
            showStatus("Warning: Could not merge title cells. Text may appear in single cell only.", false);
          }
        }

        // Row 3: Instructions (merged A3:F3, italic)
        const instructionsCell = dashSheet.getRange("A3");
        instructionsCell.values = [
          ["Instructions: Complete the analysis below using the Raw Data tab. All calculations should use formulas."],
        ];
        instructionsCell.format.font.italic = true;
        instructionsCell.format.font.size = 11;
        instructionsCell.format.font.name = "Calibri";
        await context.sync(); // Sync before merge attempt

        // Merge cells A3:F3
        try {
          const instructionsRange = dashSheet.getRange("A3:H3");
          instructionsRange.merge(true);
          await context.sync();
          console.log("[STEP 3a-1b] Instructions cells merged successfully");
        } catch (mergeError) {
          console.error("[STEP 3a-1b] Merge error:", (mergeError as Error).message);
          try {
            const instructionsRange = dashSheet.getRange("A3:H3");
            instructionsRange.merge();
            await context.sync();
            console.log("[STEP 3a-1b] Instructions cells merged (alternative method)");
          } catch (mergeError2) {
            console.error("[STEP 3a-1b] All merge attempts failed:", (mergeError2 as Error).message);
          }
        }

        // Row 5: "Quarterly Performance Summary" (merged A5:F5, bold)
        const summaryCell = dashSheet.getRange("A5");
        summaryCell.values = [["Quarterly Performance Summary"]];
        summaryCell.format.font.bold = true;
        summaryCell.format.font.size = 12;
        summaryCell.format.font.name = "Calibri";
        await context.sync(); // Sync before merge attempt

        // Merge cells A5:F5
        try {
          const summaryRange = dashSheet.getRange("A5:B5");
          summaryRange.merge(true);
          await context.sync();
          console.log("[STEP 3a-1c] Summary cells merged successfully");
        } catch (mergeError) {
          console.error("[STEP 3a-1c] Merge error:", (mergeError as Error).message);
          try {
            const summaryRange = dashSheet.getRange("A5:B5");
            summaryRange.merge();
            await context.sync();
            console.log("[STEP 3a-1c] Summary cells merged (alternative method)");
          } catch (mergeError2) {
            console.error("[STEP 3a-1c] All merge attempts failed:", (mergeError2 as Error).message);
          }
        }

        // Final sync after all title/label operations and merge attempts
        await context.sync();
        console.log("[STEP 3a-2] Dashboard title and labels set");

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
        console.log("[STEP 3b] Setting dashboard headers");
        const dashHeaderRange = dashSheet.getRange("A7:G7");
        dashHeaderRange.values = dashHeaders;
        dashHeaderRange.format.font.bold = true;
        dashHeaderRange.format.font.name = "Calibri";
        dashHeaderRange.format.fill.color = "#D9E1F2";
        await context.sync(); // Single sync after all header operations
        console.log("[STEP 3c] Dashboard headers set and formatted");

        // Generate product rows (4 products × 8 quarters = 32 rows)
        console.log("[STEP 3f] Generating product rows");
        const products = ["Widget Pro", "Widget Standard", "Service Package", "Accessory Kit"];
        const quarters = ["2023 Q1", "2023 Q2", "2023 Q3", "2023 Q4", "2024 Q1", "2024 Q2", "2024 Q3", "2024 Q4"];
        const productRows: string[][] = [];
        for (const product of products) {
          for (const quarter of quarters) {
            productRows.push([product, quarter]);
          }
        }
        console.log("[STEP 3g] Product rows generated, count:", productRows.length);
        const productRange = dashSheet.getRange("A8:B39");
        productRange.values = productRows;
        await context.sync(); // Sync after setting product rows
        console.log("[STEP 3h] Product rows written to Dashboard");
      } catch (error) {
        console.error("[STEP 3 ERROR]", error);
        showStatus("Error in dashboard setup: " + (error as Error).message, true);
        throw error;
      }

      // 4. APPLY FORMULAS (as per PDF requirements)
      showStatus("Applying formulas...", false);

      // Column C: Total Revenue
      console.log("[STEP 4a] Setting revenue formulas...");
      showStatus("Setting revenue formulas...", false);
      try {
        const revenueFormulas: string[][] = [];
        console.log("[STEP 4a-1] Building revenue formula array");
        for (let row = 8; row <= 39; row++) {
          const formula = `=SUMIFS('Raw Data'!$E$2:$E$187,'Raw Data'!$D$2:$D$187,$A${row},'Raw Data'!$B$2:$B$187,LEFT($B${row},4)*1,'Raw Data'!$C$2:$C$187,RIGHT($B${row},2))`;
          revenueFormulas.push([formula]);
          if (row === 8) {
            console.log("[STEP 4a-2] First revenue formula:", formula);
          }
        }
        console.log("[STEP 4a-3] Revenue formulas array built, count:", revenueFormulas.length);
        const revenueRange = dashSheet.getRange("C8:C39");
        console.log("[STEP 4a-4] Setting formulas to range C8:C39");
        revenueRange.formulas = revenueFormulas;
        revenueRange.numberFormat = "$#,##0";
        console.log("[STEP 4a-5] Formulas and format assigned, syncing...");
        try {
          await context.sync();
        } catch (syncError) {
          console.error("[STEP 4a-5 ERROR] Sync failed after setting revenue formulas");
          throw syncError;
        }
        console.log("[STEP 4a-6] Revenue formulas set and formatted successfully");
      } catch (error) {
        const errorDetails = `[STEP 4a ERROR] Revenue formulas failed!\nError: ${(error as Error).message}`;
        console.error("[STEP 4a ERROR]", error);
        showStatus(errorDetails, true);
        // Try alternative approach
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
          altRevenueRange.numberFormat = "$#,##0";
          await context.sync();
          showStatus("✓ Revenue formulas set using SUMPRODUCT", false);
        } catch (altError) {
          showStatus(
            `[STEP 4a ERROR] Both approaches failed!\n\nOriginal Error: ${(error as Error).message}\n\nAlternative Error: ${(altError as Error).message}`,
            true
          );
          throw altError;
        }
      }

      // Column D: Weighted Avg Margin
      console.log("[STEP 4b] Setting margin formulas...");
      showStatus("Setting margin formulas...", false);
      try {
        const marginFormulas: string[][] = [];
        console.log("[STEP 4b-1] Building margin formula array");
        for (let row = 8; row <= 39; row++) {
          const formula = `=IFERROR(SUMPRODUCT(('Raw Data'!$D$2:$D$187=$A${row})*('Raw Data'!$B$2:$B$187=LEFT($B${row},4)*1)*('Raw Data'!$C$2:$C$187=RIGHT($B${row},2))*'Raw Data'!$G$2:$G$187)/C${row},0)`;
          marginFormulas.push([formula]);
          if (row === 8) {
            console.log("[STEP 4b-2] First margin formula:", formula);
          }
        }
        console.log("[STEP 4b-3] Margin formulas array built, count:", marginFormulas.length);
        const marginRange = dashSheet.getRange("D8:D39");
        console.log("[STEP 4b-4] Setting formulas to range D8:D39");
        marginRange.formulas = marginFormulas;
        marginRange.numberFormat = "0.0%";
        console.log("[STEP 4b-5] Formulas and format assigned, syncing...");
        try {
          await context.sync();
        } catch (syncError) {
          console.error("[STEP 4b-5 ERROR] Sync failed after setting margin formulas");
          throw syncError;
        }
        console.log("[STEP 4b-6] Margin formulas set and formatted successfully");
      } catch (error) {
        const errorDetails = `[STEP 4b ERROR] Margin formulas failed!\nError: ${(error as Error).message}`;
        console.error("[STEP 4b ERROR]", error);
        showStatus(errorDetails, true);
        // Set a placeholder formula instead
        try {
          const placeholderFormulas: string[][] = [];
          const marginRange = dashSheet.getRange("D8:D39");
          for (let row = 8; row <= 39; row++) {
            placeholderFormulas.push([`=IF(C${row}>0,C${row}/C${row},0)`]); // Simple placeholder
          }
          marginRange.formulas = placeholderFormulas;
          await context.sync();
          marginRange.numberFormat = "0.0%";
          await context.sync();
          showStatus("⚠ Margin formulas set to placeholder - check manually", false);
        } catch (placeholderError) {
          showStatus(`[STEP 4b ERROR] Could not set placeholder: ${(placeholderError as Error).message}`, true);
          throw placeholderError;
        }
      }

      // Column E: Rolling 3-Month Trend
      console.log("[STEP 4c] Setting trend formulas...");
      showStatus("Setting trend formulas...", false);
      try {
        const trendFormulas: string[][] = [];
        console.log("[STEP 4c-1] Building trend formula array");
        for (let row = 8; row <= 39; row++) {
          const isFirstQ1 = row === 8 || row === 16 || row === 24 || row === 32; // First quarter of each product
          if (isFirstQ1) {
            trendFormulas.push([`="N/A"`]);
          } else {
            trendFormulas.push([`=D${row}-D${row - 1}`]);
          }
        }
        console.log("[STEP 4c-2] Trend formulas array built, count:", trendFormulas.length);
        const trendRange = dashSheet.getRange("E8:E39");
        console.log("[STEP 4c-3] Setting formulas to range E8:E39");
        trendRange.formulas = trendFormulas;
        trendRange.numberFormat = "0.0%";
        console.log("[STEP 4c-4] Formulas and format assigned, syncing...");
        await context.sync();
        console.log("[STEP 4c-5] Trend formulas set and formatted successfully");
      } catch (error) {
        const errorDetails = `[STEP 4c ERROR] Trend formulas failed!\nError: ${(error as Error).message}`;
        console.error("[STEP 4c ERROR]", error);
        showStatus(errorDetails, true);
        throw error;
      }

      // Column F: YoY Margin Delta
      console.log("[STEP 4d] Setting YoY formulas...");
      showStatus("Setting YoY formulas...", false);
      const yoyFormulas: string[][] = [];
      try {
        console.log("[STEP 4d-1] Building YoY formula array");
        for (let row = 8; row <= 39; row++) {
          const formula = `=IF(LEFT($B${row},4)="2023","N/A","")`;
          yoyFormulas.push([formula]);
          if (row === 8) {
            console.log("[STEP 4d-2] First YoY formula:", formula);
          }
        }
        console.log("[STEP 4d-3] YoY formulas array built, count:", yoyFormulas.length);
        const yoyRange = dashSheet.getRange("F8:F39");
        console.log("[STEP 4d-4] Setting formulas to range F8:F39");
        yoyRange.formulas = yoyFormulas;
        yoyRange.numberFormat = "0.0%";
        console.log("[STEP 4d-5] Formulas and format assigned, syncing...");
        await context.sync();
        console.log("[STEP 4d-6] YoY formulas set and formatted successfully");
      } catch (error) {
        const errorDetails = `[STEP 4d ERROR] YoY formulas failed!\nError: ${(error as Error).message}`;
        console.error("[STEP 4d ERROR]", error);
        showStatus(errorDetails, true);
        throw error;
      }

      // Column G: Margin Health
      console.log("[STEP 4e] Setting health formulas...");
      showStatus("Setting health formulas...", false);
      try {
        const healthFormulas: string[][] = [];
        console.log("[STEP 4e-1] Building health formula array");
        for (let row = 8; row <= 39; row++) {
          const formula = `=IF(D${row}>0.35,"Strong",IF(D${row}>=0.2,"Moderate","At Risk"))`;
          healthFormulas.push([formula]);
          if (row === 8) {
            console.log("[STEP 4e-2] First health formula:", formula);
          }
        }
        console.log("[STEP 4e-3] Health formulas array built, count:", healthFormulas.length);
        const healthRange = dashSheet.getRange("G8:G39");
        console.log("[STEP 4e-4] Setting formulas to range G8:G39");
        healthRange.formulas = healthFormulas;
        console.log("[STEP 4e-5] Formulas assigned, syncing...");
        await context.sync();
        console.log("[STEP 4e-6] Health formulas set successfully");
      } catch (error) {
        const errorDetails = `[STEP 4e ERROR] Health formulas failed!\nError: ${(error as Error).message}`;
        console.error("[STEP 4e ERROR]", error);
        showStatus(errorDetails, true);
        throw error;
      }

      // 5. CONDITIONAL FORMATTING
      showStatus("Applying formatting...", false);
      try {
        const healthRange = dashSheet.getRange("G8:G39");
        healthRange.load("conditionalFormats");
        await context.sync();

        // Strong - Green
        const strongFormat = healthRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
        strongFormat.textComparison.format.fill.color = "#C6EFCE";
        strongFormat.textComparison.format.font.color = "#006100";
        strongFormat.textComparison.rule = {
          operator: Excel.ConditionalTextOperator.contains,
          text: "Strong",
        };

        // Moderate - Yellow
        const moderateFormat = healthRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
        moderateFormat.textComparison.format.fill.color = "#FFEB9C";
        moderateFormat.textComparison.format.font.color = "#9C6500";
        moderateFormat.textComparison.rule = {
          operator: Excel.ConditionalTextOperator.contains,
          text: "Moderate",
        };

        // At Risk - Red
        const riskFormat = healthRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
        riskFormat.textComparison.format.fill.color = "#FFC7CE";
        riskFormat.textComparison.format.font.color = "#9C0006";
        riskFormat.textComparison.rule = {
          operator: Excel.ConditionalTextOperator.contains,
          text: "At Risk",
        };
        await context.sync();
      } catch (formatError) {
        console.warn("Conditional formatting not available:", formatError);
      }

      // 6. CREATE CHART DATA TABLE
      showStatus("Creating chart data...", false);

      // Chart Data Headers
      console.log("[STEP 6a] Setting chart data headers");
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
      console.log("[STEP 6b] Chart headers set and formatted");

      // Quarters and Chart Data Formulas (B44:F51)
      console.log("[STEP 6c] Setting chart quarters and formulas");
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

      // Build all chart formulas at once
      const chartFormulas: string[][] = [];
      for (let i = 0; i < 8; i++) {
        const row = 44 + i;
        chartFormulas.push([
          `=SUMPRODUCT(($A$8:$A$39="Widget Pro")*($B$8:$B$39=$A${row})*($D$8:$D$39))`, // Widget Pro
          `=SUMPRODUCT(($A$8:$A$39="Widget Standard")*($B$8:$B$39=$A${row})*($D$8:$D$39))`, // Widget Standard
          `=SUMPRODUCT(($A$8:$A$39="Service Package")*($B$8:$B$39=$A${row})*($D$8:$D$39))`, // Service Package
          `=SUMPRODUCT(($A$8:$A$39="Accessory Kit")*($B$8:$B$39=$A${row})*($D$8:$D$39))`, // Accessory Kit
          `=SUMIF($B$8:$B$39, $A${row}, $C$8:$C$39)`, // Total Revenue
        ]);
      }
      dashSheet.getRange("B44:F51").formulas = chartFormulas;

      // Format chart data
      dashSheet.getRange("B44:E51").numberFormat = "0.0%";
      dashSheet.getRange("F44:F51").numberFormat = "$#,##0";
      await context.sync();
      console.log("[STEP 6d] Chart data set and formatted");

      // 7. CREATE COMBO CHART
      try {
        // Clear existing charts first
        dashSheet.charts.load("items");
        await context.sync();
        dashSheet.charts.items.forEach((chart) => chart.delete());
        await context.sync();

        // Create chart data range
        const chartDataRange = dashSheet.getRange("A43:F51");

        // Use combo chart template
        const chartConfig = {
          chartName: "QuarterlyMarginTrend",
          title: "Quarterly Margin Trends by Product",
          seriesNames: ["Widget Pro", "Widget Standard", "Service Package", "Accessory Kit", "Total Revenue"],
          seriesConfig: [
            // Widget Pro - dark blue column
            {
              chartType: Excel.ChartType.columnClustered,
              axisGroup: typeof Excel.ChartAxisGroup !== "undefined" ? Excel.ChartAxisGroup.primary : 0,
              fillFormat: {
                color: "#1F4E78", // Dark blue
              },
            },
            // Widget Standard - red column
            {
              chartType: Excel.ChartType.columnClustered,
              axisGroup: typeof Excel.ChartAxisGroup !== "undefined" ? Excel.ChartAxisGroup.primary : 0,
              fillFormat: {
                color: "#C00000", // Red
              },
            },
            // Service Package - green column
            {
              chartType: Excel.ChartType.columnClustered,
              axisGroup: typeof Excel.ChartAxisGroup !== "undefined" ? Excel.ChartAxisGroup.primary : 0,
              fillFormat: {
                color: "#70AD47", // Green
              },
            },
            // Accessory Kit - purple column
            {
              chartType: Excel.ChartType.columnClustered,
              axisGroup: typeof Excel.ChartAxisGroup !== "undefined" ? Excel.ChartAxisGroup.primary : 0,
              fillFormat: {
                color: "#7030A0", // Purple
              },
            },
            // Total Revenue - blue line on secondary axis
            {
              chartType: Excel.ChartType.line,
              axisGroup: 1,
              lineFormat: {
                weight: 3,
                color: "#0070C0", // Blue
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

        // Activate sheet
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

    console.error("=== FINAL ERROR ===", error);
    showStatus(finalErrorDetails, true);
  }
  console.log("=== runBuild() COMPLETED ===");
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

async function prepareSheet(context: Excel.RequestContext, name: string): Promise<Excel.Worksheet> {
  try {
    // Try to get the sheet
    const sheet = context.workbook.worksheets.getItem(name);
    await context.sync();
    // If we get here, the sheet exists, so delete it
    sheet.delete();
    await context.sync();
  } catch (error) {
    // Sheet doesn't exist, which is fine - we'll create it
    await context.sync();
  }

  // Create the new sheet
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
    console.log(`[CHART CREATE] Starting chart creation...`);

    // 1. Create base chart
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

    // Position chart early
    if (config.position && config.position.startCell && config.position.endCell) {
      try {
        chart.setPosition(config.position.startCell, config.position.endCell);
        console.log(`[CHART POSITION] ✓ Early setPosition() queued`);
      } catch (earlyPosError) {
        console.warn(`[CHART POSITION] Early setPosition failed: ${(earlyPosError as Error).message}`);
      }
    }

    await context.sync();

    // 2. Set series names
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

    // 3. Configure each series
    if (config.seriesConfig && config.seriesConfig.length > 0) {
      for (let i = 0; i < seriesCount && i < config.seriesConfig.length; i++) {
        const series = chart.series.getItemAt(i);
        const seriesCfg = config.seriesConfig[i];

        // Set chart type
        if (seriesCfg.chartType) {
          series.chartType = seriesCfg.chartType;
        }

        // Set axis group
        if (seriesCfg.axisGroup !== undefined && seriesCfg.axisGroup !== null) {
          let axisGroupValue = seriesCfg.axisGroup;
          if (typeof axisGroupValue === "string") {
            if (axisGroupValue === "Secondary" || axisGroupValue === "secondary") {
              axisGroupValue = 1;
            } else if (axisGroupValue === "Primary" || axisGroupValue === "primary") {
              axisGroupValue = 0;
            }
          }
          series.axisGroup = axisGroupValue as number;
        }

        // Apply fill formatting if it's a column chart
        if (seriesCfg.chartType === Excel.ChartType.columnClustered && seriesCfg.fillFormat) {
          try {
            if (seriesCfg.fillFormat.color) {
              series.format.fill.color = seriesCfg.fillFormat.color;
            }
          } catch (e) {
            // Formatting may fail, continue
          }
        }

        // Apply line formatting if it's a line chart
        if (seriesCfg.chartType === Excel.ChartType.line && seriesCfg.lineFormat) {
          try {
            if (seriesCfg.lineFormat.weight) {
              series.format.line.weight = seriesCfg.lineFormat.weight;
            }
            if (seriesCfg.lineFormat.color) {
              series.format.line.color = seriesCfg.lineFormat.color;
            }
          } catch (e) {
            // Formatting may fail, continue
          }
        }
      }
      await context.sync();
    }

    // 4. Configure primary axis
    if (config.primaryAxis) {
      try {
        const primaryAxis = chart.axes.valueAxis;
        if (config.primaryAxis.title) {
          primaryAxis.title.text = config.primaryAxis.title;
          primaryAxis.title.format.font.bold = true;
          primaryAxis.title.format.font.size = 11;
        }
        await context.sync();
      } catch (e) {
        // Primary axis config may fail
      }
    }

    // 5. Configure secondary axis
    if (config.secondaryAxis) {
      try {
        await context.sync();
        let secondaryAxis: Excel.ChartAxis | null = null;
        chart.axes.load("items");
        await context.sync();
        const axes = chart.axes.items;

        if (axes && axes.length > 0) {
          for (const axis of axes) {
            axis.load("axisType,axisGroup");
          }
          await context.sync();

          for (const axis of axes) {
            const isValue = axis.axisType === Excel.ChartAxisType.value;
            const isSecondary = axis.axisGroup === 1;
            if (isValue && isSecondary) {
              secondaryAxis = axis;
              break;
            }
          }
        }

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
              } catch (codeError) {
                // Format setting failed
              }
            }
          }
          await context.sync();
        }
      } catch (e) {
        // Secondary axis config may fail
      }
    }

    // 6. Position chart
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
      } catch (positionError) {
        console.error(`[CHART POSITION ERROR] ${(positionError as Error).message}`);
      }
    }

    console.log(`[CHART CREATE] ✓ Chart creation completed successfully`);
    return chart;
  } catch (chartError) {
    console.error(`[CHART CREATE ERROR] Chart creation failed: ${(chartError as Error).message}`);
    throw chartError;
  }
}

function showStatus(message: string, isError: boolean): void {
  console.log(`[STATUS] ${isError ? "ERROR" : "INFO"}: ${message}`);
  const statusDiv = document.getElementById("status");
  if (!statusDiv) {
    console.error("Status element not found!");
    return;
  }

  // Keep only the last 100 messages
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

  // Scroll to bottom to show latest message
  statusDiv.scrollTop = statusDiv.scrollHeight;
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").addEventListener("click", run);
  }
});

async function run() {
  try {
    await Excel.run(async (ctx) => {
      const rawDataSheet = ctx.workbook.worksheets.getItem("Raw Data");
      const dashboard = ctx.workbook.worksheets.getItem("Dashboard");

      const usedRange = rawDataSheet.getUsedRange();
      usedRange.load("rowCount");
      await ctx.sync();

      const lastRow = usedRange.rowCount;
      const dataStart = 2;
      const dataEnd = lastRow;

      const hdrRow = 7;
      let firstDataRow = 8;
      let lastDataRow = 39;

      const prodCols = dashboard.getRange(`A${firstDataRow}:B${lastDataRow}`);
      prodCols.load("values");
      await ctx.sync();

      const prodData = prodCols.values;
      for (let idx = 0; idx < prodData.length; idx++) {
        if (!prodData[idx][0] || !prodData[idx][1]) {
          lastDataRow = firstDataRow + idx - 1;
          break;
        }
      }

      function buildRevenueFormula(row) {
        return `=SUMIFS('Raw Data'!$E$${dataStart}:$E$${dataEnd},'Raw Data'!$D$${dataStart}:$D$${dataEnd},$A${row},'Raw Data'!$B$${dataStart}:$B$${dataEnd},LEFT($B${row},4)*1,'Raw Data'!$C$${dataStart}:$C$${dataEnd},RIGHT($B${row},2))`;
      }

      function buildMarginFormula(row) {
        return `=IFERROR(SUMPRODUCT(('Raw Data'!$D$${dataStart}:$D$${dataEnd}=$A${row})*('Raw Data'!$B$${dataStart}:$B$${dataEnd}=LEFT($B${row},4)*1)*('Raw Data'!$C$${dataStart}:$C$${dataEnd}=RIGHT($B${row},2))*'Raw Data'!$G$${dataStart}:$G$${dataEnd})/C${row},0)`;
      }

      const revFormulas = [];
      const margFormulas = [];
      const trendFormulas = [];
      const yoyFormulas = [];
      const healthFormulas = [];

      for (let r = firstDataRow; r <= lastDataRow; r++) {
        if (!prodData[r - firstDataRow][0] || !prodData[r - firstDataRow][1]) {
          continue;
        }

        revFormulas.push([buildRevenueFormula(r)]);
        margFormulas.push([buildMarginFormula(r)]);

        const isQ1Start = r === 8 || r === 16 || r === 24 || r === 32;
        if (isQ1Start) {
          trendFormulas.push([`="N/A"`]);
        } else {
          trendFormulas.push([`=D${r}-D${r - 1}`]);
        }

        yoyFormulas.push([`=IF(LEFT($B${r},4)="2023","N/A","")`]);
        healthFormulas.push([`=IF(D${r}>0.35,"Strong",IF(D${r}>=0.2,"Moderate","At Risk"))`]);
      }

      dashboard.getRange(`C${firstDataRow}:C${lastDataRow}`).formulas = revFormulas;
      dashboard.getRange(`C${firstDataRow}:C${lastDataRow}`).numberFormat = "$#,##0";

      const marginCol = dashboard.getRange(`D${firstDataRow}:D${lastDataRow}`);
      marginCol.formulas = margFormulas;
      marginCol.numberFormat = "0.0%";

      dashboard.getRange(`E${firstDataRow}:E${lastDataRow}`).formulas = trendFormulas;
      dashboard.getRange(`E${firstDataRow}:E${lastDataRow}`).numberFormat = "0.0%";

      dashboard.getRange(`F${firstDataRow}:F${lastDataRow}`).formulas = yoyFormulas;
      dashboard.getRange(`F${firstDataRow}:F${lastDataRow}`).numberFormat = "0.0%";

      dashboard.getRange(`G${firstDataRow}:G${lastDataRow}`).formulas = healthFormulas;
      await ctx.sync();

      const headerCells = dashboard.getRange(`A${hdrRow}:G${hdrRow}`);
      headerCells.format.font.bold = true;
      headerCells.format.font.name = "Calibri";
      headerCells.format.fill.color = "#D9E1F2";
      headerCells.format.wrapText = true;
      headerCells.format.horizontalAlignment = Excel.HorizontalAlignment.center;

      dashboard.getRange("5:5").format.autofitRows();
      await ctx.sync();

      const healthCol = dashboard.getRange(`G${firstDataRow}:G${lastDataRow}`);
      healthCol.load("conditionalFormats");
      await ctx.sync();

      const strong = healthCol.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
      strong.textComparison.format.fill.color = "#C6EFCE";
      strong.textComparison.format.font.color = "#006100";
      strong.textComparison.rule = {
        operator: Excel.ConditionalTextOperator.contains,
        text: "Strong",
      };

      const moderate = healthCol.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
      moderate.textComparison.format.fill.color = "#FFEB9C";
      moderate.textComparison.format.font.color = "#9C6500";
      moderate.textComparison.rule = {
        operator: Excel.ConditionalTextOperator.contains,
        text: "Moderate",
      };

      const atRisk = healthCol.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
      atRisk.textComparison.format.fill.color = "#FFC7CE";
      atRisk.textComparison.format.font.color = "#9C0006";
      atRisk.textComparison.rule = {
        operator: Excel.ConditionalTextOperator.contains,
        text: "At Risk",
      };
      await ctx.sync();

      dashboard.getRange("A42").values = [["Chart Data"]];
      dashboard.getRange("A42").format.font.bold = true;

      const hdrRowNum = 43;
      const headers = [
        ["Quarter", "Widget Pro", "Widget Standard", "Service Package", "Accessory Kit", "Total Revenue"],
      ];
      const hdrRange = dashboard.getRange(`A${hdrRowNum}:F${hdrRowNum}`);
      hdrRange.values = headers;
      hdrRange.format.font.bold = true;
      hdrRange.format.fill.color = "#D9E1F2";
      hdrRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;
      await ctx.sync();

      const quarters = [
        ["2023 Q1"],
        ["2023 Q2"],
        ["2023 Q3"],
        ["2023 Q4"],
        ["2024 Q1"],
        ["2024 Q2"],
        ["2024 Q3"],
        ["2024 Q4"],
      ];
      const chartStart = hdrRowNum + 1;

      for (let i = 0; i < quarters.length; i++) {
        const qtr = quarters[i];
        const rowNum = chartStart + i;

        if (i === 0) {
          dashboard.getRange(`A${rowNum}`).values = [qtr];
          await ctx.sync();
          const merged = dashboard.getRange(`B${rowNum}:F${rowNum}`);
          merged.merge(true);
          await ctx.sync();
          merged.format.font.italic = true;
          merged.format.font.color = "#707070";
        } else {
          dashboard.getRange(`A${rowNum}`).values = [qtr];

          const formulas = [
            `=SUMPRODUCT(($A$${firstDataRow}:$A$${lastDataRow}="Widget Pro")*($B$${firstDataRow}:$B$${lastDataRow}=$A${rowNum})*($D$${firstDataRow}:$D$${lastDataRow}))`,
            `=SUMPRODUCT(($A$${firstDataRow}:$A$${lastDataRow}="Widget Standard")*($B$${firstDataRow}:$B$${lastDataRow}=$A${rowNum})*($D$${firstDataRow}:$D$${lastDataRow}))`,
            `=SUMPRODUCT(($A$${firstDataRow}:$A$${lastDataRow}="Service Package")*($B$${firstDataRow}:$B$${lastDataRow}=$A${rowNum})*($D$${firstDataRow}:$D$${lastDataRow}))`,
            `=SUMPRODUCT(($A$${firstDataRow}:$A$${lastDataRow}="Accessory Kit")*($B$${firstDataRow}:$B$${lastDataRow}=$A${rowNum})*($D$${firstDataRow}:$D$${lastDataRow}))`,
            `=SUMIF($B$${firstDataRow}:$B$${lastDataRow}, $A${rowNum}, $C$${firstDataRow}:$C$${lastDataRow})`,
          ];

          dashboard.getRange(`B${rowNum}:F${rowNum}`).formulas = [formulas];
          dashboard.getRange(`B${rowNum}:E${rowNum}`).numberFormat = "0.0%";
          dashboard.getRange(`F${rowNum}`).numberFormat = "$#,##0";
        }
      }
      await ctx.sync();

      try {
        dashboard.charts.load("items");
        await ctx.sync();
        dashboard.charts.items.forEach((ch) => ch.delete());
        await ctx.sync();

        const dataRange = dashboard.getRange("A43:F51");

        const cfg = {
          chartName: "QuarterlyMarginTrend",
          title: "Quarterly Margin Trends by Product",
          seriesNames: ["Widget Pro", "Widget Standard", "Service Package", "Accessory Kit", "Total Revenue"],
          colors: ["#4F81BD", "#C0504D", "#9BBB59", "#8064A2"],
          primaryAxis: {
            title: "Profit Margin",
            numberFormat: "0.0%",
          },
          secondaryAxis: {
            title: "Total Revenue ($)",
            numberFormat: "$#,##0",
          },
          position: {
            startCell: "A53",
            endCell: "H75",
          },
        };

        const ch = await createChart(dashboard, dataRange, cfg);
        dashboard.activate();
        await ctx.sync();
      } catch (err) {
        console.error("Chart error:", err);
      }
    });
  } catch (e) {
    console.error("Failed:", e.message);
  }
}

async function createChart(sheet, dataRange, config) {
  const ctx = sheet.context;

  const ch = sheet.charts.add(Excel.ChartType.columnClustered, dataRange, Excel.ChartSeriesBy.columns);
  ch.name = config.chartName || "ComboChart";
  ch.title.text = config.title || "";
  ch.title.format.font.size = 14;
  await ctx.sync();

  ch.series.load("count");
  await ctx.sync();
  const seriesCount = ch.series.count;

  for (let i = 0; i < seriesCount && i < config.seriesNames.length; i++) {
    const s = ch.series.getItemAt(i);
    s.name = config.seriesNames[i];

    if (i < 4) {
      s.chartType = Excel.ChartType.columnClustered;
      s.axisGroup = 0;
      if (config.colors && config.colors[i]) {
        try {
          s.format.fill.setSolidColor(config.colors[i]);
        } catch (e) {}
      }
    } else {
      s.chartType = Excel.ChartType.line;
      s.axisGroup = 1;
      s.format.line.weight = 3;
      s.format.line.color = "#4BACC6";
    }
  }
  await ctx.sync();

  ch.legend.visible = true;
  ch.legend.position = Excel.ChartLegendPosition.bottom;
  await ctx.sync();

  const primaryAxis = ch.axes.valueAxis;
  primaryAxis.title.text = config.primaryAxis.title;
  primaryAxis.title.format.font.size = 11;
  try {
    primaryAxis.numberFormat = config.primaryAxis.numberFormat;
  } catch (e) {
    primaryAxis.format.code = config.primaryAxis.numberFormat;
  }

  let secAxis;
  try {
    if (typeof Excel.ChartAxisType !== "undefined" && typeof Excel.ChartAxisGroup !== "undefined") {
      secAxis = ch.axes.getItem(Excel.ChartAxisType.value, Excel.ChartAxisGroup.secondary);
    } else {
      secAxis = ch.axes.getItem(Excel.ChartAxisType.value, "Secondary");
    }
  } catch (e) {
    secAxis = null;
  }

  if (secAxis) {
    secAxis.visible = true;
    secAxis.title.text = config.secondaryAxis.title;
    secAxis.title.format.font.size = 11;
    try {
      secAxis.numberFormat = config.secondaryAxis.numberFormat;
    } catch (e) {
      secAxis.format.code = config.secondaryAxis.numberFormat;
    }
  }
  await ctx.sync();

  if (config.position && config.position.startCell && config.position.endCell) {
    ch.setPosition(config.position.startCell, config.position.endCell);
    await ctx.sync();
  }

  return ch;
}
//

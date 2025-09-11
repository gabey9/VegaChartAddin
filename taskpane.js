/* global console, document, Excel, Office */

Office.onReady(() => {
  document.getElementById("drawChart").onclick = run;
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Get current selection
      const range = context.workbook.getSelectedRange();
      range.load("values, left, top");
      await context.sync();

      const values = range.values;
      if (values.length < 2) {
        console.warn("Need at least header row + one data row");
        return;
      }

      // First row is header
      const headers = values[0];
      const rows = values.slice(1);

      // Convert rows -> objects
      const data = rows.map(row => {
        let obj = {};
        headers.forEach((h, i) => {
          obj[h] = row[i];
        });
        return obj;
      });

      // Get chart type from dropdown (for minimal version, we'll just use line)
      const chartType = document.getElementById("chartType").value || "line";

      let spec;

      if (chartType === "line") {
        // Transform data for multi-series line chart
        const transformedData = [];
        const valueColumns = headers.slice(1);
        data.forEach(row => {
          valueColumns.forEach(colName => {
            if (row[colName] !== null && row[colName] !== undefined && row[colName] !== "") {
              transformedData.push({
                [headers[0]]: row[headers[0]], // x-axis value (first column)
                series: colName,               // series name (column header)
                value: parseFloat(row[colName]) || 0  // y-axis value
              });
            }
          });
        });

        spec = {
          $schema: "https://vega.github.io/schema/vega-lite/v6.json",
          description: "Multi-series line chart from Excel selection",
          background: "white",
          config: { view: { stroke: "transparent" }},
          data: { values: transformedData },
          mark: { 
            type: "line", 
            point: false,
            tooltip: true,
            strokeWidth: 2
          },
          encoding: {
            x: { 
              field: headers[0], 
              type: "ordinal",
              axis: {
                title: headers[0],
                labelFontSize: 12,
                titleFontSize: 14,
                labelAngle: 0
              }
            },
            y: { 
              field: "value", 
              type: "quantitative",
              axis: {
                title: "Value",
                labelFontSize: 12,
                titleFontSize: 14
              }
            },
            color: { 
              field: "series", 
              type: "nominal",
              scale: {
                scheme: "category10"
              },
              legend: {
                title: "Series",
                titleFontSize: 12,
                labelFontSize: 11
              }
            }
          },
          config: {
            font: "Segoe UI",
            axis: {
              labelColor: "#605e5c",
              titleColor: "#323130",
              gridColor: "#f3f2f1"
            },
            legend: {
              titleColor: "#323130",
              labelColor: "#605e5c"
            },
            point: {
              size: 60,
              filled: true
            }
          }
        };
      }

      // Validate spec was created
      if (!spec) {
        console.warn("No chart specification created for chart type:", chartType);
        return;
      }

      // Render hidden chart
      const hiddenDiv = document.createElement("div");
      hiddenDiv.style.display = "none";
      document.body.appendChild(hiddenDiv);

      try {
        const result = await vegaEmbed(hiddenDiv, spec, { actions: false });
        const view = result.view;

        // Export chart -> PNG
        const pngUrl = await view.toImageURL("png");
        const response = await fetch(pngUrl);
        const blob = await response.blob();

        // Use Promise-based approach for FileReader
        const base64data = await new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onloadend = () => {
            resolve(reader.result.split(",")[1]);
          };
          reader.onerror = reject;
          reader.readAsDataURL(blob);
        });

        // Insert picture at selection location
        const image = sheet.shapes.addImage(base64data);
        image.left = range.left;
        image.top = range.top;
        image.lockAspectRatio = true;

        await context.sync();
        
        console.log("Line chart generated successfully!");

      } catch (chartError) {
        console.error("Error generating chart:", chartError);
      } finally {
        // Clean up hidden div
        if (document.body.contains(hiddenDiv)) {
          document.body.removeChild(hiddenDiv);
        }
      }

    }); // End of Excel.run
  } catch (error) {
    console.error("Error in run function:", error);
  }
}
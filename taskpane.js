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

      // Get chart type from dropdown
      const chartType = document.getElementById("chartType").value;

      let spec;

      if (chartType === "waterfall") {
        spec = createWaterfallSpec(data, headers); 
      }

      else if (chartType === "pie") {
        spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        description: "Pie chart from Excel selection",
        data: { values: data },
        mark: { type: "arc", outerRadius: 120 },
        encoding: {
          theta: { field: headers[1], type: "quantitative" },
          color: { field: headers[0], type: "nominal" }
          }
        };
      } 

      else if (chartType === "radial") {
        spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        description: "Radial chart from Excel selection",
        data: { values: data },
        layer: [{
        mark: { type: "arc", innerRadius: 20, stroke: "#fff" }
        }, {
        mark: { type: "text", radiusOffset: 10 },
        encoding: {
            text: { field: headers[1], type: "quantitative" }
        }
        }],
        encoding: {
        theta: { field: headers[1], type: "quantitative", stack: true },
        radius: { 
            field: headers[1], 
            scale: { type: "sqrt", zero: true, rangeMin: 20 }
        },
        color: { field: headers[0], type: "nominal", legend: null }
        }
      };
    }

      else {
        spec = {
          $schema: "https://vega.github.io/schema/vega-lite/v6.json",
          description: "Chart from Excel selection",
          data: { values: data },
          mark: chartType,
          encoding: {
            x: { field: headers[0], type: "quantitative" },
            y: { field: headers[1], type: "quantitative" }
          }
        };
      }

      // Render hidden chart
      const hiddenDiv = document.createElement("div");
      hiddenDiv.style.display = "none";
      document.body.appendChild(hiddenDiv);

      const result = await vegaEmbed(hiddenDiv, spec, { actions: false });
      const view = result.view;

      // Export chart -> PNG
      const pngUrl = await view.toImageURL("png");
      const response = await fetch(pngUrl);
      const blob = await response.blob();

      const reader = new FileReader();
      reader.onloadend = async () => {
      const base64data = reader.result.split(",")[1];

      // Insert picture above selection
      const image = sheet.shapes.addImage(base64data);
      image.left = range.left;
      image.top = range.top;
      image.lockAspectRatio = true; // keep proportions

      await context.sync();
      };

      reader.readAsDataURL(blob);
      
      // Clean up hidden div
      document.body.removeChild(hiddenDiv);
    });
  } catch (error) {
    console.error(error);
  }
}

function createWaterfallSpec(data, headers) {
  // For waterfall charts, we expect:
  // - First column: labels (categories/time periods)
  // - Second column: amounts (positive/negative values)
  
  const labelField = headers[0];
  const amountField = headers[1];

  // Calculate dynamic dimensions based on data
  const numDataPoints = data.length;
  const dynamicWidth = Math.max(400, Math.min(1200, numDataPoints * 70));
  const maxAmount = Math.max(...data.map(d => Math.abs(d[amountField])));
  const dynamicHeight = Math.max(300, Math.min(600, maxAmount / 100 + 200));

  // Process data to ensure Begin and End points
  const processedData = processWaterfallData(data, labelField, amountField);

  return {
    $schema: "https://vega.github.io/schema/vega-lite/v6.json",
    description: "Waterfall chart from Excel selection",
    data: { values: processedData },
    width: dynamicWidth,
    height: dynamicHeight,
    transform: [
      {"window": [{"op": "sum", "field": amountField, "as": "sum"}]},
      {"window": [{"op": "lead", "field": labelField, "as": "lead"}]},
      {
        "calculate": `datum.lead === null ? datum.${labelField} : datum.lead`,
        "as": "lead"
      },
      {
        "calculate": `datum.${labelField} === '${data[data.length - 1][labelField]}' ? 0 : datum.sum - datum.${amountField}`,
        "as": "previous_sum"
      },
      {
        "calculate": `datum.${labelField} === '${data[data.length - 1][labelField]}' ? datum.sum : datum.${amountField}`,
        "as": "amount"
      },
      {
        "calculate": `datum.${labelField} === '${data[0][labelField]}' ? datum.${amountField} / 2 : datum.${labelField} === '${data[data.length - 1][labelField]}' ? datum.${amountField} / 2 : (datum.sum + datum.previous_sum) / 2`,
        "as": "center"
      },
      {
        "calculate": `datum.${labelField} === '${data[data.length - 1][labelField]}' ? datum.sum : (datum.${labelField} !== '${data[0][labelField]}' && datum.${labelField} !== '${data[data.length - 1][labelField]}' && datum.${amountField} > 0 ? '+' : '') + datum.${amountField}`,
        "as": "text_amount"
      },
      {"calculate": "(datum.sum + datum.previous_sum) / 2", "as": "center"}
    ],
    encoding: {
      x: {
        field: labelField,
        type: "ordinal",
        sort: null,
        axis: {"labelAngle": 0, "title": headers[0]},
        scale: {"paddingInner": 0.1, "paddingOuter": 0.05}
      }
    },
    layer: [
      {
        mark: {"type": "bar", "size": 60},
        encoding: {
          y: {
            field: "previous_sum",
            type: "quantitative",
            title: headers[1]
          },
          y2: {"field": "sum"},
          color: {
            condition: [
              {
                test: `datum.${labelField} === '${data[0][labelField]}' || datum.${labelField} === '${data[data.length - 1][labelField]}'`,
                value: "#f7e0b6"
              },
              {"test": "datum.sum < datum.previous_sum", "value": "#f78a64"}
            ],
            value: "#93c4aa"
          }
        }
      },
      {
        mark: {
          type: "rule",
          color: "#404040",
          opacity: 1,
          strokeWidth: 2,
          xOffset: -30,
          x2Offset: 30
        },
        encoding: {
          x2: {"field": "lead"},
          y: {"field": "sum", "type": "quantitative"}
        }
      },
      {
        "mark": {
          type: "text", 
          dy: {"expr": "datum.amount >= 0 ? -4 : 4"}, 
          baseline: {"expr": "datum.amount >= 0 ? 'bottom' : 'top'"}
        },
        encoding: {
          y: {"field": "sum", "type": "quantitative"},
          text: {"field": "sum", "type": "nominal"},
          opacity: {
            "condition": {
              "test": `datum.${labelField} === '${data[0][labelField]}' || datum.${labelField} === '${data[data.length - 1][labelField]}'`,
              "value": 0
            },
            "value": 1
          }
        }
      },
      {
        mark: {"type": "text", "fontWeight": "bold", "baseline": "middle"},
        encoding: {
          y: {"field": "center", "type": "quantitative"},
          text: {"field": "text_amount", "type": "nominal"},
          color: {
            condition: [
              {
                test: `datum.${labelField} === '${data[0][labelField]}' || datum.${labelField} === '${data[data.length - 1][labelField]}'`,
                value: "#725a30"
              }
            ],
            value: "white"
          }
        }
      }
    ],
    config: {"text": {"fontWeight": "bold", "color": "#404040"}}
  };
}

function processWaterfallData(data, labelField, amountField) {
  // Use the actual data from Excel, but set the last entry's amount to 0
  // since End should show cumulative total, not be an additional change
  const processedData = [...data];
  
  if (processedData.length > 0) {
    // Set last data point amount to 0 (End shows cumulative total)
    processedData[processedData.length - 1] = {
      ...processedData[processedData.length - 1],
      [amountField]: 0
    };
  }
  
  return processedData;
}
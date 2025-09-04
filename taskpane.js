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

      else if (chartType === "deviation") {
      spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        description: "Deviation chart from Excel selection",
        data: { values: data },
        layer: [
        {
            mark: { type: "line", tooltip: true, color: "grey" },
            encoding: {
            x: { field: headers[0], type: "ordinal" },
            y: { field: headers[1], type: "quantitative" }
            }
        },
        {
            mark: { type: "circle", size: 80, color: "grey", tooltip: true },
            encoding: {
            x: { field: headers[0], type: "ordinal" },
            y: { field: headers[1], type: "quantitative" }
            }
        },
        {
            mark: { type: "rule", strokeWidth: 2, tooltip: true },
            encoding: {
            x: { field: headers[0], type: "ordinal" },
            y: { field: headers[1], type: "quantitative" },
            y2: { field: headers[2] },
            color: {
                condition: { test: `datum["${headers[1]}"] < datum["${headers[2]}"]`, value: "red" },
                value: "green"
            }
            }
        },
        {
            mark: { type: "circle", size: 60, tooltip: true },
            encoding: {
            x: { field: headers[0], type: "ordinal" },
            y: { field: headers[2], type: "quantitative" },
            color: {
                condition: { test: `datum["${headers[1]}"] < datum["${headers[2]}"]`, value: "red" },
                value: "green"
            }
            }
        }
        ],
        encoding: {
        x: { field: headers[0], type: "ordinal", axis: null },
        y: { type: "quantitative", axis: { title: "" } }
        },
        config: {
        view: { stroke: "transparent" },
        line: { strokeWidth: 3, strokeCap: "round", strokeJoin: "round" },
        axis: {
            ticks: false,
            grid: false,
            domain: false,
            labelColor: "#605E5C",
            labelFontSize: 12
        }
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

      else if (chartType === "bump") {
      spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        description: "Bump chart from Excel selection",
        data: { values: data },
        encoding: {
        x: {
            field: headers[0],      // X-Axis (e.g. date)
            type: "ordinal",
            axis: { title: "" }
        },
        y: {
            field: headers[2],      // Rank values
            type: "ordinal",
            axis: false
        },
        order: {
            field: headers[0],
            type: "temporal"
        }
        },
        layer: [
        {
            mark: { type: "line", interpolate: "monotone" },
            encoding: {
            color: {
                field: headers[1],   // Category
                type: "nominal",
                legend: false
            }
            }
        },
        {
            mark: { type: "circle", size: 400, tooltip: true },
            encoding: {
            color: {
                field: headers[1],
                type: "nominal",
                legend: false
            }
            }
        },
        {
            mark: { type: "text", color: "white" },
            encoding: {
            text: { field: headers[2] }
            }
        },
        {
            // Left-side labels
            transform: [
            { window: [{ op: "rank", as: "rank" }], sort: [{ field: headers[0], order: "descending" }] },
            { filter: "datum.rank === 1" }
            ],
            mark: {
            type: "text",
            align: "left",
            baseline: "middle",
            dx: 15,
            fontWeight: "bold",
            fontSize: 12
            },
            encoding: {
            text: { field: headers[1], type: "nominal" },
            color: { field: headers[1], type: "nominal", legend: false }
            }
        },
        {
            // Right-side labels
            transform: [
            { window: [{ op: "rank", as: "rank" }], sort: [{ field: headers[0], order: "ascending" }] },
            { filter: "datum.rank === 1" }
            ],
            mark: {
            type: "text",
            align: "right",
            baseline: "middle",
            dx: -15,
            fontWeight: "bold",
            fontSize: 12
            },
            encoding: {
            text: { field: headers[1], type: "nominal" },
            color: { field: headers[1], type: "nominal", legend: false }
            }
        }
        ],
        config: {
        view: { stroke: "transparent" },
        line: { strokeWidth: 3, strokeCap: "round", strokeJoin: "round" },
        axis: {
            ticks: false,
            grid: false,
            domain: false,
            labelColor: "#666666",
            labelFontSize: 12
        }
        }
      };
      }

      else if (chartType === "ribbon") {
      spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        description: "Ribbon chart from Excel selection",
        data: { values: data },
        layer: [
        {
            mark: { type: "area", interpolate: "monotone", tooltip: true },
            encoding: {
            x: {
                field: headers[0],
                type: "ordinal" // temporal change to "ordinal" if your first col is not a date
            },
            y: {
                aggregate: "sum",
                field: headers[2],
                type: "quantitative",
                axis: null,
                stack: "center"
            },
            color: {
                field: headers[1],
                type: "nominal"
            },
            order: {
                aggregate: "sum",
                field: headers[2],
                type: "quantitative"
            }
            }
        }
        ],
        config: {
        view: { stroke: "transparent" },
        axis: {
            ticks: false,
            grid: true,
            gridColor: "white",
            gridWidth: 3,
            domain: false,
            labelColor: "#666666"
        },
        legend: {
            titleFont: "Segoe UI",
            titleFontWeight: "bold",
            titleColor: "#666666",
            labelFont: "Segoe UI",
            labelColor: "#666666",
            symbolType: "circle",
            symbolSize: 75
        }
        }
      };
      }

      else if (chartType === "ridgeline") {
      spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        description: "Ridgeline (Joyplot) chart from Excel selection",
        data: { values: data },
        mark: {
        type: "area",
        fillOpacity: 0.7,
        strokeOpacity: 1,
        strokeWidth: 1,
        interpolate: "monotone"
        },
        width: 400,
        height: 20,
        encoding: {
        x: {
            field: headers[0],       // date/time column
            type: "ordinal",
            title: headers[0]
        },
        y: {
            aggregate: "sum",
            field: headers[2],       // value column
            type: "quantitative",
            scale: { range: [20, -40] },
            axis: {
            title: null,
            values: [0],
            domain: false,
            labels: false,
            ticks: false
            }
        },
        row: {
            field: headers[1],       // category column
            type: "nominal",
            title: headers[1],
            header: {
            title: null,
            labelAngle: 0,
            labelOrient: "left",
            labelAlign: "left",
            labelPadding: 0
            },
            sort: { field: headers[0], op: "max", order: "ascending" }
        },
        fill: {
            field: headers[1],
            type: "nominal",
            legend: null,
            scale: { scheme: "plasma" }
        }
        },
        resolve: { scale: { y: "independent" } },
        config: {
        view: { stroke: "transparent" },
        facet: { spacing: 20 },
        header: {
            labelFontSize: 12,
            labelFontWeight: 500,
            labelAngle: 0,
            labelAnchor: "end",
            labelOrient: "top",
            labelPadding: -19
        },
        axis: {
            domain: false,
            grid: false,
            ticks: false,
            tickCount: 5,
            labelFontSize: 12,
            titleFontSize: 12,
            titleFontWeight: 400,
            titleColor: "#605E5C"
        }
        }
      };
      }

      else if (chartType === "wordcloud") {
      spec = {
        $schema: "https://vega.github.io/schema/vega/v5.json",
        description: "Word cloud from Excel selection",
        width: 800,
        height: 400,
        padding: 0,
        
        data: [
        {
            name: "table",
            values: data,
            transform: [
            {
                type: "countpattern",
                field: headers[0], // Use first column as text source
                case: "upper",
                pattern: "[\\w']{3,}",
                stopwords: "(i|me|my|myself|we|us|our|ours|ourselves|you|your|yours|yourself|yourselves|he|him|his|himself|she|her|hers|herself|it|its|itself|they|them|their|theirs|themselves|what|which|who|whom|whose|this|that|these|those|am|is|are|was|were|be|been|being|have|has|had|having|do|does|did|doing|will|would|should|can|could|ought|i'm|you're|he's|she's|it's|we're|they're|i've|you've|we've|they've|i'd|you'd|he'd|she'd|we'd|they'd|i'll|you'll|he'll|she'll|we'll|they'll|isn't|aren't|wasn't|weren't|hasn't|haven't|hadn't|doesn't|don't|didn't|won't|wouldn't|shan't|shouldn't|can't|cannot|couldn't|mustn't|let's|that's|who's|what's|here's|there's|when's|where's|why's|how's|a|an|the|and|but|if|or|because|as|until|while|of|at|by|for|with|about|against|between|into|through|during|before|after|above|below|to|from|up|upon|down|in|out|on|off|over|under|again|further|then|once|here|there|when|where|why|how|all|any|both|each|few|more|most|other|some|such|no|nor|not|only|own|same|so|than|too|very|say|says|said|shall)"
            },
            {
                type: "formula", 
                as: "angle",
                expr: "[-45, 0, 45][~~(random() * 3)]"
            },
            {
                type: "formula", 
                as: "weight",
                expr: "if(datum.count > 10, 600, 300)"
            }
            ]
        }
        ],
        
        scales: [
        {
            name: "color",
            type: "ordinal",
            domain: { data: "table", field: "text" },
            range: ["#d5a928", "#652c90", "#939597", "#2563eb", "#dc2626", "#059669"]
        }
        ],
        
        marks: [
        {
            type: "text",
            from: { data: "table" },
            encode: {
            enter: {
                text: { field: "text" },
                align: { value: "center" },
                baseline: { value: "alphabetic" },
                fill: { scale: "color", field: "text" }
            },
            update: {
                fillOpacity: { value: 1 }
            },
            hover: {
                fillOpacity: { value: 0.5 }
            }
            },
            transform: [
            {
                type: "wordcloud",
                size: [800, 400],
                text: { field: "text" },
                rotate: { field: "datum.angle" },
                font: "Helvetica Neue, Arial",
                fontSize: { field: "datum.count" },
                fontWeight: { field: "datum.weight" },
                fontSizeRange: [12, 56],
                padding: 2
            }
            ]
        }
        ]
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
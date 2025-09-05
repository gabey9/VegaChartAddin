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
        // Process waterfall data inline - set last entry's amount to 0
        const processedData = [...data];
        if (processedData.length > 0) {
          processedData[processedData.length - 1] = {
            ...processedData[processedData.length - 1],
            [headers[1]]: 0
          };
        }

        // Calculate dynamic dimensions
        const numDataPoints = data.length;
        const dynamicWidth = Math.max(400, Math.min(1200, numDataPoints * 70));
        const maxAmount = Math.max(...data.map(d => Math.abs(d[headers[1]])));
        const dynamicHeight = Math.max(300, Math.min(600, maxAmount / 100 + 200));

        spec = {
          $schema: "https://vega.github.io/schema/vega-lite/v6.json",
          description: "Waterfall chart from Excel selection",
          data: { values: processedData },
          width: dynamicWidth,
          height: dynamicHeight,
          transform: [
            {"window": [{"op": "sum", "field": headers[1], "as": "sum"}]},
            {"window": [{"op": "lead", "field": headers[0], "as": "lead"}]},
            {
              "calculate": `datum.lead === null ? datum.${headers[0]} : datum.lead`,
              "as": "lead"
            },
            {
              "calculate": `datum.${headers[0]} === '${data[data.length - 1][headers[0]]}' ? 0 : datum.sum - datum.${headers[1]}`,
              "as": "previous_sum"
            },
            {
              "calculate": `datum.${headers[0]} === '${data[data.length - 1][headers[0]]}' ? datum.sum : datum.${headers[1]}`,
              "as": "amount"
            },
            {
              "calculate": `datum.${headers[0]} === '${data[0][headers[0]]}' ? datum.${headers[1]} / 2 : datum.${headers[0]} === '${data[data.length - 1][headers[0]]}' ? datum.${headers[1]} / 2 : (datum.sum + datum.previous_sum) / 2`,
              "as": "center"
            },
            {
              "calculate": `datum.${headers[0]} === '${data[data.length - 1][headers[0]]}' ? datum.sum : (datum.${headers[0]} !== '${data[0][headers[0]]}' && datum.${headers[0]} !== '${data[data.length - 1][headers[0]]}' && datum.${headers[1]} > 0 ? '+' : '') + datum.${headers[1]}`,
              "as": "text_amount"
            },
            {"calculate": "(datum.sum + datum.previous_sum) / 2", "as": "center"}
          ],
          encoding: {
            x: {
              field: headers[0],
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
                      test: `datum.${headers[0]} === '${data[0][headers[0]]}' || datum.${headers[0]} === '${data[data.length - 1][headers[0]]}'`,
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
                    "test": `datum.${headers[0]} === '${data[0][headers[0]]}' || datum.${headers[0]} === '${data[data.length - 1][headers[0]]}'`,
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
                      test: `datum.${headers[0]} === '${data[0][headers[0]]}' || datum.${headers[0]} === '${data[data.length - 1][headers[0]]}'`,
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

      else if (chartType === "mekko") {
        spec = {
          $schema: "https://vega.github.io/schema/vega/v5.json",
          description: "Marimekko chart from Excel selection",
          width: 800,
          height: 500,
          background: "#f8f9fa",
          view: { stroke: null },
          padding: { top: 60, bottom: 80, left: 60, right: 60 },
          data: [
            {
              name: "table",
              values: data
            },
            {
              name: "categories",
              source: "table",
              transform: [
                {
                  type: "aggregate",
                  fields: [headers[2]],
                  ops: ["sum"],
                  as: ["categoryTotal"],
                  groupby: [headers[0]]
                },
                {
                  type: "stack",
                  offset: "normalize",
                  sort: { field: "categoryTotal", order: "descending" },
                  field: "categoryTotal",
                  as: ["x0", "x1"]
                },
                {
                  type: "formula",
                  as: "Percent",
                  expr: "datum.x1-datum.x0"
                },
                {
                  type: "formula",
                  as: "Label",
                  expr: `datum.${headers[0]} + ' (' + format(datum.Percent,'.1%') + ')'`
                }
              ]
            },
            {
              name: "finalTable",
              source: "table",
              transform: [
                {
                  type: "stack",
                  offset: "normalize",
                  groupby: [headers[0]],
                  sort: { field: headers[2], order: "descending" },
                  field: headers[2],
                  as: ["y0", "y1"]
                },
                {
                  type: "stack",
                  groupby: [headers[0]],
                  sort: { field: headers[2], order: "descending" },
                  field: headers[2],
                  as: ["z0", "z1"]
                },
                {
                  type: "lookup",
                  from: "categories",
                  key: headers[0],
                  values: ["x0", "x1"],
                  fields: [headers[0]]
                },
                {
                  type: "formula",
                  as: "Percent",
                  expr: "datum.y1-datum.y0"
                },
                {
                  type: "formula",
                  as: "Label",
                  expr: `[datum.${headers[1]}, format(datum.${headers[2]}, '.0f') + ' (' + format(datum.Percent, '.1%') + ')']`
                },
                {
                  type: "window",
                  sort: { field: "y0", order: "ascending" },
                  ops: ["row_number"],
                  fields: [null],
                  as: ["rank"],
                  groupby: [headers[0]]
                }
              ]
            }
          ],
          scales: [
            {
              name: "x",
              type: "linear",
              range: "width",
              domain: { data: "finalTable", field: "x1" }
            },
            {
              name: "y",
              type: "linear",
              range: "height",
              nice: false,
              zero: true,
              domain: { data: "finalTable", field: "z1" }
            },
            {
              name: "opacity",
              type: "linear",
              range: [1, 0.6],
              domain: { data: "finalTable", field: "rank" }
            },
            {
              name: "color",
              type: "ordinal",
              range: { scheme: "category20" },
              domain: {
                data: "categories",
                field: headers[0],
                sort: { field: "x0", order: "ascending", op: "sum" }
              }
            }
          ],
          axes: [
            {
              orient: "left",
              scale: "y",
              zindex: 1,
              format: "",
              tickCount: 5,
              tickSize: 15,
              labelColor: { value: "#333740" },
              labelFontWeight: { value: "normal" },
              labelFontSize: { value: 12 },
              labelFont: { value: "Segoe UI" },
              offset: 5,
              domain: false,
              encode: {
                labels: {
                  update: {
                    text: { signal: `format(datum.value, '.0f')` }
                  }
                }
              }
            }
          ],
          marks: [
            {
              type: "rect",
              name: "bars",
              from: { data: "finalTable" },
              encode: {
                update: {
                  x: { scale: "x", field: "x0" },
                  x2: { scale: "x", field: "x1" },
                  y: { scale: "y", field: "z0" },
                  y2: { scale: "y", field: "z1" },
                  fill: { scale: "color", field: headers[0] },
                  stroke: { value: "white" },
                  strokeWidth: { value: 1 },
                  fillOpacity: { scale: "opacity", field: "rank" },
                  tooltip: { signal: "datum" }
                }
              }
            },
            {
              type: "text",
              name: "labels",
              interactive: false,
              from: { data: "bars" },
              encode: {
                update: {
                  x: { signal: "(datum.x2 - datum.x)*0.5 + datum.x" },
                  align: { value: "center" },
                  text: { field: "datum.Label" },
                  y: { signal: "(datum.y2 - datum.y)*0.5 + datum.y" },
                  fill: { value: "white" },
                  font: { value: "Segoe UI" },
                  lineHeight: { value: 12 },
                  fontSize: { value: 10 },
                  opacity: { signal: "(datum.x2 - datum.x) > 0.05 && (datum.y2 - datum.y) > 20 ? 1 : 0" }
                }
              }
            },
            {
              type: "text",
              name: "categoryLabels",
              from: { data: "categories" },
              encode: {
                update: {
                  x: { scale: "x", signal: "(datum.x1-datum.x0)/2 + datum.x0" },
                  y: { signal: "-15" },
                  text: { field: headers[0] },
                  align: { value: "center" },
                  baseline: { value: "bottom" },
                  fill: { value: "#333740" },
                  fontWeight: { value: "bold" },
                  fontSize: { value: 12 },
                  font: { value: "Segoe UI" }
                }
              }
            },
            {
              type: "text",
              name: "categoryPercentages",
              from: { data: "categories" },
              encode: {
                update: {
                  x: { scale: "x", signal: "(datum.x1-datum.x0)/2 + datum.x0" },
                  y: { signal: "height + 30" },
                  text: { field: "Label" },
                  align: { value: "center" },
                  baseline: { value: "top" },
                  fill: { value: "#666666" },
                  fontWeight: { value: "normal" },
                  fontSize: { value: 10 },
                  font: { value: "Segoe UI" }
                }
              }
            }
          ]
        };
      }

      else if (chartType === "lollipop") {
        spec = {
          $schema: "https://vega.github.io/schema/vega-lite/v5.json",
          description: "Lollipop chart from Excel selection",
          data: { values: data },
          encoding: {
            y: {
              field: headers[0],
              type: "nominal",
              sort: "-x",
              axis: {
                domain: false,
                title: null,
                ticks: false,
                labelFont: "Segoe UI",
                labelFontSize: 14,
                labelPadding: 10,
                labelColor: "#605e5c"
              }
            },
            x: {
              field: headers[1],
              type: "quantitative",
              axis: {
                domain: false,
                ticks: false,
                grid: true,
                gridColor: "#e0e0e0",
                labelFont: "Segoe UI",
                labelFontSize: 12,
                labelColor: "#605e5c",
                title: headers[1],
                titleFont: "Segoe UI",
                titleFontSize: 14,
                titleColor: "#323130"
              }
            },
            color: { value: "#0078d4" }
          },
          layer: [
            {
              mark: {
                type: "rule",
                tooltip: true,
                strokeWidth: 3,
                opacity: 0.7
              }
            },
            {
              mark: {
                type: "circle",
                tooltip: true,
                size: 300,
                opacity: 0.9
              },
              encoding: {
                size: {
                  field: headers[1],
                  type: "quantitative",
                  scale: {
                    range: [200, 800]
                  },
                  legend: null
                }
              }
            }
          ],
          config: {
            autosize: {
              type: "fit",
              contains: "padding"
            },
            view: { stroke: "transparent" },
            font: "Segoe UI",
            text: { font: "Segoe UI", fontSize: 12, fill: "#605E5C" }
          }
        };
      }

      else if (chartType === "waffle") {
        spec = {
          $schema: "https://vega.github.io/schema/vega-lite/v5.json",
          description: "Waffle chart from Excel selection",
          data: { values: data },
          transform: [
            {
              joinaggregate: [{"op": "sum", "field": headers[1], "as": "TotalValue"}]
            },
            {
              calculate: `round(datum.${headers[1]}/datum.TotalValue * 100)`,
              as: "PercentOfTotal"
            },
            {
              aggregate: [{"op": "min", "field": "PercentOfTotal", "as": "Percent"}],
              groupby: [headers[0]]
            },
            {"calculate": "sequence(1,101)", "as": "Sequence"},
            {"flatten": ["Sequence"]},
            {
              calculate: `if(datum.Sequence <= datum.Percent, datum.${headers[0]},'_blank')`,
              as: "Plot"
            },
            {"calculate": "ceil (datum.Sequence / 10)", "as": "row"},
            {"calculate": "datum.Sequence - datum.row * 10", "as": "col"}
          ],
          facet: {"column": {"field": headers[0], "header": {"labelOrient": "bottom"}}},
          spec: {
            layer: [
              {
                mark: {
                  type: "circle",
                  filled: true,
                  tooltip: true,
                  stroke: "#9e9b9b",
                  strokeWidth: 0.7
                },
                encoding: {
                  x: {"field": "col", "type": "ordinal", "axis": null},
                  y: {"field": "row", "type": "ordinal", "axis": null, "sort": "-y"},
                  color: {
                    condition: {"test": "datum.Plot == '_blank'", "value": "#e6e3e3"},
                    scale: {"scheme": "set1"},
                    field: "Plot",
                    type: "nominal",
                    legend: null
                  },
                  size: {"value": 241},
                  tooltip: [{"field": headers[0], "type": "nominal"}]
                }
              },
              {
                mark: {"type": "text", "fontSize": 30, "fontWeight": "bold"},
                encoding: {
                  y: {"value": 30},
                  text: {
                    condition: {
                      test: "datum.Sequence == 1",
                      value: {"expr": "datum.Percent + '%'"}
                    }
                  },
                  color: {"scale": {"scheme": "set1"}, "field": "Plot"}
                }
              }
            ]
          },
          config: {
            view: {"stroke": "transparent"},
            font: "Segoe UI",
            text: {"font": "Segoe UI", "fontSize": 12, "fill": "#605E5C"},
            axis: {
              ticks: false,
              grid: false,
              domain: false,
              labelColor: "#605E5C",
              labelFontSize: 12
            },
            header: {
              titleFont: "Segoe UI",
              titleFontSize: 16,
              titleColor: "#757575",
              labelFont: "Segoe UI",
              labelFontSize: 13,
              labelColor: "#605E5C"
            },
            legend: {
              titleFont: "Segoe UI",
              titleFontWeight: "bold",
              titleColor: "#605E5C",
              labelFont: "Segoe UI",
              labelFontSize: 13,
              labelColor: "#605E5C",
              symbolType: "circle",
              symbolSize: 75
            }
          }
        };
      }

      else if (chartType === "violin") {
        spec = {
          $schema: "https://vega.github.io/schema/vega-lite/v5.json",
          description: "Violin chart from Excel selection",
          data: { values: data },
          spacing: 0,
          facet: {
            column: {
              field: headers[0],
              header: {
                orient: "bottom",
                title: null,
                labelFontSize: 14
              }
            }
          },
          spec: {
            width: 180,
            height: 400,
            encoding: {
              y: {
                field: headers[1],
                type: "quantitative",
                title: headers[1],
                axis: {
                  tickCount: 10,
                  titleFontSize: 16,
                  labelFontSize: 12
                }
              },
              x: {
                type: "quantitative",
                axis: {
                  labels: false,
                  title: null,
                  grid: false,
                  ticks: false
                }
              },
              color: {
                field: headers[0],
                type: "nominal",
                legend: null,
                scale: {
                  scheme: "category10"
                }
              }
            },
            layer: [
              {
                name: "KDE_PLOT",
                transform: [
                  {
                    density: headers[1],
                    groupby: [headers[0]],
                    as: ["_kde_value", "_kde_density"]
                  },
                  {
                    calculate: "datum['_kde_density'] * -1",
                    as: "_negative_kde_density"
                  }
                ],
                layer: [
                  {
                    name: "KDE_POSITIVE",
                    mark: {
                      type: "area",
                      orient: "vertical",
                      opacity: 0.6
                    },
                    encoding: {
                      y: { field: "_kde_value" },
                      x: { field: "_kde_density" }
                    }
                  },
                  {
                    name: "KDE_NEGATIVE",
                    mark: {
                      type: "area",
                      orient: "vertical",
                      opacity: 0.6
                    },
                    encoding: {
                      y: { field: "_kde_value" },
                      x: { field: "_negative_kde_density" }
                    }
                  }
                ],
                encoding: {
                  x2: { datum: 0 }
                }
              },
              {
                name: "BOX_PLOT",
                mark: {
                  type: "boxplot",
                  extent: "min-max",
                  median: {
                    color: "black",
                    strokeWidth: 2
                  },
                  size: 20
                },
                encoding: {
                  y: { field: headers[1] },
                  fill: { value: "#969696" },
                  stroke: { value: "black" }
                }
              }
            ]
          },
          config: {
            view: { stroke: "transparent" },
            font: "Segoe UI",
            text: { font: "Segoe UI", fontSize: 12, fill: "#605E5C" },
            axis: {
              ticks: false,
              grid: true,
              gridColor: "#e0e0e0",
              domain: false,
              labelColor: "#605E5C",
              labelFontSize: 12
            },
            header: {
              titleFont: "Segoe UI",
              titleFontSize: 16,
              titleColor: "#757575",
              labelFont: "Segoe UI",
              labelFontSize: 13,
              labelColor: "#605E5C"
            }
          }
        };
      }

      else if (chartType === "heatmap") {
        spec = {
          $schema: "https://vega.github.io/schema/vega-lite/v5.json",
          description: "Heatmap with marginal bars from Excel selection",
          data: { values: data },
          spacing: 15,
          bounds: "flush",
          vconcat: [
            {
              height: 60,
              mark: {
                type: "bar",
                stroke: null,
                cornerRadiusEnd: 2,
                tooltip: true,
                color: "lightgrey"
              },
              encoding: {
                x: {
                  field: headers[1],
                  type: "ordinal",
                  axis: null
                },
                y: {
                  field: headers[2],
                  aggregate: "mean",
                  type: "quantitative",
                  axis: null
                }
              }
            },
            {
              spacing: 15,
              bounds: "flush",
              hconcat: [
                {
                  mark: {
                    type: "rect",
                    stroke: "white",
                    tooltip: true
                  },
                  encoding: {
                    y: {
                      field: headers[0],
                      type: "ordinal",
                      title: headers[0],
                      axis: {
                        domain: false,
                        ticks: false,
                        labels: true,
                        labelAngle: 0,
                        labelPadding: 5
                      }
                    },
                    x: {
                      field: headers[1],
                      type: "ordinal",
                      title: headers[1],
                      axis: {
                        domain: false,
                        ticks: false,
                        labels: true,
                        labelAngle: 0
                      }
                    },
                    color: {
                      aggregate: "mean",
                      field: headers[2],
                      type: "quantitative",
                      title: headers[2],
                      scale: {
                        scheme: "blues"
                      },
                      legend: {
                        direction: "vertical",
                        gradientLength: 120
                      }
                    }
                  }
                },
                {
                  mark: {
                    type: "bar",
                    stroke: null,
                    cornerRadiusEnd: 2,
                    tooltip: true,
                    color: "lightgrey"
                  },
                  width: 60,
                  encoding: {
                    y: {
                      field: headers[0],
                      type: "ordinal",
                      axis: null
                    },
                    x: {
                      field: headers[2],
                      type: "quantitative",
                      aggregate: "mean",
                      axis: null
                    }
                  }
                }
              ]
            }
          ],
          config: {
            autosize: {
              type: "fit",
              contains: "padding"
            },
            view: { stroke: "transparent" },
            font: "Segoe UI",
            text: { font: "Segoe UI", fontSize: 12, fill: "#605E5C" },
            axis: {
              ticks: false,
              grid: false,
              domain: false,
              labelColor: "#605E5C",
              labelFontSize: 12,
              titleFontSize: 14,
              titleColor: "#323130"
            },
            legend: {
              titleFont: "Segoe UI",
              titleFontWeight: "bold",
              titleColor: "#605E5C",
              labelFont: "Segoe UI",
              labelFontSize: 12,
              labelColor: "#605E5C"
            }
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
            type: "temporal",	    
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
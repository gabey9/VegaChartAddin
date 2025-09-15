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

      if (chartType === "point") {
        spec = {
          $schema: "https://vega.github.io/schema/vega-lite/v6.json",
          description: "Colored scatter plot from Excel selection",
          background: "white",
          config: { view: { stroke: "transparent" }},
          data: { values: data },
          mark: { type: "point", size: 100, tooltip: true },
          encoding: {
            x: { 
              field: headers[0], 
              type: "quantitative",
              scale: { zero: false },
              axis: {
                title: headers[0],
                labelFontSize: 12,
                titleFontSize: 14
              }
            },
            y: { 
              field: headers[1], 
              type: "quantitative",
              scale: { zero: false },
              axis: {
                title: headers[1],
                labelFontSize: 12,
                titleFontSize: 14
              }
            },
            // Add color encoding if 3rd column exists
            ...(headers.length >= 3 && {
              color: { 
                field: headers[2], 
                type: "nominal",
                legend: {
                  title: headers[2],
                  titleFontSize: 12,
                  labelFontSize: 11
                }
              }
            }),
            // Add shape encoding if 4th column exists
            ...(headers.length >= 4 && {
              shape: { 
                field: headers[3], 
                type: "nominal",
                legend: {
                  title: headers[3],
                  titleFontSize: 12,
                  labelFontSize: 11
                }
              }
            })
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
            }
          }
        };
      }

else if (chartType === "bubble") {
  if (headers.length < 3) {
    console.warn("Bubble chart requires at least 3 columns (X values, Y values, Size values)");
    return;
  }

  spec = {
    $schema: "https://vega.github.io/schema/vega-lite/v6.json",
    description: "Bubble chart from Excel selection",
    background: "white",
    config: { view: { stroke: "transparent" }},
    data: { values: data },
    mark: { 
      type: "point", 
      tooltip: true,
      opacity: 0.7,
      stroke: "white",
      strokeWidth: 1
    },
    encoding: {
      x: { 
        field: headers[0], 
        type: "quantitative",
        scale: { zero: false },
        axis: {
          title: headers[0],
          labelFontSize: 12,
          titleFontSize: 14,
          labelColor: "#605e5c",
          titleColor: "#323130",
          gridColor: "#f3f2f1"
        }
      },
      y: { 
        field: headers[1], 
        type: "quantitative",
        scale: { zero: false },
        axis: {
          title: headers[1],
          labelFontSize: 12,
          titleFontSize: 14,
          labelColor: "#605e5c",
          titleColor: "#323130",
          gridColor: "#f3f2f1"
        }
      },
      size: {
        field: headers[2],
        type: "quantitative",
        scale: {
          type: "sqrt",
          range: [50, 1000]
        },
        legend: {
          title: headers[2],
          titleFontSize: 12,
          labelFontSize: 11,
          titleColor: "#323130",
          labelColor: "#605e5c"
        }
      },
      // Add color encoding if 4th column exists
      ...(headers.length >= 4 && {
        color: { 
          field: headers[3], 
          type: "nominal",
          legend: {
            title: headers[3],
            titleFontSize: 12,
            labelFontSize: 11,
            titleColor: "#323130",
            labelColor: "#605e5c"
          }
        }
      })
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
      }
    }
  };
}

      else if (chartType === "line") {
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

      else if (chartType === "gantt") {
      function excelDateToJSDate(serial) {
          return new Date(Math.round((serial - 25569) * 86400 * 1000));
      }

      const ganttData = rows.map(row => {
          const parentId = row[0] || null;   // col 1 = parent id
          const id = row[1];                 // col 2 = task id
          const name = row[2] || `Task ${id}`;
          if (!id) return null;

          const start = typeof row[3] === "number" ? excelDateToJSDate(row[3]) : new Date(row[3]);
          const end = typeof row[4] === "number" ? excelDateToJSDate(row[4]) : new Date(row[4]);
          if (!(start instanceof Date) || isNaN(start) || !(end instanceof Date) || isNaN(end)) return null;

          let progress = 0;
          if (row[5]) {
              if (typeof row[5] === "string" && row[5].includes("%")) {
                  progress = parseFloat(row[5]) / 100;
              } else if (row[5] > 1) {
                  progress = row[5] / 100;
              } else {
                  progress = row[5];
              }
          }

          const dependencies = row[6] ? String(row[6]).split(",").map(d => d.trim()) : [];

          return { id, parentId, name, startDate: start, endDate: end, progress, dependencies };
      }).filter(Boolean);

      // Precompute progressEnd
      ganttData.forEach(task => {
          const duration = task.endDate - task.startDate;
          task.progressEnd = new Date(task.startDate.getTime() + duration * task.progress);
      });

      // Build parent->children map
      const childrenMap = new Map();
      ganttData.forEach(task => {
          if (!childrenMap.has(task.parentId)) {
              childrenMap.set(task.parentId, []);
          }
          childrenMap.get(task.parentId).push(task);
      });

      // Sort children by startDate
      for (let [pid, childList] of childrenMap.entries()) {
          childList.sort((a, b) => a.startDate - b.startDate);
      }

      // Recursive hierarchy ordering
      function buildHierarchy(parentId = null, level = 0) {
          const ordered = [];
          const tasks = childrenMap.get(parentId) || [];
          for (const task of tasks) {
              task.level = level;
              ordered.push(task);
              ordered.push(...buildHierarchy(task.id, level + 1));
          }
          return ordered;
      }

      const orderedTasks = buildHierarchy(null);

      spec = {
          $schema: "https://vega.github.io/schema/vega-lite/v5.json",
          description: "Gantt Chart from Excel Data",
          width: 800,
          height: Math.max(300, orderedTasks.length * 30),
          data: { values: orderedTasks },
          layer: [
              {
                  mark: { type: "bar", opacity: 0.3, height: 20 },
                  encoding: {
                      y: { field: "name", type: "nominal", axis: { title: null, labelFontSize: 11 }, sort: null },
                      x: { field: "startDate", type: "temporal", axis: { title: "Timeline", format: "%b %d", labelAngle: -45 } },
                      x2: { field: "endDate", type: "temporal" },
                      color: { field: "level", type: "ordinal", scale: { scheme: "category10" }, legend: { title: "Level" } },
                      tooltip: [
                          { field: "name", type: "nominal", title: "Task" },
                          { field: "startDate", type: "temporal", title: "Start", format: "%Y-%m-%d" },
                          { field: "endDate", type: "temporal", title: "End", format: "%Y-%m-%d" },
                          { field: "progress", type: "quantitative", title: "Progress", format: ".0%" }
                      ]
                  }
              },
              {
                  mark: { type: "bar", opacity: 0.8, height: 20 },
                  encoding: {
                      y: { field: "name", type: "nominal", sort: null },
                      x: { field: "startDate", type: "temporal" },
                      x2: { field: "progressEnd", type: "temporal" },
                      color: { field: "level", type: "ordinal", scale: { scheme: "category10" } }
                  }
              },
              {
                  mark: { type: "text", align: "left", baseline: "middle", dx: 5, fontSize: 10 },
                  encoding: {
                      y: { field: "name", type: "nominal", sort: null },
                      x: { field: "endDate", type: "temporal" },
                      text: { field: "progress", type: "quantitative", format: ".0%" },
                      color: { value: "#666" }
                  }
              },
              {
                  mark: { type: "rule", strokeDash: [4, 4], opacity: 0.5 },
                  data: { values: [{ date: new Date().toISOString() }] },
                  encoding: {
                      x: { field: "date", type: "temporal" },
                      color: { value: "red" },
                      size: { value: 1 }
                  }
              }
          ],
          config: { view: { stroke: null }, axis: { grid: true, gridColor: "#f0f0f0" } }
        };
      }

      else if (chartType === "slope") {
        const timePeriods = [...new Set(data.map(d => d[headers[0]]))];
        const categories = [...new Set(data.map(d => d[headers[1]]))];
        
        // Filter data for first and last periods only
        const firstPeriod = timePeriods[0];
        const lastPeriod = timePeriods[timePeriods.length - 1];
        
        const slopeData = data.filter(d => 
            d[headers[0]] === firstPeriod || d[headers[0]] === lastPeriod
        );

        // Check if values are percentages (between -1 and 1)
        const allValues = slopeData.map(d => d[headers[2]]);
        const isPercentage = allValues.every(v => v >= -1 && v <= 1);
        const formatString = isPercentage ? ".1%" : ",.0f";

        // Calculate dynamic dimensions based on number of categories
        const dynamicHeight = Math.max(300, Math.min(600, categories.length * 40));
        const dynamicWidth = 500;

        spec = {
            $schema: "https://vega.github.io/schema/vega-lite/v6.json",
            description: "Slope chart from Excel selection",
            background: "white",
            config: { 
            view: { stroke: "transparent" },
            autosize: { type: "fit", contains: "padding" }
            },
            width: dynamicWidth,
            height: dynamicHeight,
            data: { values: slopeData },
            encoding: {
            x: {
                field: headers[0],
                type: "ordinal",
                axis: {
                title: null,
                labelFontSize: 14,
                labelFontWeight: "bold",
                labelPadding: 10,
                domain: false,
                ticks: false,
                labelColor: "#323130"
                },
                scale: { padding: 0.1 }
            },
            y: {
                field: headers[2],
                type: "quantitative",
                axis: null,
                scale: { zero: false }
            },
            color: {
                field: headers[1],
                type: "nominal",
                legend: null,
                scale: { scheme: "category10" }
            }
            },
            layer: [
            // Background grid lines
            {
                mark: {
                type: "rule",
                strokeDash: [2, 2],
                opacity: 0.3
                },
                data: { values: [{}] },
                encoding: {
                x: { datum: firstPeriod },
                x2: { datum: lastPeriod },
                y: { value: 0 },
                color: { value: "#e0e0e0" }
                }
            },
            // Slope lines
            {
                mark: {
                type: "line",
                strokeWidth: 2,
                opacity: 0.7,
                tooltip: true
                },
                encoding: {
                detail: { field: headers[1], type: "nominal" },
                tooltip: [
                    { field: headers[1], type: "nominal", title: "Category" },
                    { field: headers[0], type: "nominal", title: "Period" },
                    { field: headers[2], type: "quantitative", title: "Value", format: formatString }
                ]
                }
            },
            // Points at the ends
            {
                mark: {
                type: "circle",
                size: 100,
                opacity: 1,
                tooltip: true
                }
            },
            // Left side value labels
            {
                transform: [
                { filter: `datum['${headers[0]}'] == '${firstPeriod}'` }
                ],
                mark: {
                type: "text",
                align: "right",
                baseline: "middle",
                dx: -8,
                fontSize: 11,
                fontWeight: "normal"
                },
                encoding: {
                text: { 
                    field: headers[2], 
                    type: "quantitative",
                    format: formatString
                }
                }
            },
            // Left side category labels (for top values)
            {
                transform: [
                { filter: `datum['${headers[0]}'] == '${firstPeriod}'` },
                {
                    window: [{ op: "rank", as: "rank" }],
                    sort: [{ field: headers[2], order: "descending" }]
                },
                { filter: "datum.rank <= 3" }
                ],
                mark: {
                type: "text",
                align: "right",
                baseline: "bottom",
                dx: -8,
                dy: -12,
                fontSize: 10,
                fontWeight: "bold",
                fontStyle: "italic"
                },
                encoding: {
                text: { field: headers[1], type: "nominal" }
                }
            },
            // Right side value labels
            {
                transform: [
                { filter: `datum['${headers[0]}'] == '${lastPeriod}'` }
                ],
                mark: {
                type: "text",
                align: "left",
                baseline: "middle",
                dx: 8,
                fontSize: 11,
                fontWeight: "normal"
                },
                encoding: {
                text: { 
                    field: headers[2], 
                    type: "quantitative",
                    format: formatString
                }
                }
            },
            // Right side category labels
            {
                transform: [
                { filter: `datum['${headers[0]}'] == '${lastPeriod}'` }
                ],
                mark: {
                type: "text",
                align: "left",
                baseline: "middle",
                dx: 35,
                fontSize: 10,
                fontWeight: "bold"
                },
                encoding: {
                text: { field: headers[1], type: "nominal" }
                }
            }
            ],
            config: {
            view: { stroke: "transparent" },
            font: "Segoe UI",
            text: { 
                font: "Segoe UI", 
                fontSize: 11, 
                fill: "#605E5C" 
            },
            axis: {
                labelColor: "#605e5c",
                titleColor: "#323130",
                gridColor: "#f3f2f1"
            }
            }
        };
      }

      else if (chartType === "bullet") {
      const data = rows.map(r => ({
        title: r[0],
        ranges: [+r[1], +r[2], +r[3]],
        measures: [+r[4], +r[5]],
        markers: [+r[6]]
      }));

      spec = {
        "$schema": "https://vega.github.io/schema/vega-lite/v6.json",
        background: "white",
        config: { view: { stroke: "transparent" }},
        "data": { "values": data },
        "facet": {
        "row": {
            "field": "title", "type": "ordinal",
            "header": { "labelAngle": 0, "title": "", "labelAlign": "left" }
        }
        },
        "spacing": 10,
        "spec": {
        "encoding": {
            "x": {
            "type": "quantitative",
            "scale": { "nice": false },
            "title": null
            }
        },
        "layer": [
            { "mark": { "type": "bar", "color": "#eee" }, "encoding": { "x": { "field": "ranges[2]" } } },
            { "mark": { "type": "bar", "color": "#ddd" }, "encoding": { "x": { "field": "ranges[1]" } } },
            { "mark": { "type": "bar", "color": "#ccc" }, "encoding": { "x": { "field": "ranges[0]" } } },
            { "mark": { "type": "bar", "color": "lightsteelblue", "size": 10 }, "encoding": { "x": { "field": "measures[1]" } } },
            { "mark": { "type": "bar", "color": "steelblue", "size": 10 }, "encoding": { "x": { "field": "measures[0]" } } },
            { "mark": { "type": "tick", "color": "black" }, "encoding": { "x": { "field": "markers[0]" } } }
        ]
        },
        "resolve": { "scale": { "x": "independent" } },
        "config": { "tick": { "thickness": 2 }, "scale": { "barBandPaddingInner": 0 } }
      };
      }

      else if (chartType === "box") {
      // Expect headers: Category | Value
      const data = rows
        .filter(r => r[0] && !isNaN(+r[1]))
        .map(r => ({
          category: r[0],
          value: +r[1]
        }));

      spec = {
        "$schema": "https://vega.github.io/schema/vega-lite/v6.json",
        "description": "Box plot from Excel selection",
        "data": { "values": data },
        "mark": {
          "type": "boxplot",
          "extent": "min-max"   // show whiskers from min to max
        },
        "encoding": {
          "x": { "field": "category", "type": "nominal" },
          "y": {
            "field": "value",
            "type": "quantitative",
            "scale": { "zero": false }
          },
          "color": {
            "field": "category",
            "type": "nominal",
            "legend": null
          }
        }
      };
      }

      else if (chartType === "horizon") {
        const horizonData = data.map((row, index) => ({
            x: row[headers[0]] || index + 1,
            y: parseFloat(row[headers[1]]) || 0
        }));

        // Calculate data range and bands
        const yValues = horizonData.map(d => d.y);
        const maxY = Math.max(...yValues);
        const minY = Math.min(...yValues);
        const range = maxY - minY;
        
        // Define number of bands (typically 2-4 for horizon graphs)
        const numBands = 3;
        const bandHeight = range / (numBands * 2); // Divide by 2 for positive and negative
        const baseline = minY + range / 2; // Use middle as baseline
        
        // Calculate dynamic dimensions
        const dataPoints = horizonData.length;
        const dynamicWidth = Math.max(300, Math.min(800, dataPoints * 15));

        spec = {
            "$schema": "https://vega.github.io/schema/vega-lite/v6.json",
            "description": "Horizon Graph from Excel selection (IDL methodology)",
            "width": dynamicWidth,
            "height": 60,
            "background": "white",
            "config": { 
            "view": { "stroke": "transparent" },
            "area": {"interpolate": "monotone"}
            },
            "data": { "values": horizonData },
            "encoding": {
            "x": {
                "field": "x",
                "type": headers[0].toLowerCase().includes('date') ? "temporal" : "quantitative",
                "scale": {"zero": false, "nice": false},
                "axis": {
                "title": headers[0],
                "labelFontSize": 10,
                "titleFontSize": 12,
                "labelColor": "#605e5c",
                "titleColor": "#323130",
                "font": "Segoe UI"
                }
            },
            "y": {
                "type": "quantitative",
                "scale": {"domain": [0, bandHeight]},
                "axis": {
                "title": headers[1],
                "orient": "left",
                "labelFontSize": 10,
                "titleFontSize": 12,
                "labelColor": "#605e5c",
                "titleColor": "#323130",
                "font": "Segoe UI",
                "tickCount": 3
                }
            }
            },
            "layer": [
            // Band 1 (lightest positive)
            {
                "transform": [
                {"calculate": `max(0, min(datum.y - ${baseline}, ${bandHeight}))`, "as": "band1"}
                ],
                "mark": {
                "type": "area",
                "clip": true,
                "opacity": 0.3,
                "color": "#4a90e2",
                "interpolate": "monotone"
                },
                "encoding": {
                "y": {"field": "band1"}
                }
            },
            // Band 2 (medium positive)
            {
                "transform": [
                {"calculate": `max(0, min(datum.y - ${baseline} - ${bandHeight}, ${bandHeight}))`, "as": "band2"}
                ],
                "mark": {
                "type": "area",
                "clip": true,
                "opacity": 0.6,
                "color": "#2e7bd6",
                "interpolate": "monotone"
                },
                "encoding": {
                "y": {"field": "band2"}
                }
            },
            // Band 3 (darkest positive)
            {
                "transform": [
                {"calculate": `max(0, datum.y - ${baseline} - ${bandHeight * 2})`, "as": "band3"}
                ],
                "mark": {
                "type": "area",
                "clip": true,
                "opacity": 0.9,
                "color": "#1a5bb8",
                "interpolate": "monotone"
                },
                "encoding": {
                "y": {"field": "band3"}
                }
            },
            // Band -1 (lightest negative, mirrored)
            {
                "transform": [
                {"calculate": `max(0, min(${baseline} - datum.y, ${bandHeight}))`, "as": "nband1"}
                ],
                "mark": {
                "type": "area",
                "clip": true,
                "opacity": 0.3,
                "color": "#e74c3c",
                "interpolate": "monotone"
                },
                "encoding": {
                "y": {"field": "nband1"}
                }
            },
            // Band -2 (medium negative, mirrored)
            {
                "transform": [
                {"calculate": `max(0, min(${baseline} - datum.y - ${bandHeight}, ${bandHeight}))`, "as": "nband2"}
                ],
                "mark": {
                "type": "area",
                "clip": true,
                "opacity": 0.6,
                "color": "#c0392b",
                "interpolate": "monotone"
                },
                "encoding": {
                "y": {"field": "nband2"}
                }
            },
            // Band -3 (darkest negative, mirrored)
            {
                "transform": [
                {"calculate": `max(0, ${baseline} - datum.y - ${bandHeight * 2})`, "as": "nband3"}
                ],
                "mark": {
                "type": "area",
                "clip": true,
                "opacity": 0.9,
                "color": "#a93226",
                "interpolate": "monotone"
                },
                "encoding": {
                "y": {"field": "nband3"}
                }
            }
            ]
        };
      }

      else if (chartType === "tree") {
        const nodes = new Map();

        data.forEach((row, i) => {
            const parent = row[headers[0]] || "";
            const child = row[headers[1]] || `node_${i}`;
            const value = headers.length >= 3 ? (parseFloat(row[headers[2]]) || 1) : 1;
            
            // Add parent node if it doesn't exist and is not empty
            if (parent && !nodes.has(parent)) {
            nodes.set(parent, {
                id: parent,
                parent: "",
                name: parent,
                value: 1
            });
            }
            
            // Add child node
            if (!nodes.has(child)) {
            nodes.set(child, {
                id: child,
                parent: parent,
                name: child,
                value: value
            });
            } else {
            // Update parent and value if child already exists
            const existingNode = nodes.get(child);
            existingNode.parent = parent;
            existingNode.value = value;
            }
        });
        
        // Convert Map to array
        const treeData = Array.from(nodes.values());
        
        // Find root nodes (nodes with no parent or parent not in dataset)
        const allIds = new Set(treeData.map(d => d.id));
        treeData.forEach(node => {
            if (node.parent && !allIds.has(node.parent)) {
            node.parent = ""; // Make it a root node if parent doesn't exist
            }
        });

        // Calculate dynamic dimensions based on data size
        const nodeCount = treeData.length;
        const dynamicWidth = Math.max(600, Math.min(1200, nodeCount * 40));
        const dynamicHeight = Math.max(400, Math.min(1600, nodeCount * 30));

        spec = {
            "$schema": "https://vega.github.io/schema/vega/v6.json",
            "description": "Tree diagram from Excel selection",
            "width": dynamicWidth,
            "height": dynamicHeight,
            "padding": 20,
            "background": "white",
            "config": { "view": { "stroke": "transparent" }},

            "signals": [
            {
                "name": "layout", 
                "value": "tidy"
            },
            {
                "name": "links", 
                "value": "diagonal"
            }
            ],

            "data": [
            {
                "name": "tree",
                "values": treeData,
                "transform": [
                {
                    "type": "stratify",
                    "key": "id",
                    "parentKey": "parent"
                },
                {
                    "type": "tree",
                    "method": {"signal": "layout"},
                    "size": [{"signal": "height - 40"}, {"signal": "width - 100"}],
                    "as": ["y", "x", "depth", "children"]
                }
                ]
            },
            {
                "name": "links",
                "source": "tree",
                "transform": [
                { "type": "treelinks" },
                {
                    "type": "linkpath",
                    "orient": "horizontal",
                    "shape": {"signal": "links"}
                }
                ]
            }
            ],

            "scales": [
            {
                "name": "color",
                "type": "ordinal",
                "range": ["#0078d4", "#00bcf2", "#40e0d0", "#00cc6a", "#10893e", "#107c10", "#bad80a", "#ffb900", "#ff8c00", "#d13438"],
                "domain": {"data": "tree", "field": "depth"}
            },
            {
                "name": "size",
                "type": "linear",
                "range": [100, 400],
                "domain": {"data": "tree", "field": "value"}
            }
            ],

            "marks": [
            {
                "type": "path",
                "from": {"data": "links"},
                "encode": {
                "update": {
                    "path": {"field": "path"},
                    "stroke": {"value": "#8a8886"},
                    "strokeWidth": {"value": 2},
                    "strokeOpacity": {"value": 0.6}
                }
                }
            },
            {
                "type": "symbol",
                "from": {"data": "tree"},
                "encode": {
                "enter": {
                    "stroke": {"value": "#ffffff"},
                    "strokeWidth": {"value": 2}
                },
                "update": {
                    "x": {"field": "x"},
                    "y": {"field": "y"},
                    "size": {"scale": "size", "field": "value"},
                    "fill": {"scale": "color", "field": "depth"},
                    "fillOpacity": {"value": 0.8},
                    "tooltip": {
                    "signal": "{'Name': datum.name, 'ID': datum.id, 'Parent': datum.parent, 'Depth': datum.depth, 'Value': datum.value}"
                    }
                },
                "hover": {
                    "fillOpacity": {"value": 1.0},
                    "strokeWidth": {"value": 3}
                }
                }
            },
            {
                "type": "text",
                "from": {"data": "tree"},
                "encode": {
                "enter": {
                    "fontSize": {"value": 11},
                    "baseline": {"value": "middle"},
                    "font": {"value": "Segoe UI"},
                    "fontWeight": {"value": "bold"}
                },
                "update": {
                    "x": {"field": "x"},
                    "y": {"field": "y"},
                    "text": {"field": "name"},
                    "dx": {"signal": "datum.children ? -12 : 12"},
                    "align": {"signal": "datum.children ? 'right' : 'left'"},
                    "fill": {"value": "#323130"}
                }
                }
            }
            ]
        };
      }

      else if (chartType === "sunburst") {
        const nodes = new Map();
        data.forEach((row, i) => {
            const parent = row[headers[0]] || "";
            const child = row[headers[1]] || `node_${i}`;
            const value = headers.length >= 3 ? (parseFloat(row[headers[2]]) || 1) : 1;
            
            // Add parent node if it doesn't exist and is not empty
            if (parent && !nodes.has(parent)) {
            nodes.set(parent, {
                id: parent,
                parent: "",
                name: parent,
                size: 0 // Will be calculated later
            });
            }
            
            // Add child node
            if (!nodes.has(child)) {
            nodes.set(child, {
                id: child,
                parent: parent,
                name: child,
                size: value
            });
            } else {
            // Update parent and value if child already exists
            const existingNode = nodes.get(child);
            existingNode.parent = parent;
            existingNode.size = value;
            }
        });
        
        // Convert Map to array
        const hierarchicalData = Array.from(nodes.values());
        
        // Find root nodes (nodes with no parent or parent not in dataset)
        const allIds = new Set(hierarchicalData.map(d => d.id));
        hierarchicalData.forEach(node => {
            if (node.parent && !allIds.has(node.parent)) {
            node.parent = ""; // Make it a root node if parent doesn't exist
            }
        });

        // Calculate chart size based on data complexity
        const nodeCount = hierarchicalData.length;
        const chartSize = Math.max(400, Math.min(600, nodeCount * 15 + 300));

        spec = {
            "$schema": "https://vega.github.io/schema/vega/v6.json",
            "description": "Sunburst chart from Excel selection",
            "width": chartSize,
            "height": chartSize,
            "padding": 10,
            "autosize": "none",
            "background": "white",
            "config": { "view": { "stroke": "transparent" }},

            "signals": [
            {
                "name": "centerX",
                "update": "width / 2"
            },
            {
                "name": "centerY", 
                "update": "height / 2"
            },
            {
                "name": "outerRadius",
                "update": "min(width, height) / 2 - 10"
            }
            ],

            "data": [
            {
                "name": "tree",
                "values": hierarchicalData,
                "transform": [
                {
                    "type": "stratify",
                    "key": "id",
                    "parentKey": "parent"
                },
                {
                    "type": "partition",
                    "field": "size",
                    "sort": {"field": "size", "order": "descending"},
                    "size": [{"signal": "2 * PI"}, {"signal": "outerRadius"}],
                    "as": ["a0", "r0", "a1", "r1", "depth", "children"]
                }
                ]
            }
            ],

            "scales": [
            {
                "name": "color",
                "type": "ordinal",
                "domain": {"data": "tree", "field": "depth"},
                "range": [
                "#0078d4", "#00bcf2", "#40e0d0", "#00cc6a", "#10893e", 
                "#107c10", "#bad80a", "#ffb900", "#ff8c00", "#d13438",
                "#8764b8", "#e3008c", "#00b7c3", "#038387", "#486991"
                ]
            },
            {
                "name": "opacity",
                "type": "linear",
                "domain": {"data": "tree", "field": "depth"},
                "range": [0.8, 0.4]
            }
            ],

            "marks": [
            {
                "type": "arc",
                "from": {"data": "tree"},
                "encode": {
                "enter": {
                    "x": {"signal": "centerX"},
                    "y": {"signal": "centerY"},
                    "stroke": {"value": "white"},
                    "strokeWidth": {"value": 1}
                },
                "update": {
                    "startAngle": {"field": "a0"},
                    "endAngle": {"field": "a1"},
                    "innerRadius": {"field": "r0"},
                    "outerRadius": {"field": "r1"},
                    "fill": {"scale": "color", "field": "depth"},
                    "fillOpacity": {"scale": "opacity", "field": "depth"}
                }
                }
            },
            {
                "type": "text",
                "from": {"data": "tree"},
                "encode": {
                "enter": {
                    "x": {"signal": "centerX"},
                    "y": {"signal": "centerY"},
                    "radius": {"signal": "(datum.r0 + datum.r1) / 2"},
                    "theta": {"signal": "(datum.a0 + datum.a1) / 2"},
                    "fill": {"value": "#323130"},
                    "font": {"value": "Segoe UI"},
                    "fontSize": {"value": 10},
                    "fontWeight": {"value": "bold"},
                    "align": {"value": "center"},
                    "baseline": {"value": "middle"}
                },
                "update": {
                    "text": {
                    "signal": "(datum.r1 - datum.r0) > 20 && (datum.a1 - datum.a0) > 0.3 ? datum.name : ''"
                    },
                    "opacity": {
                    "signal": "(datum.r1 - datum.r0) > 20 && (datum.a1 - datum.a0) > 0.3 ? 1 : 0"
                    }
                }
                }
            }
            ]
        };
      }

      else if (chartType === "radar") {
        const radarData = [];
        const dimensions = headers.slice(1); // All columns except first are dimensions
        
        data.forEach((row, seriesIndex) => {
            const seriesName = row[headers[0]] || `Series ${seriesIndex + 1}`;
            
            dimensions.forEach(dimension => {
            const value = parseFloat(row[dimension]) || 0;
            radarData.push({
                series: seriesName,
                dimension: dimension,
                value: value,
                category: seriesIndex
            });
            });
        });

        // Get unique dimensions for grid
        const uniqueDimensions = [...new Set(radarData.map(d => d.dimension))];

        spec = {
            "$schema": "https://vega.github.io/schema/vega/v6.json",
            "description": "Radar chart from Excel selection",
            "width": 400,
            "height": 400,
            "padding": 60,
            "autosize": {"type": "none", "contains": "padding"},
            "background": "white",
            "config": { "view": { "stroke": "transparent" }},

            "signals": [
            {"name": "radius", "update": "width / 2"}
            ],

            "data": [
            {
                "name": "table",
                "values": radarData
            },
            {
                "name": "dimensions",
                "values": uniqueDimensions.map(d => ({dimension: d}))
            }
            ],

            "scales": [
            {
                "name": "angular",
                "type": "point",
                "range": {"signal": "[-PI, PI]"},
                "padding": 0.5,
                "domain": uniqueDimensions
            },
            {
                "name": "radial",
                "type": "linear",
                "range": {"signal": "[0, radius]"},
                "zero": true,
                "nice": true,
                "domain": {"data": "table", "field": "value"},
                "domainMin": 0
            },
            {
                "name": "color",
                "type": "ordinal",
                "domain": {"data": "table", "field": "category"},
                "range": [
                "#0078d4", "#00bcf2", "#40e0d0", "#00cc6a", "#10893e",
                "#107c10", "#bad80a", "#ffb900", "#ff8c00", "#d13438"
                ]
            }
            ],

            "encode": {
            "enter": {
                "x": {"signal": "radius"},
                "y": {"signal": "radius"}
            }
            },

            "marks": [
            {
                "type": "group",
                "name": "categories",
                "zindex": 1,
                "from": {
                "facet": {"data": "table", "name": "facet", "groupby": ["category", "series"]}
                },
                "marks": [
                {
                    "type": "line",
                    "name": "category-line",
                    "from": {"data": "facet"},
                    "encode": {
                    "enter": {
                        "interpolate": {"value": "linear-closed"},
                        "x": {"signal": "scale('radial', datum.value) * cos(scale('angular', datum.dimension))"},
                        "y": {"signal": "scale('radial', datum.value) * sin(scale('angular', datum.dimension))"},
                        "stroke": {"scale": "color", "field": "category"},
                        "strokeWidth": {"value": 2},
                        "fill": {"scale": "color", "field": "category"},
                        "fillOpacity": {"value": 0.1},
                        "strokeOpacity": {"value": 0.8}
                    }
                    }
                },
                {
                    "type": "symbol",
                    "name": "category-points",
                    "from": {"data": "facet"},
                    "encode": {
                    "enter": {
                        "x": {"signal": "scale('radial', datum.value) * cos(scale('angular', datum.dimension))"},
                        "y": {"signal": "scale('radial', datum.value) * sin(scale('angular', datum.dimension))"},
                        "size": {"value": 50},
                        "fill": {"scale": "color", "field": "category"},
                        "stroke": {"value": "white"},
                        "strokeWidth": {"value": 1}
                    }
                    }
                }
                ]
            },
            {
                "type": "rule",
                "name": "radial-grid",
                "from": {"data": "dimensions"},
                "zindex": 0,
                "encode": {
                "enter": {
                    "x": {"value": 0},
                    "y": {"value": 0},
                    "x2": {"signal": "radius * cos(scale('angular', datum.dimension))"},
                    "y2": {"signal": "radius * sin(scale('angular', datum.dimension))"},
                    "stroke": {"value": "#e1e4e8"},
                    "strokeWidth": {"value": 1}
                }
                }
            },
            {
                "type": "text",
                "name": "dimension-label",
                "from": {"data": "dimensions"},
                "zindex": 1,
                "encode": {
                "enter": {
                    "x": {"signal": "(radius + 20) * cos(scale('angular', datum.dimension))"},
                    "y": {"signal": "(radius + 20) * sin(scale('angular', datum.dimension))"},
                    "text": {"field": "dimension"},
                    "align": [
                    {
                        "test": "abs(scale('angular', datum.dimension)) > PI / 2",
                        "value": "right"
                    },
                    {
                        "value": "left"
                    }
                    ],
                    "baseline": [
                    {
                        "test": "scale('angular', datum.dimension) > 0", 
                        "value": "top"
                    },
                    {
                        "test": "scale('angular', datum.dimension) == 0", 
                        "value": "middle"
                    },
                    {
                        "value": "bottom"
                    }
                    ],
                    "fill": {"value": "#323130"},
                    "fontWeight": {"value": "bold"},
                    "font": {"value": "Segoe UI"},
                    "fontSize": {"value": 12}
                }
                }
            },
            {
                "type": "line",
                "name": "outer-line",
                "from": {"data": "radial-grid"},
                "encode": {
                "enter": {
                    "interpolate": {"value": "linear-closed"},
                    "x": {"field": "x2"},
                    "y": {"field": "y2"},
                    "stroke": {"value": "#8a8886"},
                    "strokeWidth": {"value": 2},
                    "strokeOpacity": {"value": 0.6}
                }
                }
            }
            ]
        };
      }

      else if (chartType === "waterfall") {
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
        const dynamicWidth = Math.max(400, Math.min(1600, numDataPoints * 50));
        const maxAmount = Math.max(...data.map(d => Math.abs(d[headers[1]])));
        const dynamicHeight = Math.max(300, Math.min(600, maxAmount / 100 + 200));

        spec = {
            $schema: "https://vega.github.io/schema/vega-lite/v6.json",
            description: "Waterfall chart with multiple subtotals",
            background: "white",
            data: { values: processedData },
            config: { view: { stroke: "transparent" }},
            width: dynamicWidth,
            height: dynamicHeight,
            transform: [
            { "window": [{ "op": "sum", "field": headers[1], "as": "sum" }] },
            { "window": [{ "op": "lead", "field": headers[0], "as": "lead" }] },
            {
                "calculate": `datum.lead === null ? datum.${headers[0]} : datum.lead`,
                "as": "lead"
            },
            {
                // If total → reset, else → running sum step
                "calculate": `datum.${headers[2]} == 'total' ? 0 : datum.sum - datum.${headers[1]}`,
                "as": "previous_sum"
            },
            {
                "calculate": `datum.${headers[2]} == 'total' ? datum.sum : datum.${headers[1]}`,
                "as": "amount"
            },
            {
                "calculate": `datum.${headers[2]} == 'total' ? datum.${headers[1]} / 2 : (datum.sum + datum.previous_sum) / 2`,
                "as": "center"
            },
            {
                "calculate": `datum.${headers[2]} == 'total' ? datum.sum : (datum.${headers[1]} > 0 ? '+' : '') + datum.${headers[1]}`,
                "as": "text_amount"
            },
            { "calculate": "(datum.sum + datum.previous_sum) / 2", "as": "center" },

            // Add group index for stacked handling
            {
                "window": [{ "op": "rank", "as": "group_index" }],
                "frame": [null, null],
                "groupby": [headers[0]]
            },

            // Precompute color shades
            {
                "calculate": `
                datum.${headers[2]} == 'total'
                    ? '#00B0F0'
                    : datum.amount >= 0
                    ? (datum.group_index == 1 ? '#70AD47'
                        : (datum.group_index == 2 ? '#8BC97A'
                        : (datum.group_index == 3 ? '#A7DA9D'
                        : '#C3EBC0')))
                    : (datum.group_index == 1 ? '#E15759'
                        : (datum.group_index == 2 ? '#EC7A7C'
                        : (datum.group_index == 3 ? '#F29C9D'
                        : '#F8BEBF')))
                `,
                "as": "bar_color"
            }
            ],
            encoding: {
            x: {
                field: headers[0],
                type: "ordinal",
                sort: null,
                axis: { labelAngle: -45, title: null },
                scale: { paddingInner: 0.05, paddingOuter: 0.025 }
            }
            },
            layer: [
            {
                mark: { type: "bar", size: 60},
                encoding: {
                y: { field: "previous_sum", type: "quantitative", title: null },
                y2: { field: "sum" },
                color: { field: "bar_color", type: "nominal", scale: null }
                }
            },
            {
                mark: { type: "text", fontWeight: "bold", baseline: "middle" },
                encoding: {
                y: { field: "center", type: "quantitative" },
                text: { field: "text_amount", type: "nominal" },
                color: {
                    condition: [
                    { test: `datum.${headers[2]} == 'total'`, value: "#725a30" }
                    ],
                    value: "white"
                }
                }
            }
            ],
            config: { text: { fontWeight: "bold", color: "#D9D9D9" } }
        };
      }

      else if (chartType === "pie") {
        spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        background: "white",
        config: { view: { stroke: "transparent" }},
        description: "Pie chart from Excel selection",
        data: { values: data },
        mark: { type: "arc", outerRadius: 120 },
        encoding: {
          theta: { field: headers[1], type: "quantitative" },
          color: { field: headers[0], type: "nominal" }
          }
        };
      }

      else if (chartType === "histogram") {
        // Expect a single numeric column
        const numericData = rows
          .filter(r => !isNaN(+r[0]))
          .map(r => ({ value: +r[0] }));

        // Calculate data range for better binning control
        const values = numericData.map(d => d.value);
        const minVal = Math.min(...values);
        const maxVal = Math.max(...values);
        const range = maxVal - minVal;
        
        // Calculate nice bin boundaries
        const binCount = 20;
        const binWidth = range / binCount;
        const niceMin = Math.floor(minVal / binWidth) * binWidth;
        const niceMax = Math.ceil(maxVal / binWidth) * binWidth;

        spec = {
          "$schema": "https://vega.github.io/schema/vega-lite/v6.json",
          "description": "Histogram from Excel selection",
          "background": "white",
          "config": { "view": { "stroke": "transparent" }},
          "data": { "values": numericData },
          "mark": {
            "type": "bar",
            "tooltip": true,
            "stroke": "white",
            "strokeWidth": 1
          },
          "encoding": {
            "x": {
              "field": "value",
              "bin": { 
                "extent": [niceMin, niceMax],
                "step": binWidth,
                "nice": false  // Prevent Vega from adjusting our nice boundaries
              },
              "type": "quantitative",
              "axis": { 
                "title": "Value",
                "labelFontSize": 12,
                "titleFontSize": 14,
                "labelColor": "#605e5c",
                "titleColor": "#323130"
              },
              "scale": {
                "domain": [niceMin, niceMax],
                "range": "width",
                "paddingInner": 0.05,
                "paddingOuter": 0.02
              }
            },
            "y": {
              "aggregate": "count",
              "type": "quantitative",
              "axis": { 
                "title": "Count",
                "labelFontSize": 12,
                "titleFontSize": 14,
                "labelColor": "#605e5c",
                "titleColor": "#323130",
                "gridColor": "#f3f2f1"
              }
            },
            "color": {
              "value": "#0078d4"
            }
          },
          "config": {
            "font": "Segoe UI",
            "axis": {
              "labelColor": "#605e5c",
              "titleColor": "#323130",
              "gridColor": "#f3f2f1"
            }
          }
        };
      }

      else if (chartType === "candlestick") {
        // Helper function to convert Excel dates to JS dates
        function excelDateToJSDate(serial) {
          if (typeof serial === 'number') {
            return new Date(Math.round((serial - 25569) * 86400 * 1000));
          }
          return new Date(serial);
        }

        // Process and validate data - SKIP ROWS WITH MISSING VALUES
        const candlestickData = data
          .map((row, index) => {
            // Skip if any required value is missing/null/empty
            if (!row[headers[0]] || 
                row[headers[1]] == null || row[headers[1]] === "" ||
                row[headers[2]] == null || row[headers[2]] === "" ||
                row[headers[3]] == null || row[headers[3]] === "" ||
                row[headers[4]] == null || row[headers[4]] === "") {
              return null;
            }

            const date = excelDateToJSDate(row[headers[0]]);
            const open = parseFloat(row[headers[1]]);
            const high = parseFloat(row[headers[2]]);
            const low = parseFloat(row[headers[3]]);
            const close = parseFloat(row[headers[4]]);
            
            if (isNaN(date.getTime()) || isNaN(open) || isNaN(high) || isNaN(low) || isNaN(close)) {
              return null;
            }
            
            return {
              date: date.toISOString(),
              open: open,
              high: high,
              low: low,
              close: close
            };
          })
          .filter(Boolean); // Remove null entries

        if (candlestickData.length === 0) {
          console.warn("No valid candlestick data found");
          return;
        }

        // Calculate dynamic width based on number of data points
        const dataPoints = candlestickData.length;
        const minWidth = 500;
        const maxWidth = 1400;
        const widthPerPoint = 18; // Pixels per candlestick
        const dynamicWidth = Math.max(minWidth, Math.min(maxWidth, dataPoints * widthPerPoint));

        // Format dates for display (keep original date for sorting)
        const formattedData = candlestickData.map(d => ({
          ...d,
          dateDisplay: new Date(d.date).toLocaleDateString('en-US', { 
            month: dataPoints > 50 ? 'numeric' : 'short', 
            day: 'numeric',
            year: dataPoints > 100 ? undefined : '2-digit'
          })
        }));

        spec = {
          "$schema": "https://vega.github.io/schema/vega-lite/v6.json",
          "width": dynamicWidth,
          "height": 400,
          "description": "Candlestick chart V4 from Excel selection",
          "background": "white",
          "padding": {"left": 10, "right": 10, "top": 10, "bottom": 10},
          "config": { "view": { "stroke": "transparent" }},
          "data": { "values": formattedData },
          "encoding": {
            "x": {
              "field": "dateDisplay",
              "type": "ordinal", // Ordinal scale eliminates gaps for missing dates
              "title": "Date",
              "axis": {
                "labelAngle": dataPoints > 15 ? -45 : 0,
                "labelFontSize": 10,
                "titleFontSize": 12,
                "labelColor": "#605e5c",
                "titleColor": "#323130",
                "font": "Segoe UI",
                "labelLimit": 100,
                "labelOverlap": false,
                "titlePadding": 5,
                "labelPadding": 3
              },
              "scale": {
                "type": "band",
                "paddingInner": 0.1, // 10% gap between bars
                "paddingOuter": 0.05 // 5% padding at edges
              }
            },
            "y": {
              "type": "quantitative",
              "scale": { "zero": false },
              "axis": {
                "title": "Price",
                "labelFontSize": 11,
                "titleFontSize": 12,
                "labelColor": "#605e5c",
                "titleColor": "#323130",
                "font": "Segoe UI",
                "grid": true,
                "gridColor": "#f3f2f1",
                "titlePadding": 5,
                "labelPadding": 3
              }
            },
            "color": {
              "condition": {
                "test": "datum.open < datum.close",
                "value": "#06982d" // Green for bullish (up) candles
              },
              "value": "#ae1325" // Red for bearish (down) candles
            }
          },
          "layer": [
            {
              "mark": {
                "type": "rule",
                "tooltip": true,
                "strokeWidth": 1
              },
              "encoding": {
                "y": { "field": "low" },
                "y2": { "field": "high" },
                "tooltip": [
                  { "field": "date", "type": "temporal", "title": "Date", "format": "%Y-%m-%d" },
                  { "field": "open", "type": "quantitative", "title": "Open", "format": ".2f" },
                  { "field": "high", "type": "quantitative", "title": "High", "format": ".2f" },
                  { "field": "low", "type": "quantitative", "title": "Low", "format": ".2f" },
                  { "field": "close", "type": "quantitative", "title": "Close", "format": ".2f" }
                ]
              }
            },
            {
              "mark": {
                "type": "bar",
                "tooltip": true,
                "stroke": "white",
                "strokeWidth": 0.5
              },
              "encoding": {
                "y": { "field": "open" },
                "y2": { "field": "close" },
                "tooltip": [
                  { "field": "date", "type": "temporal", "title": "Date", "format": "%Y-%m-%d" },
                  { "field": "open", "type": "quantitative", "title": "Open", "format": ".2f" },
                  { "field": "high", "type": "quantitative", "title": "High", "format": ".2f" },
                  { "field": "low", "type": "quantitative", "title": "Low", "format": ".2f" },
                  { "field": "close", "type": "quantitative", "title": "Close", "format": ".2f" }
                ]
              }
            }
          ],
          "config": {
            "font": "Segoe UI",
            "axis": {
              "labelColor": "#605e5c",
              "titleColor": "#323130",
              "gridColor": "#f3f2f1"
            }
          }
        };
      }

      else if (chartType === "map") {
      // Expect headers: Country (ISO3) | Value
      const isoToId = {
        "AFG": 4,    // Afghanistan
        "AGO": 24,   // Angola
        "ALB": 8,    // Albania
        "AND": 20,   // Andorra
        "ARE": 784,  // United Arab Emirates
        "ARG": 32,   // Argentina
        "ARM": 51,   // Armenia
        "ATA": 10,   // Antarctica
        "ATG": 28,   // Antigua and Barbuda
        "AUS": 36,   // Australia
        "AUT": 40,   // Austria
        "AZE": 31,   // Azerbaijan
        "BDI": 108,  // Burundi
        "BEL": 56,   // Belgium
        "BEN": 204,  // Benin
        "BFA": 854,  // Burkina Faso
        "BGD": 50,   // Bangladesh
        "BGR": 100,  // Bulgaria
        "BHR": 48,   // Bahrain
        "BHS": 44,   // Bahamas
        "BIH": 70,   // Bosnia and Herzegovina
        "BLR": 112,  // Belarus
        "BLZ": 84,   // Belize
        "BOL": 68,   // Bolivia
        "BRA": 76,   // Brazil
        "BRB": 52,   // Barbados
        "BRN": 96,   // Brunei
        "BTN": 64,   // Bhutan
        "BWA": 72,   // Botswana
        "CAF": 140,  // Central African Republic
        "CAN": 124,  // Canada
        "CHE": 756,  // Switzerland
        "CHL": 152,  // Chile
        "CHN": 156,  // China
        "CIV": 384,  // Côte d'Ivoire
        "CMR": 120,  // Cameroon
        "COD": 180,  // Democratic Republic of Congo
        "COG": 178,  // Congo
        "COL": 170,  // Colombia
        "COM": 174,  // Comoros
        "CPV": 132,  // Cape Verde
        "CRI": 188,  // Costa Rica
        "CUB": 192,  // Cuba
        "CYP": 196,  // Cyprus
        "CZE": 203,  // Czechia
        "DEU": 276,  // Germany
        "DJI": 262,  // Djibouti
        "DMA": 212,  // Dominica
        "DNK": 208,  // Denmark
        "DOM": 214,  // Dominican Republic
        "DZA": 12,   // Algeria
        "ECU": 218,  // Ecuador
        "EGY": 818,  // Egypt
        "ERI": 232,  // Eritrea
        "ESH": 732,  // Western Sahara
        "ESP": 724,  // Spain
        "EST": 233,  // Estonia
        "ETH": 231,  // Ethiopia
        "FIN": 246,  // Finland
        "FJI": 242,  // Fiji
        "FRA": 250,  // France
        "FSM": 583,  // Micronesia
        "GAB": 266,  // Gabon
        "GBR": 826,  // United Kingdom
        "GEO": 268,  // Georgia
        "GHA": 288,  // Ghana
        "GIN": 324,  // Guinea
        "GMB": 270,  // Gambia
        "GNB": 624,  // Guinea-Bissau
        "GNQ": 226,  // Equatorial Guinea
        "GRC": 300,  // Greece
        "GRD": 308,  // Grenada
        "GRL": 304,  // Greenland
        "GTM": 320,  // Guatemala
        "GUY": 328,  // Guyana
        "HND": 340,  // Honduras
        "HRV": 191,  // Croatia
        "HTI": 332,  // Haiti
        "HUN": 348,  // Hungary
        "IDN": 360,  // Indonesia
        "IND": 356,  // India
        "IRL": 372,  // Ireland
        "IRN": 364,  // Iran
        "IRQ": 368,  // Iraq
        "ISL": 352,  // Iceland
        "ISR": 376,  // Israel
        "ITA": 380,  // Italy
        "JAM": 388,  // Jamaica
        "JOR": 400,  // Jordan
        "JPN": 392,  // Japan
        "KAZ": 398,  // Kazakhstan
        "KEN": 404,  // Kenya
        "KGZ": 417,  // Kyrgyzstan
        "KHM": 116,  // Cambodia
        "KIR": 296,  // Kiribati
        "KNA": 659,  // Saint Kitts and Nevis
        "KOR": 410,  // South Korea
        "KWT": 414,  // Kuwait
        "LAO": 418,  // Laos
        "LBN": 422,  // Lebanon
        "LBR": 430,  // Liberia
        "LBY": 434,  // Libya
        "LCA": 662,  // Saint Lucia
        "LIE": 438,  // Liechtenstein
        "LKA": 144,  // Sri Lanka
        "LSO": 426,  // Lesotho
        "LTU": 440,  // Lithuania
        "LUX": 442,  // Luxembourg
        "LVA": 428,  // Latvia
        "MAR": 504,  // Morocco
        "MCO": 492,  // Monaco
        "MDA": 498,  // Moldova
        "MDG": 450,  // Madagascar
        "MDV": 462,  // Maldives
        "MEX": 484,  // Mexico
        "MHL": 584,  // Marshall Islands
        "MKD": 807,  // North Macedonia
        "MLI": 466,  // Mali
        "MLT": 470,  // Malta
        "MMR": 104,  // Myanmar
        "MNE": 499,  // Montenegro
        "MNG": 496,  // Mongolia
        "MOZ": 508,  // Mozambique
        "MRT": 478,  // Mauritania
        "MUS": 480,  // Mauritius
        "MWI": 454,  // Malawi
        "MYS": 458,  // Malaysia
        "NAM": 516,  // Namibia
        "NCL": 540,  // New Caledonia
        "NER": 562,  // Niger
        "NGA": 566,  // Nigeria
        "NIC": 558,  // Nicaragua
        "NLD": 528,  // Netherlands
        "NOR": 578,  // Norway
        "NPL": 524,  // Nepal
        "NRU": 520,  // Nauru
        "NZL": 554,  // New Zealand
        "OMN": 512,  // Oman
        "PAK": 586,  // Pakistan
        "PAN": 591,  // Panama
        "PER": 604,  // Peru
        "PHL": 608,  // Philippines
        "PLW": 585,  // Palau
        "PNG": 598,  // Papua New Guinea
        "POL": 616,  // Poland
        "PRI": 630,  // Puerto Rico
        "PRK": 408,  // North Korea
        "PRT": 620,  // Portugal
        "PRY": 600,  // Paraguay
        "PSE": 275,  // Palestine
        "QAT": 634,  // Qatar
        "ROU": 642,  // Romania
        "RUS": 643,  // Russia
        "RWA": 646,  // Rwanda
        "SAU": 682,  // Saudi Arabia
        "SDN": 729,  // Sudan
        "SEN": 686,  // Senegal
        "SGP": 702,  // Singapore
        "SLB": 90,   // Solomon Islands
        "SLE": 694,  // Sierra Leone
        "SLV": 222,  // El Salvador
        "SMR": 674,  // San Marino
        "SOM": 706,  // Somalia
        "SRB": 688,  // Serbia
        "SSD": 728,  // South Sudan
        "STP": 678,  // São Tomé and Príncipe
        "SUR": 740,  // Suriname
        "SVK": 703,  // Slovakia
        "SVN": 705,  // Slovenia
        "SWE": 752,  // Sweden
        "SWZ": 748,  // Eswatini
        "SYC": 690,  // Seychelles
        "SYR": 760,  // Syria
        "TCD": 148,  // Chad
        "TGO": 768,  // Togo
        "THA": 764,  // Thailand
        "TJK": 762,  // Tajikistan
        "TKM": 795,  // Turkmenistan
        "TLS": 626,  // Timor-Leste
        "TON": 776,  // Tonga
        "TTO": 780,  // Trinidad and Tobago
        "TUN": 788,  // Tunisia
        "TUR": 792,  // Turkey
        "TUV": 798,  // Tuvalu
        "TWN": 158,  // Taiwan
        "TZA": 834,  // Tanzania
        "UGA": 800,  // Uganda
        "UKR": 804,  // Ukraine
        "URY": 858,  // Uruguay
        "USA": 840,  // United States
        "UZB": 860,  // Uzbekistan
        "VAT": 336,  // Vatican City
        "VCT": 670,  // Saint Vincent and the Grenadines
        "VEN": 862,  // Venezuela
        "VNM": 704,  // Vietnam
        "VUT": 548,  // Vanuatu
        "WSM": 882,  // Samoa
        "XKX": 383,  // Kosovo
        "YEM": 887,  // Yemen
        "ZAF": 710,  // South Africa
        "ZMB": 894,  // Zambia
        "ZWE": 716   // Zimbabwe
      };

      const worldData = rows
        .filter(r => r[0] && !isNaN(+r[1]))
        .map(r => {
          const iso = (r[0] || "").toUpperCase().trim();
          const idVal = isoToId[iso];
          return {
            id: idVal,     // numeric ID matching topojson country.id
            iso: iso,       // original ISO3 for tooltip
            rate: +r[1]
          };
        })
        .filter(d => d.id); // drop rows where iso isn't in lookup

      spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        width: 800,
        height: 450,
        data: {
          url: "https://cdn.jsdelivr.net/npm/world-atlas@2/countries-110m.json",
          format: { type: "topojson", feature: "countries" }
        },
        transform: [
          {
            lookup: "id",
            from: {
              data: { values: worldData },
              key: "id",
              fields: ["rate", "iso"]
            }
          }
        ],
        projection: { type: "equalEarth" },
        mark: { type: "geoshape", stroke: "white", strokeWidth: 0.5 },
        encoding: {
          color: {
            field: "rate",
            type: "quantitative",
            scale: { scheme: "blues" }
          },
          tooltip: [
            { field: "iso", type: "nominal", title: "Country (ISO3)" },
            { field: "rate", type: "quantitative", title: "Value" }
          ]
        }
      };
      }

      else if (chartType === "mekko") {
        spec = {
          $schema: "https://vega.github.io/schema/vega/v5.json",
          description: "Marimekko chart from Excel selection",
          width: 800,
          height: 500,
          background: "white",
          config: { view: { stroke: "transparent" }},
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

      if (chartType === "marimekko") {
      if (headers.length < 3) {
        console.warn("Marimekko chart requires at least 3 columns: Category, Subcategory, Value");
        return;
      }

      spec = {
        "$schema": "https://vega.github.io/schema/vega/v5.json",
        "description": "Marimekko Chart from Excel selection",
        "width": 600,
        "height": 400,
        "background": "white",
        "config": { "view": { "stroke": "transparent" }},
        "padding": { "top": 30, "bottom": 60, "left": 60, "right": 60 },
        "data": [
          {
            "name": "table",
            "values": data,
            "transform": [
              {
                "type": "formula",
                "as": "Category",
                "expr": `datum['${headers[0]}']`
              },
              {
                "type": "formula", 
                "as": "Subcategory",
                "expr": `datum['${headers[1]}']`
              },
              {
                "type": "formula",
                "as": "Value", 
                "expr": `datum['${headers[2]}']`
              }
            ]
          },
          {
            "name": "categories",
            "source": "table",
            "transform": [
              {
                "type": "aggregate",
                "fields": ["Value"],
                "ops": ["sum"],
                "as": ["categoryTotal"],
                "groupby": ["Category"]
              },
              {
                "type": "stack",
                "offset": "normalize",
                "sort": { "field": "categoryTotal", "order": "descending" },
                "field": "categoryTotal",
                "as": ["x0", "x1"]
              },
              {
                "type": "formula",
                "as": "Percent",
                "expr": "datum.x1 - datum.x0"
              }
            ]
          },
          {
            "name": "finalTable",
            "source": "table",
            "transform": [
              {
                "type": "stack",
                "offset": "normalize",
                "groupby": ["Category"],
                "sort": { "field": "Value", "order": "descending" },
                "field": "Value",
                "as": ["y0", "y1"]
              },
              {
                "type": "lookup",
                "from": "categories",
                "key": "Category",
                "values": ["x0", "x1"],
                "fields": ["Category"]
              },
              {
                "type": "formula",
                "as": "Percent",
                "expr": "datum.y1 - datum.y0"
              }
            ]
          }
        ],
        "scales": [
          {
            "name": "x",
            "type": "linear",
            "range": "width",
            "domain": { "data": "finalTable", "field": "x1" }
          },
          {
            "name": "y",
            "type": "linear",
            "range": "height",
            "nice": false,
            "zero": true,
            "domain": { "data": "finalTable", "field": "y1" }
          },
          {
            "name": "color",
            "type": "ordinal",
            "range": { "scheme": "category20" },
            "domain": {
              "data": "categories",
              "field": "Category",
              "sort": { "field": "x0", "order": "ascending", "op": "sum" }
            }
          }
        ],
        "axes": [
          {
            "orient": "left",
            "scale": "y",
            "format": "%",
            "tickCount": 5,
            "labelColor": "#333333",
            "labelFontSize": 11,
            "domain": false
          },
          {
            "orient": "bottom",
            "scale": "x",
            "format": "%",
            "tickCount": 5,
            "labelColor": "#333333", 
            "labelFontSize": 11,
            "domain": false
          }
        ],
        "marks": [
          {
            "type": "rect",
            "name": "bars",
            "from": { "data": "finalTable" },
            "encode": {
              "update": {
                "x": { "scale": "x", "field": "x0" },
                "x2": { "scale": "x", "field": "x1" },
                "y": { "scale": "y", "field": "y0" },
                "y2": { "scale": "y", "field": "y1" },
                "fill": { "scale": "color", "field": "Category" },
                "stroke": { "value": "white" },
                "strokeWidth": { "value": 1 },
                "opacity": { "value": 0.8 },
                "tooltip": { 
                  "signal": "{'Category': datum.Category, 'Subcategory': datum.Subcategory, 'Value': datum.Value, 'Percentage': format(datum.Percent, '.1%')}" 
                }
              },
              "hover": {
                "opacity": { "value": 1.0 }
              }
            }
          },
          {
            "type": "text",
            "name": "valueLabels",
            "from": { "data": "finalTable" },
            "encode": {
              "update": {
                "x": { "scale": "x", "signal": "(datum.x1 - datum.x0)/2 + datum.x0" },
                "y": { "scale": "y", "signal": "(datum.y1 - datum.y0)/2 + datum.y0" },
                "text": { 
                  "signal": "datum.Percent > 0.027 ? [datum.Subcategory, format(datum.Value, ',.0f') + ' (' + format(datum.Percent, '.0%') + ')'] : []" 
                },
                "align": { "value": "center" },
                "baseline": { "value": "middle" },
                "fill": { "value": "white" },
                "fontSize": { "value": 10 },
                "fontWeight": { "value": "normal" },
                "font": { "value": "Segoe UI" },
                "lineHeight": { "value": 12 },
                "opacity": { "signal": "datum.Percent > 0.027 ? 1 : 0" }
              }
            }
          }
        ]
      };
      }

      else if (chartType === "arc") {
        // Transform Excel data for arc chart
        const edges = data.map((row, index) => ({
          source: row[headers[0]],
          target: row[headers[1]],
          value: headers.length >= 3 && row[headers[2]] ? row[headers[2]] : 1,
          group: headers.length >= 4 && row[headers[3]] ? row[headers[3]] : "default"
        }));

        // Get unique nodes from edges
        const nodeMap = new Map();
        edges.forEach(edge => {
          if (!nodeMap.has(edge.source)) {
            nodeMap.set(edge.source, { 
              name: edge.source, 
              group: edge.group,
              index: nodeMap.size
            });
          }
          if (!nodeMap.has(edge.target)) {
            nodeMap.set(edge.target, { 
              name: edge.target, 
              group: edge.group,
              index: nodeMap.size
            });
          }
        });

        const nodes = Array.from(nodeMap.values());

        // Transform edges to use node indices
        const edgesWithIndices = edges.map(edge => ({
          source: nodeMap.get(edge.source).index,
          target: nodeMap.get(edge.target).index,
          value: edge.value
        }));

        spec = {
          $schema: "https://vega.github.io/schema/vega/v5.json",
          description: "Arc diagram from Excel selection",
          width: Math.max(600, nodes.length * 40),
          height: 300,
          padding: { top: 20, bottom: 80, left: 20, right: 20 },
          background: "white",
          config: { view: { stroke: "transparent" }},
          data: [
            {
              name: "edges",
              values: edgesWithIndices
            },
            {
              name: "sourceDegree",
              source: "edges",
              transform: [
                { type: "aggregate", groupby: ["source"], as: ["count"] }
              ]
            },
            {
              name: "targetDegree", 
              source: "edges",
              transform: [
                { type: "aggregate", groupby: ["target"], as: ["count"] }
              ]
            },
            {
              name: "nodes",
              values: nodes,
              transform: [
                { type: "window", ops: ["rank"], as: ["order"] },
                {
                  type: "lookup", from: "sourceDegree", key: "source",
                  fields: ["index"], as: ["sourceDegree"],
                  default: { count: 0 }
                },
                {
                  type: "lookup", from: "targetDegree", key: "target", 
                  fields: ["index"], as: ["targetDegree"],
                  default: { count: 0 }
                },
                {
                  type: "formula", as: "degree",
                  expr: "(datum.sourceDegree.count || 0) + (datum.targetDegree.count || 0)"
                }
              ]
            }
          ],

          scales: [
            {
              name: "position",
              type: "band",
              domain: { data: "nodes", field: "order", sort: true },
              range: "width"
            },
            {
              name: "color",
              type: "ordinal",
              range: { scheme: "category20" },
              domain: { data: "nodes", field: "group" }
            }
          ],

          marks: [
            {
              type: "symbol",
              name: "layout",
              interactive: false,
              from: { data: "nodes" },
              encode: {
                enter: { opacity: { value: 0 } },
                update: {
                  x: { scale: "position", field: "order" },
                  y: { value: 0 },
                  size: { field: "degree", mult: 8, offset: 50 },
                  fill: { scale: "color", field: "group" }
                }
              }
            },
            {
              type: "path",
              from: { data: "edges" },
              encode: {
                update: {
                  stroke: { value: "#0078d4" },
                  strokeOpacity: { value: 0.4 },
                  strokeWidth: { field: "value", mult: 2, offset: 1 }
                }
              },
              transform: [
                {
                  type: "lookup", from: "layout", key: "datum.index",
                  fields: ["datum.source", "datum.target"],
                  as: ["sourceNode", "targetNode"]
                },
                {
                  type: "linkpath",
                  sourceX: { expr: "min(datum.sourceNode.x, datum.targetNode.x)" },
                  targetX: { expr: "max(datum.sourceNode.x, datum.targetNode.x)" },
                  sourceY: { expr: "0" },
                  targetY: { expr: "0" },
                  shape: "arc"
                }
              ]
            },
            {
              type: "symbol",
              from: { data: "layout" },
              encode: {
                update: {
                  x: { field: "x" },
                  y: { field: "y" },
                  fill: { field: "fill" },
                  size: { field: "size" },
                  stroke: { value: "white" },
                  strokeWidth: { value: 1 },
                  tooltip: { 
                    signal: "{'Node': datum.datum.name, 'Group': datum.datum.group, 'Connections': datum.datum.degree}" 
                  }
                }
              }
            },
            {
              type: "text",
              from: { data: "nodes" },
              encode: {
                update: {
                  x: { scale: "position", field: "order" },
                  y: { value: 25 },
                  fontSize: { value: 10 },
                  align: { value: "center" },
                  baseline: { value: "top" },
                  angle: { value: -45 },
                  text: { field: "name" },
                  fill: { value: "#323130" },
                  font: { value: "Segoe UI" }
                }
              }
            }
          ],
          
          config: {
            view: { stroke: "transparent" },
            font: "Segoe UI",
            text: { font: "Segoe UI", fontSize: 10, fill: "#605e5c" }
          }
        };
      }

      else if (chartType === "lollipop") {
        spec = {
          $schema: "https://vega.github.io/schema/vega-lite/v5.json",
          description: "Lollipop chart from Excel selection",
          background: "white",
          config: { view: { stroke: "transparent" }},
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

      else if (chartType === "strip") {
        // Strip plot implementation
        spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        description: "Strip plot showing distribution using tick marks",
        background: "white",
        config: { 
            view: { stroke: "transparent" },
            axis: {
            labelFontSize: 11,
            titleFontSize: 12,
            labelColor: "#605E5C",
            titleColor: "#323130"
            }
        },
        data: { values: data },
        mark: {
            type: "tick",
            thickness: 2,
            size: 15,
            color: "#0078d4",
            opacity: 0.8,
            tooltip: true
        },
        encoding: {
            y: { 
            field: headers[0],
            type: "ordinal",
            axis: {
                title: headers[0],
                labelAngle: 0
            }
            },
            x: { 
            field: headers[1],
            type: "quantitative",
            axis: {
                title: headers[1],
                grid: true,
                gridColor: "#f3f2f1",
                gridOpacity: 0.5
            }
            },
            // Add color encoding if 3rd column exists
            ...(headers.length > 2 && {
            color: {
                field: headers[2],
                type: "nominal",
                scale: { scheme: "category10" },
                legend: {
                title: headers[2],
                orient: "right",
                titleFontSize: 11,
                labelFontSize: 10
                }
            }
            }),
            tooltip: headers.map(h => ({ field: h, type: "nominal" }))
        }
        };
      }

      else if (chartType === "treemap") {
        let treeData;
        
        if (headers.length >= 3) {
          // Hierarchical data with parent column
          treeData = data.map((d, i) => ({
            id: `${d[headers[1]]}_${i}`,
            name: d[headers[1]],
            parent: d[headers[0]] || "root",
            size: parseFloat(d[headers[2]]) || 0
          }));
          
          // Add root and parent nodes
          const parents = [...new Set(treeData.map(d => d.parent))];
          parents.forEach(parent => {
            if (parent !== "root" && !treeData.find(d => d.id === parent)) {
              treeData.push({
                id: parent,
                name: parent,
                parent: "root",
                size: 0
              });
            }
          });
          
          // Add root node
          treeData.push({
            id: "root",
            name: "Root",
            parent: "",
            size: 0
          });
        } else {
          // Simple flat data - create single level hierarchy
          treeData = [
            {
              id: "root",
              name: "Root", 
              parent: "",
              size: 0
            },
            ...data.map((d, i) => ({
              id: `item_${i}`,
              name: d[headers[1]],
              parent: "root",
              size: parseFloat(d[headers[2]]) || 0
            }))
          ];
        }

        spec = {
          $schema: "https://vega.github.io/schema/vega/v5.json",
          description: "Treemap visualization from Excel data",
          background: "white",
          width: 600,
          height: 400,
          padding: 5,
          autosize: "fit",
          
          data: [
            {
              name: "tree",
              values: treeData,
              transform: [
                {
                  type: "stratify",
                  key: "id",
                  parentKey: "parent"
                },
                {
                  type: "treemap",
                  field: "size",
                  sort: { field: "value" },
                  round: true,
                  method: "squarify",
                  ratio: 1.6,
                  size: [{ signal: "width" }, { signal: "height" }]
                }
              ]
            },
            {
              name: "nodes",
              source: "tree",
              transform: [
                { type: "filter", expr: "datum.children" }
              ]
            },
            {
              name: "leaves", 
              source: "tree",
              transform: [
                { type: "filter", expr: "!datum.children" }
              ]
            }
          ],
          
          scales: [
            {
              name: "color",
              type: "ordinal",
              domain: { data: "nodes", field: "name" },
              range: [
                "#0078d4", "#00bcf2", "#40e0d0", "#00cc6a", "#10893e",
                "#107c10", "#bad80a", "#ffb900", "#ff8c00", "#d13438"
              ]
            },
            {
              name: "fontSize",
              type: "ordinal", 
              domain: [0, 1, 2, 3],
              range: [20, 16, 12, 10]
            },
            {
              name: "opacity",
              type: "ordinal",
              domain: [0, 1, 2, 3], 
              range: [0.3, 0.6, 0.8, 1.0]
            }
          ],
          
          marks: [
            {
              type: "rect",
              from: { data: "nodes" },
              interactive: false,
              encode: {
                enter: {
                  fill: { scale: "color", field: "name" },
                  fillOpacity: { scale: "opacity", field: "depth" }
                },
                update: {
                  x: { field: "x0" },
                  y: { field: "y0" },
                  x2: { field: "x1" },
                  y2: { field: "y1" },
                  stroke: { value: "#ffffff" },
                  strokeWidth: { value: 1 }
                }
              }
            },
            {
              type: "rect",
              from: { data: "leaves" },
              encode: {
                enter: {
                  stroke: { value: "#ffffff" },
                  strokeWidth: { value: 2 }
                },
                update: {
                  x: { field: "x0" },
                  y: { field: "y0" },
                  x2: { field: "x1" },
                  y2: { field: "y1" },
                  fill: { value: "transparent" },
                  tooltip: {
                    signal: `{'Category': datum.name, 'Value': datum.size, 'Parent': datum.parent}`
                  }
                },
                hover: {
                  fill: { value: "#323130" },
                  fillOpacity: { value: 0.1 }
                }
              }
            },
            {
              type: "text",
              from: { data: "leaves" },
              interactive: false,
              encode: {
                enter: {
                  font: { value: "Segoe UI, Arial, sans-serif" },
                  align: { value: "center" },
                  baseline: { value: "middle" },
                  fill: { value: "#323130" },
                  fontWeight: { value: "bold" },
                  text: { field: "name" },
                  fontSize: { scale: "fontSize", field: "depth" }
                },
                update: {
                  x: { signal: "0.5 * (datum.x0 + datum.x1)" },
                  y: { signal: "0.5 * (datum.y0 + datum.y1)" },
                  opacity: {
                    signal: "(datum.x1 - datum.x0) > 50 && (datum.y1 - datum.y0) > 20 ? 1 : 0"
                  }
                }
              }
            }
          ]
        };
      }
      
      else if (chartType === "waffle") {
        spec = {
          $schema: "https://vega.github.io/schema/vega-lite/v5.json",
          description: "Waffle chart from Excel selection",
          background: "white",
          config: { view: { stroke: "transparent" }},
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
          $schema: "https://vega.github.io/schema/vega-lite/v5.20.1.json",
          description: "Violin chart from Excel selection",
          background: "white",
          config: {
            view: { continuousWidth: 300, continuousHeight: 300, stroke: null },
            facet: { spacing: 0 }
          },
          data: { values: data },
          mark: { type: "area", orient: "horizontal" },
          encoding: {
            color: { field: headers[0], type: "nominal" },
            column: {
              field: headers[0],
              header: {
                labelOrient: "bottom",
                labelPadding: 0,
                titleOrient: "bottom"
              },
              type: "nominal"
            },
            x: {
              axis: { grid: false, labels: false, ticks: true, values: [0] },
              field: "density",
              impute: null,
              stack: "center",
              title: null,
              type: "quantitative"
            },
            y: { field: headers[1], type: "quantitative" }
          },
          transform: [
            {
              density: headers[1],
              groupby: [headers[0]],
              as: [headers[1], "density"]
            }
          ],
          width: 100
        };
      }

      else if (chartType === "heatmap") {
        spec = {
          $schema: "https://vega.github.io/schema/vega-lite/v5.json",
          description: "Heatmap with marginal bars from Excel selection",
          background: "white",
          config: { view: { stroke: "transparent" }},
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
        background: "white",
        config: { view: { stroke: "transparent" }},
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
        x: { 
            field: headers[0], 
            type: "ordinal", 
            axis: { 
                title: null,
                labelAngle: 0       // Optional: adjust label angle if needed
            } 
        },
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
        background: "white",
        config: { view: { stroke: "transparent" }},
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
      // calculate width based on number of unique x-values
      const uniqueX = [...new Set(data.map(d => d[headers[0]]))];
      const dynamicWidth = Math.max(400, uniqueX.length * 80);

      spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        description: "Bump chart from Excel selection",
        background: "white",
        config: { view: { stroke: "transparent" }},
        data: { values: data },
        width: dynamicWidth,
        height: 200,   // give it some room
        encoding: {
          x: {
            field: headers[0],
            type: "nominal",
            axis: { title: "" },
            scale: { type: "point", padding: 1 }   // padding 1 for more spread
          },
          y: {
            field: headers[2],      
            type: "ordinal",
            axis: false
          }
        },
        layer: [
          {
            mark: { type: "line", interpolate: "monotone" },
            encoding: {
              color: {
                field: headers[1],   
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

      else if (chartType === "sankey") {
      const data = rows
        .filter(r => r[0] && r[1] && !isNaN(+r[2]))
        .map(r => ({
          key: { stk1: r[0], stk2: r[1] },
          doc_count: +r[2]
        }));

      spec = {
        $schema: "https://vega.github.io/schema/vega/v5.2.json",
        height: 300,
        width: 600,
        background: "white",
        config: { view: { stroke: "transparent" }},
        data: [
          {
            name: "rawData",
            values: data,
            transform: [
              { type: "formula", expr: "datum.key.stk1", as: "stk1" },
              { type: "formula", expr: "datum.key.stk2", as: "stk2" },
              { type: "formula", expr: "datum.doc_count", as: "size" }
            ]
          },
          {
            name: "nodes",
            source: "rawData",
            transform: [
              {
                type: "filter",
                expr:
                  "!groupSelector || groupSelector.stk1 == datum.stk1 || groupSelector.stk2 == datum.stk2"
              },
              { type: "formula", expr: "datum.stk1+datum.stk2", as: "key" },
              { type: "fold", fields: ["stk1", "stk2"], as: ["stack", "grpId"] },
              {
                type: "formula",
                expr:
                  "datum.stack == 'stk1' ? datum.stk1+' '+datum.stk2 : datum.stk2+' '+datum.stk1",
                as: "sortField"
              },
              {
                type: "stack",
                groupby: ["stack"],
                sort: { field: "sortField", order: "descending" },
                field: "size"
              },
              { type: "formula", expr: "(datum.y0+datum.y1)/2", as: "yc" }
            ]
          },
          {
            name: "groups",
            source: "nodes",
            transform: [
              {
                type: "aggregate",
                groupby: ["stack", "grpId"],
                fields: ["size"],
                ops: ["sum"],
                as: ["total"]
              },
              {
                type: "stack",
                groupby: ["stack"],
                sort: { field: "grpId", order: "descending" },
                field: "total"
              },
              { type: "formula", expr: "scale('y', datum.y0)", as: "scaledY0" },
              { type: "formula", expr: "scale('y', datum.y1)", as: "scaledY1" },
              { type: "formula", expr: "datum.stack == 'stk1'", as: "rightLabel" },
              { type: "formula", expr: "datum.total/domain('y')[1]", as: "percentage" }
            ]
          },
          {
            name: "destinationNodes",
            source: "nodes",
            transform: [{ type: "filter", expr: "datum.stack == 'stk2'" }]
          },
          {
            name: "edges",
            source: "nodes",
            transform: [
              { type: "filter", expr: "datum.stack == 'stk1'" },
              {
                type: "lookup",
                from: "destinationNodes",
                key: "key",
                fields: ["key"],
                as: ["target"]
              },
              {
                type: "linkpath",
                orient: "horizontal",
                shape: "diagonal",
                sourceY: { expr: "scale('y', datum.yc)" },
                sourceX: { expr: "scale('x', 'stk1') + bandwidth('x')" },
                targetY: { expr: "scale('y', datum.target.yc)" },
                targetX: { expr: "scale('x', 'stk2')" }
              },
              { type: "formula", expr: "range('y')[0]-scale('y', datum.size)", as: "strokeWidth" },
              { type: "formula", expr: "datum.size/domain('y')[1]", as: "percentage" }
            ]
          }
        ],
        scales: [
          {
            name: "x",
            type: "band",
            range: "width",
            domain: ["stk1", "stk2"],
            paddingOuter: 0.05,
            paddingInner: 0.95
          },
          {
            name: "y",
            type: "linear",
            range: "height",
            domain: { data: "nodes", field: "y1" }
          },
          {
            name: "color",
            type: "ordinal",
            range: "category",
            domain: {
              fields: [
                { data: "rawData", field: "stk1" },
                { data: "rawData", field: "stk2" }
              ]
            }
          },
          {
            name: "stackNames",
            type: "ordinal",
            range: ["Source", "Destination"],
            domain: ["stk1", "stk2"]
          }
        ],
        axes: [
          {
            orient: "bottom",
            scale: "x",
            encode: {
              labels: { update: { text: { scale: "stackNames", field: "value" } } }
            }
          },
          { orient: "left", scale: "y" }
        ],
        marks: [
          {
            type: "path",
            name: "edgeMark",
            from: { data: "edges" },
            clip: true,
            encode: {
              update: {
                stroke: { scale: "color", field: "stk1" }, // links colored by source
                strokeWidth: { field: "strokeWidth" },
                path: { field: "path" },
                strokeOpacity: {
                  signal:
                    "!groupSelector && (groupHover.stk1 == datum.stk1 || groupHover.stk2 == datum.stk2) ? 0.9 : 0.3"
                },
                zindex: {
                  signal:
                    "!groupSelector && (groupHover.stk1 == datum.stk1 || groupHover.stk2 == datum.stk2) ? 1 : 0"
                },
                tooltip: {
                  signal:
                    "datum.stk1 + ' → ' + datum.stk2 + '    ' + format(datum.size, ',.0f') + '   (' + format(datum.percentage, '.1%') + ')'"
                }
              },
              hover: { strokeOpacity: { value: 1 } }
            }
          },
          {
            type: "rect",
            name: "groupMark",
            from: { data: "groups" },
            encode: {
              enter: {
                fill: { scale: "color", field: "grpId" }, // both source & destination use union colors
                width: { scale: "x", band: 1 }
              },
              update: {
                x: { scale: "x", field: "stack" },
                y: { field: "scaledY0" },
                y2: { field: "scaledY1" },
                fillOpacity: { value: 0.6 },
                tooltip: {
                  signal:
                    "datum.grpId + '   ' + format(datum.total, ',.0f') + '   (' + format(datum.percentage, '.1%') + ')'"
                }
              },
              hover: { fillOpacity: { value: 1 } }
            }
          },
          {
            type: "text",
            from: { data: "groups" },
            interactive: false,
            encode: {
              update: {
                x: {
                  signal:
                    "scale('x', datum.stack) + (datum.rightLabel ? bandwidth('x') + 8 : -8)"
                },
                yc: { signal: "(datum.scaledY0 + datum.scaledY1)/2" },
                align: { signal: "datum.rightLabel ? 'left' : 'right'" },
                baseline: { value: "middle" },
                fontWeight: { value: "bold" },
                text: {
                  signal: "abs(datum.scaledY0-datum.scaledY1) > 13 ? datum.grpId : ''"
                }
              }
            }
          }
        ],
        signals: [
          {
            name: "groupHover",
            value: {},
            on: [
              {
                events: "@groupMark:mouseover",
                update:
                  "{stk1:datum.stack=='stk1' && datum.grpId, stk2:datum.stack=='stk2' && datum.grpId}"
              },
              { events: "mouseout", update: "{}" }
            ]
          },
          {
            name: "groupSelector",
            value: false,
            on: [
              {
                events: "@groupMark:click!",
                update:
                  "{stack:datum.stack, stk1:datum.stack=='stk1' && datum.grpId, stk2:datum.stack=='stk2' && datum.grpId}"
              },
              {
                events: [
                  { type: "click", markname: "groupReset" },
                  { type: "dblclick" }
                ],
                update: "false"
              }
            ]
          }
        ]
      };
      }

      else if (chartType === "ribbon") {
        // Calculate dynamic dimensions based on data
        const uniquePeriods = [...new Set(data.map(d => d[headers[0]]))];
        const dynamicWidth = Math.max(600, uniquePeriods.length * 100); // More space per period
        const dynamicHeight = 400; // Adequate height for ribbon flow

        spec = {
            $schema: "https://vega.github.io/schema/vega-lite/v6.json",
            description: "Ribbon chart from Excel selection",
            background: "white",
            width: dynamicWidth,
            height: dynamicHeight,
            config: { view: { stroke: "transparent" }},
            data: { values: data },
            layer: [
            {
                mark: { 
                type: "area", 
                interpolate: "monotone", 
                tooltip: true,
                opacity: 0.8
                },
                encoding: {
                x: {
                    field: headers[0],
                    type: "ordinal", // temporal change to "ordinal" if your first col is not a date
                    scale: {
                    type: "point",
                    padding: 0.3 // Add padding between periods for more spread
                    },
                    axis: {
                    title: headers[0],
                    labelAngle: -45, // Angle labels to prevent overlap
                    labelFontSize: 12,
                    titleFontSize: 14,
                    labelPadding: 10,
                    titlePadding: 20
                    }
                },
                y: {
                    aggregate: "sum",
                    field: headers[2],
                    type: "quantitative",
                    axis: {
                    title: headers[2],
                    labelFontSize: 12,
                    titleFontSize: 14,
                    grid: true,
                    gridOpacity: 0.3
                    },
                    stack: "center"
                },
                color: {
                    field: headers[1],
                    type: "nominal",
                    legend: {
                    title: headers[1],
                    titleFontSize: 12,
                    labelFontSize: 11,
                    orient: "right"
                    }
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
            font: "Segoe UI",
            axis: {
                ticks: false,
                grid: true,
                gridColor: "#f0f0f0",
                gridOpacity: 0.5,
                gridWidth: 1,
                domain: false,
                labelColor: "#605e5c",
                titleColor: "#323130"
            },
            legend: {
                titleFont: "Segoe UI",
                titleFontWeight: "bold",
                titleColor: "#323130",
                labelFont: "Segoe UI",
                labelColor: "#605e5c",
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
        background: "white",
        config: { view: { stroke: "transparent" }},
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
        background: "white",
        config: { view: { stroke: "transparent" }},
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

      else if (chartType === "area") {
        spec = {
          $schema: "https://vega.github.io/schema/vega-lite/v6.json",
          description: "Area chart from Excel selection",
          background: "white",
          config: { view: { stroke: "transparent" }},
          data: { values: data },
          mark: { 
            type: "area", 
            tooltip: true,
            opacity: 0.7
          },
          encoding: {
            x: { 
              field: headers[0], 
              type: "ordinal",
              axis: {
                title: headers[0],
                labelFontSize: 12,
                titleFontSize: 14
              }
            },
            y: { 
              field: headers[1], 
              type: "quantitative",
              axis: {
                title: headers[1],
                labelFontSize: 12,
                titleFontSize: 14
              }
            },
            // Add color encoding for multiple areas if 3rd column exists
            ...(headers.length >= 3 && {
              color: { 
                field: headers[2], 
                type: "nominal",
                legend: {
                  title: headers[2],
                  titleFontSize: 12,
                  labelFontSize: 11
                }
              }
            })
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
            }
          }
        };
      }

      else if (chartType === "bar") {
        spec = {
          $schema: "https://vega.github.io/schema/vega-lite/v6.json",
          description: "Bar chart from Excel selection",
          background: "white",
          config: { view: { stroke: "transparent" }},
          data: { values: data },
          mark: { 
            type: "bar", 
            tooltip: true
          },
          encoding: {
            x: { 
              field: headers[0], 
              type: "nominal",
              axis: {
                title: headers[0],
                labelFontSize: 12,
                titleFontSize: 14
              }
            },
            y: { 
              field: headers[1], 
              type: "quantitative",
              axis: {
                title: headers[1],
                labelFontSize: 12,
                titleFontSize: 14
              }
            },
            // Add color encoding for grouped bars if 3rd column exists
            ...(headers.length >= 3 && {
              color: { 
                field: headers[2], 
                type: "nominal",
                legend: {
                  title: headers[2],
                  titleFontSize: 12,
                  labelFontSize: 11
                }
              }
            })
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
            }
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
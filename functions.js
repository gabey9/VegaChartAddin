/**
 * LINE custom function using the exact same specification as taskpane.js
 * Creates a multi-series line chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function LINE(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      // Convert rows -> objects (same as taskpane.js)
      const processedData = rows.map(row => {
        let obj = {};
        headers.forEach((h, i) => {
          obj[h] = row[i];
        });
        return obj;
      });

      // Transform data for multi-series line chart (exact copy from taskpane.js)
      const transformedData = [];
      const valueColumns = headers.slice(1);
      processedData.forEach(row => {
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

      // Use EXACT specification from taskpane.js line chart
      const spec = {
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

      createChart(spec, "line", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * BAR custom function using the exact same specification as taskpane.js
 * Creates a bar chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function BAR(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      // Convert rows -> objects (same as taskpane.js)
      const processedData = rows.map(row => {
        let obj = {};
        headers.forEach((h, i) => {
          obj[h] = row[i];
        });
        return obj;
      });

      // Use EXACT specification from taskpane.js bar chart
      const spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        description: "Bar chart from Excel selection",
        background: "white",
        config: { view: { stroke: "transparent" }},
        data: { values: processedData },
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

      createChart(spec, "bar", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * PIE custom function using the exact same specification as taskpane.js
 * Creates a pie chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function PIE(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 2) {
        resolve("Error: Pie chart requires 2 columns (Category, Value)");
        return;
      }

      // Validate that all values are positive numbers
      const hasInvalidValues = rows.some(row => isNaN(row[1]) || row[1] <= 0);
      if (hasInvalidValues) {
        resolve("Error: Pie chart values must be positive numbers");
        return;
      }

      // Convert rows -> objects (same as taskpane.js)
      const processedData = rows.map(row => {
        let obj = {};
        headers.forEach((h, i) => {
          obj[h] = row[i];
        });
        return obj;
      });

      // Use EXACT specification from taskpane.js pie chart
      const spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        background: "white",
        config: { view: { stroke: "transparent" }},
        description: "Pie chart from Excel selection",
        data: { values: processedData },
        mark: { type: "arc", outerRadius: 120 },
        encoding: {
          theta: { field: headers[1], type: "quantitative" },
          color: { field: headers[0], type: "nominal" }
        }
      };

      createChart(spec, "pie", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * AREA custom function using the exact same specification as taskpane.js
 * Creates an area chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function AREA(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      // Convert rows -> objects (same as taskpane.js)
      const processedData = rows.map(row => {
        let obj = {};
        headers.forEach((h, i) => {
          obj[h] = row[i];
        });
        return obj;
      });

      // Use EXACT specification from taskpane.js area chart
      const spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        description: "Area chart from Excel selection",
        background: "white",
        config: { view: { stroke: "transparent" }},
        data: { values: processedData },
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

      createChart(spec, "area", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * SCATTER custom function using the exact same specification as taskpane.js
 * Creates a scatter plot from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function SCATTER(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 2) {
        resolve("Error: Scatter plot requires at least 2 columns (X, Y values)");
        return;
      }

      // Convert rows -> objects (same as taskpane.js)
      const processedData = rows.map(row => {
        let obj = {};
        headers.forEach((h, i) => {
          obj[h] = row[i];
        });
        return obj;
      });

      // Use EXACT specification from taskpane.js point chart
      const spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        description: "Colored scatter plot from Excel selection",
        background: "white",
        config: { view: { stroke: "transparent" }},
        data: { values: processedData },
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

      createChart(spec, "scatter", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * RADIAL custom function using the exact same specification as taskpane.js
 * Creates a radial chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function RADIAL(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 2) {
        resolve("Error: Radial chart requires 2 columns (Category, Value)");
        return;
      }

      // Convert rows -> objects (same as taskpane.js)
      const processedData = rows.map(row => {
        let obj = {};
        headers.forEach((h, i) => {
          obj[h] = row[i];
        });
        return obj;
      });

      // Use EXACT specification from taskpane.js radial chart
      const spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        description: "Radial chart from Excel selection",
        background: "white",
        config: { view: { stroke: "transparent" }},
        data: { values: processedData },
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

      createChart(spec, "radial", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * BOX custom function using the exact same specification as taskpane.js
 * Creates a box plot from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function BOX(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 2) {
        resolve("Error: Box plot requires 2 columns (Category, Values)");
        return;
      }

      // Expect headers: Category | Value (same as taskpane.js)
      const processedData = rows
        .filter(r => r[0] && !isNaN(+r[1]))
        .map(r => ({
          category: r[0],
          value: +r[1]
        }));

      if (processedData.length === 0) {
        resolve("Error: No valid numeric data found for box plot");
        return;
      }

      // Use EXACT specification from taskpane.js box chart
      const spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        description: "Box plot from Excel selection",
        data: { values: processedData },
        mark: {
          type: "boxplot",
          extent: "min-max"   // show whiskers from min to max
        },
        encoding: {
          x: { field: "category", type: "nominal" },
          y: {
            field: "value",
            type: "quantitative",
            scale: { zero: false }
          },
          color: {
            field: "category",
            type: "nominal",
            legend: null
          }
        }
      };

      createChart(spec, "box", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * RADAR custom function using the exact same specification as taskpane.js
 * Creates a radar chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function RADAR(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 3) {
        resolve("Error: Radar chart requires at least 3 columns (Series, Dimension1, Dimension2, ...)");
        return;
      }

      const radarData = [];
      const dimensions = headers.slice(1); // All columns except first are dimensions
      
      rows.forEach((row, seriesIndex) => {
        const seriesName = row[headers[0]] || `Series ${seriesIndex + 1}`;
        
        dimensions.forEach(dimension => {
          const value = parseFloat(row[headers.indexOf(dimension)]) || 0;
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

      // Use EXACT specification from taskpane.js radar chart
      const spec = {
        $schema: "https://vega.github.io/schema/vega/v6.json",
        description: "Radar chart from Excel selection",
        width: 400,
        height: 400,
        padding: 60,
        autosize: {"type": "none", "contains": "padding"},
        background: "white",
        config: { view: { stroke: "transparent" }},

        signals: [
          {"name": "radius", "update": "width / 2"}
        ],

        data: [
          {
            name: "table",
            values: radarData
          },
          {
            name: "dimensions",
            values: uniqueDimensions.map(d => ({dimension: d}))
          }
        ],

        scales: [
          {
            name: "angular",
            type: "point",
            range: {"signal": "[-PI, PI]"},
            padding: 0.5,
            domain: uniqueDimensions
          },
          {
            name: "radial",
            type: "linear",
            range: {"signal": "[0, radius]"},
            zero: true,
            nice: true,
            domain: {"data": "table", "field": "value"},
            domainMin: 0
          },
          {
            name: "color",
            type: "ordinal",
            domain: {"data": "table", "field": "category"},
            range: [
              "#0078d4", "#00bcf2", "#40e0d0", "#00cc6a", "#10893e",
              "#107c10", "#bad80a", "#ffb900", "#ff8c00", "#d13438"
            ]
          }
        ],

        encode: {
          enter: {
            x: {"signal": "radius"},
            y: {"signal": "radius"}
          }
        },

        marks: [
          {
            type: "group",
            name: "categories",
            zindex: 1,
            from: {
              facet: {"data": "table", "name": "facet", "groupby": ["category", "series"]}
            },
            marks: [
              {
                type: "line",
                name: "category-line",
                from: {"data": "facet"},
                encode: {
                  enter: {
                    interpolate: {"value": "linear-closed"},
                    x: {"signal": "scale('radial', datum.value) * cos(scale('angular', datum.dimension))"},
                    y: {"signal": "scale('radial', datum.value) * sin(scale('angular', datum.dimension))"},
                    stroke: {"scale": "color", "field": "category"},
                    strokeWidth: {"value": 2},
                    fill: {"scale": "color", "field": "category"},
                    fillOpacity: {"value": 0.1},
                    strokeOpacity: {"value": 0.8}
                  }
                }
              },
              {
                type: "symbol",
                name: "category-points",
                from: {"data": "facet"},
                encode: {
                  enter: {
                    x: {"signal": "scale('radial', datum.value) * cos(scale('angular', datum.dimension))"},
                    y: {"signal": "scale('radial', datum.value) * sin(scale('angular', datum.dimension))"},
                    size: {"value": 50},
                    fill: {"scale": "color", "field": "category"},
                    stroke: {"value": "white"},
                    strokeWidth: {"value": 1}
                  }
                }
              }
            ]
          },
          {
            type: "rule",
            name: "radial-grid",
            from: {"data": "dimensions"},
            zindex: 0,
            encode: {
              enter: {
                x: {"value": 0},
                y: {"value": 0},
                x2: {"signal": "radius * cos(scale('angular', datum.dimension))"},
                y2: {"signal": "radius * sin(scale('angular', datum.dimension))"},
                stroke: {"value": "#e1e4e8"},
                strokeWidth: {"value": 1}
              }
            }
          },
          {
            type: "text",
            name: "dimension-label",
            from: {"data": "dimensions"},
            zindex: 1,
            encode: {
              enter: {
                x: {"signal": "(radius + 20) * cos(scale('angular', datum.dimension))"},
                y: {"signal": "(radius + 20) * sin(scale('angular', datum.dimension))"},
                text: {"field": "dimension"},
                align: [
                  {
                    test: "abs(scale('angular', datum.dimension)) > PI / 2",
                    value: "right"
                  },
                  {
                    value: "left"
                  }
                ],
                baseline: [
                  {
                    test: "scale('angular', datum.dimension) > 0", 
                    value: "top"
                  },
                  {
                    test: "scale('angular', datum.dimension) == 0", 
                    value: "middle"
                  },
                  {
                    value: "bottom"
                  }
                ],
                fill: {"value": "#323130"},
                fontWeight: {"value": "bold"},
                font: {"value": "Segoe UI"},
                fontSize: {"value": 12}
              }
            }
          },
          {
            type: "line",
            name: "outer-line",
            from: {"data": "radial-grid"},
            encode: {
              enter: {
                interpolate: {"value": "linear-closed"},
                x: {"field": "x2"},
                y: {"field": "y2"},
                stroke: {"value": "#8a8886"},
                strokeWidth: {"value": 2},
                strokeOpacity: {"value": 0.6}
              }
            }
          }
        ]
      };

      createChart(spec, "radar", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * WATERFALL custom function using the exact same specification as taskpane.js
 * Creates a waterfall chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function WATERFALL(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 3) {
        resolve("Error: Waterfall chart requires 3 columns (Category, Amount, Type)");
        return;
      }

      // Convert rows -> objects (same as taskpane.js)
      const processedDataRaw = rows.map(row => {
        let obj = {};
        headers.forEach((h, i) => {
          obj[h] = row[i];
        });
        return obj;
      });

      // Process waterfall data inline - set last entry's amount to 0 (same as taskpane.js)
      const processedData = [...processedDataRaw];
      if (processedData.length > 0) {
        processedData[processedData.length - 1] = {
          ...processedData[processedData.length - 1],
          [headers[1]]: 0
        };
      }

      // Calculate dynamic dimensions
      const numDataPoints = processedDataRaw.length;
      const dynamicWidth = Math.max(400, Math.min(1600, numDataPoints * 50));
      const maxAmount = Math.max(...processedDataRaw.map(d => Math.abs(d[headers[1]])));
      const dynamicHeight = Math.max(300, Math.min(600, maxAmount / 100 + 200));

      // Use EXACT specification from taskpane.js waterfall chart
      const spec = {
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

      createChart(spec, "waterfall", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * SUNBURST custom function using the exact same specification as taskpane.js
 * Creates a sunburst chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function SUNBURST(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 2) {
        resolve("Error: Sunburst chart requires at least 2 columns (Parent, Child, Value optional)");
        return;
      }

      const nodes = new Map();
      rows.forEach((row, i) => {
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

      // Use EXACT specification from taskpane.js sunburst chart
      const spec = {
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

      createChart(spec, "sunburst", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * TREEMAP custom function using the exact same specification as taskpane.js
 * Creates a treemap chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function TREEMAP(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 3) {
        resolve("Error: Treemap chart requires 3 columns (Parent, Category, Value)");
        return;
      }

      let treeData;
      
      if (headers.length >= 3) {
        // Hierarchical data with parent column
        treeData = rows.map((d, i) => ({
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
          ...rows.map((d, i) => ({
            id: `item_${i}`,
            name: d[headers[1]],
            parent: "root",
            size: parseFloat(d[headers[2]]) || 0
          }))
        ];
      }

      // Use EXACT specification from taskpane.js treemap chart
      const spec = {
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

      createChart(spec, "treemap", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * HISTOGRAM custom function using the exact same specification as taskpane.js
 * Creates a histogram from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function HISTOGRAM(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 1) {
        resolve("Error: Histogram requires at least 1 column of numeric values");
        return;
      }

      // Expect a single numeric column (same as taskpane.js)
      const processedData = rows
        .filter(r => !isNaN(+r[0]))
        .map(r => ({ value: +r[0] }));

      if (processedData.length === 0) {
        resolve("Error: No valid numeric data found for histogram");
        return;
      }

      // Use EXACT specification from taskpane.js histogram chart
      const spec = {
        "$schema": "https://vega.github.io/schema/vega-lite/v6.json",
        "description": "Histogram from Excel selection",
        "data": { "values": processedData },
        "mark": "bar",
        "encoding": {
          "x": {
            "field": "value",
            "bin": { "maxbins": 20 },   // adjust bin count here
            "type": "quantitative",
            "axis": { "title": "Value" }
          },
          "y": {
            "aggregate": "count",
            "type": "quantitative",
            "axis": { "title": "Count" }
          }
        }
      };

      createChart(spec, "histogram", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * ARC custom function using the exact same specification as taskpane.js
 * Creates an arc diagram from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function ARC(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 2) {
        resolve("Error: Arc chart requires at least 2 columns (Source, Target, Weight optional)");
        return;
      }

      // Convert rows -> objects (same as taskpane.js)
      const processedData = rows.map(row => {
        let obj = {};
        headers.forEach((h, i) => {
          obj[h] = row[i];
        });
        return obj;
      });

      // EXACT data processing from taskpane.js - Transform Excel data for arc chart
      const edges = processedData.map((row, index) => ({
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

      // Use EXACT specification from taskpane.js arc chart
      const spec = {
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

      createChart(spec, "arc", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * Generic chart creation function (same approach as taskpane.js)
 */
async function createChart(spec, chartType, headers, rows) {
  return new Promise(async (resolve, reject) => {
    try {
      const chartId = `${chartType}_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`;
      
      // Render hidden chart (same as taskpane.js)
      const hiddenDiv = document.createElement("div");
      hiddenDiv.style.display = "none";
      hiddenDiv.id = chartId;
      document.body.appendChild(hiddenDiv);

      // Load Vega-Lite if not available
      if (typeof vegaEmbed === 'undefined') {
        await loadVegaLibraries();
      }

      const result = await vegaEmbed(hiddenDiv, spec, { actions: false });
      const view = result.view;

      // Export chart -> PNG (same as taskpane.js)
      const pngUrl = await view.toImageURL("png");
      const response = await fetch(pngUrl);
      const blob = await response.blob();

      const reader = new FileReader();
      reader.onloadend = async () => {
        try {
          const base64data = reader.result.split(",")[1];

          // Insert into Excel (same approach as taskpane.js)
          await insertChartIntoExcel(base64data, chartType, chartId);
          
          // Clean up hidden div
          document.body.removeChild(hiddenDiv);
          resolve();
          
        } catch (error) {
          // Clean up on error
          if (document.body.contains(hiddenDiv)) {
            document.body.removeChild(hiddenDiv);
          }
          reject(error);
        }
      };
      
      reader.readAsDataURL(blob);

    } catch (error) {
      reject(error);
    }
  });
}

/**
 * Inserts chart into Excel using the same approach as taskpane.js
 */
async function insertChartIntoExcel(base64data, chartType, chartId) {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Remove old chart and get its position
    const oldPosition = await removeExistingCharts(context, sheet, chartType);

    let left, top, targetWidth;

    if (oldPosition) {
      // Use old chart position and size
      left = oldPosition.left;
      top = oldPosition.top;
      targetWidth = oldPosition.width;
    } else {
      // Fall back to current selection
      const range = context.workbook.getSelectedRange();
      range.load("left, top, width, height");
      await context.sync();
      left = range.left;
      top = range.top;
      targetWidth = Math.max(400, range.width * 8); // Default chart width
    }

    // Insert the new image
    const image = sheet.shapes.addImage(base64data);
    image.left = left;
    image.top = top;
    image.lockAspectRatio = true; // Set this BEFORE setting dimensions
    image.width = targetWidth; // Only set width, let Excel calculate height
    image.name = `${chartType.charAt(0).toUpperCase() + chartType.slice(1)}Chart_${chartId}`;

    await context.sync();
  });
}

/**
 * Remove existing charts of the same type (prevents duplicates)
 */
async function removeExistingCharts(context, sheet, chartType) {
  const shapes = sheet.shapes;
  shapes.load("items");
  await context.sync();

  const chartPrefix = `${chartType.charAt(0).toUpperCase() + chartType.slice(1)}Chart_`;
  let oldPosition = null;

  for (let i = shapes.items.length - 1; i >= 0; i--) {
    const shape = shapes.items[i];
    shape.load(["name", "left", "top", "width", "height"]);
  }
  await context.sync();

  for (let i = shapes.items.length - 1; i >= 0; i--) {
    const shape = shapes.items[i];
    if (shape.name && shape.name.startsWith(chartPrefix)) {
      // Save position before deleting
      oldPosition = {
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
      };
      shape.delete();
      await context.sync();
    }
  }

  return oldPosition;
}

/**
 * Load Vega libraries (same CDN versions as taskpane.html)
 */
function loadVegaLibraries() {
  return new Promise((resolve, reject) => {
    if (typeof vegaEmbed !== 'undefined') {
      resolve();
      return;
    }

    // Load libraries in sequence (same as taskpane.html)
    const scripts = [
      'https://cdn.jsdelivr.net/npm/vega@6',
      'https://cdn.jsdelivr.net/npm/vega-lite@6', 
      'https://cdn.jsdelivr.net/npm/vega-embed@6'
    ];

    let loadedCount = 0;
    
    scripts.forEach((src, index) => {
      const script = document.createElement('script');
      script.src = src;
      script.onload = () => {
        loadedCount++;
        if (loadedCount === scripts.length) {
          resolve();
        }
      };
      script.onerror = () => reject(new Error(`Failed to load ${src}`));
      document.head.appendChild(script);
    });
  });
}

// Register all custom functions
if (typeof CustomFunctions !== 'undefined') {
  CustomFunctions.associate("LINE", LINE);
  CustomFunctions.associate("BAR", BAR);
  CustomFunctions.associate("PIE", PIE);
  CustomFunctions.associate("AREA", AREA);
  CustomFunctions.associate("SCATTER", SCATTER);
  CustomFunctions.associate("RADIAL", RADIAL);
  CustomFunctions.associate("BOX", BOX);
  CustomFunctions.associate("RADAR", RADAR);
  CustomFunctions.associate("WATERFALL", WATERFALL);
  CustomFunctions.associate("SUNBURST", SUNBURST);
  CustomFunctions.associate("TREEMAP", TREEMAP);
  CustomFunctions.associate("HISTOGRAM", HISTOGRAM);
  CustomFunctions.associate("ARC", ARC);
}
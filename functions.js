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
 * BUBBLE custom function using the exact same specification as taskpane.js
 * Creates a bubble chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function BUBBLE(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 3) {
        resolve("Error: Bubble chart requires at least 3 columns (X values, Y values, Size values)");
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

      // Use EXACT specification from taskpane.js bubble chart
      const spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        description: "Bubble chart from Excel selection",
        background: "white",
        config: { view: { stroke: "transparent" }},
        data: { values: processedData },
        mark: { type: "circle", tooltip: true, opacity: 0.7 },
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
          size: {
            field: headers[2],
            type: "quantitative",
            scale: {
              type: "linear",
              range: [100, 1000]
            },
            legend: {
              title: headers[2],
              titleFontSize: 12,
              labelFontSize: 11
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

      createChart(spec, "bubble", headers, rows)
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
        resolve("Error: Sunburst chart requires at least 2 columns (Parent, Child, optional Value)");
        return;
      }

      // Build hierarchical data (same as taskpane.js)
      const nodes = new Map();
      rows.forEach((row, i) => {
        const parent = row[0] || "";
        const child = row[1] || `node_${i}`;
        const value = headers.length >= 3 ? (parseFloat(row[2]) || 1) : 1;
        
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

      // Convert rows -> objects (same as taskpane.js)
      const processedData = rows.map(row => {
        let obj = {};
        headers.forEach((h, i) => {
          obj[h] = row[i];
        });
        return obj;
      });

      // Build treemap data structure (same as taskpane.js)
      let treeData;
      
      if (headers.length >= 3) {
        // Hierarchical data with parent column
        treeData = processedData.map((d, i) => ({
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
          ...processedData.map((d, i) => ({
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

      // Expect a single numeric column
      const numericData = rows
        .filter(r => !isNaN(+r[0]))
        .map(r => ({ value: +r[0] }));

      if (numericData.length === 0) {
        resolve("Error: No valid numeric data found for histogram");
        return;
      }

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

      // Use EXACT specification from taskpane.js histogram
      const spec = {
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

      createChart(spec, "histogram", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * MAP custom function using the exact same specification as taskpane.js
 * Creates a world map chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function MAP(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 2) {
        resolve("Error: Map chart requires 2 columns (Country ISO3, Value)");
        return;
      }

      // ISO3 to numeric ID mapping (same as taskpane.js)
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

      // Process data (same as taskpane.js)
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

      if (worldData.length === 0) {
        resolve("Error: No valid country data found. Please use ISO3 country codes (USA, GBR, DEU, etc.)");
        return;
      }

      // Use EXACT specification from taskpane.js map chart
      const spec = {
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

      createChart(spec, "map", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * CANDLESTICK custom function
 * Creates a candlestick chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function CANDLESTICK(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 5) {
        resolve("Error: Candlestick chart requires 5 columns (Date, Open, High, Low, Close)");
        return;
      }

      // Convert rows -> objects
      const processedData = rows.map(row => {
        let obj = {};
        headers.forEach((h, i) => {
          obj[h] = row[i];
        });
        return obj;
      });

      // Helper function to convert Excel dates to JS dates
      function excelDateToJSDate(serial) {
        if (typeof serial === 'number') {
          return new Date(Math.round((serial - 25569) * 86400 * 1000));
        }
        return new Date(serial);
      }

      // Process and validate data - SKIP ROWS WITH MISSING VALUES
      const candlestickData = processedData
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
        resolve("Error: No valid candlestick data found");
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

      // V4 Specification - Tight spacing, no gaps, optimized layout
      const spec = {
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

      createChart(spec, "candlestick", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * TREE custom function using the exact same specification as taskpane.js
 * Creates a tree diagram from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function TREE(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 2) {
        resolve("Error: Tree chart requires at least 2 columns (Parent, Child, Value optional)");
        return;
      }

      // Process data same as taskpane.js
      const nodes = new Map();

      rows.forEach((row, i) => {
        const parent = row[0] || "";
        const child = row[1] || `node_${i}`;
        const value = headers.length >= 3 ? (parseFloat(row[2]) || 1) : 1;
        
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

      // Use EXACT specification from taskpane.js tree chart
      const spec = {
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

      createChart(spec, "tree", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * WORDCLOUD custom function using the exact same specification as taskpane.js
 * Creates a word cloud from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function WORDCLOUD(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 1) {
        resolve("Error: Wordcloud requires at least 1 column (Text data)");
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

      // Use EXACT specification from taskpane.js wordcloud chart
      const spec = {
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
            values: processedData,
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

      createChart(spec, "wordcloud", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * STRIP custom function using the exact same specification as taskpane.js
 * Creates a strip plot from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function STRIP(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 2) {
        resolve("Error: Strip plot requires at least 2 columns (Categories, Values)");
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

      // Use EXACT specification from taskpane.js strip chart
      const spec = {
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
        data: { values: processedData },
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

      createChart(spec, "strip", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * HEATMAP custom function using the exact same specification as taskpane.js
 * Creates a heatmap from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function HEATMAP(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 3) {
        resolve("Error: Heatmap requires 3 columns (Y-categories, X-categories, Values)");
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

      // Use EXACT specification from taskpane.js heatmap chart
      const spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v5.json",
        description: "Heatmap with marginal bars from Excel selection",
        background: "white",
        config: { view: { stroke: "transparent" }},
        data: { values: processedData },
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

      createChart(spec, "heatmap", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * BULLET custom function using the exact same specification as taskpane.js
 * Creates a bullet chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function BULLET(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 7) {
        resolve("Error: Bullet chart requires 7 columns (Title, Poor max, Satisfactory max, Good max, Actual, Forecast, Target)");
        return;
      }

      // Convert to bullet chart data format (same as taskpane.js)
      const processedData = rows.map(r => ({
        title: r[0],
        ranges: [+r[1], +r[2], +r[3]],
        measures: [+r[4], +r[5]],
        markers: [+r[6]]
      }));

      // Use EXACT specification from taskpane.js bullet chart
      const spec = {
        "$schema": "https://vega.github.io/schema/vega-lite/v6.json",
        background: "white",
        config: { view: { stroke: "transparent" }},
        "data": { "values": processedData },
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

      createChart(spec, "bullet", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * HORIZON custom function using the exact same specification as taskpane.js
 * Creates a horizon chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function HORIZON(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 2) {
        resolve("Error: Horizon chart requires 2 columns (X values, Y values)");
        return;
      }

      // Convert rows -> objects and transform data (same as taskpane.js)
      const horizonData = rows.map((row, index) => ({
        x: row[0] || index + 1,
        y: parseFloat(row[1]) || 0
      }));

      // Calculate data range and bands (same as taskpane.js)
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

      // Use EXACT specification from taskpane.js horizon chart
      const spec = {
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

      createChart(spec, "horizon", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * SLOPE custom function using the exact same specification as taskpane.js
 * Creates a slope chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function SLOPE(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 3) {
        resolve("Error: Slope chart requires 3 columns (Time Period, Category, Value)");
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

      const timePeriods = [...new Set(processedData.map(d => d[headers[0]]))];
      const categories = [...new Set(processedData.map(d => d[headers[1]]))];
      
      // Filter data for first and last periods only (same as taskpane.js)
      const firstPeriod = timePeriods[0];
      const lastPeriod = timePeriods[timePeriods.length - 1];
      
      const slopeData = processedData.filter(d => 
        d[headers[0]] === firstPeriod || d[headers[0]] === lastPeriod
      );

      // Check if values are percentages (between -1 and 1)
      const allValues = slopeData.map(d => d[headers[2]]);
      const isPercentage = allValues.every(v => v >= -1 && v <= 1);
      const formatString = isPercentage ? ".1%" : ",.0f";

      // Calculate dynamic dimensions based on number of categories
      const dynamicHeight = Math.max(300, Math.min(600, categories.length * 40));
      const dynamicWidth = 500;

      // Use EXACT specification from taskpane.js slope chart
      const spec = {
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

      createChart(spec, "slope", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * MEKKO custom function using the exact same specification as taskpane.js
 * Creates a Mekko chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function MEKKO(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 3) {
        resolve("Error: Mekko chart requires 3 columns (Category, Subcategory, Value)");
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

      // Use EXACT specification from taskpane.js mekko chart
      const spec = {
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
            values: processedData
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

      createChart(spec, "mekko", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * MARIMEKKO custom function using the exact same specification as taskpane.js
 * Creates a marimekko chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function MARIMEKKO(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 3) {
        resolve("Error: Marimekko chart requires at least 3 columns: Category, Subcategory, Value");
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

      // Use EXACT specification from taskpane.js marimekko chart
      const spec = {
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
            "values": processedData,
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

      createChart(spec, "marimekko", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * BUMP custom function using the exact same specification as taskpane.js
 * Creates a bump chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function BUMP(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 3) {
        resolve("Error: Bump chart requires 3 columns: Time periods, Categories, Rank values");
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

      // Calculate width based on number of unique x-values
      const uniqueX = [...new Set(processedData.map(d => d[headers[0]]))];
      const dynamicWidth = Math.max(400, uniqueX.length * 80);

      // Use EXACT specification from taskpane.js bump chart
      const spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        description: "Bump chart from Excel selection",
        background: "white",
        config: { view: { stroke: "transparent" }},
        data: { values: processedData },
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

      createChart(spec, "bump", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * WAFFLE custom function using the exact same specification as taskpane.js
 * Creates a waffle chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function WAFFLE(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 2) {
        resolve("Error: Waffle chart requires 2 columns: Category names, Values");
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

      // Use EXACT specification from taskpane.js waffle chart
      const spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v5.json",
        description: "Waffle chart from Excel selection",
        background: "white",
        config: { view: { stroke: "transparent" }},
        data: { values: processedData },
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

      createChart(spec, "waffle", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * LOLLIPOP custom function using the exact same specification as taskpane.js
 * Creates a lollipop chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function LOLLIPOP(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 2) {
        resolve("Error: Lollipop chart requires 2 columns: Category names, Values");
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

      // Use EXACT specification from taskpane.js lollipop chart
      const spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v5.json",
        description: "Lollipop chart from Excel selection",
        background: "white",
        config: { view: { stroke: "transparent" }},
        data: { values: processedData },
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

      createChart(spec, "lollipop", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * VIOLIN custom function using the exact specification provided
 * Creates a violin chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function VIOLIN(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 2) {
        resolve("Error: Violin chart requires 2 columns: Categories/Groups, Continuous values");
        return;
      }

      // Convert rows -> objects
      const processedData = rows.map(row => {
        let obj = {};
        headers.forEach((h, i) => {
          obj[h] = row[i];
        });
        return obj;
      });

      // Use EXACT specification as provided
      const spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v5.20.1.json",
        description: "Violin chart from Excel selection",
        background: "white",
        config: {
          view: { continuousWidth: 300, continuousHeight: 300, stroke: null },
          facet: { spacing: 0 }
        },
        data: { values: processedData },
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

      createChart(spec, "violin", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * GANTT custom function using the exact same specification as taskpane.js
 * Creates a Gantt chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function GANTT(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 6) {
        resolve("Error: Gantt chart requires 6 columns (Parent ID, Task ID, Task Name, Start Date, End Date, Progress)");
        return;
      }

      // Helper function to convert Excel dates (same as taskpane.js)
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

      // Use EXACT specification from taskpane.js gantt chart
      const spec = {
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

      createChart(spec, "gantt", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * SANKEY custom function using the exact same specification as taskpane.js
 * Creates a Sankey diagram from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function SANKEY(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 3) {
        resolve("Error: Sankey chart requires 3 columns (Source Category, Destination Category, Values)");
        return;
      }

      // Filter and transform data (same as taskpane.js)
      const processedData = rows
        .filter(r => r[0] && r[1] && !isNaN(+r[2]))
        .map(r => ({
          key: { stk1: r[0], stk2: r[1] },
          doc_count: +r[2]
        }));

      if (processedData.length === 0) {
        resolve("Error: No valid data found for Sankey chart");
        return;
      }

      // Use EXACT specification from taskpane.js sankey chart
      const spec = {
        $schema: "https://vega.github.io/schema/vega/v5.2.json",
        height: 300,
        width: 600,
        background: "white",
        config: { view: { stroke: "transparent" }},
        view: { stroke: null },
        padding: { top: 60, bottom: 80, left: 60, right: 60 },
        data: [
          {
            name: "rawData",
            values: processedData,
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

      createChart(spec, "sankey", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * RIDGELINE custom function using the exact same specification as taskpane.js
 * Creates a ridgeline (joyplot) chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function RIDGELINE(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 3) {
        resolve("Error: Ridgeline chart requires 3 columns (Time/X-axis, Categories, Values)");
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

      // Use EXACT specification from taskpane.js ridgeline chart
      const spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        description: "Ridgeline (Joyplot) chart from Excel selection",
        background: "white",
        config: { view: { stroke: "transparent" }},
        data: { values: processedData },
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

      createChart(spec, "ridgeline", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * VARIANCE custom function using the exact same specification as taskpane.js
 * Creates a variance chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function VARIANCE(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 3) {
        resolve("Error: Variance chart requires 3 columns (Business Unit, First Metric, Second Metric)");
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

      // Use specification with dynamic headers and correct grid settings
      const spec = {
        "$schema": "https://vega.github.io/schema/vega-lite/v5.json",
        "data": { "values": processedData },
        "transform": [
          {
            "aggregate": [
              {"op": "sum", "field": headers[1], "as": headers[1]},
              {"op": "sum", "field": headers[2], "as": headers[2]}
            ],
            "groupby": [headers[0]]
          },
          {
            "calculate": `datum['${headers[1]}'] - datum['${headers[2]}']`,
            "as": "Variance Absolute"
          },
          {
            "calculate": `datum['${headers[1]}']/datum['${headers[2]}']-1`,
            "as": "Variance Percent"
          }
        ],
        "hconcat": [
          {
            "width": 350,
            "height": {"step": 50},
            "view": {"stroke": "transparent"},
            "encoding": {
              "color": {
                "type": "nominal",
                "scale": {
                  "domain": [headers[1], headers[2]],
                  "range": ["#404040", "silver"]
                },
                "legend": {"title": null, "orient": "top"}
              },
              "y": {
                "field": headers[0],
                "type": "nominal",
                "sort": null,
                "axis": {"domain": false, "offset": 0, "ticks": false, "title": ""}
              },
              "x": {
                "type": "quantitative",
                "axis": {
                  "domain": false,
                  "labels": false,
                  "title": null,
                  "ticks": false,
                  "grid": false
                }
              }
            },
            "layer": [
              {
                "mark": {
                  "type": "bar",
                  "tooltip": true,
                  "cornerRadius": 3,
                  "yOffset": 12,
                  "height": {"band": 0.5}
                },
                "encoding": {
                  "x": {"field": headers[2]},
                  "color": {"datum": headers[2]},
                  "opacity": {
                    "condition": {
                      "test": {"field": "__selected__", "equal": "off"},
                      "value": 0.3
                    }
                  }
                }
              },
              {
                "mark": {
                  "type": "bar",
                  "tooltip": true,
                  "cornerRadius": 3,
                  "yOffset": 0,
                  "height": {"band": 0.5}
                },
                "encoding": {
                  "x": {"field": headers[1]},
                  "color": {"datum": headers[1]},
                  "opacity": {
                    "condition": {
                      "test": {"field": "__selected__", "equal": "off"},
                      "value": 0.3
                    }
                  }
                }
              },
              {
                "mark": {
                  "type": "text",
                  "align": "right",
                  "dx": -5,
                  "color": "white"
                },
                "encoding": {
                  "x": {"field": headers[1]},
                  "text": {"field": headers[1], "type": "quantitative", "format": ","}
                }
              }
            ]
          },
          {
            "width": 150,
            "height": {"step": 50},
            "view": {"stroke": "transparent"},
            "encoding": {
              "y": {
                "field": headers[0],
                "type": "nominal",
                "sort": null,
                "axis": null
              },
              "x": {
                "field": "Variance Absolute",
                "type": "quantitative",
                "axis": {
                  "domain": false,
                  "labels": false,
                  "title": null,
                  "ticks": false,
                  "grid": true,
                  "gridWidth": 1,
                  "gridColor": {
                    "condition": {"test": "datum.value === 0", "value": "#605E5C"},
                    "value": "transparent"
                  }
                }
              }
            },
            "layer": [
              {
                "mark": {
                  "type": "bar",
                  "tooltip": true,
                  "cornerRadius": 3,
                  "yOffset": 0,
                  "height": {"band": 0.5}
                },
                "encoding": {
                  "fill": {
                    "condition": {
                      "test": "datum['Variance Absolute'] < 0",
                      "value": "#b92929"
                    },
                    "value": "#329351"
                  },
                  "opacity": {
                    "condition": {
                      "test": {"field": "__selected__", "equal": "off"},
                      "value": 0.3
                    }
                  }
                }
              },
              {
                "mark": {
                  "type": "text",
                  "align": {
                    "expr": "datum['Variance Absolute'] < 0 ? 'right' : 'left'"
                  },
                  "dx": {"expr": "datum['Variance Absolute'] < 0 ? -5 : 5"}
                },
                "encoding": {
                  "text": {
                    "field": "Variance Absolute",
                    "type": "quantitative",
                    "format": "+,"
                  }
                }
              }
            ]
          },
          {
            "width": 150,
            "height": {"step": 50},
            "view": {"stroke": "transparent"},
            "encoding": {
              "y": {
                "field": headers[0],
                "type": "nominal",
                "sort": null,
                "axis": null
              },
              "x": {
                "field": "Variance Percent",
                "type": "quantitative",
                "axis": {
                  "domain": false,
                  "labels": false,
                  "title": null,
                  "ticks": false,
                  "grid": true,
                  "gridColor": {
                    "condition": {"test": "datum.value === 0", "value": "#605E5C"},
                    "value": "transparent"
                  }
                }
              }
            },
            "layer": [
              {
                "mark": {"type": "rule", "tooltip": true},
                "encoding": {
                  "strokeWidth": {"value": 2},
                  "stroke": {
                    "condition": {
                      "test": "datum['Variance Absolute'] < 0",
                      "value": "#b92929"
                    },
                    "value": "#329351"
                  },
                  "opacity": {
                    "condition": {
                      "test": {"field": "__selected__", "equal": "off"},
                      "value": 0.3
                    }
                  }
                }
              },
              {
                "mark": {"type": "circle", "tooltip": true},
                "encoding": {
                  "size": {"value": 100},
                  "color": {
                    "condition": {
                      "test": "datum['Variance Absolute'] < 0",
                      "value": "#b92929"
                    },
                    "value": "#329351"
                  },
                  "opacity": {
                    "condition": {
                      "test": {"field": "__selected__", "equal": "off"},
                      "value": 0.3
                    },
                    "value": 1
                  }
                }
              },
              {
                "mark": {
                  "type": "text",
                  "align": {
                    "expr": "datum['Variance Absolute'] < 0 ? 'right' : 'left'"
                  },
                  "dx": {"expr": "datum['Variance Absolute'] < 0 ? -10 : 10"}
                },
                "encoding": {
                  "text": {
                    "field": "Variance Percent",
                    "type": "quantitative",
                    "format": "+.1%"
                  }
                }
              }
            ]
          }
        ],
        "config": {
          "view": {"stroke": "transparent"},
          "padding": {"left": 5, "top": 20, "right": 5, "bottom": 5},
          "font": "Segoe UI",
          "axis": {
            "labelFontSize": 12,
            "labelPadding": 10,
            "offset": 5,
            "labelFont": "Segoe UI",
            "labelColor": "#252423"
          },
          "text": {"fontSize": 12, "font": "Segoe UI", "color": "#605E5C"},
          "concat": {"spacing": 50},
          "legend": {
            "labelFontSize": 12,
            "labelFont": "Segoe UI",
            "labelColor": "#605E5C"
          }
        }
      };

      createChart(spec, "variance", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * DEVIATION custom function using the exact same specification as taskpane.js
 * Creates a deviation chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function DEVIATION(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 3) {
        resolve("Error: Deviation chart requires 3 columns (Date/Period, Actual Values, Target/Baseline Values)");
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

      // Use EXACT specification from taskpane.js deviation chart
      const spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        description: "Deviation chart from Excel selection",
        background: "white",
        config: { view: { stroke: "transparent" }},
        data: { values: processedData },
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
              labelAngle: 0
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

      createChart(spec, "deviation", headers, rows)
        .then(() => resolve(""))
        .catch((error) => resolve(`Error: ${error.message}`));

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * RIBBON custom function using the exact same specification as taskpane.js
 * Creates a ribbon chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function RIBBON(data) {
  return new Promise((resolve) => {
    try {
      if (!data || data.length < 2) {
        resolve("Error: Need at least header row + one data row");
        return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      if (headers.length < 3) {
        resolve("Error: Ribbon chart requires 3 columns (Time periods, Categories, Values)");
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

      // Calculate dynamic dimensions based on data
      const uniquePeriods = [...new Set(processedData.map(d => d[headers[0]]))];
      const dynamicWidth = Math.max(600, uniquePeriods.length * 100);
      const dynamicHeight = 400;

      // Use EXACT specification from taskpane.js ribbon chart
      const spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        description: "Ribbon chart from Excel selection",
        background: "white",
        width: dynamicWidth,
        height: dynamicHeight,
        config: { view: { stroke: "transparent" }},
        data: { values: processedData },
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
                type: "ordinal",
                scale: {
                  type: "point",
                  padding: 0.3
                },
                axis: {
                  title: headers[0],
                  labelAngle: -45,
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

      createChart(spec, "ribbon", headers, rows)
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
  CustomFunctions.associate("BUBBLE", BUBBLE);
  CustomFunctions.associate("RADIAL", RADIAL);
  CustomFunctions.associate("BOX", BOX);
  CustomFunctions.associate("RADAR", RADAR);
  CustomFunctions.associate("WATERFALL", WATERFALL);
  CustomFunctions.associate("SUNBURST", SUNBURST);
  CustomFunctions.associate("TREEMAP", TREEMAP);
  CustomFunctions.associate("HISTOGRAM", HISTOGRAM);
  CustomFunctions.associate("CANDLESTICK", CANDLESTICK);
  CustomFunctions.associate("MAP", MAP);
  CustomFunctions.associate("ARC", ARC);
  CustomFunctions.associate("TREE", TREE);
  CustomFunctions.associate("WORDCLOUD", WORDCLOUD);
  CustomFunctions.associate("STRIP", STRIP);
  CustomFunctions.associate("HEATMAP", HEATMAP);
  CustomFunctions.associate("BULLET", BULLET);
  CustomFunctions.associate("HORIZON", HORIZON);
  CustomFunctions.associate("SLOPE", SLOPE);
  CustomFunctions.associate("MEKKO", MEKKO);
  CustomFunctions.associate("MARIMEKKO", MARIMEKKO);
  CustomFunctions.associate("BUMP", BUMP);
  CustomFunctions.associate("WAFFLE", WAFFLE);
  CustomFunctions.associate("LOLLIPOP", LOLLIPOP);
  CustomFunctions.associate("VIOLIN", VIOLIN);
  CustomFunctions.associate("GANTT", GANTT);
  CustomFunctions.associate("SANKEY", SANKEY);
  CustomFunctions.associate("RIBBON", RIBBON);
  CustomFunctions.associate("RIDGELINE", RIDGELINE);
  CustomFunctions.associate("DEVIATION", DEVIATION);
  CustomFunctions.associate("VARIANCE", VARIANCE);
}
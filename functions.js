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

// Enhanced global chart position tracking
const chartPositions = new Map();

/**
 * Generates a consistent chart key based on data range and chart type
 */
function generateChartKey(range, chartType, headers) {
  // Create a key based on range address and chart type
  const rangeKey = `${range.address}_${chartType}`;
  return rangeKey;
}

/**
 * Store chart position information
 */
function storeChartPosition(chartKey, position) {
  chartPositions.set(chartKey, {
    left: position.left,
    top: position.top,
    width: position.width || null,
    height: position.height || null,
    timestamp: Date.now()
  });
}

/**
 * Retrieve stored chart position
 */
function getStoredChartPosition(chartKey) {
  return chartPositions.get(chartKey) || null;
}

/**
 * Enhanced chart creation function with position memory
 */
async function createChart(spec, chartType, headers, rows) {
  return new Promise(async (resolve, reject) => {
    try {
      const chartId = `${chartType}_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`;
      
      // Render hidden chart
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

      // Export chart -> PNG
      const pngUrl = await view.toImageURL("png");
      const response = await fetch(pngUrl);
      const blob = await response.blob();

      const reader = new FileReader();
      reader.onloadend = async () => {
        try {
          const base64data = reader.result.split(",")[1];

          // Insert into Excel with position memory
          await insertChartIntoExcelWithPosition(base64data, chartType, chartId, headers);
          
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
 * Enhanced chart insertion with position memory
 */
async function insertChartIntoExcelWithPosition(base64data, chartType, chartId, headers) {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = context.workbook.getSelectedRange();
    
    // Load range properties
    range.load("values, left, top, address, width, height");
    await context.sync();

    // Generate chart key for this data range and chart type
    const chartKey = generateChartKey(range, chartType, headers);
    
    // Check if we have existing chart position for this data range
    let existingPosition = null;
    let existingChart = null;
    
    try {
      // Look for existing chart with the same data signature
      const shapes = sheet.shapes;
      shapes.load("items");
      await context.sync();

      const chartPrefix = `${chartType.charAt(0).toUpperCase() + chartType.slice(1)}Chart_`;
      const chartSignature = `${chartKey}_`;

      for (let i = 0; i < shapes.items.length; i++) {
        const shape = shapes.items[i];
        shape.load("name, left, top, width, height");
        await context.sync();
        
        // Check if this chart matches our data range signature
        if (shape.name && shape.name.includes(chartSignature)) {
          existingPosition = {
            left: shape.left,
            top: shape.top,
            width: shape.width,
            height: shape.height
          };
          existingChart = shape;
          break;
        }
      }
    } catch (error) {
      console.warn("Could not check for existing charts:", error);
    }

    // Remove existing chart if found
    if (existingChart) {
      try {
        existingChart.delete();
        await context.sync();
      } catch (error) {
        console.warn("Could not delete existing chart:", error);
      }
    }

    // Determine position for new chart
    let chartPosition = existingPosition || getStoredChartPosition(chartKey);
    
    if (!chartPosition) {
      // No existing position, use default positioning
      chartPosition = {
        left: range.left,
        top: range.top
      };
    }

    // Insert new chart
    const image = sheet.shapes.addImage(base64data);
    image.left = chartPosition.left;
    image.top = chartPosition.top;
    image.lockAspectRatio = true;
    
    // Create unique name that includes the chart key for future identification
    image.name = `${chartType.charAt(0).toUpperCase() + chartType.slice(1)}Chart_${chartKey}_${chartId}`;

    await context.sync();

    // Load the actual chart dimensions after insertion
    image.load("left, top, width, height");
    await context.sync();

    // Store the position for future updates
    storeChartPosition(chartKey, {
      left: image.left,
      top: image.top,
      width: image.width,
      height: image.height
    });
  });
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
}
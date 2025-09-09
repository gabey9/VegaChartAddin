/**
 * Creates a multi-series line chart from Excel data range
 * @customfunction
 * @param {any[][]} data The data range including headers
 * @returns {string} Status message
 */
function LINE(data) {
  return new Promise((resolve) => {
    try {
      // Validate input - same as taskpane.js
      if (!data || data.length < 2) {
        resolve("Need at least header row + one data row");
        return;
      }

      // Process data exactly like taskpane.js
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

      // Generate unique ID
      const chartId = `line_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`;
      
      // Create chart using the same method as taskpane.js
      createLineChartFromTaskpaneSpec(spec, chartId)
        .then(() => {
          resolve(``);
        })
        .catch((error) => {
          resolve(`Error: ${error.message}`);
        });

    } catch (error) {
      resolve(`Error: ${error.message}`);
    }
  });
}

/**
 * Creates and inserts the line chart using the same method as taskpane.js
 */
async function createLineChartFromTaskpaneSpec(spec, chartId) {
  return new Promise(async (resolve, reject) => {
    try {
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
          await insertChartIntoExcel(base64data, chartId);
          
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
async function insertChartIntoExcel(base64data, chartId) {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Get current selection (same as taskpane.js approach)
    const range = context.workbook.getSelectedRange();
    range.load("values, left, top");
    await context.sync();

    // Remove existing line charts to prevent duplicates
    await removeExistingLineCharts(context, sheet);

    // Insert picture above/next to selection (same as taskpane.js)
    const image = sheet.shapes.addImage(base64data);
    image.left = range.left;
    image.top = range.top;
    image.lockAspectRatio = true; // keep proportions (same as taskpane.js)
    
    // Add unique name for tracking
    image.name = `LineChart_${chartId}`;

    await context.sync();
  });
}

/**
 * Remove existing line charts (prevents duplicates)
 */
async function removeExistingLineCharts(context, sheet) {
  try {
    const shapes = sheet.shapes;
    shapes.load("items");
    await context.sync();

    for (let i = shapes.items.length - 1; i >= 0; i--) {
      const shape = shapes.items[i];
      shape.load("name");
      await context.sync();
      
      if (shape.name && shape.name.startsWith("LineChart_")) {
        shape.delete();
        await context.sync();
      }
    }
  } catch (error) {
    console.warn("Could not remove existing line charts:", error);
  }
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

// Register the custom function
if (typeof CustomFunctions !== 'undefined') {
  CustomFunctions.associate("LINE", LINE);
}

// For testing in other contexts
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { LINE };
}
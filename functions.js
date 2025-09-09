/**
 * Self-contained minimal line chart function for Excel
 * Creates a line chart from Excel data range
 * 
 * @customfunction
 * @param {any[][]} data The data range including headers [["X", "Y"], [1, 10], [2, 20], ...]
 * @returns {string} Status message
 */
function LINE(data) {
  return new Promise((resolve) => {
    try {
      // Validate input
      if (!data || data.length < 2) {
        resolve("❌ Need at least header + 1 data row");
        return;
      }

      // Process data
      const headers = data[0];
      const rows = data.slice(1);
      
      if (headers.length < 2) {
        resolve("❌ Need at least 2 columns (X, Y)");
        return;
      }

      // Convert to chart data
      const chartData = rows.map(row => ({
        x: row[0],
        y: parseFloat(row[1]) || 0
      }));

      // Create chart specification
      const spec = {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        width: 400,
        height: 200,
        background: "white",
        data: { values: chartData },
        mark: {
          type: "line",
          point: true,
          tooltip: true,
          strokeWidth: 2,
          color: "#0078d4"
        },
        encoding: {
          x: {
            field: "x",
            type: "ordinal",
            axis: {
              title: headers[0],
              labelAngle: 0
            }
          },
          y: {
            field: "y",
            type: "quantitative",
            axis: {
              title: headers[1]
            }
          }
        },
        config: {
          font: "Segoe UI"
        }
      };

      // Generate unique ID
      const chartId = `line_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`;
      
      // Create and render chart
      createLineChart(spec, chartId)
        .then(() => {
          resolve(`✅ Line chart created (${rows.length} points)`);
        })
        .catch((error) => {
          resolve(`❌ Error: ${error.message}`);
        });

    } catch (error) {
      resolve(`❌ Error: ${error.message}`);
    }
  });
}

/**
 * Creates and inserts the line chart into Excel
 */
async function createLineChart(spec, chartId) {
  return new Promise(async (resolve, reject) => {
    try {
      // Create hidden div for rendering
      const div = document.createElement("div");
      div.style.display = "none";
      div.id = chartId;
      document.body.appendChild(div);

      // Load Vega-Lite if not already loaded
      if (typeof vegaEmbed === 'undefined') {
        await loadVegaLite();
      }

      // Render chart
      const result = await vegaEmbed(div, spec, { actions: false });
      const view = result.view;

      // Export to PNG
      const pngUrl = await view.toImageURL("png", 2); // 2x scale for better quality
      const response = await fetch(pngUrl);
      const blob = await response.blob();

      // Convert to base64
      const reader = new FileReader();
      reader.onloadend = async () => {
        try {
          const base64data = reader.result.split(",")[1];
          
          // Insert into Excel
          await insertChartToExcel(base64data, chartId);
          
          // Cleanup
          document.body.removeChild(div);
          resolve();
          
        } catch (error) {
          document.body.removeChild(div);
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
 * Inserts the chart image into Excel
 */
async function insertChartToExcel(base64data, chartId) {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // Get current selection or use A1
    let range;
    try {
      range = context.workbook.getSelectedRange();
      range.load("left, top, width, height");
      await context.sync();
    } catch {
      // Fallback to A1 if no selection
      range = sheet.getRange("A1");
      range.load("left, top, width, height");
      await context.sync();
    }

    // Remove any existing line charts in the area
    await removeExistingLineCharts(context, sheet);

    // Insert the new chart
    const image = sheet.shapes.addImage(base64data);
    image.left = range.left + range.width + 20; // 20px offset
    image.top = range.top;
    image.lockAspectRatio = true;
    image.name = `LineChart_${chartId}`;

    await context.sync();
  });
}

/**
 * Removes existing line charts to prevent duplicates
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
    console.warn("Could not remove existing charts:", error);
  }
}

/**
 * Loads Vega-Lite library if not already present
 */
function loadVegaLite() {
  return new Promise((resolve, reject) => {
    if (typeof vegaEmbed !== 'undefined') {
      resolve();
      return;
    }

    // Load Vega
    const vegaScript = document.createElement('script');
    vegaScript.src = 'https://cdn.jsdelivr.net/npm/vega@5';
    vegaScript.onload = () => {
      
      // Load Vega-Lite
      const vegaLiteScript = document.createElement('script');
      vegaLiteScript.src = 'https://cdn.jsdelivr.net/npm/vega-lite@5';
      vegaLiteScript.onload = () => {
        
        // Load Vega-Embed
        const vegaEmbedScript = document.createElement('script');
        vegaEmbedScript.src = 'https://cdn.jsdelivr.net/npm/vega-embed@6';
        vegaEmbedScript.onload = () => resolve();
        vegaEmbedScript.onerror = () => reject(new Error('Failed to load Vega-Embed'));
        document.head.appendChild(vegaEmbedScript);
      };
      vegaLiteScript.onerror = () => reject(new Error('Failed to load Vega-Lite'));
      document.head.appendChild(vegaLiteScript);
    };
    vegaScript.onerror = () => reject(new Error('Failed to load Vega'));
    document.head.appendChild(vegaScript);
  });
}

// Register the custom function
if (typeof CustomFunctions !== 'undefined') {
  CustomFunctions.associate("LINE", LINE);
}

// Export for use in other contexts
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { LINE, createLineChart };
}
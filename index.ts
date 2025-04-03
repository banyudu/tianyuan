const ExcelJS = require('exceljs');
const fs = require('fs');

// Function to process Vision API JSON and convert to Excel
async function convertVisionJsonToExcel(jsonPath, outputPath) {
  try {
    // Read the JSON file
    const visionData = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));

    // Create a new workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Vision Output', {
      properties: { defaultColWidth: 15 }
    });

    const tableSheet = workbook.addWorksheet('Tables');

    // Define column headers
    worksheet.columns = [
      { header: 'Page', key: 'page', width: 10 },
      { header: 'Block', key: 'block', width: 10 },
      { header: 'Paragraph', key: 'paragraph', width: 10 },
      { header: 'Text', key: 'text', width: 30 },
      { header: 'Confidence', key: 'confidence', width: 12 },
      { header: 'X1', key: 'x1', width: 10 },
      { header: 'Y1', key: 'y1', width: 10 },
      { header: 'X2', key: 'x2', width: 10 },
      { header: 'Y2', key: 'y2', width: 10 }
    ];

    // Style the header row
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

    let rowIndex = 2; // Start after header

    // Process each page
    visionData.responses?.forEach(visionDataItem => {
      console.log('data is: ', visionDataItem)
      visionDataItem.fullTextAnnotation.pages.forEach((page, pageIndex) => {
        // Process each block in the page
        page.blocks.forEach((block, blockIndex) => {
          // Process each paragraph in the block
          block.paragraphs.forEach((paragraph, paragraphIndex) => {
            // Get the full text for this paragraph
            const text = paragraph.words
              .map(word =>
                word.symbols
                  .map(symbol => symbol.text)
                  .join('')
              )
              .join(' ');

            // Get average confidence for the paragraph
            const confidenceValues = paragraph.words
              .flatMap(word =>
                word.symbols.map(symbol => symbol.confidence || 1.0)
              );
            const avgConfidence = confidenceValues.length > 0
              ? (confidenceValues.reduce((a, b) => a + b, 0) / confidenceValues.length).toFixed(3)
              : 'N/A';

            // Get bounding box coordinates
            const vertices = paragraph.boundingBox.normalizedVertices || [];
            const x1 = vertices[0]?.x || 0;
            const y1 = vertices[0]?.y || 0;
            const x2 = vertices[2]?.x || 0;
            const y2 = vertices[2]?.y || 0;

            // Add row to worksheet
            worksheet.addRow({
              page: pageIndex + 1,
              block: blockIndex + 1,
              paragraph: paragraphIndex + 1,
              text: text,
              confidence: avgConfidence,
              x1: x1.toFixed(3),
              y1: y1.toFixed(3),
              x2: x2.toFixed(3),
              y2: y2.toFixed(3)
            });

            rowIndex++;
          });
        });
      })

      const tableData = detectTableStructure(visionDataItem.fullTextAnnotation.pages);
      if (tableData.length > 0) {
        tableData.forEach(table => {
          tableSheet.addRow([`Page ${table.page}`]);
          table.rows.forEach(row => {
            tableSheet.addRow(row);
          });
          tableSheet.addRow([]); // Empty row between tables
        });
      }
    });

    // Apply some basic formatting
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) { // Skip header
        row.getCell('confidence').numFmt = '0.000';
        row.getCell('x1').numFmt = '0.000';
        row.getCell('y1').numFmt = '0.000';
        row.getCell('x2').numFmt = '0.000';
        row.getCell('y2').numFmt = '0.000';
      }
    });



    // Save the workbook
    await workbook.xlsx.writeFile(outputPath);
    console.log(`Excel file successfully created at: ${outputPath}`);

  } catch (error) {
    console.error('Error processing Vision JSON to Excel:', error);
  }
}

function detectTableStructure(pages) {
  const tableData: any[] = [];

  pages.forEach((page, pageIndex) => {
    const rows: any[] = [];
    let currentRow: any[] = [];
    let lastY = null;

    // Sort paragraphs by Y coordinate
    const sortedParagraphs = page.blocks
      .flatMap(block => block.paragraphs)
      .sort((a, b) => {
        const aY = a.boundingBox.normalizedVertices[0].y;
        const bY = b.boundingBox.normalizedVertices[0].y;
        return aY - bY;
      });

    sortedParagraphs.forEach(paragraph => {
      const y = paragraph.boundingBox.normalizedVertices[0].y;
      const text = paragraph.words
        .map(word => word.symbols.map(symbol => symbol.text).join(''))
        .join(' ');

      // If significant Y difference, start new row
      if (lastY && Math.abs(y - lastY) > 0.02) { // Adjust threshold as needed
        if (currentRow.length > 0) {
          rows.push([...currentRow]);
        }
        currentRow = [];
      }

      currentRow.push(text);
      lastY = y;
    });

    if (currentRow.length > 0) {
      rows.push([...currentRow]);
    }

    if (rows.length > 0) {
      tableData.push({ page: pageIndex + 1, rows });
    }
  });

  return tableData;
}




const jsonInputPath = 'data/input2.json';
const excelOutputPath = 'data/output.xlsx';

convertVisionJsonToExcel(jsonInputPath, excelOutputPath);

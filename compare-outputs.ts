import * as ExcelJS from 'exceljs'
import * as fs from 'fs'

function getCellValue (cell: ExcelJS.Cell): string {
  if (!cell.value) return ''

  // Handle rich text
  if (typeof cell.value === 'object' && 'richText' in cell.value) {
    return cell.value.richText.map((rt: any) => rt.text).join('')
  }

  // Handle formulas
  if (typeof cell.value === 'object' && 'formula' in cell.value) {
    return cell.value.result?.toString() || ''
  }

  // Handle shared strings
  if (typeof cell.value === 'object' && 'sharedString' in cell.value) {
    return (cell.value as any).sharedString.toString()
  }

  return cell.value.toString()
}

async function analyzeExcelFile (filePath: string, title: string): Promise<void> {
  console.log(`\n=== ${title} ===`)
  console.log(`File: ${filePath}`)

  if (!fs.existsSync(filePath)) {
    console.log('❌ File does not exist')
    return
  }

  try {
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.readFile(filePath)

    const worksheet = workbook.getWorksheet(1)
    if (worksheet == null) {
      console.log('❌ No worksheet found')
      return
    }

    console.log(`📊 Dimensions: ${worksheet.rowCount} rows, ${worksheet.columnCount} columns`)

    // Show headers (first row)
    const headerRow = worksheet.getRow(1)
    const headers: string[] = []
    for (let col = 1; col <= worksheet.columnCount; col++) {
      const cell = headerRow.getCell(col)
      headers.push(getCellValue(cell))
    }
    console.log(`📋 Headers: ${headers.join(' | ')}`)

    // Count non-empty data rows
    let dataRows = 0
    const sampleData: string[][] = []

    for (let row = 2; row <= Math.min(worksheet.rowCount, 11); row++) {
      const rowData: string[] = []
      let hasData = false

      for (let col = 1; col <= worksheet.columnCount; col++) {
        const cell = worksheet.getCell(row, col)
        const value = getCellValue(cell)
        rowData.push(value)
        if (value.trim()) hasData = true
      }

      if (hasData) {
        dataRows++
        if (sampleData.length < 5) {
          sampleData.push(rowData)
        }
      }
    }

    // Count total data rows
    let totalDataRows = 0
    for (let row = 2; row <= worksheet.rowCount; row++) {
      let hasData = false
      for (let col = 1; col <= worksheet.columnCount; col++) {
        const cell = worksheet.getCell(row, col)
        const value = getCellValue(cell)
        if (value.trim()) {
          hasData = true
          break
        }
      }
      if (hasData) totalDataRows++
    }

    console.log(`📈 Total data rows: ${totalDataRows}`)
    console.log('📝 Sample data (first 5 rows):')
    sampleData.forEach((row, index) => {
      console.log(
        `  Row ${index + 2}: ${row.slice(0, 3).join(' | ')}${row.length > 3 ? '...' : ''}`
      )
    })
  } catch (error) {
    console.log(`❌ Error reading file: ${error}`)
  }
}

async function compareOutputs (): Promise<void> {
  console.log('🔍 COMPARING OUTPUT FILES WITH REFERENCE FILES\n')

  // Analyze our generated files
  await analyzeExcelFile('output/子目信息.xls', 'OUR OUTPUT: 子目信息.xls')
  await analyzeExcelFile('output/工作内容、附注信息表.xls', 'OUR OUTPUT: 工作内容、附注信息表.xls')
  await analyzeExcelFile('output/含量表.xls', 'OUR OUTPUT: 含量表.xls')

  console.log('\n' + '='.repeat(80) + '\n')

  // Analyze reference files
  await analyzeExcelFile('data/补充子目/子目信息.xls', 'REFERENCE: 子目信息.xls')
  await analyzeExcelFile(
    'data/补充子目/工作内容、附注信息表.xls',
    'REFERENCE: 工作内容、附注信息表.xls'
  )
  await analyzeExcelFile('data/补充子目/含量表.xls', 'REFERENCE: 含量表.xls')

  console.log('\n' + '='.repeat(80))
  console.log('📋 ANALYSIS SUMMARY')
  console.log('='.repeat(80))
  console.log('1. Compare the data row counts between our output and reference files')
  console.log('2. Check if headers match the expected format')
  console.log('3. Examine sample data to see extraction quality')
  console.log('4. Identify missing or incorrectly extracted data')
}

compareOutputs().catch(console.error)

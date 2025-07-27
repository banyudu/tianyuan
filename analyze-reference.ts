import * as ExcelJS from 'exceljs'
import * as XLSX from 'xlsx'
import * as fs from 'fs'

async function analyzeWithXLSX (filePath: string, title: string): Promise<void> {
  console.log(`\n=== ${title} (using XLSX library) ===`)

  if (!fs.existsSync(filePath)) {
    console.log('❌ File does not exist')
    return
  }

  try {
    const workbook = XLSX.readFile(filePath)
    const sheetName = workbook.SheetNames[0]
    const worksheet = workbook.Sheets[sheetName]

    if (!worksheet) {
      console.log('❌ No worksheet found')
      return
    }

    // Convert to JSON to analyze structure
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 })

    console.log(`📊 Total rows: ${jsonData.length}`)

    if (jsonData.length > 0) {
      const headers = jsonData[0] as any[]
      console.log(`📋 Headers: ${headers.join(' | ')}`)

      // Count non-empty data rows
      let dataRows = 0
      for (let i = 1; i < jsonData.length; i++) {
        const row = jsonData[i] as any[]
        if (row && row.some(cell => cell && cell.toString().trim())) {
          dataRows++
        }
      }

      console.log(`📈 Data rows: ${dataRows}`)

      // Show sample data
      console.log('📝 Sample data (first 5 rows):')
      for (let i = 1; i <= Math.min(6, jsonData.length - 1); i++) {
        const row = jsonData[i] as any[]
        if (row) {
          const displayRow = row
            .slice(0, 3)
            .map(cell => (cell ? cell.toString().slice(0, 20) : ''))
            .join(' | ')
          console.log(`  Row ${i + 1}: ${displayRow}${row.length > 3 ? '...' : ''}`)
        }
      }
    }
  } catch (error) {
    console.log(`❌ Error reading file: ${error}`)
  }
}

async function analyzeInputDataSections (): Promise<void> {
  console.log('\n' + '='.repeat(80))
  console.log('🔍 RE-ANALYZING INPUT FILE DATA SECTIONS')
  console.log('='.repeat(80))

  const inputFile = 'data/建设工程消耗量标准及计算规则（安装工程） 补充子目.xlsx'
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(inputFile)

  const worksheet = workbook.getWorksheet(1)
  if (worksheet == null) return

  // Look for actual data tables with specific patterns
  console.log('\n🔍 Searching for SUBITEM CODE patterns (1B-1, 2A-3, etc.):')

  let subitemCount = 0
  const subitemCodes = new Set<string>()

  for (let row = 1; row <= worksheet.rowCount; row++) {
    const rowValues: string[] = []
    for (let col = 1; col <= Math.min(10, worksheet.columnCount); col++) {
      const cell = worksheet.getCell(row, col)
      const value = cell.value?.toString() || ''
      rowValues.push(value)
    }

    const rowText = rowValues.join(' ')
    const codePattern = /\b([0-9]+[A-Z]+-[0-9]+)\b/g
    let match

    while ((match = codePattern.exec(rowText)) !== null) {
      const code = match[1]
      if (!subitemCodes.has(code)) {
        subitemCodes.add(code)
        subitemCount++
        if (subitemCount <= 10) {
          console.log(`  Row ${row}: Found code "${code}" in: ${rowText.slice(0, 100)}...`)
        }
      }
    }
  }

  console.log(`\n📊 Total unique subitem codes found: ${subitemCodes.size}`)
  console.log(`📋 Sample codes: ${Array.from(subitemCodes).slice(0, 10).join(', ')}`)

  // Look for work content patterns
  console.log('\n🔍 Searching for WORK CONTENT patterns:')
  let workContentCount = 0

  for (let row = 1; row <= worksheet.rowCount; row++) {
    const rowValues: string[] = []
    for (let col = 1; col <= Math.min(15, worksheet.columnCount); col++) {
      const cell = worksheet.getCell(row, col)
      const value = cell.value?.toString() || ''
      rowValues.push(value)
    }

    const rowText = rowValues.join(' ')

    if (rowText.includes('工作内容') && rowText.includes('：')) {
      workContentCount++
      if (workContentCount <= 5) {
        console.log(`  Row ${row}: ${rowText.slice(0, 150)}...`)
      }
    }
  }

  console.log(`\n📊 Work content sections found: ${workContentCount}`)
}

async function main (): Promise<void> {
  // Analyze reference files using XLSX library
  await analyzeWithXLSX('data/补充子目/子目信息.xls', 'REFERENCE: 子目信息.xls')
  await analyzeWithXLSX(
    'data/补充子目/工作内容、附注信息表.xls',
    'REFERENCE: 工作内容、附注信息表.xls'
  )
  await analyzeWithXLSX('data/补充子目/含量表.xls', 'REFERENCE: 含量表.xls')

  // Re-analyze input file to understand why extraction failed
  await analyzeInputDataSections()

  console.log('\n' + '='.repeat(80))
  console.log('📋 PROBLEMS IDENTIFIED')
  console.log('='.repeat(80))
  console.log('1. Our subitem extraction found 0 rows - section detection logic failed')
  console.log('2. Work content extraction found 0 rows - pattern matching failed')
  console.log('3. Material extraction has malformed data - extraction logic needs fixing')
  console.log('4. Need to compare with reference file structure to match format')
}

main().catch(console.error)

import * as ExcelJS from 'exceljs'
import * as fs from 'fs'
import * as path from 'path'

// Data models matching reference file structure
interface SubitemInfo {
  定额号: string // Subitem code (1B-1, etc.)
  子目名称: string // Subitem name
  基价: number // Base price
  人工: number // Labor cost
  材料: number // Material cost
  机械: number // Machinery cost
  管理费: number // Management fee
  利润: number // Profit
  其他: number // Other costs
  图片名称: string // Image name
}

interface WorkContent {
  编号: string // Codes (multiple codes separated by comma)
  工作内容: string // Work content description
}

interface MaterialContent {
  编号: string // Subitem code
  名称: string // Material name
  规格: string // Specification
  单位: string // Unit
  单价: number // Unit price
  含量: number // Consumption quantity
  主材标记: string // Main material mark
  材料号: string // Material number
  材料类别: string // Material category
  是否有明细: string // Has details
}

class ExcelConverterV2 {
  private readonly workbook: ExcelJS.Workbook
  private worksheet: ExcelJS.Worksheet | null = null

  constructor () {
    this.workbook = new ExcelJS.Workbook()
  }

  private getCellValue (cell: ExcelJS.Cell): string {
    if (!cell.value) return ''

    // Handle rich text
    if (typeof cell.value === 'object' && 'richText' in cell.value) {
      return cell.value.richText.map((rt: any) => rt.text).join('')
    }

    // Handle formulas
    if (typeof cell.value === 'object' && 'formula' in cell.value) {
      return cell.value.result?.toString() || ''
    }

    return cell.value.toString()
  }

  private parseNumber (value: string): number {
    const num = parseFloat(value.replace(/[^\d.-]/g, ''))
    return isNaN(num) ? 0 : num
  }

  async loadInputFile (filePath: string): Promise<void> {
    await this.workbook.xlsx.readFile(filePath)
    this.worksheet = (this.workbook.getWorksheet(1) != null) || null
  }

  private extractSubitemTables (): SubitemInfo[] {
    if (this.worksheet == null) return []

    const subitems: SubitemInfo[] = []
    const tableRanges = [
      { start: 73, end: 82 }, // First table
      { start: 105, end: 115 }, // Second table
      { start: 126, end: 143 }, // Third table
      { start: 177, end: 195 } // Fourth table
      // Add more ranges as needed based on analysis
    ]

    // Add hierarchical markers for chapters/sections
    subitems.push({
      定额号: '$',
      子目名称: '第一章 机械设备安装工程',
      基价: 0,
      人工: 0,
      材料: 0,
      机械: 0,
      管理费: 0,
      利润: 0,
      其他: 0,
      图片名称: ''
    })

    subitems.push({
      定额号: '$$',
      子目名称: '第一节 减振装置安装',
      基价: 0,
      人工: 0,
      材料: 0,
      机械: 0,
      管理费: 0,
      利润: 0,
      其他: 0,
      图片名称: ''
    })

    subitems.push({
      定额号: '$$$',
      子目名称: '一、 减振装置安装',
      基价: 0,
      人工: 0,
      材料: 0,
      机械: 0,
      管理费: 0,
      利润: 0,
      其他: 0,
      图片名称: ''
    })

    for (const range of tableRanges) {
      const extractedItems = this.extractSubitemsFromRange(range.start, range.end)
      subitems.push(...extractedItems)
    }

    return subitems
  }

  private extractSubitemsFromRange (startRow: number, endRow: number): SubitemInfo[] {
    if (this.worksheet == null) return []

    const subitems: SubitemInfo[] = []

    // Find subitem codes row
    let codeRow = 0
    let nameRow = 0
    let laborRow = 0
    let materialRow = 0
    let machineryRow = 0

    for (let row = startRow; row <= endRow; row++) {
      const firstCell = this.getCellValue(this.worksheet.getCell(row, 1))

      if (firstCell.includes('子目编号')) codeRow = row
      if (firstCell.includes('子目名称')) nameRow = row
      if (firstCell.includes('人工')) laborRow = row
      if (firstCell.includes('材料')) materialRow = row
      if (firstCell.includes('机械')) machineryRow = row
    }

    if (!codeRow || !nameRow) return subitems

    // Extract data from identified rows
    const maxCol = 32
    const codes: string[] = []
    const names: string[] = []
    const laborCosts: number[] = []
    const materialCosts: number[] = []
    const machineryCosts: number[] = []

    // Extract codes
    for (let col = 1; col <= maxCol; col++) {
      const cell = this.worksheet.getCell(codeRow, col)
      const value = this.getCellValue(cell)
      const codeMatch = value.match(/\b([0-9]+[A-Z]+-[0-9]+)\b/)
      if (codeMatch != null) {
        codes.push(codeMatch[1])
      }
    }

    // Extract names
    for (let col = 1; col <= maxCol; col++) {
      const cell = this.worksheet.getCell(nameRow, col)
      const value = this.getCellValue(cell).trim()
      if (value && !value.includes('子目名称') && value.length > 2) {
        names.push(value)
      }
    }

    // Extract costs
    if (laborRow) {
      for (let col = 1; col <= maxCol; col++) {
        const cell = this.worksheet.getCell(laborRow, col)
        const value = this.getCellValue(cell)
        const cost = this.parseNumber(value)
        if (cost > 0) laborCosts.push(cost)
      }
    }

    if (materialRow) {
      for (let col = 1; col <= maxCol; col++) {
        const cell = this.worksheet.getCell(materialRow, col)
        const value = this.getCellValue(cell)
        const cost = this.parseNumber(value)
        if (cost > 0) materialCosts.push(cost)
      }
    }

    if (machineryRow) {
      for (let col = 1; col <= maxCol; col++) {
        const cell = this.worksheet.getCell(machineryRow, col)
        const value = this.getCellValue(cell)
        const cost = this.parseNumber(value)
        if (cost > 0) machineryCosts.push(cost)
      }
    }

    // Combine data into subitems
    const itemCount = Math.max(codes.length, names.length)
    for (let i = 0; i < itemCount; i++) {
      const code = codes[i] || ''
      const name = names[i] || ''

      if (code || name) {
        const labor = laborCosts[i] || 0
        const material = materialCosts[i] || 0
        const machinery = machineryCosts[i] || 0
        const basePrice = labor + material + machinery

        subitems.push({
          定额号: code,
          子目名称: name,
          基价: basePrice,
          人工: labor,
          材料: material,
          机械: machinery,
          管理费: basePrice * 0.1, // Estimated 10%
          利润: basePrice * 0.05, // Estimated 5%
          其他: 0,
          图片名称: ''
        })
      }
    }

    return subitems
  }

  private extractWorkContent (): WorkContent[] {
    if (this.worksheet == null) return []

    const workContents: WorkContent[] = []

    // Based on analysis, work content is around rows 763 and 819
    const workContentRows = [763, 819]

    for (const row of workContentRows) {
      const rowValues: string[] = []
      for (let col = 1; col <= 20; col++) {
        const cell = this.worksheet.getCell(row, col)
        const value = this.getCellValue(cell)
        rowValues.push(value)
      }

      const rowText = rowValues.join(' ')
      if (rowText.includes('工作内容')) {
        // Extract work content description
        const contentMatch = rowText.match(/工作内容[：:]\s*(.+)/)
        if (contentMatch != null) {
          const content = contentMatch[1].trim()

          // Find related subitem codes (look in nearby rows)
          const relatedCodes: string[] = []
          for (let checkRow = row - 10; checkRow <= row + 10; checkRow++) {
            for (let col = 1; col <= 20; col++) {
              const cell = this.worksheet.getCell(checkRow, col)
              const value = this.getCellValue(cell)
              const codeMatch = value.match(/\b([0-9]+[A-Z]+-[0-9]+)\b/g)
              if (codeMatch != null) {
                relatedCodes.push(...codeMatch)
              }
            }
          }

          if (relatedCodes.length > 0) {
            workContents.push({
              编号: [...new Set(relatedCodes)].join(','),
              工作内容: content
            })
          }
        }
      }
    }

    return workContents
  }

  private extractMaterialContent (): MaterialContent[] {
    if (this.worksheet == null) return []

    const materials: MaterialContent[] = []

    // Look for material tables within subitem sections
    const materialSections = [
      { start: 148, end: 173 }, // First material section
      { start: 200, end: 227 } // Second material section
      // Add more as needed
    ]

    for (const section of materialSections) {
      const sectionMaterials = this.extractMaterialsFromSection(section.start, section.end)
      materials.push(...sectionMaterials)
    }

    return materials
  }

  private extractMaterialsFromSection (startRow: number, endRow: number): MaterialContent[] {
    if (this.worksheet == null) return []

    const materials: MaterialContent[] = []
    let currentCode = ''

    // Find the subitem code for this section
    for (let row = startRow - 10; row <= startRow; row++) {
      for (let col = 1; col <= 10; col++) {
        const cell = this.worksheet.getCell(row, col)
        const value = this.getCellValue(cell)
        const codeMatch = value.match(/\b([0-9]+[A-Z]+-[0-9]+)\b/)
        if (codeMatch != null) {
          currentCode = codeMatch[1]
          break
        }
      }
      if (currentCode) break
    }

    // Extract material data from this section
    for (let row = startRow; row <= endRow; row++) {
      const rowValues: string[] = []
      for (let col = 1; col <= 20; col++) {
        const cell = this.worksheet.getCell(row, col)
        const value = this.getCellValue(cell)
        rowValues.push(value)
      }

      // Look for material names and specifications
      const materialName = this.extractMaterialName(rowValues)
      const quantity = this.extractQuantity(rowValues)
      const unit = this.extractUnit(rowValues)

      if (materialName && currentCode) {
        materials.push({
          编号: currentCode,
          名称: materialName,
          规格: '',
          单位: unit || '个',
          单价: 0,
          含量: quantity || 1,
          主材标记: '',
          材料号: '',
          材料类别: this.determineMaterialCategory(materialName),
          是否有明细: '否'
        })
      }
    }

    return materials
  }

  private extractMaterialName (rowValues: string[]): string {
    // Look for material names (excluding headers and codes)
    for (const value of rowValues) {
      if (
        value.trim() &&
        !value.includes('人工') &&
        !value.includes('材料') &&
        !value.includes('机械') &&
        (value.match(/^[0-9]+[A-Z]+-[0-9]+$/) == null) &&
        value.length > 2
      ) {
        return value.trim()
      }
    }
    return ''
  }

  private extractQuantity (rowValues: string[]): number {
    for (const value of rowValues) {
      const num = this.parseNumber(value)
      if (num > 0 && num < 1000) {
        // Reasonable quantity range
        return num
      }
    }
    return 1
  }

  private extractUnit (rowValues: string[]): string {
    const units = ['个', '台', '套', 'm', 'kg', '只', '根', '块', '张', '副', 'm²', 'm³']
    for (const value of rowValues) {
      for (const unit of units) {
        if (value.includes(unit)) {
          return unit
        }
      }
    }
    return '个'
  }

  private determineMaterialCategory (materialName: string): string {
    if (materialName.includes('用工') || materialName.includes('人工')) return '人工'
    if (materialName.includes('机械') || materialName.includes('机器')) return '机械'
    return '材料'
  }

  private async generateSubitemInfoFile (subitems: SubitemInfo[]): Promise<void> {
    const outputWorkbook = new ExcelJS.Workbook()
    const worksheet = outputWorkbook.addWorksheet('子目信息')

    // Add headers matching reference format
    worksheet.addRow([
      '',
      '定额号',
      '子目名称',
      '基价',
      '人工',
      '材料',
      '机械',
      '管理费',
      '利润',
      '其他',
      '图片名称'
    ])

    // Add data rows
    for (const item of subitems) {
      worksheet.addRow([
        '',
        item.定额号,
        item.子目名称,
        item.基价,
        item.人工,
        item.材料,
        item.机械,
        item.管理费,
        item.利润,
        item.其他,
        item.图片名称
      ])
    }

    const outputDir = 'output'
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true })
    }

    await outputWorkbook.xlsx.writeFile(path.join(outputDir, '子目信息.xls'))
  }

  private async generateWorkContentFile (workContents: WorkContent[]): Promise<void> {
    const outputWorkbook = new ExcelJS.Workbook()
    const worksheet = outputWorkbook.addWorksheet('工作内容、附注信息')

    // Add headers
    worksheet.addRow(['编号', '工作内容'])

    // Add data rows
    for (const item of workContents) {
      worksheet.addRow([item.编号, item.工作内容])
    }

    await outputWorkbook.xlsx.writeFile(path.join('output', '工作内容、附注信息表.xls'))
  }

  private async generateMaterialContentFile (materials: MaterialContent[]): Promise<void> {
    const outputWorkbook = new ExcelJS.Workbook()
    const worksheet = outputWorkbook.addWorksheet('含量表')

    // Add headers matching reference format
    worksheet.addRow([
      '编号',
      '名称',
      '规格',
      '单位',
      '单价',
      '含量',
      '主材标记',
      '材料号',
      '材料类别',
      '是否有明细'
    ])

    // Add data rows
    for (const item of materials) {
      worksheet.addRow([
        item.编号,
        item.名称,
        item.规格,
        item.单位,
        item.单价,
        item.含量,
        item.主材标记,
        item.材料号,
        item.材料类别,
        item.是否有明细
      ])
    }

    await outputWorkbook.xlsx.writeFile(path.join('output', '含量表.xls'))
  }

  async convert (inputFilePath: string): Promise<void> {
    console.log('🔄 Starting conversion with improved extraction logic...\n')

    // Load input file
    await this.loadInputFile(inputFilePath)
    console.log('✅ Input file loaded')

    // Extract data using new logic
    const subitems = this.extractSubitemTables()
    console.log(`✅ Extracted ${subitems.length} subitem records`)

    const workContents = this.extractWorkContent()
    console.log(`✅ Extracted ${workContents.length} work content records`)

    const materials = this.extractMaterialContent()
    console.log(`✅ Extracted ${materials.length} material records`)

    // Generate output files
    await this.generateSubitemInfoFile(subitems)
    console.log('✅ Generated 子目信息.xls')

    await this.generateWorkContentFile(workContents)
    console.log('✅ Generated 工作内容、附注信息表.xls')

    await this.generateMaterialContentFile(materials)
    console.log('✅ Generated 含量表.xls')

    console.log('\n🎉 Conversion completed! Check the output directory.')
  }
}

// Run conversion
async function main () {
  const converter = new ExcelConverterV2()
  const inputFile = 'data/建设工程消耗量标准及计算规则（安装工程） 补充子目.xlsx'

  try {
    await converter.convert(inputFile)
  } catch (error) {
    console.error('❌ Conversion failed:', error)
  }
}

main().catch(console.error)

package demo.write.sheet

import no.vedaadata.excel.*

import demo.write.*

import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook

@main def main(num: Int) =

  val items = Person.generator.generate(num)
  
  given workbook: Workbook = new XSSFWorkbook

  createSheet(items)(using SheetWriter.fromLayout(Person.layout))
  
  Excel.writeWorkbookToFile(workbook, "temp/demo-write-sheet.xlsx")
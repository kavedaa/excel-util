package demo.write.pivot

import no.vedaadata.excel.*

import demo.write.*

import org.apache.poi.ss.usermodel.DataConsolidateFunction
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.ss.util.CellReference

@main def main(num: Int) =

  val items = Person.generator.generate(num)
  
  given workbook: XSSFWorkbook = new XSSFWorkbook

  val area = createSheet(items, Some("Data"))(using SheetWriter.fromLayout(Person.layout))

  val pivotSheet = workbook.createSheet("Pivot")
  val pivotTable = pivotSheet.createPivotTable(area.toAreaReference, CellReference(0, 0))
  pivotTable.addRowLabel(0)
  pivotTable.addRowLabel(1)
  pivotTable.addColumnLabel(DataConsolidateFunction.COUNT, 2, "Count of age")
  pivotTable.addColumnLabel(DataConsolidateFunction.COUNT, 3, "Count of fortune")
  pivotTable.addColumnLabel(DataConsolidateFunction.AVERAGE, 2, "Average age", "0.0")
  pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 3, "Total fortune")
  
  Excel.writeWorkbookToFile(workbook, "temp/demo-write-pivot.xlsx")
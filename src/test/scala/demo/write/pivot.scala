package demo.write.pivot

import java.time.LocalDate

import org.apache.poi.ss.usermodel.DataConsolidateFunction
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.ss.util.CellReference

import no.vedaadata.generator.*

import no.vedaadata.excel.*

case class Person(
  name: String,
  age: Int,
  fortune: BigDecimal,
  birthDate: LocalDate)

object Person:

  val generator = 
    (Generator("Alex", "Bob", "Charlie", "David", "Eve", "Frank", "Grace", "Helen", "Ivan", "Jane"), 
    Generator.between(20, 50), 
    Generator.between(1000, 100000).map(BigDecimal.apply),
    Generator.between(LocalDate.of(2000, 1, 1), LocalDate.of(2005, 1, 1)))
    .mapN(Person.apply)

object PersonLayout extends Layout[Person]:

  val Name = Column(25)("Name", _.name)
  val Age = Column("Age", _.age)
  val Fortune = Column("Fortune", _.fortune)
  val BirthDate = Column(15)("Date of birth", _.birthDate)

  def columns = List(Name, Age, Fortune, BirthDate)


@main def main(num: Int) =

  val items = Person.generator.generate(num)
  
  given workbook: XSSFWorkbook = new XSSFWorkbook

  val area = createSheet(items, Some("Data"))(using SheetWriter.fromLayout(PersonLayout))

  val pivotSheet = workbook.createSheet("Pivot")
  val pivotTable = pivotSheet.createPivotTable(area.toAreaReference, CellReference(0, 0))

  pivotTable.addRowLabel(PersonLayout.indexOf(_.Name))
  pivotTable.addRowLabel(PersonLayout.indexOf(_.Age))
  pivotTable.addColumnLabel(DataConsolidateFunction.COUNT, PersonLayout.indexOf(_.Age), "Count of age")
  pivotTable.addColumnLabel(DataConsolidateFunction.COUNT, PersonLayout.indexOf(_.Fortune), "Count of fortune")
  pivotTable.addColumnLabel(DataConsolidateFunction.AVERAGE, PersonLayout.indexOf(_.Age), "Average age", "0.0")
  pivotTable.addColumnLabel(DataConsolidateFunction.SUM, PersonLayout.indexOf(_.Fortune), "Total fortune")
  
  Excel.writeWorkbookToFile(workbook, "temp/demo-write-pivot.xlsx")
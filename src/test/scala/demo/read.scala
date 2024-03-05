package demo.read

import scala.util.*
import scala.deriving.Mirror
import scala.compiletime.*

import java.time.LocalDate

import no.vedaadata.excel.*

import org.apache.poi.ss.usermodel.*

case class Person(
  firstName: String,
  lastName: String,
  birthDate: Option[LocalDate],
  city: Option[String])

@main def main =

  val exampleFileName = "data/example-for-read.xlsx"

  def print(persons: Try[List[Person]]) =
    persons match
      case Success(xs) => xs.foreach(println)
      case Failure(ex) => println(ex.getMessage)

  //  1

  {
    println("--- 1 ---")

    val rowReader1 = new RowReader[Person]:
      val firstNameReader = summon[CellReader[String]]
      val lastNameReader = summon[CellReader[String]]
      val birthDateReader = summon[CellReader[Option[LocalDate]]]
      val cityReader = summon[CellReader[Option[String]]]
      def read(row: Row) =       
        for
          firstName <- firstNameReader.read(row.getCell(2))
          lastName <- lastNameReader.read(row.getCell(3))
          birthDate <- birthDateReader.read(row.getCell(4))
          city <- cityReader.read(row.getCell(5))
        yield Person(firstName, lastName, birthDate, city)

    val sheetReader1 = SheetReader.fromRowReader(rowReader1)

    val persons1 = Excel.readFile(exampleFileName)(using sheetReader1)

    print(persons1)
  }

  //  2

  {
    println("--- 2 ---")

    val rowReader2 = new RowReader[Person]:
      val firstNameColumn = IndexedColumn[String](2)
      val lastNameColumn = IndexedColumn[String](3)
      val birthDateColumn = IndexedColumn[Option[LocalDate]](4)
      val cityColumn = IndexedColumn[Option[String]](5)
      def read(row: Row) =       
        for
          firstName <- firstNameColumn.read(row)
          lastName <- lastNameColumn.read(row)
          birthDate <- birthDateColumn.read(row)
          city <- cityColumn.read(row)
        yield Person(firstName, lastName, birthDate, city)

    val sheetReader2 = SheetReader.fromRowReader(rowReader2)

    val persons2 = Excel.readFile(exampleFileName)(using sheetReader2)

    print(persons2)
  }

  //  3

  {
    println("--- 3 ---")

    val rowReaderFactory3: (Detection.IndexedLabels ?=> RowReader[Person]) = new RowReader:
      val firstNameColumn = LabeledColumn[String]("First name")
      val lastNameColumn = LabeledColumn[String]("Last name")
      val birthDateColumn = LabeledColumn[Option[LocalDate]]("Birth date")
      val cityColumn = LabeledColumn[Option[String]]("City")
      def read(row: Row) =       
        for
          firstName <- firstNameColumn.read(row)
          lastName <- lastNameColumn.read(row)
          birthDate <- birthDateColumn.read(row)
          city <- cityColumn.read(row)
        yield Person(firstName, lastName, birthDate, city)

    val sheetReader3 = SheetReader.fromRowReaderFactory(rowReaderFactory3)

    val persons3 = Excel.readFile(exampleFileName)(using sheetReader3)

    print(persons3)
  }

  //  4

  {
    println("--- 4 ---")

    val labels4 = ("First name", "Last name", "Birth date", "City")

    val rowReaderFactory4: Detection.RowReaderFactory[Person] = Detection.RowReaderFactory.derived(labels4)

    val sheetReader4 = SheetReader.fromRowReaderFactory(rowReaderFactory4)

    val persons4 = Excel.readFile(exampleFileName)(using sheetReader4)

    print(persons4)
  }

  //  5

  {
    println("--- 5 ---")

    val exampleHeaderPolicyFileName = "data/example-for-read-header-policy.xlsx"

    val labels5 = ("First name", "Last name", "Birth date", "City")

    val rowReaderFactory5: Detection.RowReaderFactory[Person] = Detection.RowReaderFactory.derived(labels5)

    given HeaderPolicy.Header = HeaderPolicy.Header(1)

    val sheetReader5 = SheetReader.fromRowReaderFactory(rowReaderFactory5)

    val persons5 = Excel.readFile(exampleHeaderPolicyFileName)(using sheetReader5)

    print(persons5)
  }
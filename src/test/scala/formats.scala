package no.vedaadata.excel

import scala.util.*

import org.scalatest.funsuite._
import org.scalatest.matchers.should._

case class Data(
  firstName: String,
  lastName: Option[String],
  age: Option[Int])

class FormatsTest extends AnyFunSuite with Matchers:

  val labels = ("First name", "Last name", "Age")

  val rowReaderFactory: Detection.RowReaderFactory[Data] = Detection.RowReaderFactory.derived(labels)

  val sheetReader = SheetReader.fromRowReaderFactory(rowReaderFactory)

  val expected = List(
    Data("Alice", Some("Doe"), Some(24)),
    Data("Bob", Some("Smith"), None))


  test("XLSX"):

    val persons = Excel.readFile("data/test-for-read.xlsx")(using sheetReader)

    persons shouldEqual Success(expected)


  test("XLS"):

    val persons = Excel.readFile("data/test-for-read.xls")(using sheetReader)

    persons shouldEqual Success(expected)


  test("XLSM"):

    val persons = Excel.readFile("data/test-for-read.xlsm")(using sheetReader)

    persons shouldEqual Success(expected)




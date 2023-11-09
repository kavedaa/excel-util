package no.vedaadata.excel

import scala.util.*
import scala.deriving.*
import scala.compiletime.*

import java.time.*

import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.*

import no.vedaadata.text.LabelTransformer
import java.time.format.DateTimeFormatter

trait CellReader[A]:
  def read(cell: Cell): Try[A]

extension [A](cellReader: CellReader[A])
  def orElse(that: CellReader[A]): CellReader[A] = new CellReader:
    def read(cell: Cell) = cellReader.read(cell).recoverWith { case _ => that.read(cell) }

object CellReader:

  def apply[A](using cellReader: CellReader[A]) = cellReader

  def read[A](cell: Cell)(using cellReader: CellReader[A]): Try[A] =
    cellReader.read(cell)

  given boolean: CellReader[Boolean] with
    def read(cell: Cell) =
      Try(cell.getBooleanCellValue)

  given string: CellReader[String] with
    def read(cell: Cell) =
      Try(cell.getStringCellValue).orElse:
        Try(cell.getNumericCellValue.toString)

  given char: CellReader[Char] with
    def read(cell: Cell) =
      Try(cell.getStringCellValue.charAt(0))

  given byte: CellReader[Byte] with
    def read(cell: Cell) =
      Try(cell.getNumericCellValue.toByte)

  given short: CellReader[Short] with
    def read(cell: Cell) =
      Try(cell.getNumericCellValue.toShort)

  given int: CellReader[Int] with
    def read(cell: Cell) =
      Try(cell.getNumericCellValue.toInt)

  given long: CellReader[Long] with
    def read(cell: Cell) =
      Try(cell.getNumericCellValue.toLong)

  given float: CellReader[Float] with
    def read(cell: Cell) =
      Try(cell.getNumericCellValue.toFloat)

  given double: CellReader[Double] with
    def read(cell: Cell) =
      Try(cell.getNumericCellValue)

  given bigint: CellReader[BigInt] with
    def read(cell: Cell) =
      Try(BigInt(cell.getNumericCellValue.toLong))

  given bigdecimal: CellReader[BigDecimal] with
    def read(cell: Cell) =
      Try(BigDecimal(cell.getNumericCellValue))

  given localDate: CellReader[LocalDate] with
    def read(cell: Cell) =
      Try(cell.getLocalDateTimeCellValue.toLocalDate)

  class StringLocalDateReader(formats: List[DateTimeFormatter]) extends CellReader[LocalDate]:
    def read(cell: Cell) =
      Try(cell.getStringCellValue).flatMap: string =>
        formats
          .iterator
          .map(format => Try(LocalDate.parse(string, format)))
          .collectFirst { case success @ Success(_) => success }
          .getOrElse(Failure(new Exception(s"Could not parse $string as date")))

  given localTime: CellReader[LocalTime] with
    def read(cell: Cell) =
      Try {
        val doubleValue = cell.getNumericCellValue
        val nanoOfDay = doubleValue * (24L * 60 * 60 * 1000 * 1000 * 1000)
        LocalTime.ofNanoOfDay(nanoOfDay.toLong)
      }

  given localDateTime: CellReader[LocalDateTime] with
    def read(cell: Cell) =
      Try(cell.getLocalDateTimeCellValue)

  given option[A](using inner: CellReader[A]): CellReader[Option[A]] with
    def read(cell: Cell) =
      if (cell == null) || (cell.getCellType == CellType.BLANK) then Success(None)
      else inner.read(cell).map(Some.apply)

end CellReader

trait RowReader[A]:
  def read(row: Row): Try[A]

trait SheetReader[A]:
  def read(sheet: Sheet)(using HeaderPolicy[A]): Try[List[A]]

object SheetReader:

  def read[A](sheet: Sheet)(using sheetReader: SheetReader[A])(using HeaderPolicy[A]) =
    sheetReader.read(sheet)

  def fromRowReader[A](rowReader: RowReader[A]) =
    new SheetReader[A]:
      def read(sheet: Sheet)(using headerPolicy: HeaderPolicy[A]) =
        val rowNums = (sheet.getFirstRowNum + headerPolicy.offset) to sheet.getLastRowNum
        val validRowNums = rowNums filter { rowNum =>
          Option(sheet.getRow(rowNum)).isDefined
        } 
        val items = validRowNums map { rowNum =>
          val row = sheet.getRow(rowNum)
          rowReader.read(row)
        }
        Try(items.map(_.get).toList)

  private def defaultIndexedReaders(cellReaders: List[CellReader[Any]]): Try[List[(CellReader[Any], (String, Int))]] = 
    Success(cellReaders.zipWithIndex map { (cellReader, index) => (cellReader, ("", index)) } )

  inline given derived[P <: Product](using m: Mirror.ProductOf[P]): SheetReader[P] =
    new SheetReader[P]:
      type CellReaders = Tuple.Map[m.MirroredElemTypes, CellReader]
      val cellReaders = summonAll[CellReaders].toList.asInstanceOf[List[CellReader[Any]]]
      def read(sheet: Sheet)(using headerPolicy: HeaderPolicy[P]) =
        val headerRowNum = headerPolicy.offset
        val headerRow = sheet.getRow(headerRowNum)
        val indexedReaders = defaultIndexedReaders(cellReaders)
        indexedReaders.flatMap { xs =>
          val rowNums = (sheet.getFirstRowNum + headerPolicy.offset) to sheet.getLastRowNum
          val validRowNums = rowNums filter { rowNum =>
            Option(sheet.getRow(rowNum)).isDefined
          } 
          val items = validRowNums map { rowNum =>
            val row = sheet.getRow(rowNum)
            val values = xs map { case (cellReader, (label, index)) =>
              cellReader.read(row.getCell(index)).recover { ex => s"Error at column '$label': ${ex.getMessage}"}
            }
            Try(values.map(_.get)) map { list =>
              val tuple = list.foldRight[Tuple](EmptyTuple)(_ *: _)
              m.fromProduct(tuple)            
            }
          }
          Try(items.map(_.get).toList)
        }

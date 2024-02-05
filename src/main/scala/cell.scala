package no.vedaadata.excel

import scala.util.*

import java.time.*
import java.time.format.DateTimeFormatter

import org.apache.poi.ss.usermodel.*

trait CellReader[A]:
  def read(cell: Cell): Try[A]

extension [A](cellReader: CellReader[A])
  def orElse(that: CellReader[A]): CellReader[A] = new CellReader:
    def read(cell: Cell) = cellReader.read(cell).recoverWith { case _ => that.read(cell) }

object CellReader:

  def apply[A](using cellReader: CellReader[A]) = cellReader

  def read[A](cell: Cell)(using cellReader: CellReader[A]): Try[A] =
    cellReader.read(cell)

  private def checkNull(cell: Cell): Try[Cell] =
    if cell == null then Failure(Exception("Cell is null"))
    else Success(cell)

  given boolean: CellReader[Boolean] with
    def read(cell: Cell) =
      checkNull(cell).map(_.getBooleanCellValue)

  given string: CellReader[String] with
    def read(cell: Cell) =
      checkNull(cell).map(_.getStringCellValue)

  val stringFromNumeric: CellReader[String] = new CellReader: 
    def read(cell: Cell) =
      checkNull(cell).map(_.getNumericCellValue.toInt.toString)

  given char: CellReader[Char] with
    def read(cell: Cell) =
      checkNull(cell).map(_.getStringCellValue.charAt(0))

  given byte: CellReader[Byte] with
    def read(cell: Cell) =
      checkNull(cell).map(_.getNumericCellValue.toByte)

  given short: CellReader[Short] with
    def read(cell: Cell) =
      checkNull(cell).map(_.getNumericCellValue.toShort)

  given int: CellReader[Int] with
    def read(cell: Cell) =
      checkNull(cell).map(_.getNumericCellValue.toInt)

  given long: CellReader[Long] with
    def read(cell: Cell) =
      checkNull(cell).map(_.getNumericCellValue.toLong)

  given float: CellReader[Float] with
    def read(cell: Cell) =
      checkNull(cell).map(_.getNumericCellValue.toFloat)

  given double: CellReader[Double] with
    def read(cell: Cell) =
      checkNull(cell).map(_.getNumericCellValue)

  given bigint: CellReader[BigInt] with
    def read(cell: Cell) =
      checkNull(cell).map(x => BigInt(x.getNumericCellValue.toLong))

  given bigdecimal: CellReader[BigDecimal] with
    def read(cell: Cell) =
      checkNull(cell).map(x => BigDecimal(x.getNumericCellValue))

  given localDate: CellReader[LocalDate] with
    def read(cell: Cell) =
      checkNull(cell).map(_.getLocalDateTimeCellValue.toLocalDate)

  class StringLocalDateReader(formats: List[DateTimeFormatter]) extends CellReader[LocalDate]:
    def read(cell: Cell) =
      checkNull(cell).map(_.getStringCellValue).flatMap: string =>
        formats
          .iterator
          .map(format => Try(LocalDate.parse(string, format)))
          .collectFirst { case success @ Success(_) => success }
          .getOrElse(Failure(new Exception(s"Could not parse $string as date")))

  given localTime: CellReader[LocalTime] with
    def read(cell: Cell) =
      checkNull(cell).map: x =>
        val doubleValue = x.getNumericCellValue
        val nanoOfDay = doubleValue * (24L * 60 * 60 * 1000 * 1000 * 1000)
        LocalTime.ofNanoOfDay(nanoOfDay.toLong)

  given localDateTime: CellReader[LocalDateTime] with
    def read(cell: Cell) =
      checkNull(cell).map(_.getLocalDateTimeCellValue)

  given option[A](using inner: CellReader[A]): CellReader[Option[A]] with
    def read(cell: Cell) =
      if (cell == null) || (cell.getCellType == CellType.BLANK) then Success(None)
      else inner.read(cell).map(Some.apply)
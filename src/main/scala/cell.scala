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

  class StringBooleanReader(trueTexts: List[String], falseTexts: List[String]) extends CellReader[Boolean]:
    def read(cell: Cell) =
      checkNull(cell).map(_.getStringCellValue).flatMap: x =>
        if trueTexts.contains(x) then Success(true)
        else if falseTexts.contains(x) then Success(false)
        else Failure(Exception(s"Could not parse $x as boolean"))

  object DefaultStringBooleanReader extends StringBooleanReader(List("true", "yes", "TRUE", "YES"), List("false", "no", "FALSE", "NO"))

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
      checkNull(cell).map(_.getStringCellValue).flatMap: x =>
        formats
          .iterator
          .map(format => Try(LocalDate.parse(x.trim, format)))
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
      if (cell == null) || (cell.getCellType == CellType.BLANK) || ((cell.getCellType == CellType.STRING) && (cell.getStringCellValue.isBlank)) then Success(None)
      else inner.read(cell).map(Some.apply)

end CellReader


trait CellWriter[A]:
  def write(cell: Cell, value: A): Unit

object CellWriter:

  def apply[A](using cellWriter: CellWriter[A]) = cellWriter

  given boolean: CellWriter[Boolean] with
    def write(cell: Cell, value: Boolean) =
      cell.setCellValue(value)

  given string: CellWriter[String] with
    def write(cell: Cell, value: String) =
      cell.setCellValue(value)

  given char: CellWriter[Char] with
    def write(cell: Cell, value: Char) =
      cell.setCellValue(value.toString)

  given numeric[N](using n: Numeric[N]): CellWriter[N] with
    def write(cell: Cell, value: N) =
      cell.setCellValue(n.toDouble(value))

  given localDate: CellWriter[LocalDate] with
    def write(cell: Cell, value: LocalDate) =
      cell.setCellValue(value)

  given localTime: CellWriter[LocalTime] with
    def write(cell: Cell, value: LocalTime) =
      //  for some reason POI does not support time values directly 
      val doubleValue = value.toNanoOfDay.toDouble / (24L * 60 * 60 * 1000 * 1000 * 1000)
      cell.setCellValue(doubleValue)

  given localDateTime: CellWriter[LocalDateTime] with
    def write(cell: Cell, value: LocalDateTime) =
      cell.setCellValue(value)

  given option[A](using inner: CellWriter[A]): CellWriter[Option[A]] with
    def write(cell: Cell, value: Option[A]) =
      value match
        case Some(x) => inner.write(cell, x)
        case None => cell.setBlank()

  class StringBooleanWriter(trueText: String, falseText: String) extends CellWriter[Boolean]:
    def write(cell: Cell, value: Boolean) =
      cell.setCellValue(if value then trueText else falseText)



package no.vedaadata.excel

import java.time.*

import org.apache.poi.ss.usermodel.*

trait CellStyleProvider[-A]:
  def provide(baseStyle: CellStyle)(using wb: Workbook): CellStyle

object CellStyleProvider:

  given default[A]: CellStyleProvider[A] with
    def provide(baseStyle: CellStyle)(using wb: Workbook) = baseStyle

  given option[A](using inner: CellStyleProvider[A]): CellStyleProvider[Option[A]] with
    def provide(baseStyle: CellStyle)(using wb: Workbook) = inner.provide(baseStyle)

  given localDate: CellStyleProvider[LocalDate] with
    def provide(baseStyle: CellStyle)(using wb: Workbook) =
      val cellStyle = wb.createCellStyle
      cellStyle.cloneStyleFrom(baseStyle)
      cellStyle.setDataFormat(0x0e)        //  magic Excel number, "m/d/yy"
      cellStyle

  given localTime: CellStyleProvider[LocalTime] with
    def provide(baseStyle: CellStyle)(using wb: Workbook) =
      val cellStyle = wb.createCellStyle
      cellStyle.cloneStyleFrom(baseStyle)
      cellStyle.setDataFormat(0x15)        //  magic Excel number, "h:mm:ss"
      cellStyle

  given localDateTime: CellStyleProvider[LocalDateTime] with
    def provide(baseStyle: CellStyle)(using wb: Workbook) =
      val cellStyle = wb.createCellStyle
      cellStyle.cloneStyleFrom(baseStyle)
      cellStyle.setDataFormat(0x16)        //  magic Excel number, "m/d/yy h:mm"
      cellStyle

trait CellStyleFactory extends (Workbook => CellStyle)

trait CellStyleFactoryCompanion:

  trait BoldCellStyleFactory extends CellStyleFactory:
    def apply(wb: Workbook) =
      val cellStyle = wb.createCellStyle()
      val font = wb.createFont()
      font.setBold(true)
      cellStyle.setFont(font)
      cellStyle

  trait ItalicCellStyleFactory extends CellStyleFactory:
    def apply(wb: Workbook) =
      val cellStyle = wb.createCellStyle()
      val font = wb.createFont()
      font.setItalic(true)
      cellStyle.setFont(font)
      cellStyle

trait HeaderCellStyleFactory extends CellStyleFactory

object HeaderCellStyleFactory extends CellStyleFactoryCompanion:

  given default: HeaderCellStyleFactory = _.createCellStyle()

  object Bold extends HeaderCellStyleFactory with BoldCellStyleFactory
  object Italic extends HeaderCellStyleFactory with ItalicCellStyleFactory


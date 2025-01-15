package no.vedaadata.excel

import scala.util.*
import scala.deriving.*
import scala.compiletime.*

import java.time.*

import org.apache.poi.ss.usermodel.*

import no.vedaadata.text.LabelTransformer



def createCell[A](index: Int, value: A)(using row: Row)(using cellWriter: CellWriter[A])(using cellStyle: CellStyle) =
  val cell = row.createCell(index)
  cell.setCellStyle(cellStyle)
  cellWriter.write(cell, value)

trait SheetWriter[A]:

  /**
    *   Writes a collection of values to a sheet.
    *   Returns the area of the sheet that was written to, including any header.
    */
  def write(xs: Iterable[A])(using sheet: Sheet)(using HeaderPolicy)(using headerCellStyleFactory: HeaderCellStyleFactory)(using wb: Workbook): Area

object SheetWriter:

  class Layout[A](val columns: Layout.Column[A, ?]*)

  object Layout:
    case class Column[A, B](label: String, f: A => B, width: Option[Int])(using val cellWriter: CellWriter[B], val cellStyleProvider: CellStyleProvider[B])
    object Column:
      def apply[A, B](label: String, f: A => B)(using CellWriter[B], CellStyleProvider[B]): Column[A, B] = Column(label, f, None)
      def apply[A, B](width: Int)(label: String, f: A => B)(using CellWriter[B], CellStyleProvider[B]): Column[A, B] = Column(label, f, Some(width))

  def fromLayout[A](layout: Layout[A]) = new SheetWriter[A]:
    def write(xs: Iterable[A])(using sheet: Sheet)(using headerPolicy: HeaderPolicy)(using headerCellStyleFactory: HeaderCellStyleFactory)(using wb: Workbook): Area =
      //  column widths
      layout.columns.map(_.width).zipWithIndex.foreach: (columnWidth, index) =>
        columnWidth.foreach(width => sheet.setColumnWidth(index, width * 256))
      //  headers
      headerPolicy match
        case header: HeaderPolicy.Header =>
          val row = sheet.createRow(0)
          val cellStyle = headerCellStyleFactory(wb)
          layout.columns.toList.zipWithIndex.foreach: (column, columnIndex) =>
            val cell = row.createCell(columnIndex)
            cell.setCellValue(column.label)
            cell.setCellStyle(cellStyle)
        case _ =>
      //  rows          
      val baseCellStyle = wb.createCellStyle
      val cellStyles = layout.columns.map(_.cellStyleProvider.provide(baseCellStyle))
      xs.zipWithIndex.foreach: (x, rowIndex) =>
        val row = sheet.createRow(headerPolicy.dataRowIndex + rowIndex)
        layout.columns.toList.zip(cellStyles).zipWithIndex.foreach: 
          case ((column, cellStyle), columnIndex) =>          
            val cell = row.createCell(columnIndex)
            createCell(columnIndex, column.f(x))(using row)(using column.cellWriter)(using cellStyle)
      Area(sheet, 0, headerPolicy.firstRowIndex, layout.columns.size - 1, headerPolicy.dataRowIndex + xs.size - 1)

  inline def derived[P <: Product](using m: Mirror.ProductOf[P])(using labelTransformer: LabelTransformer): SheetWriter[P] =
    new SheetWriter:
      type CellWriters = Tuple.Map[m.MirroredElemTypes, CellWriter]
      type CellStyleProviders = Tuple.Map[m.MirroredElemTypes, CellStyleProvider]
      val cellWriters = summonAll[CellWriters].toList.asInstanceOf[List[CellWriter[Any]]]
      val cellStyleProviders = summonAll[CellStyleProviders].toList.asInstanceOf[List[CellStyleProvider[Any]]]
      val labels = constValueTuple[m.MirroredElemLabels].toList.asInstanceOf[List[String]].map(labelTransformer)
      def write(xs: Iterable[P])(using sheet: Sheet)(using headerPolicy: HeaderPolicy)(using headerCellStyleFactory: HeaderCellStyleFactory)(using wb: Workbook) =
        // columnWidths.xs.zipWithIndex foreach { (columnWidth, index) =>
        //   columnWidth foreach { width =>
        //     sheet.setColumnWidth(index, width * 256)  
        //   }  
        // }
        headerPolicy match
          case header: HeaderPolicy.Header =>
            val row = sheet.createRow(0)
            val cellStyle = headerCellStyleFactory(wb)
            labels.zipWithIndex foreach: (label, index) =>
              val cell = row.createCell(index)
              cell.setCellValue(label)
              cell.setCellStyle(cellStyle)
          case _ =>
        val baseCellStyle = wb.createCellStyle
        val cellStyles = cellStyleProviders.map(_.provide(baseCellStyle))
        xs.zipWithIndex.foreach: (x, index) =>
          val elements = x.productIterator.toList
          val row = sheet.createRow(headerPolicy.dataRowIndex + index)
          elements.zip(cellWriters).zip(cellStyles).zipWithIndex.foreach:
            case (((element, cellWriter), cellStyle), index) =>
              createCell(index, element)(using row)(using cellWriter)(using cellStyle)
        Area(sheet, 0, headerPolicy.firstRowIndex, cellWriters.size - 1, headerPolicy.dataRowIndex + xs.size - 1)
      
  inline given [P <: Product](using m: Mirror.ProductOf[P]): SheetWriter[P] = derived
    
def createSheet[A](value: Iterable[A], name: Option[String] = None)(using sheetWriter: SheetWriter[A])(using headerPolicy: HeaderPolicy)(using wb: Workbook) =
  val sheet = name match
    case Some(name) => wb.createSheet(name)
    case None => wb.createSheet
  sheetWriter.write(value)(using sheet)


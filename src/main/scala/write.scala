package no.vedaadata.excel

import scala.util.*
import scala.deriving.*
import scala.compiletime.*

import java.time.*

import org.apache.poi.ss.usermodel.*

import no.vedaadata.text.LabelTransformer

// TODO maybe not have these as top level

def createCell[A](index: Int, value: A)(using row: Row)(using cellWriter: CellWriter[A])(using cellStyle: CellStyle) =
  val cell = row.createCell(index)
  cell.setCellStyle(cellStyle)
  cellWriter.write(cell, value)

def createSheet[A](value: Iterable[A], name: Option[String] = None)(using sheetWriter: SheetWriter[A])(using HeaderPolicy)(using wb: Workbook) =
  val sheet = name match
    case Some(name) => wb.createSheet(name)
    case None => wb.createSheet
  sheetWriter.write(value)(using sheet)


trait SheetWriter[A]:

  /**
    *   Writes a collection of values to a `Sheet`.
    *   Returns the area of the sheet that was written to, including any header.
    */
  def write(xs: Iterable[A])(using Sheet)(using HeaderPolicy)(using HeaderCellStyleFactory)(using Workbook): Area

object SheetWriter:

  /**
    *   Creates a `SheetWriter` from a `Layout`.
    */
  def fromLayout[A](layout: Layout[A]): SheetWriter[A] = layout.toSheetWriter

  inline def derived[P <: Product](using m: Mirror.ProductOf[P])(using labelTransformer: LabelTransformer): SheetWriter[P] =
    new SheetWriter:
      type CellWriters = Tuple.Map[m.MirroredElemTypes, CellWriter]
      type CellStyleProviders = Tuple.Map[m.MirroredElemTypes, CellStyleProvider]
      val cellWriters = summonAll[CellWriters].toList.asInstanceOf[List[CellWriter[Any]]]
      val cellStyleProviders = summonAll[CellStyleProviders].toList.asInstanceOf[List[CellStyleProvider[Any]]]
      val labels = constValueTuple[m.MirroredElemLabels].toList.asInstanceOf[List[String]].map(labelTransformer)
      def write(xs: Iterable[P])(using sheet: Sheet)(using headerPolicy: HeaderPolicy)(using headerCellStyleFactory: HeaderCellStyleFactory)(using wb: Workbook) =
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
  
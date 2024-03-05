package no.vedaadata.excel

import scala.util.*
import scala.deriving.*
import scala.compiletime.*

import org.apache.poi.ss.usermodel.*

import no.vedaadata.text.LabelTransformer

class IndexedColumn[A](index: Int)(using cellReader: CellReader[A]):
  def read(row: Row) = Try(row.getCell(index)).flatMap(cellReader.read).recoverWith:
    case ex => Failure(Exception(s"Row ${row.getRowNum + 1} column ${index + 1}: " + ex.getMessage, ex))

class LabeledColumn[A](label: String)(using cellReader: CellReader[A])(using lim: Detection.IndexedLabels):
  val indexedColumn = lim.resolve(label).map(index => IndexedColumn(index)(using cellReader))
  def read(row: Row): Try[A] = indexedColumn.flatMap(_.read(row)).recoverWith:
    case ex => Failure(Exception(s"Can not read data for '$label': " + ex.getMessage, ex))

trait RowReader[A]:
  def read(row: Row): Try[A]

trait SheetReader[A]:
  def read(sheet: Sheet)(using HeaderPolicy): Try[List[A]]

object SheetReader:

  def read[A](sheet: Sheet)(using sheetReader: SheetReader[A])(using HeaderPolicy) =
    sheetReader.read(sheet)

  def fromRowReader[A](rowReader: RowReader[A]) =
    new SheetReader[A]:
      def read(sheet: Sheet)(using headerPolicy: HeaderPolicy) =
        val rowNums = (sheet.getFirstRowNum + headerPolicy.dataRowIndex) to sheet.getLastRowNum
        val validRowNums = rowNums.filter: rowNum =>
          Option(sheet.getRow(rowNum)).isDefined
        val items = validRowNums.map: rowNum =>
          val row = sheet.getRow(rowNum)
          rowReader.read(row)
        Try(items.map(_.get).toList)

  def fromRowReaderFactory[A](rowReaderFactory: Detection.IndexedLabels ?=> RowReader[A])(using header: HeaderPolicy.Header) =
    new SheetReader[A]:
      def read(sheet: Sheet)(using headerPolicy: HeaderPolicy) =
        Detection.detectLabels(sheet)(header.headerRowIndex).flatMap: detected =>
          val rowReader = rowReaderFactory(using detected)
          fromRowReader(rowReader).read(sheet)

  private def defaultIndexedReaders(cellReaders: List[CellReader[Any]]): Try[List[(CellReader[Any], (String, Int))]] = 
    Success(cellReaders.zipWithIndex map { (cellReader, index) => (cellReader, ("", index)) } )

  inline given derived[P <: Product](using m: Mirror.ProductOf[P]): SheetReader[P] =
    new SheetReader[P]:
      type CellReaders = Tuple.Map[m.MirroredElemTypes, CellReader]
      val cellReaders = summonAll[CellReaders].toList.asInstanceOf[List[CellReader[Any]]]
      def read(sheet: Sheet)(using headerPolicy: HeaderPolicy) =
        val indexedReaders = defaultIndexedReaders(cellReaders)
        indexedReaders.flatMap { xs =>
          val rowNums = (sheet.getFirstRowNum + headerPolicy.dataRowIndex) to sheet.getLastRowNum
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

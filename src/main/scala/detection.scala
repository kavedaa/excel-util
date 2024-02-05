package no.vedaadata.excel

import scala.deriving.Mirror
import scala.compiletime.*

import org.apache.poi.ss.usermodel.*

import scala.util.*

object Detection:

  opaque type IndexedLabels = List[(Int, String)]

  extension (lim: IndexedLabels) 
    def resolve(label: String): Try[Int] = lim.filter((i, l) => l == label) match
      case x :: Nil =>
        val (index, _) = x
        Success(index)
      case xs if xs.nonEmpty => 
        val indexes = xs.map((index, _) => index)
        Failure(Exception(s"Multiple columns with label '$label' found, at indexes ${indexes.mkString(", ")}"))
      case _ =>
        Failure(Exception(s"Column $label not found"))

  def detectLabels(sheet: Sheet)(headerRowNum: Int): IndexedLabels =
    Excel.columnLabelsInSheet(sheet)(headerRowNum)

  type RowReaderFactory[A] = IndexedLabels ?=> RowReader[A]

  //  TODO might do an index-based derivation as well, and share some code
  object RowReaderFactory:
    inline def derived[Labels <: Tuple, P <: Product]
      (labels: Labels)
      (using m: Mirror.ProductOf[P])
      (using Tuple.Union[Labels] <:< String, Tuple.Size[Labels] =:= Tuple.Size[m.MirroredElemTypes]):
      RowReaderFactory[P] =
        new RowReader[P]:
          type CellReaders = Tuple.Map[m.MirroredElemTypes, CellReader]
          val cellReaders = summonAll[CellReaders].toList.asInstanceOf[List[CellReader[Any]]]
          val labelsList = labels.toList.asInstanceOf[List[String]]
          val labeledColumns = (labelsList.zip(cellReaders)).map: (label, cellReader) =>
            LabeledColumn(label)(using cellReader)
          def read(row: Row) =
            val values = labeledColumns.map(_.read(row))
            Try(values.map(_.get)).map: xs =>
              val tuple = xs.foldRight[Tuple](EmptyTuple)(_ *: _)
              m.fromProduct(tuple)

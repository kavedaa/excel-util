package no.vedaadata.excel

import scala.deriving.*
import scala.compiletime.*

import java.time.*

opaque type ColumnWidth[-A] <: Option[Int] = Option[Int]

object ColumnWidth:
  given ColumnWidth[Any] = None
  given ColumnWidth[String] = Some(25)
  given ColumnWidth[Boolean] = Some(25)
  given ColumnWidth[Byte] = Some(10)
  given ColumnWidth[Short] = Some(15)
  given ColumnWidth[Int] = Some(20)
  given ColumnWidth[Long] = Some(25)
  given ColumnWidth[Float] = Some(20)
  given ColumnWidth[Double] = Some(25)
  given ColumnWidth[BigDecimal] = Some(25)
  given ColumnWidth[BigInt] = Some(25)
  given ColumnWidth[LocalDate] = Some(15)
  given ColumnWidth[LocalTime] = Some(15)
  given ColumnWidth[LocalDateTime] = Some(20)

// class ColumnWidths[+A](val xs: Seq[Option[Int]])

// object ColumnWidths:

//   def from[Tup <: Tuple](tup: Tup)(using Tuple.Union[Tup] <:< Int) =
//     ColumnWidths[Nothing](tup.toList.asInstanceOf[List[Int]].map(Option.apply))

//   def fromLabelsAndWidths[Tup <: Tuple](tup: Tup)(using Tuple.Union[Tup] <:< (String, Int)) =
//     ColumnWidths[Nothing](tup.toList.asInstanceOf[List[(String, Int)]].map(_._2).map(Option.apply))

//   inline given derived[P <: Product](using m: Mirror.ProductOf[P]): ColumnWidths[P] =
//     ColumnWidths(summonAll[Tuple.Map[m.MirroredElemTypes, ColumnWidth]].toList.asInstanceOf[List[ColumnWidth[Any]]])


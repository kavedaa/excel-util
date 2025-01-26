package no.vedaadata.excel

import org.apache.poi.ss.usermodel.*

abstract class Layout[A]:

  case class Column[B](label: String, f: A => B, width: Option[Int])(using val cellWriter: CellWriter[B])(using val cellStyleProvider: CellStyleProvider[B])
  
  object Column:
    def apply[B](label: String, f: A => B)(using CellWriter[B], CellStyleProvider[B]): Column[B] = Column(label, f, None)
    def apply[B](width: Int)(label: String, f: A => B)(using CellWriter[B], CellStyleProvider[B]): Column[B] = Column(label, f, Some(width))

  def columns: List[Column[?]]

  def indexOf(column: this.type => Column[?]) = columns.indexWhere(_.label == column(this).label)

  def toSheetWriter = new SheetWriter[A]:
    def write(xs: Iterable[A])(using sheet: Sheet)(using headerPolicy: HeaderPolicy)(using headerCellStyleFactory: HeaderCellStyleFactory)(using wb: Workbook): Area =
      //  column widths
      columns.map(_.width).zipWithIndex.foreach: (columnWidth, index) =>
        columnWidth.foreach(width => sheet.setColumnWidth(index, width * 256))
      //  headers
      headerPolicy match
        case header: HeaderPolicy.Header =>
          val row = sheet.createRow(0)
          val cellStyle = headerCellStyleFactory(wb)
          columns.toList.zipWithIndex.foreach: (column, columnIndex) =>
            val cell = row.createCell(columnIndex)
            cell.setCellValue(column.label)
            cell.setCellStyle(cellStyle)
        case _ =>
      //  rows          
      val baseCellStyle = wb.createCellStyle
      val cellStyles = columns.map(_.cellStyleProvider.provide(baseCellStyle))
      xs.zipWithIndex.foreach: (x, rowIndex) =>
        val row = sheet.createRow(headerPolicy.dataRowIndex + rowIndex)
        columns.toList.zip(cellStyles).zipWithIndex.foreach: 
          case ((column, cellStyle), columnIndex) =>          
            val cell = row.createCell(columnIndex)
            createCell(columnIndex, column.f(x))(using row)(using column.cellWriter)(using cellStyle)
      Area(sheet, 0, headerPolicy.firstRowIndex, columns.size - 1, headerPolicy.dataRowIndex + xs.size - 1)

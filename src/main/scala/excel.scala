package no.vedaadata.excel

import scala.util.*

import java.io.*

import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.*

object Excel:

  def withWorkbook[A](file: File)(f: Workbook => Try[A]): Try[A] = 
    try
      val wb = WorkbookFactory.create(file)
      val res = f(wb)
      wb.close()
      res
    catch 
      case ex => Failure(ex)

  def sheetsInWorkbook(wb: Workbook): List[Sheet] =
    (0 until wb.getNumberOfSheets).map(wb.getSheetAt).toList

  def withSheetsInFile[A](file: File)(f: List[Sheet] => Try[A]): Try[A] =
    withWorkbook(file): wb =>
      f(sheetsInWorkbook(wb))

  def columnLabelsInSheet(sheet: Sheet)(headerRowIndex: Int): Try[List[(Int, String)]] =
    Option(sheet.getRow(headerRowIndex)) match
      case Some(headerRow) =>
        val cellNums = headerRow.getFirstCellNum.toInt to headerRow.getLastCellNum.toInt
        val list = cellNums.toList.flatMap: cellNum =>
          val label = Try(headerRow.getCell(cellNum).getStringCellValue)
          label.toOption.filter(_.nonEmpty).map(cellNum -> _)
        Success(list)
      case None =>
        Failure(Exception(s"No row found at specified header row index $headerRowIndex"))

  def readFile[A](file: File)(using sheetReader: SheetReader[A])(using HeaderPolicy): Try[List[A]] =
    withWorkbook(file): wb =>
      try
        val sheet = wb.getSheetAt(0)
        sheetReader.read(sheet)
      catch 
        case ex => Failure(ex)

  def readFile[A](filename: String)(using sheetReader: SheetReader[A])(using HeaderPolicy): Try[List[A]] = 
    readFile(new File(filename))

  def writeFile[A](file: File, xs: Iterable[A])(using sheetWriter: SheetWriter[A])(using HeaderPolicy): Try[File] =
    given wb: Workbook = new XSSFWorkbook
    createSheet(xs)
    try
      val fos = new FileOutputStream(file) 
      try
        wb.write(fos)
        wb.close()
        Success(file)
      finally
        fos.close()
    catch
      case ex => Failure(ex)            

  def writeFile[A](filename: String, xs: Iterable[A])(using sheetWriter: SheetWriter[A])(using HeaderPolicy): Try[File] =
    writeFile(new File(filename), xs)


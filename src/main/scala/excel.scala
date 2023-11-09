package no.vedaadata.excel

import scala.util.*

import java.io.*

import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.*

object Excel:

  def withWorkbook[A](file: File)(f: Workbook => Try[A]): Try[A] = 
    try
      val fis = new FileInputStream(file) 
      try
        val wb = new XSSFWorkbook(fis)
        val res = f(wb)
        wb.close()
        res
      finally
        fis.close()
    catch 
      case ex => Failure(ex)

  def sheetsInWorkbook(wb: Workbook): List[Sheet] =
    (0 until wb.getNumberOfSheets).map(wb.getSheetAt).toList

  def withSheetsInFile[A](file: File)(f: List[Sheet] => Try[A]): Try[A] =
    withWorkbook(file): wb =>
      f(sheetsInWorkbook(wb))

  def columnLabelsInSheet(sheet: Sheet)(headerRowNum: Int): List[(Int, String)] =
    val headerRow = sheet.getRow(headerRowNum)
    val cellNums = headerRow.getFirstCellNum.toInt to headerRow.getLastCellNum.toInt
    cellNums.toList.flatMap: cellNum =>
      val label = Try(headerRow.getCell(cellNum).getStringCellValue)
      label.toOption.filter(_.nonEmpty).map(cellNum -> _)

  def readFile[A](file: File)(using sheetReader: SheetReader[A])(using HeaderPolicy[A]): Try[List[A]] =
    withWorkbook(file): wb =>
      try
        val sheet = wb.getSheetAt(0)
        sheetReader.read(sheet)
      catch 
        case ex => Failure(ex)

  def readFile[A](filename: String)(using sheetReader: SheetReader[A])(using HeaderPolicy[A]): Try[List[A]] = 
    readFile(new File(filename))

  def writeFile[A](file: File, xs: Iterable[A])(using sheetWriter: SheetWriter[A])(using HeaderPolicy[A]): Try[File] =
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

  def writeFile[A](filename: String, xs: Iterable[A])(using sheetWriter: SheetWriter[A])(using HeaderPolicy[A]): Try[File] =
    writeFile(new File(filename), xs)


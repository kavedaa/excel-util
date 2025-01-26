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

  //  writing

  def writeWorkbookToStream(wb: Workbook, os: OutputStream): Try[Unit] =
    Try(wb.write(os))

  def writeWorkbookToByteArray(wb: Workbook): Try[Array[Byte]] =
    try
      val bos = ByteArrayOutputStream()
      val res = writeWorkbookToStream(wb, bos)
      bos.close()
      res.map(_ => bos.toByteArray)
    catch
      case ex => Failure(ex)
      
  def writeWorkbookToFile(wb: Workbook, file: File): Try[File] =
    try
      val fos = new FileOutputStream(file) 
      val res = writeWorkbookToStream(wb, fos)
      fos.close()
      res.map(_ => file)
    catch
      case ex => Failure(ex)

  def writeWorkbookToFile(wb: Workbook, filename: String): Try[File] =
    writeWorkbookToFile(wb, new File(filename))

  def writeToStream[A](xs: Iterable[A], os: OutputStream)(using sheetWriter: SheetWriter[A])(using HeaderPolicy): Try[Unit] =
    given wb: Workbook = new XSSFWorkbook
    createSheet(xs)
    writeWorkbookToStream(wb, os)

  def writeByteArray[A](xs: Iterable[A])(using sheetWriter: SheetWriter[A])(using HeaderPolicy): Try[Array[Byte]] =
    given wb: Workbook = new XSSFWorkbook
    createSheet(xs)
    writeWorkbookToByteArray(wb)

  def writeFile[A](xs: Iterable[A], file: File)(using sheetWriter: SheetWriter[A])(using HeaderPolicy): Try[File] =
    given wb: Workbook = new XSSFWorkbook
    createSheet(xs)
    writeWorkbookToFile(wb, file)

  def writeFile[A](xs: Iterable[A], filename: String)(using sheetWriter: SheetWriter[A])(using HeaderPolicy): Try[File] =
    writeFile(xs, new File(filename))




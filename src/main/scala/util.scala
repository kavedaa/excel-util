package no.vedaadata.excel

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellReference
import org.apache.poi.ss.util.AreaReference
import org.apache.poi.ss.SpreadsheetVersion

case class Area(
  sheet: Sheet,
  firstColumnIndex: Int,
  firstRowIndex: Int,
  lastColumnIndex: Int,
  lastRowIndex: Int):
  def toAreaReference: AreaReference =
    val upperLeft = CellReference(sheet.getSheetName, firstRowIndex, firstColumnIndex, true, true)
    val lowerRight = CellReference(sheet.getSheetName, lastRowIndex, lastColumnIndex, true, true)
    AreaReference(upperLeft, lowerRight, SpreadsheetVersion.EXCEL2007)
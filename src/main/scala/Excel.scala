package app.excelconverter

import java.io._
import org.apache.poi.hssf.usermodel._
import org.apache.poi.ss.usermodel._

class Excel private (workbook: Workbook, filepath: String) {

  val wb = workbook
  val path = filepath
  var saveDir: String = null

  def setSaveDirectory(dir: String) {
    saveDir = if (dir.endsWith("/")) dir else dir + "/"

  }
  def getFileName(index: Int, sheet: String): String = {
    if (saveDir == null) throw new Exception("directory for saved data is not set")
    val filename = path.split("/").last.replaceAll("""\..+$""", "")
    "%s%s_%d_%s.csv" format (saveDir, filename, index, sheet)
  }

  //  * 区切り文字はカンマで、値はダブルクオートで囲む
  //  * 値内のダブルクオートは、二重にする
  def convertToCsv(): List[(String, String)] = {
    val data = for {
      sheetIndex <- 0 until wb.getNumberOfSheets
      sheet = wb.getSheetAt(sheetIndex)
    } yield {
      val filename = getFileName(sheetIndex, sheet.getSheetName)
      val content = (for {
        rowIndex <- sheet.getFirstRowNum to sheet.getLastRowNum
        row = sheet.getRow(rowIndex)
      } yield (for {
        cellIndex <- row.getFirstCellNum to row.getLastCellNum
        cell = row.getCell(cellIndex)
      } yield addQuote(getCellValue(cell))).mkString(",")).mkString("\n")
      Tuple2(filename, content)
    }
    return data.toList
  }

  // 文字列内のクオーテーションをエスケープし、文字列全体をクオーテーションで囲む
  def addQuote(str: String) = "\"" + str.replaceAll("""\"""", "\"\"") + "\""

  // セルの種別毎に適切な文字列型を返す
  def getCellValue(cell: Cell): String = {
    return if (cell == null) { "" } else {
      cell.getCellType() match {
        case Cell.CELL_TYPE_STRING => cell.getStringCellValue()
        case Cell.CELL_TYPE_NUMERIC => cell.getNumericCellValue().toString()
        case Cell.CELL_TYPE_FORMULA => cell.getCellFormula()
        case Cell.CELL_TYPE_BOOLEAN => cell.getBooleanCellValue().toString()
        case Cell.CELL_TYPE_ERROR => cell.getErrorCellValue().toString()
        case Cell.CELL_TYPE_BLANK => ""
        case _ => throw new Exception("cell type is unknown")
      }
    }
  }
}

// コンパニオンオブジェクト
object Excel {

  def apply(path: String): Excel = {
    var in: FileInputStream = null;
    var result: Workbook = null;
    try {
      in = new FileInputStream(path);
      result = WorkbookFactory.create(in)
    } catch {
      case e => throw e
    } finally {
      try { in.close } catch { case e => ; }
    }
    return new Excel(result, path)
  }

}


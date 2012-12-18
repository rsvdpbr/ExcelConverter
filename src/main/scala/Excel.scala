package app.excelconverter

import java.io._
import org.apache.poi.hssf.usermodel._
import org.apache.poi.ss.usermodel._

class Excel private (workbook: Workbook) {

  val wb: Workbook = workbook
  var saveDir = ""

  def setSaveDirectory(dir: String) {
    saveDir = dir
  }

  def convertToCsv() {
    // ＠未実装
    // 捕捉せずに例外を投げると、変換処理が失敗として扱われる
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
    return new Excel(result)
  }

}


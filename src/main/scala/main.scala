package app.excelconverter

import java.io._
import org.apache.poi.hssf.usermodel._
import org.apache.poi.ss.usermodel._

object Main extends App {

  println("[ExcelConverter ver 1.0]")
  val timer = new Timer

  // コマンドライン引数にファイルが指定されていなければ入力を受付
  timer.start()
  val pathList = if (args.length > 0) { args } else {
    readLine(" Input excel file path to be converted to csv format: ").split(" ")
  }
  timer.stop()

  // 入力からエクセルファイルをリスト化
  println(" List up excel files to be converted to csv")
  val fileList = (for (path <- pathList) yield {
    try {
      new File(path).listFiles.map(_.toString).toList
    } catch {
      case e => List(path)
    }
  }).toList.flatten.filter(_.matches("""^.*\.xlsx?$""")).filter(new File(_).exists)
  if (fileList.length > 0) {
    printStringListWithIndex(fileList)
  } else {
    println("    -> no excel files")
  }

  // 処理の続行確認
  timer.start()
  val continueToProcess = readLine(" Continue to process? [y/n] : ")
  timer.stop()
  // 変換処理
  if (continueToProcess == "y") {
    // 出力先ディレクトリの指定の入力を受付
    timer.start()
    val saveDir = readLine(" Input the directory in which generated csv files are saved: ./")
    timer.stop()
    if (!saveDir.isEmpty) {
      (new File(saveDir)).mkdir()
    }
    // エクセルをCSVに変換
    val results = for (path <- fileList) yield {
      // エクセルオブジェクトの生成
      val error = try {
        val excel = Excel(path)
        excel.setSaveDirectory(saveDir)
        excel.convertToCsv()
        ""
      } catch {
        case e => e.toString
      }
      val result = if (error.isEmpty) "Success" else "Failure"
      Map(
        "path" -> path,
        "file" -> path.split("/").last,
        "result" -> result,
        "error" -> error)
    }
    // 結果を出力
    println(" The result of converting excel to csv")

    printStringListWithIndex(for (result <- results) yield {
      "%s : %s" format (result("result"), result("file"))
    })
  }

  // 終了処理
  val totalTime = timer.totalTime
  val processTime = timer.processTime
  println("[Finish] total time: %.3f s, processing time: %.3f s" format (totalTime, processTime))

  // 文字列リストをインデックスと共に出力
  def printStringListWithIndex(list: List[String]) = list.zipWithIndex.foreach {
    case (str: String, i: Int) => println("    -> [%03d] %s" format (i, str))
  }
}


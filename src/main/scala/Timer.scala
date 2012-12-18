package app.excelconverter

class Timer() {

  val startTimestamp = getNow
  var waitTime = 0.0
  var tmpWaitStart = 0L

  def getNow() = System.currentTimeMillis
  def start() {
    if (tmpWaitStart > 0L) throw new Exception("Timer is already started")
    tmpWaitStart = getNow()
  }
  def stop() {
    if (tmpWaitStart == 0L) throw new Exception("Timer is not yet started")
    waitTime += (getNow() - tmpWaitStart).toDouble / 1000
    tmpWaitStart = 0L
  }

  def totalTime = (getNow() - startTimestamp).toDouble / 1000
  def processTime = totalTime - waitTime

}

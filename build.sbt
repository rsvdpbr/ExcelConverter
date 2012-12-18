
name := "ExcelConverter"

version := "1.0"

scalaVersion := "2.9.2"

mainClass := Some("app.excelconverter.Main")

libraryDependencies += "org.apache.poi" % "poi" % "3.8"

libraryDependencies += "org.apache.poi" % "poi-ooxml" % "3.8"

libraryDependencies += "org.apache.poi" % "ooxml-schemas" % "1.1"

libraryDependencies += "org.apache.poi" % "poi-ooxml-schemas" % "3.8"
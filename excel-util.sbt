name := "excel-util"

version := "0.9.1.14"

organization := "no.vedaadata"

scalaVersion := "3.3.3"

resolvers += "Vedaa Data Public" at "https://mymavenrepo.com/repo/UulFGWFKTwklJGmfuD8D/"

libraryDependencies += "no.vedaadata" %% "text-util" % "0.9.3"

libraryDependencies ++= Seq(
	"org.apache.poi" % "poi" % "5.3.0",
	"org.apache.poi" % "poi-ooxml" % "5.3.0"
)

libraryDependencies += "no.vedaadata" %% "generator-util" % "0.9.4" % "test"

libraryDependencies += "org.scalatest" %% "scalatest" % "3.2.14" % "test"

publishTo := Some("Vedaa Data Public publisher" at "https://mymavenrepo.com/repo/zPAvi2SoOMk6Bj2jtxNA/")
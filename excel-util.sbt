name := "excel-util"

version := "0.9.1.12"

organization := "no.vedaadata"

scalaVersion := "3.3.1"

resolvers += "Vedaa Data Public" at "https://mymavenrepo.com/repo/UulFGWFKTwklJGmfuD8D/"

libraryDependencies += "no.vedaadata" %% "text-util" % "0.9.3"

libraryDependencies ++= Seq(
	"org.apache.poi" % "poi" % "5.2.3",
	"org.apache.poi" % "poi-ooxml" % "5.2.3"
)

libraryDependencies += "org.scalatest" %% "scalatest" % "3.2.14" % "test"

publishTo := Some("Vedaa Data Public publisher" at "https://mymavenrepo.com/repo/zPAvi2SoOMk6Bj2jtxNA/")
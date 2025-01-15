package demo.write.file

import no.vedaadata.excel.*

import demo.write.*

@main def main(num: Int) =

  val items = Person.generator.generate(num)

  //  derived

  Excel.writeFile("temp/demo-write-file-derived.xlsx", items)

  //  manual layout

  Excel.writeFile("temp/demo-write-file-layout.xlsx", items)(using SheetWriter.fromLayout(Person.layout))
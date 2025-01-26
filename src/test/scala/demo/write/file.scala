package demo.write.file

import no.vedaadata.excel.*

import demo.write.*

@main def main(num: Int) =

  val items = Person.generator.generate(num)

  //  derived

  Excel.writeFile(items, "temp/demo-write-file-derived.xlsx")

  //  manual layout

  Excel.writeFile(items, "temp/demo-write-file-layout.xlsx")(using SheetWriter.fromLayout(PersonLayout))
package demo

import java.time.LocalDate

import no.vedaadata.excel.*

case class Person(
  firstName: String,
  lastName: Option[String],
  age: Int,
  fortune: Option[BigDecimal],
  birthDate: Option[LocalDate])

object Person:

  val bob = Person("Bob", None, 23, Some(100000), Some(LocalDate.of(2000, 1, 1)))
  val tom = Person("Tom", Some("Smith"), 34, None, None)

  val items = List(bob, tom)

@main def main =

  //  derived

  Excel.writeFile("demo-write-derived.xlsx", Person.items)

  //  manual layout

  import SheetWriter.Layout

  val layout = Layout[Person](
    Layout.Column(25)("Fornavn", _.firstName),
    Layout.Column(30)("Etternavn", _.lastName),
    Layout.Column("Alder", _.age),
    Layout.Column("Formue", _.fortune),
    Layout.Column(15)("Fødselsdato", _.birthDate))

  Excel.writeFile("demo-write-layout.xlsx", Person.items)(using SheetWriter.fromLayout(layout))
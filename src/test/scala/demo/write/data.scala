package demo.write

import java.time.LocalDate

import no.vedaadata.generator.Generator

case class Person(
  firstName: String,
  lastName: Option[String],
  age: Int,
  fortune: Option[BigDecimal],
  birthDate: Option[LocalDate])

object Person:

  val generator = 
    (Generator("Alex", "Bob", "Charlie"), 
    Generator("Doe", "Smith", "Johnson").andThen[Option], 
    Generator.between(20, 50), 
    Generator.between(1000, 100000).map(BigDecimal.apply).andThen[Option],
    Generator.between(LocalDate.of(2000, 1, 1), LocalDate.of(2005, 1, 1)).andThen[Option])
    .mapN(Person.apply)

  import no.vedaadata.excel.SheetWriter.Layout

  val layout = Layout[Person](
    Layout.Column(25)("First name", _.firstName),
    Layout.Column(30)("Last name", _.lastName),
    Layout.Column("Age", _.age),
    Layout.Column("Fortune", _.fortune),
    Layout.Column(15)("Date of birth", _.birthDate))



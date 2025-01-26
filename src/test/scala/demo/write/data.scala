package demo.write

import java.time.LocalDate

import no.vedaadata.generator.Generator

import no.vedaadata.excel.*

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


object PersonLayout extends Layout[Person]:
  def columns = List(
    Column(25)("First name", _.firstName),
    Column(30)("Last name", _.lastName),
    Column("Age", _.age),
    Column("Fortune", _.fortune),
    Column(15)("Date of birth", _.birthDate))



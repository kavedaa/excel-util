## Get it

```scala
resolvers += "Vedaa Data Public" at "https://mymavenrepo.com/repo/UulFGWFKTwklJGmfuD8D/"

libraryDependencies += "no.vedaadata" %% "excel-util" % "0.9.1.12"
```

## Use it

### Writing

```scala
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
```

```scala
//  automatically derived

Excel.writeFile("demo-write-derived.xlsx", Person.items)
```

```scala
//  manual layout

import SheetWriter.Layout

val layout = Layout[Person](
  Layout.Column(25)("Fornavn", _.firstName),
  Layout.Column(30)("Etternavn", _.lastName),
  Layout.Column("Alder", _.age),
  Layout.Column("Formue", _.fortune),
  Layout.Column(15)("Fødselsdato", _.birthDate))

Excel.writeFile("demo-write-layout.xlsx", Person.items)(using SheetWriter.fromLayout(layout))
```

### Reading

```scala
case class Person(
  firstName: String,
  lastName: String,
  birthDate: Option[LocalDate],
  city: Option[String])

val labels = ("First name", "Last name", "Birth date", "City")

val rowReaderFactory: Detection.RowReaderFactory[Person] = Detection.RowReaderFactory.derived(labels)

given HeaderPolicy.Header = HeaderPolicy.Header(1)  //  specify header row, zero-based, default is 0

val sheetReader = SheetReader.fromRowReaderFactory(rowReaderFactory)

val persons = Excel.readFile("data/example-for-read-header-policy.xlsx")(using sheetReader)
```
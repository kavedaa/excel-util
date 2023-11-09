package no.vedaadata.excel

import scala.util.*

import java.time.*

import org.scalatest.funsuite._
import org.scalatest.matchers.should._

class ReadbackTest extends AnyFunSuite with Matchers:

  case class Data(
    boolean: Boolean,
    string: String,
    char: Char,
    byte: Byte,
    short: Short,
    int: Int,
    long: Long,
    float: Float,
    double: Double,
    bigInt: BigInt,
    bigDecimal: BigDecimal,
    localDate: LocalDate,
    localTime: LocalTime,
    localDateTime: LocalDateTime)
    derives SheetWriter, SheetReader

  case class OptionData(
    boolean: Option[Boolean],
    string: Option[String],
    char: Option[Char],
    byte: Option[Byte],
    short: Option[Short],
    int: Option[Int],
    long: Option[Long],
    float: Option[Float],
    double: Option[Double],
    bigInt: Option[BigInt],
    bigDecimal: Option[BigDecimal],
    localDate: Option[LocalDate],
    localTime: Option[LocalTime],
    localDateTime: Option[LocalDateTime])
    derives SheetWriter, SheetReader

  test("base data types"):

    val data = Data(
      true,
      "abc",
      'A',
      12,
      1234,
      1234567,
      1234567890,
      1.2345,
      1.23456789,
      BigInt("1234567890"),
      BigDecimal("1234567890.12345"),
      LocalDate.now,
      LocalTime.now,
      LocalDateTime.now)

    val filename = "test-base.xlsx"

    Excel.writeFile(filename, Seq(data))
    val res = Excel.readFile[Data](filename)

    res.isSuccess shouldBe true
    res.get shouldEqual Seq(data)

  test("option types, when some"):

    val data = OptionData(
      Some(true),
      Some("abc"),
      Some('A'),
      Some(12),
      Some(1234),
      Some(1234567),
      Some(1234567890),
      Some(1.2345f),
      Some(1.23456789),
      Some(BigInt("1234567890")),
      Some(BigDecimal("1234567890.12345")),
      Some(LocalDate.now),
      Some(LocalTime.now),
      Some(LocalDateTime.now))

    val filename = "test-option-some.xlsx"

    Excel.writeFile(filename, Seq(data))
    val res = Excel.readFile[OptionData](filename)

    res.isSuccess shouldBe true
    res.get shouldEqual Seq(data)

  test("option types, when none"):

    val data = OptionData(
      None,
      None,
      None,
      None,
      None,
      None,
      None,
      None,
      None,
      None,
      None,
      None,
      None,
      None)

    val filename = "test-option-none.xlsx"

    Excel.writeFile(filename, Seq(data))
    val res = Excel.readFile[OptionData](filename)

    res.isSuccess shouldBe true
    res.get shouldEqual Seq(data)


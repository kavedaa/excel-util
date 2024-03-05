package no.vedaadata.excel

trait DataPolicy:
  def dataRowIndex: Int

sealed trait HeaderPolicy extends DataPolicy

// sealed trait StyledHeaderPolicy extends HeaderPolicy:
//   def cellStyleFactory: HeaderCellStyleFactory

object HeaderPolicy:

  object NoHeader extends HeaderPolicy:
    def dataRowIndex = 0

  class Header(val headerRowIndex: Int, val dataRowIndex: Int) extends HeaderPolicy:
    def this(headerRowIndex: Int)  = this(headerRowIndex, headerRowIndex + 1)

  object Header:
    given default: Header = Header(0, 1)
  
  given default: HeaderPolicy = Header.default
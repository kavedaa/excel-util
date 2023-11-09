package no.vedaadata.excel

import org.apache.poi.ss.usermodel.*

import no.vedaadata.text.LabelTransformer

sealed trait HeaderPolicy[+A](val offset: Int)

object HeaderPolicy:

  object NoHeader extends HeaderPolicy[Nothing](0)
  class SomeHeader[A](gap: Int)(using val cellStyleFactory: HeaderCellStyleFactory) extends HeaderPolicy(gap + 1)
  
  def Header[A](using HeaderCellStyleFactory) = SomeHeader(0)
  def HeaderGap[A](gap: Int)(using HeaderCellStyleFactory, LabelTransformer) = SomeHeader(gap)

  given [A](using HeaderCellStyleFactory): HeaderPolicy[A] = Header
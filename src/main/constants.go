package main

import "github.com/icewind666/html-to-excel-renderer/src/types"

// StyleAttrName Html style constants
const StyleAttrName = "style"

// ColspanAttrName Colspan attribute name
const ColspanAttrName = "colspan"

// TextAlignStyleAttr Text align attribute name
const TextAlignStyleAttr = "text-align"

// WordWrapStyleAttr Word wrap attribute name
const WordWrapStyleAttr = "word-wrap"

// BreakWordWrapStyleAttrValue Word wrap attribute name
const BreakWordWrapStyleAttrValue = "break-word"

// FontSizeStyleAttr Font size attribute name
const FontSizeStyleAttr = "font-size"

// FontWeightStyleAttr Font weight attribute name
const FontWeightStyleAttr = "font-weight"

// BorderStyleAttr Border style attribute name
const BorderStyleAttr = "border-style"

// BorderStyleAttrValue Border style attribute value
const BorderStyleAttrValue = "solid"

// BorderInheritanceStyleAttr Border inheritance style attribute name
const BorderInheritanceStyleAttr = "border-inheritance-type"

// BorderInheritanceStyleAttrValue Border inheritance style attribute value
const BorderInheritanceStyleAttrValue = "solid"

// WidthStyleAttr Width attribute name
const WidthStyleAttr = "width"

// MinWidthStyleAttr Minimum width attribute name
const MinWidthStyleAttr = "min-width"

// MaxWidthStyleAttr Maximum width attribute name
const MaxWidthStyleAttr = "max-width"

// HeightStyleAttr Height attribute name
const HeightStyleAttr = "height"

// MinHeightStyleAttr Minimum height attribute name
const MinHeightStyleAttr = "min-height"

// MaxHeightStyleAttr Maximum height attribute name
const MaxHeightStyleAttr = "max-height"
const TextVerticalAlignStyleAttrValue = "center"
const ExcelBorderTypeValue = "thin"
const TextVerticalAlignStyleAttr = "vertical-align" // values can be: top | middle | bottom | baseline
const TextVerticalAlignStyleMiddle = "center"
const TextVerticalAlignStyleTop = "top"
const TextVerticalAlignStyleBottom = "bottom"
const ValueTypeAttrName = "cell-type"
const BackgroundColorAttrName = "background-color"

const(
	FloatValueType types.ValueType = "float"
	BooleanValueType types.ValueType = "bool"
	StringValueType types.ValueType = "string"
)
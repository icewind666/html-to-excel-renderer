package types

// ValueType Cell formatting type indication
type ValueType string

// HtmlStyle Mapped style from html element
type HtmlStyle struct {
	TextAlign         string
	WordWrap          bool
	Width             float64
	Height            float64
	BorderInheritance bool
	BorderStyle       bool
	FontSize          float64
	IsBold            bool
	Colspan           int
	VerticalAlign     string
	CellValueType	  ValueType
	BackgroundColor   string
}

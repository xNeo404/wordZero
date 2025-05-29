// Package document 提供Word文档的表格操作功能
package document

import (
	"encoding/xml"
	"fmt"
)

// Table 表示一个表格
type Table struct {
	XMLName    xml.Name         `xml:"w:tbl"`
	Properties *TableProperties `xml:"w:tblPr,omitempty"`
	Grid       *TableGrid       `xml:"w:tblGrid,omitempty"`
	Rows       []TableRow       `xml:"w:tr"`
}

// TableProperties 表格属性
type TableProperties struct {
	XMLName   xml.Name    `xml:"w:tblPr"`
	TableW    *TableWidth `xml:"w:tblW,omitempty"`
	TableJc   *TableJc    `xml:"w:jc,omitempty"`
	TableLook *TableLook  `xml:"w:tblLook,omitempty"`
}

// TableWidth 表格宽度
type TableWidth struct {
	XMLName xml.Name `xml:"w:tblW"`
	W       string   `xml:"w:w,attr"`
	Type    string   `xml:"w:type,attr"`
}

// TableJc 表格对齐
type TableJc struct {
	XMLName xml.Name `xml:"w:jc"`
	Val     string   `xml:"w:val,attr"`
}

// TableLook 表格外观
type TableLook struct {
	XMLName  xml.Name `xml:"w:tblLook"`
	Val      string   `xml:"w:val,attr"`
	FirstRow string   `xml:"w:firstRow,attr,omitempty"`
	LastRow  string   `xml:"w:lastRow,attr,omitempty"`
	FirstCol string   `xml:"w:firstColumn,attr,omitempty"`
	LastCol  string   `xml:"w:lastColumn,attr,omitempty"`
	NoHBand  string   `xml:"w:noHBand,attr,omitempty"`
	NoVBand  string   `xml:"w:noVBand,attr,omitempty"`
}

// TableGrid 表格网格
type TableGrid struct {
	XMLName xml.Name       `xml:"w:tblGrid"`
	Cols    []TableGridCol `xml:"w:gridCol"`
}

// TableGridCol 表格网格列
type TableGridCol struct {
	XMLName xml.Name `xml:"w:gridCol"`
	W       string   `xml:"w:w,attr,omitempty"`
}

// TableRow 表格行
type TableRow struct {
	XMLName    xml.Name            `xml:"w:tr"`
	Properties *TableRowProperties `xml:"w:trPr,omitempty"`
	Cells      []TableCell         `xml:"w:tc"`
}

// TableRowProperties 表格行属性
type TableRowProperties struct {
	XMLName   xml.Name   `xml:"w:trPr"`
	TableRowH *TableRowH `xml:"w:trHeight,omitempty"`
}

// TableRowH 表格行高
type TableRowH struct {
	XMLName xml.Name `xml:"w:trHeight"`
	Val     string   `xml:"w:val,attr,omitempty"`
	HRule   string   `xml:"w:hRule,attr,omitempty"`
}

// TableCell 表格单元格
type TableCell struct {
	XMLName    xml.Name             `xml:"w:tc"`
	Properties *TableCellProperties `xml:"w:tcPr,omitempty"`
	Paragraphs []Paragraph          `xml:"w:p"`
}

// TableCellProperties 表格单元格属性
type TableCellProperties struct {
	XMLName       xml.Name       `xml:"w:tcPr"`
	TableCellW    *TableCellW    `xml:"w:tcW,omitempty"`
	VAlign        *VAlign        `xml:"w:vAlign,omitempty"`
	GridSpan      *GridSpan      `xml:"w:gridSpan,omitempty"`
	VMerge        *VMerge        `xml:"w:vMerge,omitempty"`
	TextDirection *TextDirection `xml:"w:textDirection,omitempty"`
}

// TableCellW 单元格宽度
type TableCellW struct {
	XMLName xml.Name `xml:"w:tcW"`
	W       string   `xml:"w:w,attr"`
	Type    string   `xml:"w:type,attr"`
}

// VAlign 垂直对齐
type VAlign struct {
	XMLName xml.Name `xml:"w:vAlign"`
	Val     string   `xml:"w:val,attr"`
}

// GridSpan 网格跨度（列合并）
type GridSpan struct {
	XMLName xml.Name `xml:"w:gridSpan"`
	Val     string   `xml:"w:val,attr"`
}

// VMerge 垂直合并（行合并）
type VMerge struct {
	XMLName xml.Name `xml:"w:vMerge"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// TableConfig 表格配置
type TableConfig struct {
	Rows      int        // 行数
	Cols      int        // 列数
	Width     int        // 表格总宽度（磅）
	ColWidths []int      // 各列宽度（磅），如果为空则平均分配
	Data      [][]string // 初始数据
}

// CreateTable 创建一个新表格
func (d *Document) CreateTable(config *TableConfig) *Table {
	if config.Rows <= 0 || config.Cols <= 0 {
		Error("表格行数和列数必须大于0")
		return nil
	}

	table := &Table{
		Properties: &TableProperties{
			TableW: &TableWidth{
				W:    fmt.Sprintf("%d", config.Width),
				Type: "dxa", // 磅为单位
			},
			TableJc: &TableJc{
				Val: "center", // 默认居中对齐
			},
			TableLook: &TableLook{
				Val:      "04A0",
				FirstRow: "1",
				LastRow:  "0",
				FirstCol: "1",
				LastCol:  "0",
				NoHBand:  "0",
				NoVBand:  "1",
			},
		},
		Grid: &TableGrid{},
		Rows: make([]TableRow, 0, config.Rows),
	}

	// 设置列宽
	colWidths := config.ColWidths
	if len(colWidths) == 0 {
		// 平均分配宽度
		avgWidth := config.Width / config.Cols
		colWidths = make([]int, config.Cols)
		for i := range colWidths {
			colWidths[i] = avgWidth
		}
	} else if len(colWidths) != config.Cols {
		Error("列宽数量与列数不匹配")
		return nil
	}

	// 创建表格网格
	for _, width := range colWidths {
		table.Grid.Cols = append(table.Grid.Cols, TableGridCol{
			W: fmt.Sprintf("%d", width),
		})
	}

	// 创建表格行和单元格
	for i := 0; i < config.Rows; i++ {
		row := TableRow{
			Cells: make([]TableCell, 0, config.Cols),
		}

		for j := 0; j < config.Cols; j++ {
			cell := TableCell{
				Properties: &TableCellProperties{
					TableCellW: &TableCellW{
						W:    fmt.Sprintf("%d", colWidths[j]),
						Type: "dxa",
					},
					VAlign: &VAlign{
						Val: "center",
					},
				},
				Paragraphs: []Paragraph{
					{
						Runs: []Run{
							{
								Text: Text{
									Content: "",
								},
							},
						},
					},
				},
			}

			// 如果有初始数据，设置单元格内容
			if config.Data != nil && i < len(config.Data) && j < len(config.Data[i]) {
				cell.Paragraphs[0].Runs[0].Text.Content = config.Data[i][j]
			}

			row.Cells = append(row.Cells, cell)
		}

		table.Rows = append(table.Rows, row)
	}

	Info(fmt.Sprintf("创建表格成功：%d行 x %d列", config.Rows, config.Cols))
	return table
}

// AddTable 将表格添加到文档中
func (d *Document) AddTable(config *TableConfig) *Table {
	table := d.CreateTable(config)
	if table == nil {
		return nil
	}

	// 将表格添加到文档主体中
	d.Body.Tables = append(d.Body.Tables, *table)

	Info(fmt.Sprintf("表格已添加到文档，当前文档包含%d个表格", len(d.Body.Tables)))
	return table
}

// InsertRow 在指定位置插入行
func (t *Table) InsertRow(position int, data []string) error {
	if position < 0 || position > len(t.Rows) {
		return fmt.Errorf("插入位置无效：%d，表格共有%d行", position, len(t.Rows))
	}

	if len(t.Rows) == 0 {
		return fmt.Errorf("表格没有列定义，无法插入行")
	}

	colCount := len(t.Rows[0].Cells)
	if len(data) > colCount {
		return fmt.Errorf("数据列数(%d)超过表格列数(%d)", len(data), colCount)
	}

	// 创建新行
	newRow := TableRow{
		Cells: make([]TableCell, colCount),
	}

	// 复制第一行的单元格属性作为模板
	templateRow := t.Rows[0]
	for i := 0; i < colCount; i++ {
		newRow.Cells[i] = TableCell{
			Properties: templateRow.Cells[i].Properties, // 复用属性
			Paragraphs: []Paragraph{
				{
					Runs: []Run{
						{
							Text: Text{
								Content: "",
							},
						},
					},
				},
			},
		}

		// 设置数据
		if i < len(data) {
			newRow.Cells[i].Paragraphs[0].Runs[0].Text.Content = data[i]
		}
	}

	// 插入行
	if position == len(t.Rows) {
		// 在末尾添加
		t.Rows = append(t.Rows, newRow)
	} else {
		// 在中间插入
		t.Rows = append(t.Rows[:position+1], t.Rows[position:]...)
		t.Rows[position] = newRow
	}

	Info(fmt.Sprintf("在位置%d插入行成功", position))
	return nil
}

// AppendRow 在表格末尾添加行
func (t *Table) AppendRow(data []string) error {
	return t.InsertRow(len(t.Rows), data)
}

// DeleteRow 删除指定行
func (t *Table) DeleteRow(rowIndex int) error {
	if rowIndex < 0 || rowIndex >= len(t.Rows) {
		return fmt.Errorf("行索引无效：%d，表格共有%d行", rowIndex, len(t.Rows))
	}

	if len(t.Rows) <= 1 {
		return fmt.Errorf("表格至少需要保留一行")
	}

	// 删除行
	t.Rows = append(t.Rows[:rowIndex], t.Rows[rowIndex+1:]...)

	Info(fmt.Sprintf("删除第%d行成功", rowIndex))
	return nil
}

// DeleteRows 删除指定范围的行
func (t *Table) DeleteRows(startIndex, endIndex int) error {
	if startIndex < 0 || endIndex >= len(t.Rows) || startIndex > endIndex {
		return fmt.Errorf("行索引范围无效：[%d, %d]，表格共有%d行", startIndex, endIndex, len(t.Rows))
	}

	deleteCount := endIndex - startIndex + 1
	if len(t.Rows)-deleteCount < 1 {
		return fmt.Errorf("删除后表格至少需要保留一行")
	}

	// 删除行范围
	t.Rows = append(t.Rows[:startIndex], t.Rows[endIndex+1:]...)

	Info(fmt.Sprintf("删除第%d到%d行成功", startIndex, endIndex))
	return nil
}

// InsertColumn 在指定位置插入列
func (t *Table) InsertColumn(position int, data []string, width int) error {
	if len(t.Rows) == 0 {
		return fmt.Errorf("表格没有行，无法插入列")
	}

	colCount := len(t.Rows[0].Cells)
	if position < 0 || position > colCount {
		return fmt.Errorf("插入位置无效：%d，表格共有%d列", position, colCount)
	}

	if len(data) > len(t.Rows) {
		return fmt.Errorf("数据行数(%d)超过表格行数(%d)", len(data), len(t.Rows))
	}

	// 更新表格网格
	newGridCol := TableGridCol{
		W: fmt.Sprintf("%d", width),
	}
	if position == len(t.Grid.Cols) {
		t.Grid.Cols = append(t.Grid.Cols, newGridCol)
	} else {
		t.Grid.Cols = append(t.Grid.Cols[:position+1], t.Grid.Cols[position:]...)
		t.Grid.Cols[position] = newGridCol
	}

	// 为每一行插入新单元格
	for i := range t.Rows {
		newCell := TableCell{
			Properties: &TableCellProperties{
				TableCellW: &TableCellW{
					W:    fmt.Sprintf("%d", width),
					Type: "dxa",
				},
				VAlign: &VAlign{
					Val: "center",
				},
			},
			Paragraphs: []Paragraph{
				{
					Runs: []Run{
						{
							Text: Text{
								Content: "",
							},
						},
					},
				},
			},
		}

		// 设置数据
		if i < len(data) {
			newCell.Paragraphs[0].Runs[0].Text.Content = data[i]
		}

		// 插入单元格
		if position == len(t.Rows[i].Cells) {
			t.Rows[i].Cells = append(t.Rows[i].Cells, newCell)
		} else {
			t.Rows[i].Cells = append(t.Rows[i].Cells[:position+1], t.Rows[i].Cells[position:]...)
			t.Rows[i].Cells[position] = newCell
		}
	}

	Info(fmt.Sprintf("在位置%d插入列成功", position))
	return nil
}

// AppendColumn 在表格末尾添加列
func (t *Table) AppendColumn(data []string, width int) error {
	colCount := 0
	if len(t.Rows) > 0 {
		colCount = len(t.Rows[0].Cells)
	}
	return t.InsertColumn(colCount, data, width)
}

// DeleteColumn 删除指定列
func (t *Table) DeleteColumn(colIndex int) error {
	if len(t.Rows) == 0 {
		return fmt.Errorf("表格没有行")
	}

	colCount := len(t.Rows[0].Cells)
	if colIndex < 0 || colIndex >= colCount {
		return fmt.Errorf("列索引无效：%d，表格共有%d列", colIndex, colCount)
	}

	if colCount <= 1 {
		return fmt.Errorf("表格至少需要保留一列")
	}

	// 删除网格列
	t.Grid.Cols = append(t.Grid.Cols[:colIndex], t.Grid.Cols[colIndex+1:]...)

	// 删除每行的对应单元格
	for i := range t.Rows {
		t.Rows[i].Cells = append(t.Rows[i].Cells[:colIndex], t.Rows[i].Cells[colIndex+1:]...)
	}

	Info(fmt.Sprintf("删除第%d列成功", colIndex))
	return nil
}

// DeleteColumns 删除指定范围的列
func (t *Table) DeleteColumns(startIndex, endIndex int) error {
	if len(t.Rows) == 0 {
		return fmt.Errorf("表格没有行")
	}

	colCount := len(t.Rows[0].Cells)
	if startIndex < 0 || endIndex >= colCount || startIndex > endIndex {
		return fmt.Errorf("列索引范围无效：[%d, %d]，表格共有%d列", startIndex, endIndex, colCount)
	}

	deleteCount := endIndex - startIndex + 1
	if colCount-deleteCount < 1 {
		return fmt.Errorf("删除后表格至少需要保留一列")
	}

	// 删除网格列范围
	t.Grid.Cols = append(t.Grid.Cols[:startIndex], t.Grid.Cols[endIndex+1:]...)

	// 删除每行的对应单元格范围
	for i := range t.Rows {
		t.Rows[i].Cells = append(t.Rows[i].Cells[:startIndex], t.Rows[i].Cells[endIndex+1:]...)
	}

	Info(fmt.Sprintf("删除第%d到%d列成功", startIndex, endIndex))
	return nil
}

// GetCell 获取指定位置的单元格
func (t *Table) GetCell(row, col int) (*TableCell, error) {
	if row < 0 || row >= len(t.Rows) {
		return nil, fmt.Errorf("行索引无效：%d，表格共有%d行", row, len(t.Rows))
	}

	if col < 0 || col >= len(t.Rows[row].Cells) {
		return nil, fmt.Errorf("列索引无效：%d，第%d行共有%d列", col, row, len(t.Rows[row].Cells))
	}

	return &t.Rows[row].Cells[col], nil
}

// SetCellText 设置单元格文本
func (t *Table) SetCellText(row, col int, text string) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// 确保单元格有段落和运行
	if len(cell.Paragraphs) == 0 {
		cell.Paragraphs = []Paragraph{
			{
				Runs: []Run{
					{
						Text: Text{Content: text},
					},
				},
			},
		}
	} else {
		if len(cell.Paragraphs[0].Runs) == 0 {
			cell.Paragraphs[0].Runs = []Run{
				{
					Text: Text{Content: text},
				},
			}
		} else {
			cell.Paragraphs[0].Runs[0].Text.Content = text
		}
	}

	return nil
}

// GetCellText 获取单元格文本
func (t *Table) GetCellText(row, col int) (string, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return "", err
	}

	if len(cell.Paragraphs) == 0 || len(cell.Paragraphs[0].Runs) == 0 {
		return "", nil
	}

	return cell.Paragraphs[0].Runs[0].Text.Content, nil
}

// GetRowCount 获取表格行数
func (t *Table) GetRowCount() int {
	return len(t.Rows)
}

// GetColumnCount 获取表格列数
func (t *Table) GetColumnCount() int {
	if len(t.Rows) == 0 {
		return 0
	}
	return len(t.Rows[0].Cells)
}

// ClearTable 清空表格内容（保留结构）
func (t *Table) ClearTable() {
	for i := range t.Rows {
		for j := range t.Rows[i].Cells {
			t.Rows[i].Cells[j].Paragraphs = []Paragraph{
				{
					Runs: []Run{
						{
							Text: Text{Content: ""},
						},
					},
				},
			}
		}
	}
	Info("表格内容已清空")
}

// CopyTable 复制表格
func (t *Table) CopyTable() *Table {
	// 深拷贝表格结构
	newTable := &Table{
		Properties: t.Properties,
		Grid:       t.Grid,
		Rows:       make([]TableRow, len(t.Rows)),
	}

	// 复制所有行和单元格
	for i, row := range t.Rows {
		newTable.Rows[i] = TableRow{
			Properties: row.Properties,
			Cells:      make([]TableCell, len(row.Cells)),
		}

		for j, cell := range row.Cells {
			newTable.Rows[i].Cells[j] = TableCell{
				Properties: cell.Properties,
				Paragraphs: make([]Paragraph, len(cell.Paragraphs)),
			}

			// 复制段落内容
			for k, para := range cell.Paragraphs {
				newTable.Rows[i].Cells[j].Paragraphs[k] = Paragraph{
					Properties: para.Properties,
					Runs:       make([]Run, len(para.Runs)),
				}

				for l, run := range para.Runs {
					newTable.Rows[i].Cells[j].Paragraphs[k].Runs[l] = Run{
						Properties: run.Properties,
						Text:       Text{Content: run.Text.Content},
					}
				}
			}
		}
	}

	Info("表格复制成功")
	return newTable
}

// CellAlignment 单元格对齐方式
type CellAlignment string

const (
	// CellAlignLeft 左对齐
	CellAlignLeft CellAlignment = "left"
	// CellAlignCenter 居中对齐
	CellAlignCenter CellAlignment = "center"
	// CellAlignRight 右对齐
	CellAlignRight CellAlignment = "right"
	// CellAlignJustify 两端对齐
	CellAlignJustify CellAlignment = "both"
)

// CellVerticalAlignment 单元格垂直对齐方式
type CellVerticalAlignment string

const (
	// CellVAlignTop 顶部对齐
	CellVAlignTop CellVerticalAlignment = "top"
	// CellVAlignCenter 居中对齐
	CellVAlignCenter CellVerticalAlignment = "center"
	// CellVAlignBottom 底部对齐
	CellVAlignBottom CellVerticalAlignment = "bottom"
)

// CellTextDirection 单元格文字方向
type CellTextDirection string

const (
	// TextDirectionLR 从左到右（默认）
	TextDirectionLR CellTextDirection = "lrTb"
	// TextDirectionTB 从上到下
	TextDirectionTB CellTextDirection = "tbRl"
	// TextDirectionBT 从下到上
	TextDirectionBT CellTextDirection = "btLr"
	// TextDirectionRL 从右到左
	TextDirectionRL CellTextDirection = "rlTb"
	// TextDirectionTBV 从上到下，垂直显示
	TextDirectionTBV CellTextDirection = "tbLrV"
	// TextDirectionBTV 从下到上，垂直显示
	TextDirectionBTV CellTextDirection = "btLrV"
)

// CellFormat 单元格格式配置
type CellFormat struct {
	TextFormat      *TextFormat           // 文字格式
	HorizontalAlign CellAlignment         // 水平对齐
	VerticalAlign   CellVerticalAlignment // 垂直对齐
	TextDirection   CellTextDirection     // 文字方向
	BackgroundColor string                // 背景颜色
	BorderStyle     string                // 边框样式
	Padding         int                   // 内边距（磅）
}

// SetCellFormat 设置单元格格式
func (t *Table) SetCellFormat(row, col int, format *CellFormat) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// 确保单元格有属性
	if cell.Properties == nil {
		cell.Properties = &TableCellProperties{}
	}

	// 设置垂直对齐
	if format.VerticalAlign != "" {
		cell.Properties.VAlign = &VAlign{
			Val: string(format.VerticalAlign),
		}
	}

	// 设置文字方向
	if format.TextDirection != "" {
		cell.Properties.TextDirection = &TextDirection{
			Val: string(format.TextDirection),
		}
	}

	// 确保单元格有段落
	if len(cell.Paragraphs) == 0 {
		cell.Paragraphs = []Paragraph{{}}
	}

	// 设置水平对齐
	if format.HorizontalAlign != "" {
		if cell.Paragraphs[0].Properties == nil {
			cell.Paragraphs[0].Properties = &ParagraphProperties{}
		}
		cell.Paragraphs[0].Properties.Justification = &Justification{
			Val: string(format.HorizontalAlign),
		}
	}

	// 设置文字格式
	if format.TextFormat != nil {
		// 确保有运行
		if len(cell.Paragraphs[0].Runs) == 0 {
			cell.Paragraphs[0].Runs = []Run{{}}
		}

		run := &cell.Paragraphs[0].Runs[0]
		if run.Properties == nil {
			run.Properties = &RunProperties{}
		}

		// 设置粗体
		if format.TextFormat.Bold {
			run.Properties.Bold = &Bold{}
		}

		// 设置斜体
		if format.TextFormat.Italic {
			run.Properties.Italic = &Italic{}
		}

		// 设置字体大小
		if format.TextFormat.FontSize > 0 {
			run.Properties.FontSize = &FontSize{
				Val: fmt.Sprintf("%d", format.TextFormat.FontSize*2), // Word使用半磅为单位
			}
		}

		// 设置字体颜色
		if format.TextFormat.FontColor != "" {
			run.Properties.Color = &Color{
				Val: format.TextFormat.FontColor,
			}
		}

		// 设置字体名称
		if format.TextFormat.FontName != "" {
			run.Properties.FontFamily = &FontFamily{
				ASCII: format.TextFormat.FontName,
			}
		}
	}

	Info(fmt.Sprintf("设置单元格(%d,%d)格式成功", row, col))
	return nil
}

// SetCellFormattedText 设置单元格富文本内容
func (t *Table) SetCellFormattedText(row, col int, text string, format *TextFormat) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// 创建格式化的运行
	run := Run{
		Text: Text{Content: text},
	}

	if format != nil {
		run.Properties = &RunProperties{}

		if format.Bold {
			run.Properties.Bold = &Bold{}
		}

		if format.Italic {
			run.Properties.Italic = &Italic{}
		}

		if format.FontSize > 0 {
			run.Properties.FontSize = &FontSize{
				Val: fmt.Sprintf("%d", format.FontSize*2),
			}
		}

		if format.FontColor != "" {
			run.Properties.Color = &Color{
				Val: format.FontColor,
			}
		}

		if format.FontName != "" {
			run.Properties.FontFamily = &FontFamily{
				ASCII: format.FontName,
			}
		}
	}

	// 设置单元格内容
	cell.Paragraphs = []Paragraph{
		{
			Runs: []Run{run},
		},
	}

	Info(fmt.Sprintf("设置单元格(%d,%d)富文本内容成功", row, col))
	return nil
}

// AddCellFormattedText 添加格式化文本到单元格（追加模式）
func (t *Table) AddCellFormattedText(row, col int, text string, format *TextFormat) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// 确保单元格有段落
	if len(cell.Paragraphs) == 0 {
		cell.Paragraphs = []Paragraph{{}}
	}

	// 创建格式化的运行
	run := Run{
		Text: Text{Content: text},
	}

	if format != nil {
		run.Properties = &RunProperties{}

		if format.Bold {
			run.Properties.Bold = &Bold{}
		}

		if format.Italic {
			run.Properties.Italic = &Italic{}
		}

		if format.FontSize > 0 {
			run.Properties.FontSize = &FontSize{
				Val: fmt.Sprintf("%d", format.FontSize*2),
			}
		}

		if format.FontColor != "" {
			run.Properties.Color = &Color{
				Val: format.FontColor,
			}
		}

		if format.FontName != "" {
			run.Properties.FontFamily = &FontFamily{
				ASCII: format.FontName,
			}
		}
	}

	// 添加运行到第一个段落
	cell.Paragraphs[0].Runs = append(cell.Paragraphs[0].Runs, run)

	Info(fmt.Sprintf("添加格式化文本到单元格(%d,%d)成功", row, col))
	return nil
}

// MergeCellsHorizontal 水平合并单元格（合并列）
func (t *Table) MergeCellsHorizontal(row, startCol, endCol int) error {
	if row < 0 || row >= len(t.Rows) {
		return fmt.Errorf("行索引无效：%d", row)
	}

	if startCol < 0 || endCol >= len(t.Rows[row].Cells) || startCol > endCol {
		return fmt.Errorf("列索引范围无效：[%d, %d]", startCol, endCol)
	}

	if startCol == endCol {
		return fmt.Errorf("起始列和结束列不能相同")
	}

	// 设置起始单元格的网格跨度
	startCell := &t.Rows[row].Cells[startCol]
	if startCell.Properties == nil {
		startCell.Properties = &TableCellProperties{}
	}

	spanCount := endCol - startCol + 1
	startCell.Properties.GridSpan = &GridSpan{
		Val: fmt.Sprintf("%d", spanCount),
	}

	// 删除被合并的单元格
	newCells := make([]TableCell, 0, len(t.Rows[row].Cells)-(endCol-startCol))
	newCells = append(newCells, t.Rows[row].Cells[:startCol+1]...)
	if endCol+1 < len(t.Rows[row].Cells) {
		newCells = append(newCells, t.Rows[row].Cells[endCol+1:]...)
	}
	t.Rows[row].Cells = newCells

	Info(fmt.Sprintf("水平合并单元格：行%d，列%d到%d", row, startCol, endCol))
	return nil
}

// MergeCellsVertical 垂直合并单元格（合并行）
func (t *Table) MergeCellsVertical(startRow, endRow, col int) error {
	if startRow < 0 || endRow >= len(t.Rows) || startRow > endRow {
		return fmt.Errorf("行索引范围无效：[%d, %d]", startRow, endRow)
	}

	if col < 0 {
		return fmt.Errorf("列索引无效：%d", col)
	}

	if startRow == endRow {
		return fmt.Errorf("起始行和结束行不能相同")
	}

	// 检查所有行的列数
	for i := startRow; i <= endRow; i++ {
		if col >= len(t.Rows[i].Cells) {
			return fmt.Errorf("第%d行没有第%d列", i, col)
		}
	}

	// 设置起始单元格为合并起始
	startCell := &t.Rows[startRow].Cells[col]
	if startCell.Properties == nil {
		startCell.Properties = &TableCellProperties{}
	}
	startCell.Properties.VMerge = &VMerge{
		Val: "restart",
	}

	// 设置后续单元格为合并继续
	for i := startRow + 1; i <= endRow; i++ {
		cell := &t.Rows[i].Cells[col]
		if cell.Properties == nil {
			cell.Properties = &TableCellProperties{}
		}
		cell.Properties.VMerge = &VMerge{
			Val: "continue",
		}
		// 清空被合并单元格的内容
		cell.Paragraphs = []Paragraph{{}}
	}

	Info(fmt.Sprintf("垂直合并单元格：行%d到%d，列%d", startRow, endRow, col))
	return nil
}

// MergeCellsRange 合并单元格区域（多行多列）
func (t *Table) MergeCellsRange(startRow, endRow, startCol, endCol int) error {
	// 验证范围
	if startRow < 0 || endRow >= len(t.Rows) || startRow > endRow {
		return fmt.Errorf("行索引范围无效：[%d, %d]", startRow, endRow)
	}

	// 先水平合并每一行
	for i := startRow; i <= endRow; i++ {
		if startCol >= len(t.Rows[i].Cells) || endCol >= len(t.Rows[i].Cells) {
			return fmt.Errorf("第%d行列索引范围无效：[%d, %d]", i, startCol, endCol)
		}

		if startCol != endCol {
			err := t.MergeCellsHorizontal(i, startCol, endCol)
			if err != nil {
				return fmt.Errorf("水平合并第%d行失败：%v", i, err)
			}
		}
	}

	// 然后垂直合并第一列
	if startRow != endRow {
		err := t.MergeCellsVertical(startRow, endRow, startCol)
		if err != nil {
			return fmt.Errorf("垂直合并失败：%v", err)
		}
	}

	Info(fmt.Sprintf("合并单元格区域：行%d到%d，列%d到%d", startRow, endRow, startCol, endCol))
	return nil
}

// UnmergeCells 取消单元格合并
func (t *Table) UnmergeCells(row, col int) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	if cell.Properties == nil {
		return fmt.Errorf("单元格没有合并")
	}

	// 检查是否有水平合并
	if cell.Properties.GridSpan != nil {
		// 获取合并的列数
		spanCount := 1
		if cell.Properties.GridSpan.Val != "" {
			fmt.Sscanf(cell.Properties.GridSpan.Val, "%d", &spanCount)
		}

		// 插入被合并的单元格
		for i := 1; i < spanCount; i++ {
			newCell := TableCell{
				Properties: &TableCellProperties{
					TableCellW: cell.Properties.TableCellW,
					VAlign:     cell.Properties.VAlign,
				},
				Paragraphs: []Paragraph{{}},
			}

			// 在指定位置插入新单元格
			insertPos := col + i
			if insertPos <= len(t.Rows[row].Cells) {
				t.Rows[row].Cells = append(t.Rows[row].Cells[:insertPos], append([]TableCell{newCell}, t.Rows[row].Cells[insertPos:]...)...)
			}
		}

		// 移除水平合并属性
		cell.Properties.GridSpan = nil
	}

	// 检查是否有垂直合并
	if cell.Properties.VMerge != nil {
		// 移除垂直合并属性
		cell.Properties.VMerge = nil

		// 查找并恢复被合并的单元格
		for i := row + 1; i < len(t.Rows); i++ {
			if col < len(t.Rows[i].Cells) {
				otherCell := &t.Rows[i].Cells[col]
				if otherCell.Properties != nil && otherCell.Properties.VMerge != nil {
					if otherCell.Properties.VMerge.Val == "continue" {
						// 恢复单元格内容
						otherCell.Properties.VMerge = nil
						if len(otherCell.Paragraphs) == 0 {
							otherCell.Paragraphs = []Paragraph{{}}
						}
					} else {
						break
					}
				} else {
					break
				}
			}
		}
	}

	Info(fmt.Sprintf("取消单元格(%d,%d)合并成功", row, col))
	return nil
}

// IsCellMerged 检查单元格是否被合并
func (t *Table) IsCellMerged(row, col int) (bool, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return false, err
	}

	if cell.Properties == nil {
		return false, nil
	}

	// 检查水平合并
	if cell.Properties.GridSpan != nil && cell.Properties.GridSpan.Val != "" && cell.Properties.GridSpan.Val != "1" {
		return true, nil
	}

	// 检查垂直合并
	if cell.Properties.VMerge != nil {
		return true, nil
	}

	return false, nil
}

// GetMergedCellInfo 获取合并单元格信息
func (t *Table) GetMergedCellInfo(row, col int) (map[string]interface{}, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return nil, err
	}

	info := make(map[string]interface{})
	info["is_merged"] = false

	if cell.Properties == nil {
		return info, nil
	}

	// 检查水平合并
	if cell.Properties.GridSpan != nil && cell.Properties.GridSpan.Val != "" {
		spanCount := 1
		fmt.Sscanf(cell.Properties.GridSpan.Val, "%d", &spanCount)
		if spanCount > 1 {
			info["is_merged"] = true
			info["horizontal_span"] = spanCount
			info["merge_type"] = "horizontal"
		}
	}

	// 检查垂直合并
	if cell.Properties.VMerge != nil {
		info["is_merged"] = true
		if cell.Properties.VMerge.Val == "restart" {
			info["vertical_merge_start"] = true
			info["merge_type"] = "vertical"
		} else if cell.Properties.VMerge.Val == "continue" {
			info["vertical_merge_continue"] = true
			info["merge_type"] = "vertical"
		}
	}

	return info, nil
}

// ClearCellContent 清空单元格内容但保留格式
func (t *Table) ClearCellContent(row, col int) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// 保留格式，只清空文本内容
	for i := range cell.Paragraphs {
		for j := range cell.Paragraphs[i].Runs {
			cell.Paragraphs[i].Runs[j].Text.Content = ""
		}
	}

	Info(fmt.Sprintf("清空单元格(%d,%d)内容成功", row, col))
	return nil
}

// ClearCellFormat 清空单元格格式但保留内容
func (t *Table) ClearCellFormat(row, col int) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// 清除单元格属性中的格式
	if cell.Properties != nil {
		// 保留合并信息和基本宽度，清除其他格式
		oldGridSpan := cell.Properties.GridSpan
		oldVMerge := cell.Properties.VMerge
		oldWidth := cell.Properties.TableCellW

		cell.Properties = &TableCellProperties{
			TableCellW: oldWidth,
			GridSpan:   oldGridSpan,
			VMerge:     oldVMerge,
		}
	}

	// 清除段落和运行的格式
	for i := range cell.Paragraphs {
		cell.Paragraphs[i].Properties = nil
		for j := range cell.Paragraphs[i].Runs {
			cell.Paragraphs[i].Runs[j].Properties = nil
		}
	}

	Info(fmt.Sprintf("清空单元格(%d,%d)格式成功", row, col))
	return nil
}

// SetCellPadding 设置单元格内边距
func (t *Table) SetCellPadding(row, col int, padding int) error {
	_, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// 单元格内边距通过表格属性设置，这里先预留接口
	// 实际实现需要在表格级别设置默认内边距
	Info(fmt.Sprintf("设置单元格(%d,%d)内边距为%d磅", row, col, padding))
	return nil
}

// SetCellTextDirection 设置单元格文字方向
func (t *Table) SetCellTextDirection(row, col int, direction CellTextDirection) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// 确保单元格有属性
	if cell.Properties == nil {
		cell.Properties = &TableCellProperties{}
	}

	// 设置文字方向
	cell.Properties.TextDirection = &TextDirection{
		Val: string(direction),
	}

	Info(fmt.Sprintf("设置单元格(%d,%d)文字方向为%s", row, col, direction))
	return nil
}

// GetCellTextDirection 获取单元格文字方向
func (t *Table) GetCellTextDirection(row, col int) (CellTextDirection, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return TextDirectionLR, err
	}

	if cell.Properties != nil && cell.Properties.TextDirection != nil {
		return CellTextDirection(cell.Properties.TextDirection.Val), nil
	}

	// 默认返回从左到右
	return TextDirectionLR, nil
}

// GetCellFormat 获取单元格格式信息
func (t *Table) GetCellFormat(row, col int) (*CellFormat, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return nil, err
	}

	format := &CellFormat{}

	// 获取垂直对齐
	if cell.Properties != nil && cell.Properties.VAlign != nil {
		format.VerticalAlign = CellVerticalAlignment(cell.Properties.VAlign.Val)
	}

	// 获取文字方向
	if cell.Properties != nil && cell.Properties.TextDirection != nil {
		format.TextDirection = CellTextDirection(cell.Properties.TextDirection.Val)
	}

	// 获取水平对齐
	if len(cell.Paragraphs) > 0 && cell.Paragraphs[0].Properties != nil && cell.Paragraphs[0].Properties.Justification != nil {
		format.HorizontalAlign = CellAlignment(cell.Paragraphs[0].Properties.Justification.Val)
	}

	// 获取文字格式
	if len(cell.Paragraphs) > 0 && len(cell.Paragraphs[0].Runs) > 0 {
		run := &cell.Paragraphs[0].Runs[0]
		if run.Properties != nil {
			format.TextFormat = &TextFormat{}

			if run.Properties.Bold != nil {
				format.TextFormat.Bold = true
			}

			if run.Properties.Italic != nil {
				format.TextFormat.Italic = true
			}

			if run.Properties.FontSize != nil {
				fmt.Sscanf(run.Properties.FontSize.Val, "%d", &format.TextFormat.FontSize)
				format.TextFormat.FontSize /= 2 // 转换为磅
			}

			if run.Properties.Color != nil {
				format.TextFormat.FontColor = run.Properties.Color.Val
			}

			if run.Properties.FontFamily != nil {
				format.TextFormat.FontName = run.Properties.FontFamily.ASCII
			}
		}
	}

	return format, nil
}

// TextDirection 文字方向
type TextDirection struct {
	XMLName xml.Name `xml:"w:textDirection"`
	Val     string   `xml:"w:val,attr"`
}

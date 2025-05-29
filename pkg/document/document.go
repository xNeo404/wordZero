// Package document 提供Word文档的核心操作功能
package document

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/ZeroHawkeye/wordZero/pkg/style"
)

// Document 表示一个Word文档
type Document struct {
	// 文档的主要内容
	Body *Body
	// 文档关系
	relationships *Relationships
	// 内容类型
	contentTypes *ContentTypes
	// 样式管理器
	styleManager *style.StyleManager
	// 临时存储文档部件
	parts map[string][]byte
}

// Body 表示文档主体
type Body struct {
	XMLName  xml.Name      `xml:"w:body"`
	Elements []interface{} `xml:"-"` // 不序列化此字段，使用自定义方法
}

// MarshalXML 自定义XML序列化，按照元素顺序输出
func (b *Body) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// 开始元素
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// 序列化每个元素，保持顺序
	for _, element := range b.Elements {
		if err := e.Encode(element); err != nil {
			return err
		}
	}

	// 结束元素
	return e.EncodeToken(start.End())
}

// BodyElement 文档主体元素接口
type BodyElement interface {
	ElementType() string
}

// ElementType 返回段落元素类型
func (p *Paragraph) ElementType() string {
	return "paragraph"
}

// ElementType 返回表格元素类型
func (t *Table) ElementType() string {
	return "table"
}

// Paragraph 表示一个段落
type Paragraph struct {
	XMLName    xml.Name             `xml:"w:p"`
	Properties *ParagraphProperties `xml:"w:pPr,omitempty"`
	Runs       []Run                `xml:"w:r"`
}

// ParagraphProperties 段落属性
type ParagraphProperties struct {
	XMLName        xml.Name        `xml:"w:pPr"`
	ParagraphStyle *ParagraphStyle `xml:"w:pStyle,omitempty"`
	Spacing        *Spacing        `xml:"w:spacing,omitempty"`
	Justification  *Justification  `xml:"w:jc,omitempty"`
	Indentation    *Indentation    `xml:"w:ind,omitempty"`
}

// Spacing 间距设置
type Spacing struct {
	XMLName xml.Name `xml:"w:spacing"`
	Before  string   `xml:"w:before,attr,omitempty"`
	After   string   `xml:"w:after,attr,omitempty"`
	Line    string   `xml:"w:line,attr,omitempty"`
}

// Justification 对齐方式
type Justification struct {
	XMLName xml.Name `xml:"w:jc"`
	Val     string   `xml:"w:val,attr"`
}

// Run 表示一段文本
type Run struct {
	XMLName    xml.Name       `xml:"w:r"`
	Properties *RunProperties `xml:"w:rPr,omitempty"`
	Text       Text           `xml:"w:t"`
}

// RunProperties 文本属性
type RunProperties struct {
	XMLName    xml.Name    `xml:"w:rPr"`
	Bold       *Bold       `xml:"w:b,omitempty"`
	Italic     *Italic     `xml:"w:i,omitempty"`
	FontSize   *FontSize   `xml:"w:sz,omitempty"`
	Color      *Color      `xml:"w:color,omitempty"`
	FontFamily *FontFamily `xml:"w:rFonts,omitempty"`
}

// Bold 粗体
type Bold struct {
	XMLName xml.Name `xml:"w:b"`
}

// Italic 斜体
type Italic struct {
	XMLName xml.Name `xml:"w:i"`
}

// FontSize 字体大小
type FontSize struct {
	XMLName xml.Name `xml:"w:sz"`
	Val     string   `xml:"w:val,attr"`
}

// Color 颜色
type Color struct {
	XMLName xml.Name `xml:"w:color"`
	Val     string   `xml:"w:val,attr"`
}

// Text 文本内容
type Text struct {
	XMLName xml.Name `xml:"w:t"`
	Space   string   `xml:"xml:space,attr,omitempty"`
	Content string   `xml:",chardata"`
}

// Relationships 文档关系
type Relationships struct {
	XMLName       xml.Name       `xml:"Relationships"`
	Xmlns         string         `xml:"xmlns,attr"`
	Relationships []Relationship `xml:"Relationship"`
}

// Relationship 单个关系
type Relationship struct {
	ID     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
}

// ContentTypes 内容类型
type ContentTypes struct {
	XMLName   xml.Name   `xml:"Types"`
	Xmlns     string     `xml:"xmlns,attr"`
	Defaults  []Default  `xml:"Default"`
	Overrides []Override `xml:"Override"`
}

// Default 默认内容类型
type Default struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// Override 覆盖内容类型
type Override struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// FontFamily 字体族
type FontFamily struct {
	XMLName xml.Name `xml:"w:rFonts"`
	ASCII   string   `xml:"w:ascii,attr,omitempty"`
}

// TextFormat 文本格式配置
type TextFormat struct {
	Bold      bool   // 是否粗体
	Italic    bool   // 是否斜体
	FontSize  int    // 字体大小（磅）
	FontColor string // 字体颜色（十六进制，如 "FF0000" 表示红色）
	FontName  string // 字体名称
}

// AlignmentType 对齐类型
type AlignmentType string

const (
	// AlignLeft 左对齐
	AlignLeft AlignmentType = "left"
	// AlignCenter 居中对齐
	AlignCenter AlignmentType = "center"
	// AlignRight 右对齐
	AlignRight AlignmentType = "right"
	// AlignJustify 两端对齐
	AlignJustify AlignmentType = "both"
)

// SpacingConfig 间距配置
type SpacingConfig struct {
	LineSpacing     float64 // 行间距（倍数，如1.5表示1.5倍行距）
	BeforePara      int     // 段前间距（磅）
	AfterPara       int     // 段后间距（磅）
	FirstLineIndent int     // 首行缩进（磅）
}

// Indentation 缩进设置
type Indentation struct {
	XMLName   xml.Name `xml:"w:ind"`
	FirstLine string   `xml:"w:firstLine,attr,omitempty"`
	Left      string   `xml:"w:left,attr,omitempty"`
	Right     string   `xml:"w:right,attr,omitempty"`
}

// ParagraphStyle 段落样式引用
type ParagraphStyle struct {
	XMLName xml.Name `xml:"w:pStyle"`
	Val     string   `xml:"w:val,attr"`
}

// New 创建一个新的空文档
func New() *Document {
	Debugf("创建新文档")

	doc := &Document{
		Body: &Body{
			Elements: make([]interface{}, 0),
		},
		styleManager: style.NewStyleManager(),
		parts:        make(map[string][]byte),
	}

	doc.initializeStructure()
	return doc
}

// Open 打开一个现有的Word文档。
//
// 参数 filename 是要打开的 .docx 文件路径。
// 该函数会解析整个文档结构，包括文本内容、格式和属性。
//
// 如果文件不存在、格式错误或解析失败，会返回相应的错误。
//
// 示例:
//
//	doc, err := document.Open("existing.docx")
//	if err != nil {
//		log.Fatal(err)
//	}
//
//	// 打印所有段落内容
//	for i, para := range doc.Body.Paragraphs {
//		fmt.Printf("段落 %d: ", i+1)
//		for _, run := range para.Runs {
//			fmt.Print(run.Text.Content)
//		}
//		fmt.Println()
//	}
func Open(filename string) (*Document, error) {
	Infof("正在打开文档: %s", filename)

	reader, err := zip.OpenReader(filename)
	if err != nil {
		Errorf("无法打开文件: %s", filename)
		return nil, WrapErrorWithContext("open_file", err, filename)
	}
	defer reader.Close()

	doc := &Document{
		parts: make(map[string][]byte),
	}

	// 读取所有文件部件
	for _, file := range reader.File {
		rc, err := file.Open()
		if err != nil {
			Errorf("无法打开文件部件: %s", file.Name)
			return nil, WrapErrorWithContext("open_part", err, file.Name)
		}

		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			Errorf("无法读取文件部件: %s", file.Name)
			return nil, WrapErrorWithContext("read_part", err, file.Name)
		}

		doc.parts[file.Name] = data
		Debugf("已读取文件部件: %s (%d 字节)", file.Name, len(data))
	}

	// 初始化样式管理器
	doc.styleManager = style.NewStyleManager()

	// 解析主文档
	if err := doc.parseDocument(); err != nil {
		Errorf("解析文档失败: %s", filename)
		return nil, WrapErrorWithContext("parse_document", err, filename)
	}

	Infof("成功打开文档: %s", filename)
	return doc, nil
}

// Save 将文档保存到指定的文件路径。
//
// 参数 filename 是保存文件的路径，包含文件名和扩展名。
// 如果目录不存在，会自动创建所需的目录结构。
//
// 保存过程包括序列化所有文档内容、压缩为ZIP格式，
// 并写入到文件系统。
//
// 示例:
//
//	doc := document.New()
//	doc.AddParagraph("示例内容")
//
//	// 保存到当前目录
//	err := doc.Save("example.docx")
//
//	// 保存到子目录（会自动创建目录）
//	err = doc.Save("output/documents/example.docx")
//
//	if err != nil {
//		log.Fatal(err)
//	}
func (d *Document) Save(filename string) error {
	Infof("正在保存文档: %s", filename)

	// 确保目录存在
	dir := filepath.Dir(filename)
	if err := os.MkdirAll(dir, 0755); err != nil {
		Errorf("无法创建目录: %s", dir)
		return WrapErrorWithContext("create_dir", err, dir)
	}

	// 创建文件
	file, err := os.Create(filename)
	if err != nil {
		Errorf("无法创建文件: %s", filename)
		return WrapErrorWithContext("create_file", err, filename)
	}
	defer file.Close()

	// 创建ZIP写入器
	zipWriter := zip.NewWriter(file)
	defer zipWriter.Close()

	// 序列化主文档
	if err := d.serializeDocument(); err != nil {
		Errorf("序列化文档失败")
		return WrapError("serialize_document", err)
	}

	// 序列化样式
	if err := d.serializeStyles(); err != nil {
		Errorf("序列化样式失败")
		return WrapError("serialize_styles", err)
	}

	// 序列化内容类型
	d.serializeContentTypes()

	// 序列化关系
	d.serializeRelationships()

	// 序列化文档关系
	d.serializeDocumentRelationships()

	// 写入所有部件
	for name, data := range d.parts {
		writer, err := zipWriter.Create(name)
		if err != nil {
			Errorf("无法创建ZIP条目: %s", name)
			return WrapErrorWithContext("create_zip_entry", err, name)
		}

		if _, err := writer.Write(data); err != nil {
			Errorf("无法写入ZIP条目: %s", name)
			return WrapErrorWithContext("write_zip_entry", err, name)
		}

		Debugf("已写入ZIP条目: %s (%d 字节)", name, len(data))
	}

	Infof("成功保存文档: %s", filename)
	return nil
}

// AddParagraph 向文档添加一个普通段落。
//
// 参数 text 是段落的文本内容。段落会使用默认格式，
// 可以后续通过返回的 Paragraph 指针设置格式和属性。
//
// 返回新创建段落的指针，可用于进一步格式化。
//
// 示例:
//
//	doc := document.New()
//
//	// 添加普通段落
//	para := doc.AddParagraph("这是一个段落")
//
//	// 设置段落属性
//	para.SetAlignment(document.AlignCenter)
//	para.SetSpacing(&document.SpacingConfig{
//		LineSpacing: 1.5,
//		BeforePara:  12,
//	})
func (d *Document) AddParagraph(text string) *Paragraph {
	Debugf("添加段落: %s", text)
	p := &Paragraph{
		Runs: []Run{
			{
				Text: Text{
					Content: text,
					Space:   "preserve",
				},
			},
		},
	}

	d.Body.Elements = append(d.Body.Elements, p)
	return p
}

// AddFormattedParagraph 向文档添加一个格式化段落。
//
// 参数 text 是段落的文本内容。
// 参数 format 指定文本格式，如果为 nil 则使用默认格式。
//
// 返回新创建段落的指针，可用于进一步设置段落属性。
//
// 示例:
//
//	doc := document.New()
//
//	// 创建格式配置
//	titleFormat := &document.TextFormat{
//		Bold:      true,
//		FontSize:  18,
//		FontColor: "FF0000", // 红色
//		FontName:  "微软雅黑",
//	}
//
//	// 添加格式化标题
//	title := doc.AddFormattedParagraph("文档标题", titleFormat)
//	title.SetAlignment(document.AlignCenter)
func (d *Document) AddFormattedParagraph(text string, format *TextFormat) *Paragraph {
	Debugf("添加格式化段落: %s", text)

	// 创建运行属性
	runProps := &RunProperties{}

	if format != nil {
		if format.Bold {
			runProps.Bold = &Bold{}
		}

		if format.Italic {
			runProps.Italic = &Italic{}
		}

		if format.FontSize > 0 {
			// Word中字体大小是半磅为单位，所以需要乘以2
			runProps.FontSize = &FontSize{Val: strconv.Itoa(format.FontSize * 2)}
		}

		if format.FontColor != "" {
			// 确保颜色格式正确（移除#前缀）
			color := strings.TrimPrefix(format.FontColor, "#")
			runProps.Color = &Color{Val: color}
		}

		if format.FontName != "" {
			runProps.FontFamily = &FontFamily{ASCII: format.FontName}
		}
	}

	p := &Paragraph{
		Runs: []Run{
			{
				Properties: runProps,
				Text: Text{
					Content: text,
					Space:   "preserve",
				},
			},
		},
	}

	d.Body.Elements = append(d.Body.Elements, p)
	return p
}

// SetAlignment 设置段落的对齐方式。
//
// 参数 alignment 指定对齐类型，支持以下值：
//   - AlignLeft: 左对齐（默认）
//   - AlignCenter: 居中对齐
//   - AlignRight: 右对齐
//   - AlignJustify: 两端对齐
//
// 示例:
//
//	para := doc.AddParagraph("居中标题")
//	para.SetAlignment(document.AlignCenter)
//
//	para2 := doc.AddParagraph("右对齐文本")
//	para2.SetAlignment(document.AlignRight)
func (p *Paragraph) SetAlignment(alignment AlignmentType) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	p.Properties.Justification = &Justification{Val: string(alignment)}
	Debugf("设置段落对齐方式: %s", alignment)
}

// SetSpacing 设置段落的间距配置。
//
// 参数 config 包含各种间距设置，如果为 nil 则不进行任何设置。
// 配置选项包括：
//   - LineSpacing: 行间距倍数（如 1.5 表示1.5倍行距）
//   - BeforePara: 段前间距（磅）
//   - AfterPara: 段后间距（磅）
//   - FirstLineIndent: 首行缩进（磅）
//
// 注意：间距值会自动转换为 Word 内部使用的 TWIPs 单位（1磅=20TWIPs）。
//
// 示例:
//
//	para := doc.AddParagraph("带间距的段落")
//
//	// 设置复杂间距
//	para.SetSpacing(&document.SpacingConfig{
//		LineSpacing:     1.5, // 1.5倍行距
//		BeforePara:      12,  // 段前12磅
//		AfterPara:       6,   // 段后6磅
//		FirstLineIndent: 24,  // 首行缩进24磅
//	})
//
//	// 只设置行间距
//	para2 := doc.AddParagraph("双倍行距")
//	para2.SetSpacing(&document.SpacingConfig{
//		LineSpacing: 2.0,
//	})
func (p *Paragraph) SetSpacing(config *SpacingConfig) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	if config != nil {
		spacing := &Spacing{}

		if config.BeforePara > 0 {
			// 转换为TWIPs (1/20磅)
			spacing.Before = strconv.Itoa(config.BeforePara * 20)
		}

		if config.AfterPara > 0 {
			// 转换为TWIPs (1/20磅)
			spacing.After = strconv.Itoa(config.AfterPara * 20)
		}

		if config.LineSpacing > 0 {
			// 行间距，240表示单倍行距
			spacing.Line = strconv.Itoa(int(config.LineSpacing * 240))
		}

		p.Properties.Spacing = spacing

		if config.FirstLineIndent > 0 {
			if p.Properties.Indentation == nil {
				p.Properties.Indentation = &Indentation{}
			}
			// 转换为TWIPs (1/20磅)
			p.Properties.Indentation.FirstLine = strconv.Itoa(config.FirstLineIndent * 20)
		}

		Debugf("设置段落间距: 段前=%d, 段后=%d, 行距=%.1f, 首行缩进=%d",
			config.BeforePara, config.AfterPara, config.LineSpacing, config.FirstLineIndent)
	}
}

// AddFormattedText 向段落添加格式化的文本内容。
//
// 此方法允许在一个段落中混合使用不同格式的文本。
// 新的文本会作为一个新的 Run 添加到段落中。
//
// 参数 text 是要添加的文本内容。
// 参数 format 指定文本格式，如果为 nil 则使用默认格式。
//
// 示例:
//
//	para := doc.AddParagraph("这个段落包含")
//
//	// 添加粗体红色文本
//	para.AddFormattedText("粗体红色", &document.TextFormat{
//		Bold: true,
//		FontColor: "FF0000",
//	})
//
//	// 添加普通文本
//	para.AddFormattedText("和普通文本", nil)
//
//	// 添加斜体蓝色文本
//	para.AddFormattedText("以及斜体蓝色", &document.TextFormat{
//		Italic: true,
//		FontColor: "0000FF",
//		FontSize: 14,
//	})
func (p *Paragraph) AddFormattedText(text string, format *TextFormat) {
	// 创建运行属性
	runProps := &RunProperties{}

	if format != nil {
		if format.Bold {
			runProps.Bold = &Bold{}
		}

		if format.Italic {
			runProps.Italic = &Italic{}
		}

		if format.FontSize > 0 {
			runProps.FontSize = &FontSize{Val: strconv.Itoa(format.FontSize * 2)}
		}

		if format.FontColor != "" {
			color := strings.TrimPrefix(format.FontColor, "#")
			runProps.Color = &Color{Val: color}
		}

		if format.FontName != "" {
			runProps.FontFamily = &FontFamily{ASCII: format.FontName}
		}
	}

	run := Run{
		Properties: runProps,
		Text: Text{
			Content: text,
			Space:   "preserve",
		},
	}

	p.Runs = append(p.Runs, run)
	Debugf("向段落添加格式化文本: %s", text)
}

// AddHeadingParagraph 向文档添加一个标题段落。
//
// 参数 text 是标题的文本内容。
// 参数 level 是标题级别（1-9），对应 Heading1 到 Heading9。
//
// 返回新创建段落的指针，可用于进一步设置段落属性。
// 此方法会自动设置正确的样式引用，确保标题能被 Word 导航窗格识别。
//
// 示例:
//
//	doc := document.New()
//
//	// 添加一级标题
//	h1 := doc.AddHeadingParagraph("第一章：概述", 1)
//
//	// 添加二级标题
//	h2 := doc.AddHeadingParagraph("1.1 背景", 2)
//
//	// 添加三级标题
//	h3 := doc.AddHeadingParagraph("1.1.1 研究目标", 3)
func (d *Document) AddHeadingParagraph(text string, level int) *Paragraph {
	if level < 1 || level > 9 {
		Debugf("标题级别 %d 超出范围，使用默认级别 1", level)
		level = 1
	}

	styleID := fmt.Sprintf("Heading%d", level)
	Debugf("添加标题段落: %s (级别: %d, 样式: %s)", text, level, styleID)

	// 获取样式管理器中的样式
	headingStyle := d.styleManager.GetStyle(styleID)
	if headingStyle == nil {
		Debugf("警告：找不到样式 %s，使用默认样式", styleID)
		return d.AddParagraph(text)
	}

	// 创建运行属性，应用样式中的字符格式
	runProps := &RunProperties{}
	if headingStyle.RunPr != nil {
		if headingStyle.RunPr.Bold != nil {
			runProps.Bold = &Bold{}
		}
		if headingStyle.RunPr.Italic != nil {
			runProps.Italic = &Italic{}
		}
		if headingStyle.RunPr.FontSize != nil {
			runProps.FontSize = &FontSize{Val: headingStyle.RunPr.FontSize.Val}
		}
		if headingStyle.RunPr.Color != nil {
			runProps.Color = &Color{Val: headingStyle.RunPr.Color.Val}
		}
		if headingStyle.RunPr.FontFamily != nil {
			runProps.FontFamily = &FontFamily{ASCII: headingStyle.RunPr.FontFamily.ASCII}
		}
	}

	// 创建段落属性，应用样式中的段落格式
	paraProps := &ParagraphProperties{
		ParagraphStyle: &ParagraphStyle{Val: styleID},
	}

	// 应用样式中的段落格式
	if headingStyle.ParagraphPr != nil {
		if headingStyle.ParagraphPr.Spacing != nil {
			paraProps.Spacing = &Spacing{
				Before: headingStyle.ParagraphPr.Spacing.Before,
				After:  headingStyle.ParagraphPr.Spacing.After,
				Line:   headingStyle.ParagraphPr.Spacing.Line,
			}
		}
		if headingStyle.ParagraphPr.Justification != nil {
			paraProps.Justification = &Justification{
				Val: headingStyle.ParagraphPr.Justification.Val,
			}
		}
		if headingStyle.ParagraphPr.Indentation != nil {
			paraProps.Indentation = &Indentation{
				FirstLine: headingStyle.ParagraphPr.Indentation.FirstLine,
				Left:      headingStyle.ParagraphPr.Indentation.Left,
				Right:     headingStyle.ParagraphPr.Indentation.Right,
			}
		}
	}

	// 创建段落
	p := &Paragraph{
		Properties: paraProps,
		Runs: []Run{
			{
				Properties: runProps,
				Text: Text{
					Content: text,
					Space:   "preserve",
				},
			},
		},
	}

	d.Body.Elements = append(d.Body.Elements, p)
	return p
}

// SetStyle 设置段落的样式。
//
// 参数 styleID 是要应用的样式ID，如 "Heading1"、"Normal" 等。
// 此方法会设置段落的样式引用，确保段落使用指定的样式。
//
// 示例:
//
//	para := doc.AddParagraph("这是一个段落")
//	para.SetStyle("Heading2")  // 设置为二级标题样式
func (p *Paragraph) SetStyle(styleID string) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	p.Properties.ParagraphStyle = &ParagraphStyle{Val: styleID}
	Debugf("设置段落样式: %s", styleID)
}

// GetStyleManager 获取文档的样式管理器。
//
// 返回文档的样式管理器，可用于访问和管理样式。
//
// 示例:
//
//	doc := document.New()
//	styleManager := doc.GetStyleManager()
//	headingStyle := styleManager.GetStyle("Heading1")
func (d *Document) GetStyleManager() *style.StyleManager {
	return d.styleManager
}

// initializeStructure 初始化文档基础结构
func (d *Document) initializeStructure() {
	// 初始化 content types
	d.contentTypes = &ContentTypes{
		Xmlns: "http://schemas.openxmlformats.org/package/2006/content-types",
		Defaults: []Default{
			{Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml"},
			{Extension: "xml", ContentType: "application/xml"},
		},
		Overrides: []Override{
			{PartName: "/word/document.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"},
			{PartName: "/word/styles.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"},
		},
	}

	// 初始化主关系
	d.relationships = &Relationships{
		Xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
		Relationships: []Relationship{
			{
				ID:     "rId1",
				Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
				Target: "word/document.xml",
			},
		},
	}

	// 添加基础部件
	d.serializeContentTypes()
	d.serializeRelationships()
	d.serializeDocumentRelationships()
}

// parseDocument 解析文档内容
func (d *Document) parseDocument() error {
	Debugf("开始解析文档内容")

	// 解析主文档
	docData, ok := d.parts["word/document.xml"]
	if !ok {
		return WrapError("parse_document", ErrDocumentNotFound)
	}

	// 创建一个临时结构来解析完整的文档
	var doc struct {
		XMLName xml.Name `xml:"document"`
		Body    struct {
			XMLName    xml.Name `xml:"body"`
			Paragraphs []struct {
				XMLName    xml.Name `xml:"p"`
				Properties *struct {
					XMLName        xml.Name `xml:"pPr"`
					ParagraphStyle *struct {
						XMLName xml.Name `xml:"pStyle"`
						Val     string   `xml:"val,attr"`
					} `xml:"pStyle,omitempty"`
					Spacing *struct {
						XMLName xml.Name `xml:"spacing"`
						Before  string   `xml:"before,attr,omitempty"`
						After   string   `xml:"after,attr,omitempty"`
						Line    string   `xml:"line,attr,omitempty"`
					} `xml:"spacing,omitempty"`
					Justification *struct {
						XMLName xml.Name `xml:"jc"`
						Val     string   `xml:"val,attr"`
					} `xml:"jc,omitempty"`
					Indentation *struct {
						XMLName   xml.Name `xml:"ind"`
						FirstLine string   `xml:"firstLine,attr,omitempty"`
						Left      string   `xml:"left,attr,omitempty"`
						Right     string   `xml:"right,attr,omitempty"`
					} `xml:"ind,omitempty"`
				} `xml:"pPr,omitempty"`
				Runs []struct {
					XMLName    xml.Name `xml:"r"`
					Properties *struct {
						XMLName xml.Name `xml:"rPr"`
						Bold    *struct {
							XMLName xml.Name `xml:"b"`
						} `xml:"b,omitempty"`
						Italic *struct {
							XMLName xml.Name `xml:"i"`
						} `xml:"i,omitempty"`
						FontSize *struct {
							XMLName xml.Name `xml:"sz"`
							Val     string   `xml:"val,attr"`
						} `xml:"sz,omitempty"`
						Color *struct {
							XMLName xml.Name `xml:"color"`
							Val     string   `xml:"val,attr"`
						} `xml:"color,omitempty"`
						FontFamily *struct {
							XMLName xml.Name `xml:"rFonts"`
							ASCII   string   `xml:"ascii,attr,omitempty"`
						} `xml:"rFonts,omitempty"`
					} `xml:"rPr,omitempty"`
					Text struct {
						XMLName xml.Name `xml:"t"`
						Space   string   `xml:"space,attr,omitempty"`
						Content string   `xml:",chardata"`
					} `xml:"t"`
				} `xml:"r"`
			} `xml:"p"`
		} `xml:"body"`
	}

	if err := xml.Unmarshal(docData, &doc); err != nil {
		Errorf("XML解析失败: %v", err)
		return WrapError("unmarshal_xml", err)
	}

	// 转换为内部结构
	d.Body = &Body{
		Elements: make([]interface{}, len(doc.Body.Paragraphs)),
	}

	for i, p := range doc.Body.Paragraphs {
		paragraph := &Paragraph{
			Runs: make([]Run, len(p.Runs)),
		}

		// 转换段落属性
		if p.Properties != nil {
			paragraph.Properties = &ParagraphProperties{}

			if p.Properties.ParagraphStyle != nil {
				paragraph.Properties.ParagraphStyle = &ParagraphStyle{
					Val: p.Properties.ParagraphStyle.Val,
				}
			}

			if p.Properties.Spacing != nil {
				paragraph.Properties.Spacing = &Spacing{
					Before: p.Properties.Spacing.Before,
					After:  p.Properties.Spacing.After,
					Line:   p.Properties.Spacing.Line,
				}
			}

			if p.Properties.Justification != nil {
				paragraph.Properties.Justification = &Justification{
					Val: p.Properties.Justification.Val,
				}
			}

			if p.Properties.Indentation != nil {
				paragraph.Properties.Indentation = &Indentation{
					FirstLine: p.Properties.Indentation.FirstLine,
					Left:      p.Properties.Indentation.Left,
					Right:     p.Properties.Indentation.Right,
				}
			}
		}

		for j, r := range p.Runs {
			run := Run{
				Text: Text{
					Content: r.Text.Content,
				},
			}

			if r.Properties != nil {
				run.Properties = &RunProperties{}

				if r.Properties.Bold != nil {
					run.Properties.Bold = &Bold{}
				}

				if r.Properties.Italic != nil {
					run.Properties.Italic = &Italic{}
				}

				if r.Properties.FontSize != nil {
					run.Properties.FontSize = &FontSize{
						Val: r.Properties.FontSize.Val,
					}
				}

				if r.Properties.Color != nil {
					run.Properties.Color = &Color{
						Val: r.Properties.Color.Val,
					}
				}

				if r.Properties.FontFamily != nil {
					run.Properties.FontFamily = &FontFamily{
						ASCII: r.Properties.FontFamily.ASCII,
					}
				}
			}

			paragraph.Runs[j] = run
		}

		d.Body.Elements[i] = paragraph
	}

	Infof("解析完成，共 %d 个元素", len(d.Body.Elements))
	return nil
}

// serializeDocument 序列化文档内容
func (d *Document) serializeDocument() error {
	Debugf("开始序列化文档")

	// 创建文档结构
	type documentXML struct {
		XMLName xml.Name `xml:"w:document"`
		Xmlns   string   `xml:"xmlns:w,attr"`
		Body    *Body    `xml:"w:body"`
	}

	doc := documentXML{
		Xmlns: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		Body:  d.Body,
	}

	// 序列化为XML
	data, err := xml.MarshalIndent(doc, "", "  ")
	if err != nil {
		Errorf("XML序列化失败: %v", err)
		return WrapError("marshal_xml", err)
	}

	// 添加XML声明
	d.parts["word/document.xml"] = append([]byte(xml.Header), data...)

	Debugf("文档序列化完成")
	return nil
}

// serializeContentTypes 序列化内容类型
func (d *Document) serializeContentTypes() {
	data, _ := xml.MarshalIndent(d.contentTypes, "", "  ")
	d.parts["[Content_Types].xml"] = append([]byte(xml.Header), data...)
}

// serializeRelationships 序列化关系
func (d *Document) serializeRelationships() {
	data, _ := xml.MarshalIndent(d.relationships, "", "  ")
	d.parts["_rels/.rels"] = append([]byte(xml.Header), data...)
}

// serializeDocumentRelationships 序列化文档关系
func (d *Document) serializeDocumentRelationships() {
	// 创建文档关系
	docRels := &Relationships{
		Xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
		Relationships: []Relationship{
			{
				ID:     "rId1",
				Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
				Target: "styles.xml",
			},
		},
	}

	data, _ := xml.MarshalIndent(docRels, "", "  ")
	d.parts["word/_rels/document.xml.rels"] = append([]byte(xml.Header), data...)
}

// serializeStyles 序列化样式
func (d *Document) serializeStyles() error {
	Debugf("开始序列化样式")

	// 创建样式结构，包含完整的命名空间
	type stylesXML struct {
		XMLName     xml.Name       `xml:"w:styles"`
		XmlnsW      string         `xml:"xmlns:w,attr"`
		XmlnsMC     string         `xml:"xmlns:mc,attr"`
		XmlnsO      string         `xml:"xmlns:o,attr"`
		XmlnsR      string         `xml:"xmlns:r,attr"`
		XmlnsM      string         `xml:"xmlns:m,attr"`
		XmlnsV      string         `xml:"xmlns:v,attr"`
		XmlnsW14    string         `xml:"xmlns:w14,attr"`
		XmlnsW10    string         `xml:"xmlns:w10,attr"`
		XmlnsSL     string         `xml:"xmlns:sl,attr"`
		XmlnsWPS    string         `xml:"xmlns:wpsCustomData,attr"`
		MCIgnorable string         `xml:"mc:Ignorable,attr"`
		Styles      []*style.Style `xml:"w:style"`
	}

	doc := stylesXML{
		XmlnsW:      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		XmlnsMC:     "http://schemas.openxmlformats.org/markup-compatibility/2006",
		XmlnsO:      "urn:schemas-microsoft-com:office:office",
		XmlnsR:      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		XmlnsM:      "http://schemas.openxmlformats.org/officeDocument/2006/math",
		XmlnsV:      "urn:schemas-microsoft-com:vml",
		XmlnsW14:    "http://schemas.microsoft.com/office/word/2010/wordml",
		XmlnsW10:    "urn:schemas-microsoft-com:office:word",
		XmlnsSL:     "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
		XmlnsWPS:    "http://www.wps.cn/officeDocument/2013/wpsCustomData",
		MCIgnorable: "w14",
		Styles:      d.styleManager.GetAllStyles(),
	}

	// 序列化为XML
	data, err := xml.MarshalIndent(doc, "", "  ")
	if err != nil {
		Errorf("XML序列化失败: %v", err)
		return WrapError("marshal_xml", err)
	}

	// 添加XML声明
	d.parts["word/styles.xml"] = append([]byte(xml.Header), data...)

	Debugf("样式序列化完成")
	return nil
}

// ToBytes 将文档转换为字节数组
func (d *Document) ToBytes() ([]byte, error) {
	buf := new(bytes.Buffer)
	zipWriter := zip.NewWriter(buf)

	// 序列化文档
	if err := d.serializeDocument(); err != nil {
		return nil, err
	}

	// 序列化样式
	if err := d.serializeStyles(); err != nil {
		return nil, err
	}

	// 序列化内容类型
	d.serializeContentTypes()

	// 序列化关系
	d.serializeRelationships()

	// 序列化文档关系
	d.serializeDocumentRelationships()

	// 写入所有部件
	for name, data := range d.parts {
		writer, err := zipWriter.Create(name)
		if err != nil {
			return nil, err
		}
		if _, err := writer.Write(data); err != nil {
			return nil, err
		}
	}

	if err := zipWriter.Close(); err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}

// GetParagraphs 获取所有段落
func (b *Body) GetParagraphs() []*Paragraph {
	paragraphs := make([]*Paragraph, 0)
	for _, element := range b.Elements {
		if p, ok := element.(*Paragraph); ok {
			paragraphs = append(paragraphs, p)
		}
	}
	return paragraphs
}

// GetTables 获取所有表格
func (b *Body) GetTables() []*Table {
	tables := make([]*Table, 0)
	for _, element := range b.Elements {
		if t, ok := element.(*Table); ok {
			tables = append(tables, t)
		}
	}
	return tables
}

// AddElement 添加元素到文档主体
func (b *Body) AddElement(element interface{}) {
	b.Elements = append(b.Elements, element)
}

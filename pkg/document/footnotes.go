// Package document 提供Word文档脚注和尾注操作功能
package document

import (
	"encoding/xml"
	"fmt"
	"strconv"
)

// FootnoteType 脚注类型
type FootnoteType string

const (
	// FootnoteTypeFootnote 脚注
	FootnoteTypeFootnote FootnoteType = "footnote"
	// FootnoteTypeEndnote 尾注
	FootnoteTypeEndnote FootnoteType = "endnote"
)

// Footnotes 脚注集合
type Footnotes struct {
	XMLName   xml.Name    `xml:"w:footnotes"`
	Xmlns     string      `xml:"xmlns:w,attr"`
	Footnotes []*Footnote `xml:"w:footnote"`
}

// Endnotes 尾注集合
type Endnotes struct {
	XMLName  xml.Name   `xml:"w:endnotes"`
	Xmlns    string     `xml:"xmlns:w,attr"`
	Endnotes []*Endnote `xml:"w:endnote"`
}

// Footnote 脚注结构
type Footnote struct {
	XMLName    xml.Name     `xml:"w:footnote"`
	Type       string       `xml:"w:type,attr,omitempty"`
	ID         string       `xml:"w:id,attr"`
	Paragraphs []*Paragraph `xml:"w:p"`
}

// Endnote 尾注结构
type Endnote struct {
	XMLName    xml.Name     `xml:"w:endnote"`
	Type       string       `xml:"w:type,attr,omitempty"`
	ID         string       `xml:"w:id,attr"`
	Paragraphs []*Paragraph `xml:"w:p"`
}

// FootnoteReference 脚注引用
type FootnoteReference struct {
	XMLName xml.Name `xml:"w:footnoteReference"`
	ID      string   `xml:"w:id,attr"`
}

// EndnoteReference 尾注引用
type EndnoteReference struct {
	XMLName xml.Name `xml:"w:endnoteReference"`
	ID      string   `xml:"w:id,attr"`
}

// FootnoteConfig 脚注配置
type FootnoteConfig struct {
	NumberFormat FootnoteNumberFormat // 编号格式
	StartNumber  int                  // 起始编号
	RestartEach  FootnoteRestart      // 重新开始规则
	Position     FootnotePosition     // 位置
}

// FootnoteNumberFormat 脚注编号格式
type FootnoteNumberFormat string

const (
	// FootnoteFormatDecimal 十进制数字
	FootnoteFormatDecimal FootnoteNumberFormat = "decimal"
	// FootnoteFormatLowerRoman 小写罗马数字
	FootnoteFormatLowerRoman FootnoteNumberFormat = "lowerRoman"
	// FootnoteFormatUpperRoman 大写罗马数字
	FootnoteFormatUpperRoman FootnoteNumberFormat = "upperRoman"
	// FootnoteFormatLowerLetter 小写字母
	FootnoteFormatLowerLetter FootnoteNumberFormat = "lowerLetter"
	// FootnoteFormatUpperLetter 大写字母
	FootnoteFormatUpperLetter FootnoteNumberFormat = "upperLetter"
	// FootnoteFormatSymbol 符号
	FootnoteFormatSymbol FootnoteNumberFormat = "symbol"
)

// FootnoteRestart 脚注重新开始规则
type FootnoteRestart string

const (
	// FootnoteRestartContinuous 连续编号
	FootnoteRestartContinuous FootnoteRestart = "continuous"
	// FootnoteRestartEachSection 每节重新开始
	FootnoteRestartEachSection FootnoteRestart = "eachSect"
	// FootnoteRestartEachPage 每页重新开始
	FootnoteRestartEachPage FootnoteRestart = "eachPage"
)

// FootnotePosition 脚注位置
type FootnotePosition string

const (
	// FootnotePositionPageBottom 页面底部
	FootnotePositionPageBottom FootnotePosition = "pageBottom"
	// FootnotePositionBeneathText 文本下方
	FootnotePositionBeneathText FootnotePosition = "beneathText"
	// FootnotePositionSectionEnd 节末尾
	FootnotePositionSectionEnd FootnotePosition = "sectEnd"
	// FootnotePositionDocumentEnd 文档末尾
	FootnotePositionDocumentEnd FootnotePosition = "docEnd"
)

// 全局脚注/尾注管理器
var globalFootnoteManager *FootnoteManager

// FootnoteManager 脚注管理器
type FootnoteManager struct {
	nextFootnoteID int
	nextEndnoteID  int
	footnotes      map[string]*Footnote
	endnotes       map[string]*Endnote
}

// getFootnoteManager 获取全局脚注管理器
func getFootnoteManager() *FootnoteManager {
	if globalFootnoteManager == nil {
		globalFootnoteManager = &FootnoteManager{
			nextFootnoteID: 1,
			nextEndnoteID:  1,
			footnotes:      make(map[string]*Footnote),
			endnotes:       make(map[string]*Endnote),
		}
	}
	return globalFootnoteManager
}

// DefaultFootnoteConfig 返回默认脚注配置
func DefaultFootnoteConfig() *FootnoteConfig {
	return &FootnoteConfig{
		NumberFormat: FootnoteFormatDecimal,
		StartNumber:  1,
		RestartEach:  FootnoteRestartContinuous,
		Position:     FootnotePositionPageBottom,
	}
}

// AddFootnote 添加脚注
func (d *Document) AddFootnote(text string, footnoteText string) error {
	return d.addFootnoteOrEndnote(text, footnoteText, FootnoteTypeFootnote)
}

// AddEndnote 添加尾注
func (d *Document) AddEndnote(text string, endnoteText string) error {
	return d.addFootnoteOrEndnote(text, endnoteText, FootnoteTypeEndnote)
}

// addFootnoteOrEndnote 添加脚注或尾注的通用方法
func (d *Document) addFootnoteOrEndnote(text string, noteText string, noteType FootnoteType) error {
	manager := getFootnoteManager()

	// 确保脚注/尾注系统已初始化
	d.ensureFootnoteInitialized(noteType)

	var noteID string
	if noteType == FootnoteTypeFootnote {
		noteID = strconv.Itoa(manager.nextFootnoteID)
		manager.nextFootnoteID++
	} else {
		noteID = strconv.Itoa(manager.nextEndnoteID)
		manager.nextEndnoteID++
	}

	// 创建包含脚注引用的段落
	paragraph := &Paragraph{}

	// 添加正文文本
	if text != "" {
		textRun := Run{
			Text: Text{Content: text},
		}
		paragraph.Runs = append(paragraph.Runs, textRun)
	}

	// 添加脚注/尾注引用
	refRun := Run{
		Properties: &RunProperties{},
	}

	if noteType == FootnoteTypeFootnote {
		// 简化处理：在文本中插入脚注标记
		refRun.Text = Text{Content: fmt.Sprintf("[%s]", noteID)}
	} else {
		// 简化处理：在文本中插入尾注标记
		refRun.Text = Text{Content: fmt.Sprintf("[尾注%s]", noteID)}
	}

	paragraph.Runs = append(paragraph.Runs, refRun)
	d.Body.Elements = append(d.Body.Elements, paragraph)

	// 创建脚注/尾注内容
	if err := d.createNoteContent(noteID, noteText, noteType); err != nil {
		return fmt.Errorf("创建%s内容失败: %v", noteType, err)
	}

	return nil
}

// AddFootnoteToRun 在现有Run中添加脚注引用
func (d *Document) AddFootnoteToRun(run *Run, footnoteText string) error {
	manager := getFootnoteManager()
	d.ensureFootnoteInitialized(FootnoteTypeFootnote)

	noteID := strconv.Itoa(manager.nextFootnoteID)
	manager.nextFootnoteID++

	// 在当前Run后添加脚注引用
	refText := fmt.Sprintf("[%s]", noteID)
	run.Text.Content += refText

	// 创建脚注内容
	return d.createNoteContent(noteID, footnoteText, FootnoteTypeFootnote)
}

// SetFootnoteConfig 设置脚注配置
func (d *Document) SetFootnoteConfig(config *FootnoteConfig) error {
	if config == nil {
		config = DefaultFootnoteConfig()
	}

	// 创建脚注属性
	// 这里需要创建脚注设置的XML结构
	// 简化处理，实际需要在document.xml中添加脚注属性设置

	return nil
}

// ensureFootnoteInitialized 确保脚注/尾注系统已初始化
func (d *Document) ensureFootnoteInitialized(noteType FootnoteType) {
	if noteType == FootnoteTypeFootnote {
		if _, exists := d.parts["word/footnotes.xml"]; !exists {
			d.initializeFootnotes()
		}
	} else {
		if _, exists := d.parts["word/endnotes.xml"]; !exists {
			d.initializeEndnotes()
		}
	}
}

// initializeFootnotes 初始化脚注系统
func (d *Document) initializeFootnotes() {
	footnotes := &Footnotes{
		Xmlns:     "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		Footnotes: []*Footnote{},
	}

	// 添加默认的分隔符脚注
	separatorFootnote := &Footnote{
		Type: "separator",
		ID:   "-1",
		Paragraphs: []*Paragraph{
			{
				Runs: []Run{
					{
						Text: Text{Content: ""},
					},
				},
			},
		},
	}
	footnotes.Footnotes = append(footnotes.Footnotes, separatorFootnote)

	// 序列化脚注
	footnotesXML, err := xml.MarshalIndent(footnotes, "", "  ")
	if err != nil {
		return
	}

	// 添加XML声明
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	d.parts["word/footnotes.xml"] = append(xmlDeclaration, footnotesXML...)

	// 添加内容类型
	d.addContentType("word/footnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")

	// 添加关系
	d.addFootnoteRelationship()
}

// initializeEndnotes 初始化尾注系统
func (d *Document) initializeEndnotes() {
	endnotes := &Endnotes{
		Xmlns:    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		Endnotes: []*Endnote{},
	}

	// 添加默认的分隔符尾注
	separatorEndnote := &Endnote{
		Type: "separator",
		ID:   "-1",
		Paragraphs: []*Paragraph{
			{
				Runs: []Run{
					{
						Text: Text{Content: ""},
					},
				},
			},
		},
	}
	endnotes.Endnotes = append(endnotes.Endnotes, separatorEndnote)

	// 序列化尾注
	endnotesXML, err := xml.MarshalIndent(endnotes, "", "  ")
	if err != nil {
		return
	}

	// 添加XML声明
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	d.parts["word/endnotes.xml"] = append(xmlDeclaration, endnotesXML...)

	// 添加内容类型
	d.addContentType("word/endnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml")

	// 添加关系
	d.addEndnoteRelationship()
}

// createNoteContent 创建脚注/尾注内容
func (d *Document) createNoteContent(noteID string, noteText string, noteType FootnoteType) error {
	manager := getFootnoteManager()

	// 创建脚注/尾注段落
	noteParagraph := &Paragraph{
		Runs: []Run{
			{
				Text: Text{Content: noteText},
			},
		},
	}

	if noteType == FootnoteTypeFootnote {
		// 创建脚注
		footnote := &Footnote{
			ID:         noteID,
			Paragraphs: []*Paragraph{noteParagraph},
		}
		manager.footnotes[noteID] = footnote

		// 更新脚注文件
		d.updateFootnotesFile()
	} else {
		// 创建尾注
		endnote := &Endnote{
			ID:         noteID,
			Paragraphs: []*Paragraph{noteParagraph},
		}
		manager.endnotes[noteID] = endnote

		// 更新尾注文件
		d.updateEndnotesFile()
	}

	return nil
}

// updateFootnotesFile 更新脚注文件
func (d *Document) updateFootnotesFile() {
	manager := getFootnoteManager()

	footnotes := &Footnotes{
		Xmlns:     "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		Footnotes: []*Footnote{},
	}

	// 添加默认分隔符
	separatorFootnote := &Footnote{
		Type: "separator",
		ID:   "-1",
		Paragraphs: []*Paragraph{
			{
				Runs: []Run{
					{
						Text: Text{Content: ""},
					},
				},
			},
		},
	}
	footnotes.Footnotes = append(footnotes.Footnotes, separatorFootnote)

	// 添加所有脚注
	for _, footnote := range manager.footnotes {
		footnotes.Footnotes = append(footnotes.Footnotes, footnote)
	}

	// 序列化
	footnotesXML, err := xml.MarshalIndent(footnotes, "", "  ")
	if err != nil {
		return
	}

	// 添加XML声明
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	d.parts["word/footnotes.xml"] = append(xmlDeclaration, footnotesXML...)
}

// updateEndnotesFile 更新尾注文件
func (d *Document) updateEndnotesFile() {
	manager := getFootnoteManager()

	endnotes := &Endnotes{
		Xmlns:    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		Endnotes: []*Endnote{},
	}

	// 添加默认分隔符
	separatorEndnote := &Endnote{
		Type: "separator",
		ID:   "-1",
		Paragraphs: []*Paragraph{
			{
				Runs: []Run{
					{
						Text: Text{Content: ""},
					},
				},
			},
		},
	}
	endnotes.Endnotes = append(endnotes.Endnotes, separatorEndnote)

	// 添加所有尾注
	for _, endnote := range manager.endnotes {
		endnotes.Endnotes = append(endnotes.Endnotes, endnote)
	}

	// 序列化
	endnotesXML, err := xml.MarshalIndent(endnotes, "", "  ")
	if err != nil {
		return
	}

	// 添加XML声明
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	d.parts["word/endnotes.xml"] = append(xmlDeclaration, endnotesXML...)
}

// addFootnoteRelationship 添加脚注关系
func (d *Document) addFootnoteRelationship() {
	relationshipID := fmt.Sprintf("rId%d", len(d.relationships.Relationships)+1)

	relationship := Relationship{
		ID:     relationshipID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
		Target: "footnotes.xml",
	}
	d.relationships.Relationships = append(d.relationships.Relationships, relationship)
}

// addEndnoteRelationship 添加尾注关系
func (d *Document) addEndnoteRelationship() {
	relationshipID := fmt.Sprintf("rId%d", len(d.relationships.Relationships)+1)

	relationship := Relationship{
		ID:     relationshipID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
		Target: "endnotes.xml",
	}
	d.relationships.Relationships = append(d.relationships.Relationships, relationship)
}

// GetFootnoteCount 获取脚注数量
func (d *Document) GetFootnoteCount() int {
	manager := getFootnoteManager()
	return len(manager.footnotes)
}

// GetEndnoteCount 获取尾注数量
func (d *Document) GetEndnoteCount() int {
	manager := getFootnoteManager()
	return len(manager.endnotes)
}

// RemoveFootnote 删除指定脚注
func (d *Document) RemoveFootnote(footnoteID string) error {
	manager := getFootnoteManager()

	if _, exists := manager.footnotes[footnoteID]; !exists {
		return fmt.Errorf("脚注 %s 不存在", footnoteID)
	}

	delete(manager.footnotes, footnoteID)
	d.updateFootnotesFile()

	return nil
}

// RemoveEndnote 删除指定尾注
func (d *Document) RemoveEndnote(endnoteID string) error {
	manager := getFootnoteManager()

	if _, exists := manager.endnotes[endnoteID]; !exists {
		return fmt.Errorf("尾注 %s 不存在", endnoteID)
	}

	delete(manager.endnotes, endnoteID)
	d.updateEndnotesFile()

	return nil
}

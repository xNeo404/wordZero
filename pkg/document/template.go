// Package document 模板功能实现
package document

import (
	"fmt"
	"reflect"
	"regexp"
	"strconv"
	"strings"
	"sync"
)

// 模板相关错误
var (
	// ErrTemplateNotFound 模板未找到
	ErrTemplateNotFound = NewDocumentError("template_not_found", fmt.Errorf("template not found"), "")

	// ErrTemplateSyntaxError 模板语法错误
	ErrTemplateSyntaxError = NewDocumentError("template_syntax_error", fmt.Errorf("template syntax error"), "")

	// ErrTemplateRenderError 模板渲染错误
	ErrTemplateRenderError = NewDocumentError("template_render_error", fmt.Errorf("template render error"), "")

	// ErrInvalidTemplateData 无效模板数据
	ErrInvalidTemplateData = NewDocumentError("invalid_template_data", fmt.Errorf("invalid template data"), "")

	// ErrBlockNotFound 块未找到
	ErrBlockNotFound = NewDocumentError("block_not_found", fmt.Errorf("block not found"), "")

	// ErrInvalidBlockDefinition 无效块定义
	ErrInvalidBlockDefinition = NewDocumentError("invalid_block_definition", fmt.Errorf("invalid block definition"), "")
)

// TemplateEngine 模板引擎
type TemplateEngine struct {
	cache    map[string]*Template // 模板缓存
	mutex    sync.RWMutex         // 读写锁
	basePath string               // 基础路径
}

// Template 模板结构
type Template struct {
	Name          string                    // 模板名称
	Content       string                    // 模板内容
	BaseDoc       *Document                 // 基础文档
	Variables     map[string]string         // 模板变量
	Blocks        []*TemplateBlock          // 模板块列表
	Parent        *Template                 // 父模板（用于继承）
	DefinedBlocks map[string]*TemplateBlock // 定义的块映射
}

// TemplateBlock 模板块
type TemplateBlock struct {
	Type           string                 // 块类型：variable, if, each, inherit, block
	Name           string                 // 块名称（block类型使用）
	Content        string                 // 块内容
	Condition      string                 // 条件（if块使用）
	Variable       string                 // 变量名（each块使用）
	Children       []*TemplateBlock       // 子块
	Data           map[string]interface{} // 块数据
	DefaultContent string                 // 默认内容（用于可选重写）
	IsOverridden   bool                   // 是否被重写
}

// TemplateData 模板数据
type TemplateData struct {
	Variables  map[string]interface{}   // 变量数据
	Lists      map[string][]interface{} // 列表数据
	Conditions map[string]bool          // 条件数据
}

// NewTemplateEngine 创建新的模板引擎
func NewTemplateEngine() *TemplateEngine {
	return &TemplateEngine{
		cache: make(map[string]*Template),
		mutex: sync.RWMutex{},
	}
}

// SetBasePath 设置模板基础路径
func (te *TemplateEngine) SetBasePath(path string) {
	te.mutex.Lock()
	defer te.mutex.Unlock()
	te.basePath = path
}

// LoadTemplate 从字符串加载模板
func (te *TemplateEngine) LoadTemplate(name, content string) (*Template, error) {
	te.mutex.Lock()
	defer te.mutex.Unlock()

	template := &Template{
		Name:          name,
		Content:       content,
		Variables:     make(map[string]string),
		Blocks:        make([]*TemplateBlock, 0),
		DefinedBlocks: make(map[string]*TemplateBlock),
	}

	// 解析模板内容
	if err := te.parseTemplate(template); err != nil {
		return nil, WrapErrorWithContext("load_template", err, name)
	}

	// 缓存模板
	te.cache[name] = template

	return template, nil
}

// LoadTemplateFromDocument 从现有文档创建模板
func (te *TemplateEngine) LoadTemplateFromDocument(name string, doc *Document) (*Template, error) {
	te.mutex.Lock()
	defer te.mutex.Unlock()

	// 将文档内容转换为模板字符串
	content, err := te.documentToTemplateString(doc)
	if err != nil {
		return nil, WrapErrorWithContext("load_template_from_document", err, name)
	}

	template := &Template{
		Name:          name,
		Content:       content,
		BaseDoc:       doc,
		Variables:     make(map[string]string),
		Blocks:        make([]*TemplateBlock, 0),
		DefinedBlocks: make(map[string]*TemplateBlock),
	}

	// 解析模板内容
	if err := te.parseTemplate(template); err != nil {
		return nil, WrapErrorWithContext("load_template_from_document", err, name)
	}

	// 缓存模板
	te.cache[name] = template

	return template, nil
}

// GetTemplate 获取缓存的模板
func (te *TemplateEngine) GetTemplate(name string) (*Template, error) {
	te.mutex.RLock()
	defer te.mutex.RUnlock()

	if template, exists := te.cache[name]; exists {
		return template, nil
	}

	return nil, WrapErrorWithContext("get_template", ErrTemplateNotFound.Cause, name)
}

// getTemplateInternal 获取缓存的模板（内部方法，不加锁）
func (te *TemplateEngine) getTemplateInternal(name string) (*Template, error) {
	if template, exists := te.cache[name]; exists {
		return template, nil
	}

	return nil, WrapErrorWithContext("get_template", ErrTemplateNotFound.Cause, name)
}

// ClearCache 清空模板缓存
func (te *TemplateEngine) ClearCache() {
	te.mutex.Lock()
	defer te.mutex.Unlock()
	te.cache = make(map[string]*Template)
}

// RemoveTemplate 移除指定模板
func (te *TemplateEngine) RemoveTemplate(name string) {
	te.mutex.Lock()
	defer te.mutex.Unlock()
	delete(te.cache, name)
}

// parseTemplate 解析模板内容
func (te *TemplateEngine) parseTemplate(template *Template) error {
	content := template.Content

	// 解析变量: {{变量名}}
	varPattern := regexp.MustCompile(`\{\{(\w+)\}\}`)
	varMatches := varPattern.FindAllStringSubmatch(content, -1)
	for _, match := range varMatches {
		if len(match) >= 2 {
			varName := match[1]
			template.Variables[varName] = ""
		}
	}

	// 解析块定义: {{#block "blockName"}}...{{/block}}
	blockPattern := regexp.MustCompile(`(?s)\{\{#block\s+"([^"]+)"\}\}(.*?)\{\{/block\}\}`)
	blockMatches := blockPattern.FindAllStringSubmatch(content, -1)
	for _, match := range blockMatches {
		if len(match) >= 3 {
			blockName := match[1]
			blockContent := match[2]

			block := &TemplateBlock{
				Type:           "block",
				Name:           blockName,
				Content:        blockContent,
				DefaultContent: blockContent,
				Children:       make([]*TemplateBlock, 0),
			}

			template.Blocks = append(template.Blocks, block)
			template.DefinedBlocks[blockName] = block
		}
	}

	// 解析条件语句: {{#if 条件}}...{{/if}} (修复：添加 (?s) 标志以匹配换行符)
	ifPattern := regexp.MustCompile(`(?s)\{\{#if\s+(\w+)\}\}(.*?)\{\{/if\}\}`)
	ifMatches := ifPattern.FindAllStringSubmatch(content, -1)
	for _, match := range ifMatches {
		if len(match) >= 3 {
			condition := match[1]
			blockContent := match[2]

			block := &TemplateBlock{
				Type:      "if",
				Condition: condition,
				Content:   blockContent,
				Children:  make([]*TemplateBlock, 0),
			}

			template.Blocks = append(template.Blocks, block)
		}
	}

	// 解析循环语句: {{#each 列表}}...{{/each}} (修复：添加 (?s) 标志以匹配换行符)
	eachPattern := regexp.MustCompile(`(?s)\{\{#each\s+(\w+)\}\}(.*?)\{\{/each\}\}`)
	eachMatches := eachPattern.FindAllStringSubmatch(content, -1)
	for _, match := range eachMatches {
		if len(match) >= 3 {
			listVar := match[1]
			blockContent := match[2]

			block := &TemplateBlock{
				Type:     "each",
				Variable: listVar,
				Content:  blockContent,
				Children: make([]*TemplateBlock, 0),
			}

			template.Blocks = append(template.Blocks, block)
		}
	}

	// 解析继承: {{extends "base_template"}}
	extendsPattern := regexp.MustCompile(`\{\{extends\s+"([^"]+)"\}\}`)
	extendsMatches := extendsPattern.FindStringSubmatch(content)
	if len(extendsMatches) >= 2 {
		baseName := extendsMatches[1]
		baseTemplate, err := te.getTemplateInternal(baseName)
		if err == nil {
			template.Parent = baseTemplate
			// 处理块重写
			te.processBlockOverrides(template, baseTemplate)
		}
	}

	return nil
}

// processBlockOverrides 处理块重写
func (te *TemplateEngine) processBlockOverrides(childTemplate, parentTemplate *Template) {
	// 遍历子模板的块定义，检查是否重写父模板的块
	for blockName, childBlock := range childTemplate.DefinedBlocks {
		if parentBlock, exists := parentTemplate.DefinedBlocks[blockName]; exists {
			// 标记父模板块被重写
			parentBlock.IsOverridden = true
			parentBlock.Content = childBlock.Content
		}
	}

	// 递归处理父模板的父模板
	if parentTemplate.Parent != nil {
		te.processBlockOverrides(childTemplate, parentTemplate.Parent)
	}
}

// RenderToDocument 渲染模板到新文档
func (te *TemplateEngine) RenderToDocument(templateName string, data *TemplateData) (*Document, error) {
	template, err := te.GetTemplate(templateName)
	if err != nil {
		return nil, WrapErrorWithContext("render_to_document", err, templateName)
	}

	// 创建新文档
	var doc *Document
	if template.BaseDoc != nil {
		// 基于基础文档创建
		doc = te.cloneDocument(template.BaseDoc)
	} else {
		// 创建新文档
		doc = New()
	}

	// 渲染模板内容
	renderedContent, err := te.renderTemplate(template, data)
	if err != nil {
		return nil, WrapErrorWithContext("render_to_document", err, templateName)
	}

	// 将渲染内容应用到文档
	if err := te.applyRenderedContentToDocument(doc, renderedContent); err != nil {
		return nil, WrapErrorWithContext("render_to_document", err, templateName)
	}

	return doc, nil
}

// renderTemplate 渲染模板
func (te *TemplateEngine) renderTemplate(template *Template, data *TemplateData) (string, error) {
	var content string

	// 处理继承：如果有父模板，使用父模板作为基础
	if template.Parent != nil {
		// 渲染父模板作为基础内容
		parentContent, err := te.renderTemplate(template.Parent, data)
		if err != nil {
			return "", err
		}
		content = parentContent

		// 应用子模板的块重写到父模板内容中
		content = te.applyBlockOverrides(content, template)
	} else {
		// 没有父模板，直接使用当前模板内容
		content = template.Content
	}

	// 渲染块定义
	content = te.renderBlocks(content, template, data)

	// 渲染变量
	content = te.renderVariables(content, data.Variables)

	// 渲染循环语句（先处理循环，循环内部会处理条件语句）
	content = te.renderLoops(content, data.Lists)

	// 渲染条件语句（处理非循环内的条件语句）
	content = te.renderConditionals(content, data.Conditions)

	return content, nil
}

// applyBlockOverrides 将子模板的块重写应用到父模板内容中
func (te *TemplateEngine) applyBlockOverrides(content string, template *Template) string {
	// 将子模板的块内容替换父模板中对应的块占位符
	blockPattern := regexp.MustCompile(`(?s)\{\{#block\s+"([^"]+)"\}\}.*?\{\{/block\}\}`)

	return blockPattern.ReplaceAllStringFunc(content, func(match string) string {
		matches := blockPattern.FindStringSubmatch(match)
		if len(matches) >= 2 {
			blockName := matches[1]
			// 如果子模板中定义了这个块，使用子模板的内容
			if childBlock, exists := template.DefinedBlocks[blockName]; exists {
				return childBlock.Content
			}
		}
		return match // 保持原样
	})
}

// renderBlocks 渲染块定义
func (te *TemplateEngine) renderBlocks(content string, template *Template, data *TemplateData) string {
	blockPattern := regexp.MustCompile(`(?s)\{\{#block\s+"([^"]+)"\}\}(.*?)\{\{/block\}\}`)

	return blockPattern.ReplaceAllStringFunc(content, func(match string) string {
		matches := blockPattern.FindStringSubmatch(match)
		if len(matches) >= 3 {
			blockName := matches[1]
			blockContent := matches[2]

			// 检查是否有定义的块
			if block, exists := template.DefinedBlocks[blockName]; exists {
				// 如果块被重写，使用重写的内容，否则使用默认内容
				if block.IsOverridden {
					return block.Content
				}
				return block.DefaultContent
			}

			// 如果没有定义块，使用原始内容
			return blockContent
		}
		return match
	})
}

// renderVariables 渲染变量
func (te *TemplateEngine) renderVariables(content string, variables map[string]interface{}) string {
	varPattern := regexp.MustCompile(`\{\{(\w+)\}\}`)

	return varPattern.ReplaceAllStringFunc(content, func(match string) string {
		varName := varPattern.FindStringSubmatch(match)[1]
		if value, exists := variables[varName]; exists {
			return te.interfaceToString(value)
		}
		return match // 保持原样
	})
}

// renderConditionals 渲染条件语句
func (te *TemplateEngine) renderConditionals(content string, conditions map[string]bool) string {
	ifPattern := regexp.MustCompile(`(?s)\{\{#if\s+(\w+)\}\}(.*?)\{\{/if\}\}`)

	return ifPattern.ReplaceAllStringFunc(content, func(match string) string {
		matches := ifPattern.FindStringSubmatch(match)
		if len(matches) >= 3 {
			condition := matches[1]
			blockContent := matches[2]

			if condValue, exists := conditions[condition]; exists && condValue {
				return blockContent
			}
		}
		return "" // 条件不满足，返回空字符串
	})
}

// renderLoops 渲染循环语句
func (te *TemplateEngine) renderLoops(content string, lists map[string][]interface{}) string {
	eachPattern := regexp.MustCompile(`(?s)\{\{#each\s+(\w+)\}\}(.*?)\{\{/each\}\}`)

	return eachPattern.ReplaceAllStringFunc(content, func(match string) string {
		matches := eachPattern.FindStringSubmatch(match)
		if len(matches) >= 3 {
			listVar := matches[1]
			blockContent := matches[2]

			if listData, exists := lists[listVar]; exists {
				var result strings.Builder
				for i, item := range listData {
					// 创建循环上下文变量
					loopContent := strings.ReplaceAll(blockContent, "{{this}}", te.interfaceToString(item))
					loopContent = strings.ReplaceAll(loopContent, "{{@index}}", strconv.Itoa(i))
					loopContent = strings.ReplaceAll(loopContent, "{{@first}}", strconv.FormatBool(i == 0))
					loopContent = strings.ReplaceAll(loopContent, "{{@last}}", strconv.FormatBool(i == len(listData)-1))

					// 如果item是map，处理属性访问
					if itemMap, ok := item.(map[string]interface{}); ok {
						for key, value := range itemMap {
							placeholder := fmt.Sprintf("{{%s}}", key)
							loopContent = strings.ReplaceAll(loopContent, placeholder, te.interfaceToString(value))
						}

						// 处理循环内部的条件语句
						loopContent = te.renderLoopConditionals(loopContent, itemMap)
					}

					result.WriteString(loopContent)
				}
				return result.String()
			}
		}
		return match // 保持原样
	})
}

// renderLoopConditionals 渲染循环内部的条件语句
func (te *TemplateEngine) renderLoopConditionals(content string, itemData map[string]interface{}) string {
	ifPattern := regexp.MustCompile(`(?s)\{\{#if\s+(\w+)\}\}(.*?)\{\{/if\}\}`)

	return ifPattern.ReplaceAllStringFunc(content, func(match string) string {
		matches := ifPattern.FindStringSubmatch(match)
		if len(matches) >= 3 {
			condition := matches[1]
			blockContent := matches[2]

			// 检查条件是否在当前循环项的数据中
			if condValue, exists := itemData[condition]; exists {
				// 转换为布尔值
				switch v := condValue.(type) {
				case bool:
					if v {
						return blockContent
					}
				case string:
					if v == "true" || v == "1" || v == "yes" || v != "" {
						return blockContent
					}
				case int:
					if v != 0 {
						return blockContent
					}
				case int64:
					if v != 0 {
						return blockContent
					}
				case float64:
					if v != 0.0 {
						return blockContent
					}
				default:
					// 对于其他类型，如果不为nil就认为是true
					if v != nil {
						return blockContent
					}
				}
			}
		}
		return "" // 条件不满足，返回空字符串
	})
}

// interfaceToString 将interface{}转换为字符串
func (te *TemplateEngine) interfaceToString(value interface{}) string {
	if value == nil {
		return ""
	}

	switch v := value.(type) {
	case string:
		return v
	case int:
		return strconv.Itoa(v)
	case int64:
		return strconv.FormatInt(v, 10)
	case float64:
		return strconv.FormatFloat(v, 'f', -1, 64)
	case bool:
		return strconv.FormatBool(v)
	default:
		return fmt.Sprintf("%v", v)
	}
}

// ValidateTemplate 验证模板语法
func (te *TemplateEngine) ValidateTemplate(template *Template) error {
	content := template.Content

	// 检查括号配对
	if err := te.validateBrackets(content); err != nil {
		return WrapErrorWithContext("validate_template", err, template.Name)
	}

	// 检查块语句配对
	if err := te.validateBlockStatements(content); err != nil {
		return WrapErrorWithContext("validate_template", err, template.Name)
	}

	// 检查if语句配对
	if err := te.validateIfStatements(content); err != nil {
		return WrapErrorWithContext("validate_template", err, template.Name)
	}

	// 检查each语句配对
	if err := te.validateEachStatements(content); err != nil {
		return WrapErrorWithContext("validate_template", err, template.Name)
	}

	return nil
}

// validateBrackets 验证括号配对
func (te *TemplateEngine) validateBrackets(content string) error {
	openCount := strings.Count(content, "{{")
	closeCount := strings.Count(content, "}}")

	if openCount != closeCount {
		return NewValidationError("brackets", content, "mismatched template brackets")
	}

	return nil
}

// validateBlockStatements 验证块语句配对
func (te *TemplateEngine) validateBlockStatements(content string) error {
	blockCount := len(regexp.MustCompile(`\{\{#block\s+"[^"]+"\}\}`).FindAllString(content, -1))
	endblockCount := len(regexp.MustCompile(`\{\{/block\}\}`).FindAllString(content, -1))

	if blockCount != endblockCount {
		return NewValidationError("block_statements", content, "mismatched block/endblock statements")
	}

	return nil
}

// validateIfStatements 验证if语句配对
func (te *TemplateEngine) validateIfStatements(content string) error {
	ifCount := len(regexp.MustCompile(`\{\{#if\s+\w+\}\}`).FindAllString(content, -1))
	endifCount := len(regexp.MustCompile(`\{\{/if\}\}`).FindAllString(content, -1))

	if ifCount != endifCount {
		return NewValidationError("if_statements", content, "mismatched if/endif statements")
	}

	return nil
}

// validateEachStatements 验证each语句配对
func (te *TemplateEngine) validateEachStatements(content string) error {
	eachCount := len(regexp.MustCompile(`\{\{#each\s+\w+\}\}`).FindAllString(content, -1))
	endeachCount := len(regexp.MustCompile(`\{\{/each\}\}`).FindAllString(content, -1))

	if eachCount != endeachCount {
		return NewValidationError("each_statements", content, "mismatched each/endeach statements")
	}

	return nil
}

// documentToTemplateString 将文档转换为模板字符串
func (te *TemplateEngine) documentToTemplateString(doc *Document) (string, error) {
	// 这里不再转换为纯字符串，而是保持原始文档结构
	// 实际上我们应该直接在原文档上进行变量替换
	return "", nil // 将在新的方法中处理
}

// cloneDocument 深度复制文档所有元素和属性
func (te *TemplateEngine) cloneDocument(source *Document) *Document {
	// 创建新文档
	doc := New()

	// 复制样式管理器
	if source.styleManager != nil {
		doc.styleManager = source.styleManager
	}

	// 复制所有文档部件，保持完整的文档结构
	if source.parts != nil {
		doc.parts = make(map[string][]byte)
		for name, data := range source.parts {
			// 深度复制每个部件的数据
			copiedData := make([]byte, len(data))
			copy(copiedData, data)
			doc.parts[name] = copiedData
		}
	}

	// 复制文档关系
	if source.relationships != nil {
		doc.relationships = source.relationships
	}

	// 复制文档级关系
	if source.documentRelationships != nil {
		doc.documentRelationships = source.documentRelationships
	}

	// 复制内容类型
	if source.contentTypes != nil {
		doc.contentTypes = source.contentTypes
	}

	// 深度复制文档体元素
	for _, element := range source.Body.Elements {
		switch elem := element.(type) {
		case *Paragraph:
			// 深度复制段落
			newPara := te.cloneParagraph(elem)
			doc.Body.Elements = append(doc.Body.Elements, newPara)

		case *Table:
			// 深度复制表格
			newTable := te.cloneTable(elem)
			doc.Body.Elements = append(doc.Body.Elements, newTable)
		}
	}

	return doc
}

// cloneParagraph 深度复制段落
func (te *TemplateEngine) cloneParagraph(source *Paragraph) *Paragraph {
	newPara := &Paragraph{
		Properties: te.cloneParagraphProperties(source.Properties),
		Runs:       make([]Run, len(source.Runs)),
	}

	for i, run := range source.Runs {
		newPara.Runs[i] = te.cloneRun(&run)
	}

	return newPara
}

// cloneParagraphProperties 深度复制段落属性
func (te *TemplateEngine) cloneParagraphProperties(source *ParagraphProperties) *ParagraphProperties {
	if source == nil {
		return nil
	}

	props := &ParagraphProperties{}

	// 复制段落样式
	if source.ParagraphStyle != nil {
		props.ParagraphStyle = &ParagraphStyle{
			Val: source.ParagraphStyle.Val,
		}
	}

	// 复制编号属性
	if source.NumberingProperties != nil {
		props.NumberingProperties = &NumberingProperties{}
		if source.NumberingProperties.ILevel != nil {
			props.NumberingProperties.ILevel = &ILevel{Val: source.NumberingProperties.ILevel.Val}
		}
		if source.NumberingProperties.NumID != nil {
			props.NumberingProperties.NumID = &NumID{Val: source.NumberingProperties.NumID.Val}
		}
	}

	// 复制间距
	if source.Spacing != nil {
		props.Spacing = &Spacing{
			Before: source.Spacing.Before,
			After:  source.Spacing.After,
			Line:   source.Spacing.Line,
		}
	}

	// 复制对齐方式
	if source.Justification != nil {
		props.Justification = &Justification{
			Val: source.Justification.Val,
		}
	}

	// 复制缩进
	if source.Indentation != nil {
		props.Indentation = &Indentation{
			FirstLine: source.Indentation.FirstLine,
			Left:      source.Indentation.Left,
			Right:     source.Indentation.Right,
		}
	}

	// 复制制表符
	if source.Tabs != nil {
		props.Tabs = &Tabs{
			Tabs: make([]TabDef, len(source.Tabs.Tabs)),
		}
		for i, tab := range source.Tabs.Tabs {
			props.Tabs.Tabs[i] = TabDef{
				Val:    tab.Val,
				Leader: tab.Leader,
				Pos:    tab.Pos,
			}
		}
	}

	return props
}

// cloneRun 深度复制文本运行
func (te *TemplateEngine) cloneRun(source *Run) Run {
	newRun := Run{
		Properties: te.cloneRunProperties(source.Properties),
		Text:       Text{Content: source.Text.Content, Space: source.Text.Space},
	}

	// 复制图像（如果有）
	if source.Drawing != nil {
		// 暂时保持简单复制，图像的深度复制比较复杂
		newRun.Drawing = source.Drawing
	}

	// 复制域字符（如果有）
	if source.FieldChar != nil {
		newRun.FieldChar = source.FieldChar
	}

	// 复制指令文本（如果有）
	if source.InstrText != nil {
		newRun.InstrText = source.InstrText
	}

	return newRun
}

// cloneRunProperties 深度复制文本运行属性
func (te *TemplateEngine) cloneRunProperties(source *RunProperties) *RunProperties {
	if source == nil {
		return nil
	}

	props := &RunProperties{}

	// 复制粗体
	if source.Bold != nil {
		props.Bold = &Bold{}
	}

	// 复制斜体
	if source.Italic != nil {
		props.Italic = &Italic{}
	}

	// 复制字体大小
	if source.FontSize != nil {
		props.FontSize = &FontSize{
			Val: source.FontSize.Val,
		}
	}

	// 复制颜色
	if source.Color != nil {
		props.Color = &Color{
			Val: source.Color.Val,
		}
	}

	// 完整复制字体族属性，包括所有字体设置
	if source.FontFamily != nil {
		props.FontFamily = &FontFamily{
			ASCII:    source.FontFamily.ASCII,
			HAnsi:    source.FontFamily.HAnsi,
			EastAsia: source.FontFamily.EastAsia,
			CS:       source.FontFamily.CS,
			Hint:     source.FontFamily.Hint,
		}
	}

	return props
}

// cloneTable 深度复制表格
func (te *TemplateEngine) cloneTable(source *Table) *Table {
	newTable := &Table{
		Properties: te.cloneTableProperties(source.Properties),
		Grid:       te.cloneTableGrid(source.Grid),
		Rows:       make([]TableRow, len(source.Rows)),
	}

	for i, row := range source.Rows {
		newTable.Rows[i] = *te.cloneTableRow(&row)
	}

	return newTable
}

// cloneTableProperties 深度复制表格属性
func (te *TemplateEngine) cloneTableProperties(source *TableProperties) *TableProperties {
	if source == nil {
		return nil
	}

	props := &TableProperties{}

	// 复制表格宽度
	if source.TableW != nil {
		props.TableW = &TableWidth{
			W:    source.TableW.W,
			Type: source.TableW.Type,
		}
	}

	// 复制表格对齐
	if source.TableJc != nil {
		props.TableJc = &TableJc{
			Val: source.TableJc.Val,
		}
	}

	// 复制表格外观
	if source.TableLook != nil {
		props.TableLook = &TableLook{
			Val:      source.TableLook.Val,
			FirstRow: source.TableLook.FirstRow,
			LastRow:  source.TableLook.LastRow,
			FirstCol: source.TableLook.FirstCol,
			LastCol:  source.TableLook.LastCol,
			NoHBand:  source.TableLook.NoHBand,
			NoVBand:  source.TableLook.NoVBand,
		}
	}

	// 复制表格样式
	if source.TableStyle != nil {
		props.TableStyle = &TableStyle{
			Val: source.TableStyle.Val,
		}
	}

	// 复制表格边框
	if source.TableBorders != nil {
		props.TableBorders = te.cloneTableBorders(source.TableBorders)
	}

	// 复制表格底纹
	if source.Shd != nil {
		props.Shd = &TableShading{
			Val:       source.Shd.Val,
			Color:     source.Shd.Color,
			Fill:      source.Shd.Fill,
			ThemeFill: source.Shd.ThemeFill,
		}
	}

	// 复制表格单元格边距
	if source.TableCellMar != nil {
		props.TableCellMar = te.cloneTableCellMargins(source.TableCellMar)
	}

	// 复制表格布局
	if source.TableLayout != nil {
		props.TableLayout = &TableLayoutType{
			Type: source.TableLayout.Type,
		}
	}

	// 复制表格缩进
	if source.TableInd != nil {
		props.TableInd = &TableIndentation{
			W:    source.TableInd.W,
			Type: source.TableInd.Type,
		}
	}

	return props
}

// cloneTableBorders 深度复制表格边框
func (te *TemplateEngine) cloneTableBorders(source *TableBorders) *TableBorders {
	if source == nil {
		return nil
	}

	borders := &TableBorders{}

	if source.Top != nil {
		borders.Top = &TableBorder{
			Val:        source.Top.Val,
			Sz:         source.Top.Sz,
			Space:      source.Top.Space,
			Color:      source.Top.Color,
			ThemeColor: source.Top.ThemeColor,
		}
	}

	if source.Left != nil {
		borders.Left = &TableBorder{
			Val:        source.Left.Val,
			Sz:         source.Left.Sz,
			Space:      source.Left.Space,
			Color:      source.Left.Color,
			ThemeColor: source.Left.ThemeColor,
		}
	}

	if source.Bottom != nil {
		borders.Bottom = &TableBorder{
			Val:        source.Bottom.Val,
			Sz:         source.Bottom.Sz,
			Space:      source.Bottom.Space,
			Color:      source.Bottom.Color,
			ThemeColor: source.Bottom.ThemeColor,
		}
	}

	if source.Right != nil {
		borders.Right = &TableBorder{
			Val:        source.Right.Val,
			Sz:         source.Right.Sz,
			Space:      source.Right.Space,
			Color:      source.Right.Color,
			ThemeColor: source.Right.ThemeColor,
		}
	}

	if source.InsideH != nil {
		borders.InsideH = &TableBorder{
			Val:        source.InsideH.Val,
			Sz:         source.InsideH.Sz,
			Space:      source.InsideH.Space,
			Color:      source.InsideH.Color,
			ThemeColor: source.InsideH.ThemeColor,
		}
	}

	if source.InsideV != nil {
		borders.InsideV = &TableBorder{
			Val:        source.InsideV.Val,
			Sz:         source.InsideV.Sz,
			Space:      source.InsideV.Space,
			Color:      source.InsideV.Color,
			ThemeColor: source.InsideV.ThemeColor,
		}
	}

	return borders
}

// cloneTableCellMargins 深度复制表格单元格边距
func (te *TemplateEngine) cloneTableCellMargins(source *TableCellMargins) *TableCellMargins {
	if source == nil {
		return nil
	}

	margins := &TableCellMargins{}

	if source.Top != nil {
		margins.Top = &TableCellSpace{
			W:    source.Top.W,
			Type: source.Top.Type,
		}
	}

	if source.Left != nil {
		margins.Left = &TableCellSpace{
			W:    source.Left.W,
			Type: source.Left.Type,
		}
	}

	if source.Bottom != nil {
		margins.Bottom = &TableCellSpace{
			W:    source.Bottom.W,
			Type: source.Bottom.Type,
		}
	}

	if source.Right != nil {
		margins.Right = &TableCellSpace{
			W:    source.Right.W,
			Type: source.Right.Type,
		}
	}

	return margins
}

// cloneTableGrid 深度复制表格网格
func (te *TemplateEngine) cloneTableGrid(source *TableGrid) *TableGrid {
	if source == nil {
		return nil
	}

	grid := &TableGrid{
		Cols: make([]TableGridCol, len(source.Cols)),
	}

	for i, col := range source.Cols {
		grid.Cols[i] = TableGridCol{
			W: col.W,
		}
	}

	return grid
}

// cloneTableCellMarginsCell 深度复制表格单元格边距（单元格级别）
func (te *TemplateEngine) cloneTableCellMarginsCell(source *TableCellMarginsCell) *TableCellMarginsCell {
	if source == nil {
		return nil
	}

	margins := &TableCellMarginsCell{}

	if source.Top != nil {
		margins.Top = &TableCellSpaceCell{
			W:    source.Top.W,
			Type: source.Top.Type,
		}
	}

	if source.Left != nil {
		margins.Left = &TableCellSpaceCell{
			W:    source.Left.W,
			Type: source.Left.Type,
		}
	}

	if source.Bottom != nil {
		margins.Bottom = &TableCellSpaceCell{
			W:    source.Bottom.W,
			Type: source.Bottom.Type,
		}
	}

	if source.Right != nil {
		margins.Right = &TableCellSpaceCell{
			W:    source.Right.W,
			Type: source.Right.Type,
		}
	}

	return margins
}

// cloneTableCellBorders 深度复制表格单元格边框
func (te *TemplateEngine) cloneTableCellBorders(source *TableCellBorders) *TableCellBorders {
	if source == nil {
		return nil
	}

	borders := &TableCellBorders{}

	if source.Top != nil {
		borders.Top = &TableCellBorder{
			Val:        source.Top.Val,
			Sz:         source.Top.Sz,
			Space:      source.Top.Space,
			Color:      source.Top.Color,
			ThemeColor: source.Top.ThemeColor,
		}
	}

	if source.Left != nil {
		borders.Left = &TableCellBorder{
			Val:        source.Left.Val,
			Sz:         source.Left.Sz,
			Space:      source.Left.Space,
			Color:      source.Left.Color,
			ThemeColor: source.Left.ThemeColor,
		}
	}

	if source.Bottom != nil {
		borders.Bottom = &TableCellBorder{
			Val:        source.Bottom.Val,
			Sz:         source.Bottom.Sz,
			Space:      source.Bottom.Space,
			Color:      source.Bottom.Color,
			ThemeColor: source.Bottom.ThemeColor,
		}
	}

	if source.Right != nil {
		borders.Right = &TableCellBorder{
			Val:        source.Right.Val,
			Sz:         source.Right.Sz,
			Space:      source.Right.Space,
			Color:      source.Right.Color,
			ThemeColor: source.Right.ThemeColor,
		}
	}

	if source.InsideH != nil {
		borders.InsideH = &TableCellBorder{
			Val:        source.InsideH.Val,
			Sz:         source.InsideH.Sz,
			Space:      source.InsideH.Space,
			Color:      source.InsideH.Color,
			ThemeColor: source.InsideH.ThemeColor,
		}
	}

	if source.InsideV != nil {
		borders.InsideV = &TableCellBorder{
			Val:        source.InsideV.Val,
			Sz:         source.InsideV.Sz,
			Space:      source.InsideV.Space,
			Color:      source.InsideV.Color,
			ThemeColor: source.InsideV.ThemeColor,
		}
	}

	if source.TL2BR != nil {
		borders.TL2BR = &TableCellBorder{
			Val:        source.TL2BR.Val,
			Sz:         source.TL2BR.Sz,
			Space:      source.TL2BR.Space,
			Color:      source.TL2BR.Color,
			ThemeColor: source.TL2BR.ThemeColor,
		}
	}

	if source.TR2BL != nil {
		borders.TR2BL = &TableCellBorder{
			Val:        source.TR2BL.Val,
			Sz:         source.TR2BL.Sz,
			Space:      source.TR2BL.Space,
			Color:      source.TR2BL.Color,
			ThemeColor: source.TR2BL.ThemeColor,
		}
	}

	return borders
}

// cloneTableRow 深度复制表格行
func (te *TemplateEngine) cloneTableRow(source *TableRow) *TableRow {
	newRow := &TableRow{
		Properties: te.cloneTableRowProperties(source.Properties),
		Cells:      make([]TableCell, len(source.Cells)),
	}

	for i, cell := range source.Cells {
		newRow.Cells[i] = te.cloneTableCell(&cell)
	}

	return newRow
}

// cloneTableRowProperties 深度复制表格行属性
func (te *TemplateEngine) cloneTableRowProperties(source *TableRowProperties) *TableRowProperties {
	if source == nil {
		return nil
	}

	props := &TableRowProperties{}

	// 复制行高
	if source.TableRowH != nil {
		props.TableRowH = &TableRowH{
			Val:   source.TableRowH.Val,
			HRule: source.TableRowH.HRule,
		}
	}

	// 复制禁止跨页分割
	if source.CantSplit != nil {
		props.CantSplit = &CantSplit{
			Val: source.CantSplit.Val,
		}
	}

	// 复制标题行重复
	if source.TblHeader != nil {
		props.TblHeader = &TblHeader{
			Val: source.TblHeader.Val,
		}
	}

	return props
}

// cloneTableCell 深度复制表格单元格
func (te *TemplateEngine) cloneTableCell(source *TableCell) TableCell {
	newCell := TableCell{
		Properties: te.cloneTableCellProperties(source.Properties),
		Paragraphs: make([]Paragraph, len(source.Paragraphs)),
	}

	for i, para := range source.Paragraphs {
		newCell.Paragraphs[i] = *te.cloneParagraph(&para)
	}

	return newCell
}

// cloneTableCellProperties 深度复制表格单元格属性
func (te *TemplateEngine) cloneTableCellProperties(source *TableCellProperties) *TableCellProperties {
	if source == nil {
		return nil
	}

	props := &TableCellProperties{}

	// 复制单元格宽度
	if source.TableCellW != nil {
		props.TableCellW = &TableCellW{
			W:    source.TableCellW.W,
			Type: source.TableCellW.Type,
		}
	}

	// 复制单元格边距
	if source.TcMar != nil {
		props.TcMar = te.cloneTableCellMarginsCell(source.TcMar)
	}

	// 复制单元格边框
	if source.TcBorders != nil {
		props.TcBorders = te.cloneTableCellBorders(source.TcBorders)
	}

	// 复制单元格底纹
	if source.Shd != nil {
		props.Shd = &TableCellShading{
			Val:       source.Shd.Val,
			Color:     source.Shd.Color,
			Fill:      source.Shd.Fill,
			ThemeFill: source.Shd.ThemeFill,
		}
	}

	// 复制单元格垂直对齐
	if source.VAlign != nil {
		props.VAlign = &VAlign{
			Val: source.VAlign.Val,
		}
	}

	// 复制网格跨度
	if source.GridSpan != nil {
		props.GridSpan = &GridSpan{
			Val: source.GridSpan.Val,
		}
	}

	// 复制垂直合并
	if source.VMerge != nil {
		props.VMerge = &VMerge{
			Val: source.VMerge.Val,
		}
	}

	// 复制文字方向
	if source.TextDirection != nil {
		props.TextDirection = &TextDirection{
			Val: source.TextDirection.Val,
		}
	}

	// 复制禁止换行
	if source.NoWrap != nil {
		props.NoWrap = &NoWrap{
			Val: source.NoWrap.Val,
		}
	}

	// 复制隐藏标记
	if source.HideMark != nil {
		props.HideMark = &HideMark{
			Val: source.HideMark.Val,
		}
	}

	return props
}

// applyRenderedContentToDocument 将渲染内容应用到文档
func (te *TemplateEngine) applyRenderedContentToDocument(doc *Document, content string) error {
	// 这个方法将被新的结构化处理方法替代
	return nil
}

// RenderTemplateToDocument 渲染模板到新文档（新的主要方法）
func (te *TemplateEngine) RenderTemplateToDocument(templateName string, data *TemplateData) (*Document, error) {
	template, err := te.GetTemplate(templateName)
	if err != nil {
		return nil, WrapErrorWithContext("render_template_to_document", err, templateName)
	}

	// 如果有基础文档，克隆它并在其上进行变量替换
	if template.BaseDoc != nil {
		doc := te.cloneDocument(template.BaseDoc)

		// 在文档结构中直接进行变量替换
		err := te.replaceVariablesInDocument(doc, data)
		if err != nil {
			return nil, WrapErrorWithContext("render_template_to_document", err, templateName)
		}

		return doc, nil
	}

	// 如果没有基础文档，使用原有的方式
	return te.RenderToDocument(templateName, data)
}

// replaceVariablesInDocument 在文档结构中直接替换变量
func (te *TemplateEngine) replaceVariablesInDocument(doc *Document, data *TemplateData) error {
	for _, element := range doc.Body.Elements {
		switch elem := element.(type) {
		case *Paragraph:
			// 处理段落中的变量替换
			err := te.replaceVariablesInParagraph(elem, data)
			if err != nil {
				return err
			}

		case *Table:
			// 处理表格中的变量替换和模板语法
			err := te.replaceVariablesInTable(elem, data)
			if err != nil {
				return err
			}
		}
	}

	return nil
}

// replaceVariablesInParagraph 在段落中替换变量（改进版本，更好地保持样式）
func (te *TemplateEngine) replaceVariablesInParagraph(para *Paragraph, data *TemplateData) error {
	// 首先识别所有变量占位符的位置
	fullText := ""
	runInfos := make([]struct {
		startIndex int
		endIndex   int
		run        *Run
	}, 0)

	currentIndex := 0
	for i := range para.Runs {
		runText := para.Runs[i].Text.Content
		if runText != "" {
			runInfos = append(runInfos, struct {
				startIndex int
				endIndex   int
				run        *Run
			}{
				startIndex: currentIndex,
				endIndex:   currentIndex + len(runText),
				run:        &para.Runs[i],
			})
			fullText += runText
			currentIndex += len(runText)
		}
	}

	// 如果没有文本内容，直接返回
	if fullText == "" {
		return nil
	}

	// 使用新的逐个变量替换方法
	newRuns, hasChanges := te.replaceVariablesSequentially(runInfos, fullText, data)

	// 如果有变化，更新段落的Run
	if hasChanges {
		para.Runs = newRuns
	}

	return nil
}

// replaceVariablesSequentially 逐个替换变量，保持样式
func (te *TemplateEngine) replaceVariablesSequentially(originalRunInfos []struct {
	startIndex int
	endIndex   int
	run        *Run
}, originalText string, data *TemplateData) ([]Run, bool) {

	// 找到所有变量位置
	varPattern := regexp.MustCompile(`\{\{(\w+)\}\}`)
	varMatches := varPattern.FindAllStringSubmatchIndex(originalText, -1)

	if len(varMatches) == 0 {
		// 没有变量，检查条件语句
		return te.processConditionals(originalRunInfos, originalText, data)
	}

	newRuns := make([]Run, 0)
	currentPos := 0
	hasChanges := false

	for _, varMatch := range varMatches {
		varStart := varMatch[0]
		varEnd := varMatch[1]
		varNameStart := varMatch[2]
		varNameEnd := varMatch[3]

		// 添加变量前的文本（保持原样式）
		if varStart > currentPos {
			beforeText := originalText[currentPos:varStart]
			beforeRuns := te.extractRunsForSegment(originalRunInfos, currentPos, varStart, beforeText)
			newRuns = append(newRuns, beforeRuns...)
		}

		// 处理变量替换
		varName := originalText[varNameStart:varNameEnd]
		if value, exists := data.Variables[varName]; exists {
			replacementText := te.interfaceToString(value)

			// 为变量选择合适的样式（使用覆盖变量位置的Run样式）
			varRun := te.findRunForPosition(originalRunInfos, varStart)
			if varRun != nil {
				newRun := te.cloneRun(varRun)
				newRun.Text.Content = replacementText
				newRuns = append(newRuns, newRun)
				hasChanges = true
			}
		} else {
			// 变量不存在，保持原始占位符
			varText := originalText[varStart:varEnd]
			varRun := te.findRunForPosition(originalRunInfos, varStart)
			if varRun != nil {
				newRun := te.cloneRun(varRun)
				newRun.Text.Content = varText
				newRuns = append(newRuns, newRun)
			}
		}

		currentPos = varEnd
	}

	// 添加最后剩余的文本
	if currentPos < len(originalText) {
		afterText := originalText[currentPos:]
		afterRuns := te.extractRunsForSegment(originalRunInfos, currentPos, len(originalText), afterText)
		newRuns = append(newRuns, afterRuns...)
	}

	// 如果没有找到任何变量但文本发生了变化，处理条件语句
	if !hasChanges {
		return te.processConditionals(originalRunInfos, originalText, data)
	}

	// 对结果处理条件语句（但要保持每个Run的独立性）
	if hasChanges {
		finalRuns := te.processConditionalsPreservingRuns(newRuns, data)
		return finalRuns, true
	}

	return newRuns, hasChanges
}

// processConditionalsPreservingRuns 处理条件语句但保持Run的独立性
func (te *TemplateEngine) processConditionalsPreservingRuns(runs []Run, data *TemplateData) []Run {
	finalRuns := make([]Run, 0)

	for _, run := range runs {
		originalContent := run.Text.Content
		processedContent := te.renderConditionals(originalContent, data.Conditions)

		// 如果内容发生变化，更新这个Run
		if processedContent != originalContent {
			newRun := run // 复制Run结构
			newRun.Text.Content = processedContent
			finalRuns = append(finalRuns, newRun)
		} else {
			// 内容没有变化，保持原样
			finalRuns = append(finalRuns, run)
		}
	}

	return finalRuns
}

// processConditionals 处理条件语句
func (te *TemplateEngine) processConditionals(originalRunInfos []struct {
	startIndex int
	endIndex   int
	run        *Run
}, originalText string, data *TemplateData) ([]Run, bool) {

	processedText := te.renderConditionals(originalText, data.Conditions)

	if processedText == originalText {
		// 没有变化，返回原始Runs
		newRuns := make([]Run, len(originalRunInfos))
		for i, runInfo := range originalRunInfos {
			newRuns[i] = te.cloneRun(runInfo.run)
		}
		return newRuns, false
	}

	// 有条件语句被处理，简化处理
	if len(originalRunInfos) == 1 {
		newRun := te.cloneRun(originalRunInfos[0].run)
		newRun.Text.Content = processedText
		return []Run{newRun}, true
	}

	// 多个Run的情况，使用第一个Run的样式
	newRun := te.cloneRun(originalRunInfos[0].run)
	newRun.Text.Content = processedText
	return []Run{newRun}, true
}

// extractRunsForSegment 为文本片段提取相应的Run（改进版本）
func (te *TemplateEngine) extractRunsForSegment(originalRunInfos []struct {
	startIndex int
	endIndex   int
	run        *Run
}, segmentStart, segmentEnd int, segmentText string) []Run {
	runs := make([]Run, 0)

	for _, runInfo := range originalRunInfos {
		// 检查Run是否与文本段有重叠
		if runInfo.endIndex > segmentStart && runInfo.startIndex < segmentEnd {
			overlapStart := max(runInfo.startIndex, segmentStart)
			overlapEnd := min(runInfo.endIndex, segmentEnd)

			if overlapEnd > overlapStart {
				newRun := te.cloneRun(runInfo.run)
				// 计算在分段文本中的相对位置
				relativeStart := overlapStart - segmentStart
				relativeEnd := overlapEnd - segmentStart

				// 确保索引在有效范围内
				if relativeStart >= 0 && relativeEnd <= len(segmentText) && relativeStart < relativeEnd {
					newRun.Text.Content = segmentText[relativeStart:relativeEnd]
					if newRun.Text.Content != "" {
						runs = append(runs, newRun)
					}
				}
			}
		}
	}

	return runs
}

// findRunForPosition 找到覆盖指定位置的Run
func (te *TemplateEngine) findRunForPosition(originalRunInfos []struct {
	startIndex int
	endIndex   int
	run        *Run
}, position int) *Run {
	for _, runInfo := range originalRunInfos {
		if position >= runInfo.startIndex && position < runInfo.endIndex {
			return runInfo.run
		}
	}
	// 如果没找到，返回第一个Run
	if len(originalRunInfos) > 0 {
		return originalRunInfos[0].run
	}
	return nil
}

// max 返回两个整数中的较大值
func max(a, b int) int {
	if a > b {
		return a
	}
	return b
}

// min 返回两个整数中的较小值
func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}

// replaceVariablesInTable 在表格中替换变量和处理表格模板
func (te *TemplateEngine) replaceVariablesInTable(table *Table, data *TemplateData) error {
	// 检查是否有表格循环模板
	if len(table.Rows) > 0 && te.isTableTemplate(table) {
		return te.renderTableTemplate(table, data)
	}

	// 普通表格变量替换
	for i := range table.Rows {
		for j := range table.Rows[i].Cells {
			for k := range table.Rows[i].Cells[j].Paragraphs {
				err := te.replaceVariablesInParagraph(&table.Rows[i].Cells[j].Paragraphs[k], data)
				if err != nil {
					return err
				}
			}
		}
	}

	return nil
}

// isTableTemplate 检查表格是否包含模板语法
func (te *TemplateEngine) isTableTemplate(table *Table) bool {
	if len(table.Rows) == 0 {
		return false
	}

	// 检查所有行是否包含循环语法
	for _, row := range table.Rows {
		for _, cell := range row.Cells {
			for _, para := range cell.Paragraphs {
				for _, run := range para.Runs {
					if run.Text.Content != "" && te.containsTemplateLoop(run.Text.Content) {
						return true
					}
				}
			}
		}
	}

	return false
}

// containsTemplateLoop 检查文本是否包含循环模板语法
func (te *TemplateEngine) containsTemplateLoop(text string) bool {
	eachPattern := regexp.MustCompile(`\{\{#each\s+\w+\}\}`)
	return eachPattern.MatchString(text)
}

// renderTableTemplate 渲染表格模板
func (te *TemplateEngine) renderTableTemplate(table *Table, data *TemplateData) error {
	if len(table.Rows) == 0 {
		return nil
	}

	// 找到模板行（包含循环语法的行）
	templateRowIndex := -1
	var listVarName string

	for i, row := range table.Rows {
		found := false
		for _, cell := range row.Cells {
			for _, para := range cell.Paragraphs {
				for _, run := range para.Runs {
					if run.Text.Content != "" {
						eachPattern := regexp.MustCompile(`\{\{#each\s+(\w+)\}\}`)
						matches := eachPattern.FindStringSubmatch(run.Text.Content)
						if len(matches) > 1 {
							templateRowIndex = i
							listVarName = matches[1]
							found = true
							break
						}
					}
				}
				if found {
					break
				}
			}
			if found {
				break
			}
		}
		if found {
			break
		}
	}

	if templateRowIndex < 0 || listVarName == "" {
		return nil
	}

	// 获取列表数据
	listData, exists := data.Lists[listVarName]
	if !exists || len(listData) == 0 {
		// 删除模板行
		table.Rows = append(table.Rows[:templateRowIndex], table.Rows[templateRowIndex+1:]...)
		return nil
	}

	// 保存模板行
	templateRow := table.Rows[templateRowIndex]
	newRows := make([]TableRow, 0)

	// 保留模板行之前的行
	newRows = append(newRows, table.Rows[:templateRowIndex]...)

	// 为每个数据项生成新行
	for _, item := range listData {
		newRow := te.cloneTableRow(&templateRow)

		// 在新行中替换变量
		if itemMap, ok := item.(map[string]interface{}); ok {
			for i := range newRow.Cells {
				for j := range newRow.Cells[i].Paragraphs {
					for k := range newRow.Cells[i].Paragraphs[j].Runs {
						if newRow.Cells[i].Paragraphs[j].Runs[k].Text.Content != "" {
							// 移除模板语法标记
							content := newRow.Cells[i].Paragraphs[j].Runs[k].Text.Content
							content = regexp.MustCompile(`\{\{#each\s+\w+\}\}`).ReplaceAllString(content, "")
							content = regexp.MustCompile(`\{\{/each\}\}`).ReplaceAllString(content, "")

							// 替换变量
							for key, value := range itemMap {
								placeholder := fmt.Sprintf("{{%s}}", key)
								content = strings.ReplaceAll(content, placeholder, te.interfaceToString(value))
							}

							// 处理条件语句
							content = te.renderLoopConditionals(content, itemMap)

							newRow.Cells[i].Paragraphs[j].Runs[k].Text.Content = content
						}
					}
				}
			}
		}

		newRows = append(newRows, *newRow)
	}

	// 保留模板行之后的行
	newRows = append(newRows, table.Rows[templateRowIndex+1:]...)

	// 更新表格行
	table.Rows = newRows

	return nil
}

// NewTemplateData 创建新的模板数据
func NewTemplateData() *TemplateData {
	return &TemplateData{
		Variables:  make(map[string]interface{}),
		Lists:      make(map[string][]interface{}),
		Conditions: make(map[string]bool),
	}
}

// SetVariable 设置变量
func (td *TemplateData) SetVariable(name string, value interface{}) {
	td.Variables[name] = value
}

// SetList 设置列表
func (td *TemplateData) SetList(name string, list []interface{}) {
	td.Lists[name] = list
}

// SetCondition 设置条件
func (td *TemplateData) SetCondition(name string, value bool) {
	td.Conditions[name] = value
}

// SetVariables 批量设置变量
func (td *TemplateData) SetVariables(variables map[string]interface{}) {
	for name, value := range variables {
		td.Variables[name] = value
	}
}

// GetVariable 获取变量
func (td *TemplateData) GetVariable(name string) (interface{}, bool) {
	value, exists := td.Variables[name]
	return value, exists
}

// GetList 获取列表
func (td *TemplateData) GetList(name string) ([]interface{}, bool) {
	list, exists := td.Lists[name]
	return list, exists
}

// GetCondition 获取条件
func (td *TemplateData) GetCondition(name string) (bool, bool) {
	condition, exists := td.Conditions[name]
	return condition, exists
}

// Merge 合并模板数据
func (td *TemplateData) Merge(other *TemplateData) {
	// 合并变量
	for name, value := range other.Variables {
		td.Variables[name] = value
	}

	// 合并列表
	for name, list := range other.Lists {
		td.Lists[name] = list
	}

	// 合并条件
	for name, condition := range other.Conditions {
		td.Conditions[name] = condition
	}
}

// Clear 清空模板数据
func (td *TemplateData) Clear() {
	td.Variables = make(map[string]interface{})
	td.Lists = make(map[string][]interface{})
	td.Conditions = make(map[string]bool)
}

// FromStruct 从结构体生成模板数据
func (td *TemplateData) FromStruct(data interface{}) error {
	value := reflect.ValueOf(data)
	if value.Kind() == reflect.Ptr {
		value = value.Elem()
	}

	if value.Kind() != reflect.Struct {
		return NewValidationError("data_type", "struct", "expected struct type")
	}

	typ := value.Type()
	for i := 0; i < value.NumField(); i++ {
		field := typ.Field(i)
		fieldValue := value.Field(i)

		// 跳过不可导出的字段
		if !fieldValue.CanInterface() {
			continue
		}

		fieldName := strings.ToLower(field.Name)
		td.Variables[fieldName] = fieldValue.Interface()
	}

	return nil
}

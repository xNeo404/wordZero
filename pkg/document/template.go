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
	var content strings.Builder

	// 遍历文档的所有段落，生成模板字符串
	for _, element := range doc.Body.Elements {
		switch elem := element.(type) {
		case *Paragraph:
			// 处理段落
			for _, run := range elem.Runs {
				content.WriteString(run.Text.Content)
			}
			content.WriteString("\n")
		case *Table:
			// 处理表格
			content.WriteString("{{#table}}\n")
			for i, row := range elem.Rows {
				content.WriteString(fmt.Sprintf("{{#row_%d}}\n", i))
				for j := range row.Cells {
					content.WriteString(fmt.Sprintf("{{cell_%d_%d}}", i, j))
				}
				content.WriteString("{{/row}}\n")
			}
			content.WriteString("{{/table}}\n")
		}
	}

	return content.String(), nil
}

// cloneDocument 克隆文档
func (te *TemplateEngine) cloneDocument(source *Document) *Document {
	// 简单实现：创建新文档并复制基本结构
	doc := New()

	// 复制样式管理器
	if source.styleManager != nil {
		doc.styleManager = source.styleManager
	}

	return doc
}

// applyRenderedContentToDocument 将渲染内容应用到文档
func (te *TemplateEngine) applyRenderedContentToDocument(doc *Document, content string) error {
	// 将渲染后的内容按行分割并添加到文档
	lines := strings.Split(content, "\n")

	for _, line := range lines {
		line = strings.TrimSpace(line)
		if line != "" {
			doc.AddParagraph(line)
		}
	}

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

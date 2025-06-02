// Package test 模板继承功能测试
package test

import (
	"os"
	"testing"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
)

// TestTemplateInheritanceComplete 完整的模板继承测试
func TestTemplateInheritanceComplete(t *testing.T) {
	// 确保输出目录存在
	outputDir := "output"
	if _, err := os.Stat(outputDir); os.IsNotExist(err) {
		err = os.Mkdir(outputDir, 0755)
		if err != nil {
			t.Fatalf("Failed to create output directory: %v", err)
		}
	}

	engine := document.NewTemplateEngine()

	// 测试1: 基础模板继承
	t.Run("基础模板继承", func(t *testing.T) {
		testBasicInheritance(t, engine)
	})

	// 测试2: 块重写功能
	t.Run("块重写功能", func(t *testing.T) {
		testBlockOverride(t, engine)
	})

	// 测试3: 多级继承
	t.Run("多级继承", func(t *testing.T) {
		testMultiLevelInheritance(t, engine)
	})

	// 测试4: 块默认内容
	t.Run("块默认内容", func(t *testing.T) {
		testBlockDefaultContent(t, engine)
	})
}

// testBasicInheritance 测试基础模板继承功能
func testBasicInheritance(t *testing.T, engine *document.TemplateEngine) {
	// 创建基础模板
	baseTemplate := `{{companyName}} 官方文档

版本：{{version}}
创建时间：{{createTime}}

{{#block "content"}}
默认内容区域
{{/block}}

{{#block "footer"}}
版权所有 © {{year}} {{companyName}}
{{/block}}`

	_, err := engine.LoadTemplate("base", baseTemplate)
	if err != nil {
		t.Fatalf("Failed to load base template: %v", err)
	}

	// 创建子模板
	childTemplate := `{{extends "base"}}

{{#block "content"}}
用户手册

第一章：快速开始
欢迎使用我们的产品！

第二章：功能介绍
详细的功能说明...
{{/block}}`

	_, err = engine.LoadTemplate("user_manual", childTemplate)
	if err != nil {
		t.Fatalf("Failed to load child template: %v", err)
	}

	// 准备数据
	data := document.NewTemplateData()
	data.SetVariable("companyName", "WordZero科技")
	data.SetVariable("version", "v1.0.0")
	data.SetVariable("createTime", "2024-12-01")
	data.SetVariable("year", "2024")

	// 渲染模板
	doc, err := engine.RenderToDocument("user_manual", data)
	if err != nil {
		t.Fatalf("Failed to render template: %v", err)
	}

	// 保存文档
	err = doc.Save("output/test_basic_inheritance.docx")
	if err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// 验证文档内容
	if len(doc.Body.Elements) == 0 {
		t.Error("Expected document to have content")
	}
}

// testBlockOverride 测试块重写功能
func testBlockOverride(t *testing.T, engine *document.TemplateEngine) {
	// 创建基础模板
	baseTemplate := `企业报告模板

{{#block "header"}}
标准报告头部
{{/block}}

{{#block "main_content"}}
标准内容区域
{{/block}}

{{#block "sidebar"}}
标准侧边栏
{{/block}}

{{#block "footer"}}
标准页脚
{{/block}}`

	_, err := engine.LoadTemplate("report_base", baseTemplate)
	if err != nil {
		t.Fatalf("Failed to load base template: %v", err)
	}

	// 创建特定报告模板，重写部分块
	salesReportTemplate := `{{extends "report_base"}}

{{#block "header"}}
销售业绩报告
报告期间：{{reportPeriod}}
{{/block}}

{{#block "main_content"}}
销售数据分析

总销售额：{{totalSales}} 元
增长率：{{growthRate}}%

{{#each regions}}
- {{name}}: {{sales}} 元
{{/each}}
{{/block}}

{{#block "sidebar"}}
快速统计
- 新客户：{{newCustomers}} 人
- 回头客：{{returningCustomers}} 人
- 平均订单：{{averageOrder}} 元
{{/block}}`

	_, err = engine.LoadTemplate("sales_report", salesReportTemplate)
	if err != nil {
		t.Fatalf("Failed to load sales report template: %v", err)
	}

	// 准备数据
	data := document.NewTemplateData()
	data.SetVariable("reportPeriod", "2024年11月")
	data.SetVariable("totalSales", "1,250,000")
	data.SetVariable("growthRate", "15.8")
	data.SetVariable("newCustomers", "158")
	data.SetVariable("returningCustomers", "432")
	data.SetVariable("averageOrder", "2,890")

	regions := []interface{}{
		map[string]interface{}{"name": "华东区", "sales": "450,000"},
		map[string]interface{}{"name": "华北区", "sales": "380,000"},
		map[string]interface{}{"name": "华南区", "sales": "420,000"},
	}
	data.SetList("regions", regions)

	// 渲染模板
	doc, err := engine.RenderToDocument("sales_report", data)
	if err != nil {
		t.Fatalf("Failed to render template: %v", err)
	}

	// 保存文档
	err = doc.Save("output/test_block_override.docx")
	if err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// 验证文档内容
	if len(doc.Body.Elements) == 0 {
		t.Error("Expected document to have content")
	}
}

// testMultiLevelInheritance 测试多级继承
func testMultiLevelInheritance(t *testing.T, engine *document.TemplateEngine) {
	// 第一级：基础文档模板
	baseDocTemplate := `{{#block "document_header"}}
文档标题：{{title}}
{{/block}}

{{#block "main_body"}}
主要内容区域
{{/block}}

{{#block "document_footer"}}
页脚信息
{{/block}}`

	_, err := engine.LoadTemplate("base_doc", baseDocTemplate)
	if err != nil {
		t.Fatalf("Failed to load base doc template: %v", err)
	}

	// 第二级：技术文档模板
	techDocTemplate := `{{extends "base_doc"}}

{{#block "document_header"}}
技术文档：{{title}}
版本：{{version}}
作者：{{author}}
{{/block}}

{{#block "main_body"}}
{{#block "abstract"}}
文档摘要
{{/block}}

{{#block "technical_content"}}
技术内容
{{/block}}

{{#block "references"}}
参考资料
{{/block}}
{{/block}}`

	_, err = engine.LoadTemplate("tech_doc", techDocTemplate)
	if err != nil {
		t.Fatalf("Failed to load tech doc template: %v", err)
	}

	// 第三级：API文档模板
	apiDocTemplate := `{{extends "tech_doc"}}

{{#block "abstract"}}
API接口文档摘要
本文档描述了{{apiName}}的使用方法和接口规范。
{{/block}}

{{#block "technical_content"}}
API接口列表

{{#each endpoints}}
### {{method}} {{path}}
描述：{{description}}
参数：{{parameters}}

{{/each}}
{{/block}}

{{#block "references"}}
相关文档：
- API设计规范
- 认证机制说明
- 错误码参考
{{/block}}`

	_, err = engine.LoadTemplate("api_doc", apiDocTemplate)
	if err != nil {
		t.Fatalf("Failed to load API doc template: %v", err)
	}

	// 准备数据
	data := document.NewTemplateData()
	data.SetVariable("title", "WordZero API文档")
	data.SetVariable("version", "v2.1.0")
	data.SetVariable("author", "技术团队")
	data.SetVariable("apiName", "WordZero API")

	endpoints := []interface{}{
		map[string]interface{}{
			"method":      "POST",
			"path":        "/api/documents",
			"description": "创建新文档",
			"parameters":  "title, content, format",
		},
		map[string]interface{}{
			"method":      "GET",
			"path":        "/api/documents/{id}",
			"description": "获取文档详情",
			"parameters":  "id (路径参数)",
		},
	}
	data.SetList("endpoints", endpoints)

	// 渲染模板
	doc, err := engine.RenderToDocument("api_doc", data)
	if err != nil {
		t.Fatalf("Failed to render template: %v", err)
	}

	// 保存文档
	err = doc.Save("output/test_multi_level_inheritance.docx")
	if err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// 验证文档内容
	if len(doc.Body.Elements) == 0 {
		t.Error("Expected document to have content")
	}
}

// testBlockDefaultContent 测试块默认内容
func testBlockDefaultContent(t *testing.T, engine *document.TemplateEngine) {
	// 创建有默认内容的基础模板
	baseTemplate := `产品文档

{{#block "intro"}}
默认介绍内容
这是产品的基本介绍。
{{/block}}

{{#block "features"}}
默认功能列表
- 基础功能1
- 基础功能2
{{/block}}

{{#block "contact"}}
默认联系方式
邮箱：default@example.com
电话：000-0000-0000
{{/block}}`

	_, err := engine.LoadTemplate("product_base", baseTemplate)
	if err != nil {
		t.Fatalf("Failed to load base template: %v", err)
	}

	// 创建子模板，只重写部分块
	productTemplate := `{{extends "product_base"}}

{{#block "intro"}}
WordZero产品介绍
WordZero是一个强大的Word文档处理库。
{{/block}}

{{#block "features"}}
WordZero功能特性
- 文档创建和编辑
- 模板渲染
- 表格处理
- 图片插入
- 样式管理
{{/block}}`

	_, err = engine.LoadTemplate("wordzero_product", productTemplate)
	if err != nil {
		t.Fatalf("Failed to load product template: %v", err)
	}

	// 准备数据
	data := document.NewTemplateData()

	// 渲染模板
	doc, err := engine.RenderToDocument("wordzero_product", data)
	if err != nil {
		t.Fatalf("Failed to render template: %v", err)
	}

	// 保存文档
	err = doc.Save("output/test_block_default_content.docx")
	if err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// 验证文档内容
	if len(doc.Body.Elements) == 0 {
		t.Error("Expected document to have content")
	}

	// 验证默认内容保持不变（contact块应该使用默认内容）
	// 这里我们可以通过渲染的内容来验证
	renderedContent, err := engine.GetTemplate("wordzero_product")
	if err != nil {
		t.Fatalf("Failed to get template: %v", err)
	}

	// 验证模板结构
	if renderedContent.Parent == nil {
		t.Error("Expected template to have parent")
	}

	if len(renderedContent.DefinedBlocks) != 2 {
		t.Errorf("Expected 2 defined blocks, got %d", len(renderedContent.DefinedBlocks))
	}
}

// TestTemplateInheritanceValidation 测试模板继承验证
func TestTemplateInheritanceValidation(t *testing.T) {
	engine := document.NewTemplateEngine()

	// 测试块语法验证
	t.Run("块语法验证", func(t *testing.T) {
		// 正确的块语法
		validTemplate := `{{#block "test"}}内容{{/block}}`
		template, err := engine.LoadTemplate("valid_block", validTemplate)
		if err != nil {
			t.Fatalf("Failed to load valid template: %v", err)
		}

		err = engine.ValidateTemplate(template)
		if err != nil {
			t.Errorf("Valid template should pass validation: %v", err)
		}

		// 错误的块语法 - 缺少结束标签
		invalidTemplate := `{{#block "test"}}内容`
		template2, err := engine.LoadTemplate("invalid_block", invalidTemplate)
		if err != nil {
			t.Fatalf("Failed to load invalid template: %v", err)
		}

		err = engine.ValidateTemplate(template2)
		if err == nil {
			t.Error("Invalid template should fail validation")
		}
	})

	// 测试继承链验证
	t.Run("继承链验证", func(t *testing.T) {
		// 创建基础模板
		baseTemplate := `{{#block "content"}}基础内容{{/block}}`
		_, err := engine.LoadTemplate("inheritance_base", baseTemplate)
		if err != nil {
			t.Fatalf("Failed to load base template: %v", err)
		}

		// 创建子模板
		childTemplate := `{{extends "inheritance_base"}}
{{#block "content"}}子模板内容{{/block}}`
		child, err := engine.LoadTemplate("inheritance_child", childTemplate)
		if err != nil {
			t.Fatalf("Failed to load child template: %v", err)
		}

		// 验证继承关系
		if child.Parent == nil {
			t.Error("Child template should have parent")
		}

		if child.Parent.Name != "inheritance_base" {
			t.Errorf("Expected parent name 'inheritance_base', got '%s'", child.Parent.Name)
		}
	})
}

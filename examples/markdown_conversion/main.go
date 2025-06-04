package main

import (
	"fmt"
	"os"

	"github.com/ZeroHawkeye/wordZero/pkg/markdown"
)

func main() {
	fmt.Println("WordZero Markdown双向转换完整示例")
	fmt.Println("===================================")

	// 确保输出目录存在
	outputDir := "examples/output"
	err := os.MkdirAll(outputDir, 0755)
	if err != nil {
		fmt.Printf("❌ 创建输出目录失败: %v\n", err)
		os.Exit(1)
	}

	// 演示1: Markdown转Word
	demonstrateMarkdownToWord(outputDir)

	// 演示2: Word转Markdown（反向转换）
	demonstrateWordToMarkdown(outputDir)

	// 演示3: 双向转换器使用
	demonstrateBidirectionalConverter(outputDir)

	// 演示4: 批量转换功能
	demonstrateBatchConversion(outputDir)

	fmt.Println("\n🎉 所有转换示例运行完成！")
}

// demonstrateMarkdownToWord 演示Markdown转Word功能
func demonstrateMarkdownToWord(outputDir string) {
	fmt.Println("\n📝 演示1: Markdown → Word 转换")
	fmt.Println("================================")

	// 创建示例Markdown内容
	markdownContent := `# WordZero Markdown双向转换功能

欢迎使用WordZero库的Markdown和Word文档双向转换功能！

## 功能特性概览

WordZero现在支持**完整的双向转换**：

### 🚀 Markdown → Word 转换
- **goldmark解析引擎**: 基于CommonMark 0.31.2规范
- **完整语法支持**: 标题、格式化、列表、表格、图片、链接
- **智能样式映射**: 自动应用Word标准样式
- **可配置选项**: GitHub风味Markdown、脚注、错误处理

### 🔄 Word → Markdown 反向转换
- **结构完整保持**: 保持原文档的层次结构
- **格式智能识别**: 自动识别并转换文本格式
- **图片导出支持**: 提取图片并生成引用
- **多种导出模式**: GFM表格、Setext标题等选项

### 文本格式化示例
- **粗体文本**展示
- *斜体文本*展示
- ` + "`行内代码`" + `展示

### 列表支持示例

#### 无序列表
- 功能A: 基础Markdown语法
- 功能B: GitHub风味扩展
- 功能C: 自定义配置选项

#### 有序列表
1. 安装WordZero库
2. 创建转换器实例
3. 调用转换方法
4. 处理转换结果

### 引用块示例

> 这是一个引用块示例，演示引用文本的转换效果。
> 
> 引用块中可以包含多行内容，在Word中会以特殊格式显示。

### 代码块示例

` + "```" + `go
// WordZero双向转换示例代码
package main

import "github.com/ZeroHawkeye/wordZero/pkg/markdown"

func main() {
    // Markdown转Word
    converter := markdown.NewConverter(markdown.DefaultOptions())
    doc, _ := converter.ConvertString(markdownText, nil)
    doc.Save("output.docx")
    
    // Word转Markdown
    exporter := markdown.NewExporter(markdown.DefaultExportOptions())
    exporter.ExportToFile("input.docx", "output.md", nil)
}
` + "```" + `

---

## 技术实现亮点

### 🔧 核心技术栈
- **goldmark**: 高性能Markdown解析器
- **WordZero**: 原生Go Word文档处理
- **双向转换**: 无缝的格式转换支持

### 📋 支持的配置选项
- ✅ GitHub Flavored Markdown扩展
- ✅ 脚注和尾注支持
- ✅ 表格格式转换（待完善）
- ✅ 任务列表支持（待实现）
- ✅ 图片处理和路径解析
- ✅ 错误处理和进度报告

### 🎯 使用场景
1. **技术文档转换**: 从Markdown快速生成Word文档
2. **报告自动化**: 将Word报告转换为Markdown
3. **版本控制友好**: Word文档转为可diff的Markdown
4. **批量处理**: 大量文档的格式转换

## 总结

WordZero的双向转换功能为现代文档工作流提供了强大支持，
无论是从轻量级的Markdown到专业的Word文档，
还是反向的格式转换，都能满足不同场景的需求。`

	// 创建转换器（使用高质量配置）
	opts := markdown.HighQualityOptions()
	opts.GenerateTOC = true
	opts.TOCMaxLevel = 3
	converter := markdown.NewConverter(opts)

	fmt.Println("📝 正在转换Markdown内容...")

	// 转换为Word文档
	doc, err := converter.ConvertString(markdownContent, nil)
	if err != nil {
		fmt.Printf("❌ 转换失败: %v\n", err)
		os.Exit(1)
	}

	// 保存Word文档
	outputPath := outputDir + "/markdown_to_word_demo.docx"
	err = doc.Save(outputPath)
	if err != nil {
		fmt.Printf("❌ 保存文档失败: %v\n", err)
		os.Exit(1)
	}

	fmt.Printf("✅ Markdown转Word成功！输出: %s\n", outputPath)

	// 同时保存Markdown源文件供后续演示使用
	markdownPath := outputDir + "/source_document.md"
	err = os.WriteFile(markdownPath, []byte(markdownContent), 0644)
	if err != nil {
		fmt.Printf("❌ 保存Markdown文件失败: %v\n", err)
		os.Exit(1)
	}
}

// demonstrateWordToMarkdown 演示Word转Markdown功能
func demonstrateWordToMarkdown(outputDir string) {
	fmt.Println("\n📄 演示2: Word → Markdown 反向转换")
	fmt.Println("===================================")

	// 使用上一步生成的Word文档
	wordPath := outputDir + "/markdown_to_word_demo.docx"
	markdownOutputPath := outputDir + "/word_to_markdown_result.md"

	// 创建导出器（使用高质量配置）
	exportOpts := markdown.HighQualityExportOptions()
	exportOpts.ExtractImages = true
	exportOpts.ImageOutputDir = outputDir + "/extracted_images"
	exportOpts.UseGFMTables = true
	exportOpts.IncludeMetadata = true

	exporter := markdown.NewExporter(exportOpts)

	fmt.Println("📄 正在将Word文档转换为Markdown...")

	// 执行反向转换
	err := exporter.ExportToFile(wordPath, markdownOutputPath, nil)
	if err != nil {
		fmt.Printf("❌ Word转Markdown失败: %v\n", err)
		return
	}

	fmt.Printf("✅ Word转Markdown成功！输出: %s\n", markdownOutputPath)

	// 显示转换结果预览
	content, err := os.ReadFile(markdownOutputPath)
	if err == nil && len(content) > 0 {
		preview := string(content)
		if len(preview) > 300 {
			preview = preview[:300] + "..."
		}
		fmt.Printf("📋 转换结果预览:\n%s\n", preview)
	}
}

// demonstrateBidirectionalConverter 演示双向转换器
func demonstrateBidirectionalConverter(outputDir string) {
	fmt.Println("\n🔄 演示3: 双向转换器统一接口")
	fmt.Println("===============================")

	// 创建双向转换器
	converter := markdown.NewBidirectionalConverter(
		markdown.HighQualityOptions(),
		markdown.HighQualityExportOptions(),
	)

	// 测试自动类型检测转换
	testCases := []struct {
		input  string
		output string
		desc   string
	}{
		{
			input:  outputDir + "/source_document.md",
			output: outputDir + "/auto_converted.docx",
			desc:   "Markdown自动转换为Word",
		},
		{
			input:  outputDir + "/markdown_to_word_demo.docx",
			output: outputDir + "/auto_converted.md",
			desc:   "Word自动转换为Markdown",
		},
	}

	for i, tc := range testCases {
		fmt.Printf("🔄 测试%d: %s\n", i+1, tc.desc)

		err := converter.AutoConvert(tc.input, tc.output)
		if err != nil {
			fmt.Printf("❌ 自动转换失败: %v\n", err)
			continue
		}

		fmt.Printf("✅ 自动转换成功: %s\n", tc.output)
	}
}

// demonstrateBatchConversion 演示批量转换功能
func demonstrateBatchConversion(outputDir string) {
	fmt.Println("\n📦 演示4: 批量转换功能")
	fmt.Println("=======================")

	// 创建多个测试文件
	testMarkdownFiles := []string{
		outputDir + "/test1.md",
		outputDir + "/test2.md",
		outputDir + "/test3.md",
	}

	testContents := []string{
		"# 测试文档1\n\n这是第一个测试文档。\n\n## 内容\n- 项目A\n- 项目B",
		"# 测试文档2\n\n这是第二个测试文档。\n\n> 引用内容示例",
		"# 测试文档3\n\n这是第三个测试文档。\n\n```go\nfmt.Println(\"Hello\")\n```",
	}

	// 创建测试文件
	for i, content := range testContents {
		err := os.WriteFile(testMarkdownFiles[i], []byte(content), 0644)
		if err != nil {
			fmt.Printf("❌ 创建测试文件失败: %v\n", err)
			return
		}
	}

	// 执行批量转换
	converter := markdown.NewConverter(markdown.DefaultOptions())
	batchOutputDir := outputDir + "/batch_output"

	fmt.Println("📦 正在执行批量Markdown转Word...")

	err := converter.BatchConvert(testMarkdownFiles, batchOutputDir, &markdown.ConvertOptions{
		ProgressCallback: func(current, total int) {
			fmt.Printf("📊 批量转换进度: %d/%d\n", current, total)
		},
		ErrorCallback: func(err error) {
			fmt.Printf("⚠️ 转换警告: %v\n", err)
		},
	})

	if err != nil {
		fmt.Printf("❌ 批量转换失败: %v\n", err)
		return
	}

	fmt.Printf("✅ 批量转换完成！输出目录: %s\n", batchOutputDir)
}

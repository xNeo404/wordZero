/*
Package document 提供了用于创建、编辑和操作 Microsoft Word 文档的 Golang 库。

WordZero 专注于现代的 Office Open XML (OOXML) 格式（.docx 文件），
提供了简单易用的 API 来创建和修改 Word 文档。

# 主要功能

- 创建新的 Word 文档
- 打开和解析现有的 .docx 文件
- 添加和格式化文本内容
- 设置段落样式和对齐方式
- 配置字体、颜色和文本格式
- 设置行间距和段落间距
- 错误处理和日志记录

# 快速开始

创建一个简单的文档：

	doc := document.New()
	doc.AddParagraph("Hello, World!")
	err := doc.Save("hello.docx")

创建带格式的文档：

	doc := document.New()

	// 添加格式化标题
	titleFormat := &document.TextFormat{
		Bold:      true,
		FontSize:  18,
		FontColor: "FF0000", // 红色
		FontName:  "微软雅黑",
	}
	title := doc.AddFormattedParagraph("文档标题", titleFormat)
	title.SetAlignment(document.AlignCenter)

	// 添加正文段落
	para := doc.AddParagraph("这是正文内容...")
	para.SetSpacing(&document.SpacingConfig{
		LineSpacing:     1.5, // 1.5倍行距
		BeforePara:      12,  // 段前12磅
		AfterPara:       6,   // 段后6磅
		FirstLineIndent: 24,  // 首行缩进24磅
	})

	err := doc.Save("formatted.docx")

打开现有文档：

	doc, err := document.Open("existing.docx")
	if err != nil {
		log.Fatal(err)
	}

	// 读取段落内容
	for i, para := range doc.Body.Paragraphs {
		fmt.Printf("段落 %d: ", i+1)
		for _, run := range para.Runs {
			fmt.Print(run.Text.Content)
		}
		fmt.Println()
	}

# 错误处理

库提供了统一的错误处理机制：

	doc, err := document.Open("nonexistent.docx")
	if err != nil {
		var docErr *document.DocumentError
		if errors.As(err, &docErr) {
			fmt.Printf("操作: %s, 错误: %v\n", docErr.Operation, docErr.Cause)
		}
	}

# 日志记录

可以配置日志级别来控制输出：

	// 设置为调试模式
	document.SetGlobalLevel(document.LogLevelDebug)

	// 只显示错误
	document.SetGlobalLevel(document.LogLevelError)

# 文本格式

TextFormat 结构体支持多种文本格式选项：

	format := &document.TextFormat{
		Bold:      true,           // 粗体
		Italic:    true,           // 斜体
		FontSize:  14,             // 字体大小（磅）
		FontColor: "0000FF",       // 字体颜色（十六进制）
		FontName:  "Times New Roman", // 字体名称
	}

# 段落对齐

支持四种对齐方式：

	para.SetAlignment(document.AlignLeft)     // 左对齐
	para.SetAlignment(document.AlignCenter)   // 居中对齐
	para.SetAlignment(document.AlignRight)    // 右对齐
	para.SetAlignment(document.AlignJustify)  // 两端对齐

# 间距配置

可以精确控制段落间距：

	config := &document.SpacingConfig{
		LineSpacing:     1.5, // 行间距（倍数）
		BeforePara:      12,  // 段前间距（磅）
		AfterPara:       6,   // 段后间距（磅）
		FirstLineIndent: 24,  // 首行缩进（磅）
	}
	para.SetSpacing(config)

# 注意事项

- 字体大小以磅为单位，内部会自动转换为 Word 的半磅单位
- 颜色值使用十六进制格式，不需要 # 前缀
- 间距值以磅为单位，内部会转换为 TWIPs（1磅=20TWIPs）
- 所有文本内容都使用 UTF-8 编码

更多详细信息和示例，请参阅各个类型和函数的文档。
*/
package document

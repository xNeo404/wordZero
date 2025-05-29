package main

import (
	"flag"
	"fmt"
	"log"
	"os"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
)

func main() {
	// 定义命令行参数
	action := flag.String("action", "demo", "要执行的操作: demo, create, open")
	file := flag.String("file", "output.docx", "文档文件路径")
	text := flag.String("text", "Hello, World!", "要添加的文本")
	debug := flag.Bool("debug", false, "启用调试模式")
	flag.Parse()

	// 设置日志级别
	if *debug {
		document.SetGlobalLevel(document.LogLevelDebug)
	} else {
		document.SetGlobalLevel(document.LogLevelInfo)
	}

	switch *action {
	case "demo":
		createDemoDocument()
	case "create":
		createSimpleDocument(*file, *text)
	case "open":
		openAndReadDocument(*file)
	default:
		fmt.Printf("未知操作: %s\n", *action)
		flag.Usage()
		os.Exit(1)
	}
}

// createDemoDocument 创建一个演示文档，展示所有格式化功能
func createDemoDocument() {
	fmt.Println("创建演示文档...")

	doc := document.New()

	// 1. 添加标题
	titleFormat := &document.TextFormat{
		Bold:      true,
		FontSize:  20,
		FontColor: "FF0000", // 红色
		FontName:  "微软雅黑",
	}
	title := doc.AddFormattedParagraph("WordZero 功能演示", titleFormat)
	title.SetAlignment(document.AlignCenter)

	// 2. 添加副标题
	subtitleFormat := &document.TextFormat{
		Bold:     true,
		FontSize: 16,
		FontName: "微软雅黑",
	}
	subtitle := doc.AddFormattedParagraph("文本格式化示例", subtitleFormat)
	subtitle.SetAlignment(document.AlignCenter)
	subtitle.SetSpacing(&document.SpacingConfig{
		BeforePara: 12,
		AfterPara:  6,
	})

	// 3. 添加正文段落
	doc.AddParagraph("WordZero 是一个使用 Golang 实现的 Word 文档操作库，提供丰富的文档创建和编辑功能。")

	// 4. 添加带间距的段落
	para := doc.AddParagraph("这个段落演示了间距和缩进设置：首行缩进、1.5倍行距、段前段后间距。")
	para.SetSpacing(&document.SpacingConfig{
		LineSpacing:     1.5,
		BeforePara:      12,
		AfterPara:       6,
		FirstLineIndent: 24,
	})
	para.SetAlignment(document.AlignJustify)

	// 5. 添加混合格式段落
	mixed := doc.AddParagraph("格式化文本示例：")
	mixed.AddFormattedText("粗体红色", &document.TextFormat{
		Bold: true, FontColor: "FF0000"})
	mixed.AddFormattedText("、", nil)
	mixed.AddFormattedText("斜体蓝色", &document.TextFormat{
		Italic: true, FontColor: "0000FF"})
	mixed.AddFormattedText("、", nil)
	mixed.AddFormattedText("大号绿色", &document.TextFormat{
		FontSize: 16, FontColor: "00AA00"})
	mixed.AddFormattedText("。", nil)

	// 6. 添加不同对齐方式的段落
	left := doc.AddParagraph("左对齐文本")
	left.SetAlignment(document.AlignLeft)

	center := doc.AddParagraph("居中对齐文本")
	center.SetAlignment(document.AlignCenter)

	right := doc.AddParagraph("右对齐文本")
	right.SetAlignment(document.AlignRight)

	justify := doc.AddParagraph("两端对齐文本，当文本较长时效果更明显。这是一个较长的段落，用来演示两端对齐的效果。")
	justify.SetAlignment(document.AlignJustify)

	// 保存文档
	filename := "demo_document.docx"
	err := doc.Save(filename)
	if err != nil {
		log.Fatalf("保存文档失败: %v", err)
	}

	fmt.Printf("演示文档已创建: %s\n", filename)
}

// createSimpleDocument 创建简单文档
func createSimpleDocument(filename, text string) {
	fmt.Printf("创建文档: %s\n", filename)

	doc := document.New()
	doc.AddParagraph(text)

	err := doc.Save(filename)
	if err != nil {
		log.Fatalf("保存文档失败: %v", err)
	}

	fmt.Printf("文档已保存: %s\n", filename)
}

// openAndReadDocument 打开并读取文档
func openAndReadDocument(filename string) {
	fmt.Printf("打开文档: %s\n", filename)

	doc, err := document.Open(filename)
	if err != nil {
		log.Fatalf("打开文档失败: %v", err)
	}

	fmt.Printf("文档包含 %d 个段落\n", len(doc.Body.Paragraphs))

	for i, para := range doc.Body.Paragraphs {
		fmt.Printf("段落 %d: ", i+1)
		for _, run := range para.Runs {
			fmt.Print(run.Text.Content)
		}
		fmt.Println()
	}
}

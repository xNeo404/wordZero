package main

import (
	"log"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
)

func main() {
	// 设置日志级别为调试模式
	document.SetGlobalLevel(document.LogLevelDebug)

	// 创建新文档
	doc := document.New()

	// 1. 添加普通段落
	doc.AddParagraph("这是一个普通段落")

	// 2. 添加格式化段落 - 标题样式
	titleFormat := &document.TextFormat{
		Bold:      true,
		FontSize:  18,
		FontColor: "FF0000", // 红色
		FontName:  "微软雅黑",
	}
	p2 := doc.AddFormattedParagraph("这是一个格式化标题", titleFormat)
	p2.SetAlignment(document.AlignCenter) // 居中对齐

	// 3. 添加带间距的段落
	p3 := doc.AddParagraph("这个段落有特定的间距设置")
	spacingConfig := &document.SpacingConfig{
		LineSpacing:     1.5, // 1.5倍行距
		BeforePara:      12,  // 段前12磅
		AfterPara:       6,   // 段后6磅
		FirstLineIndent: 24,  // 首行缩进24磅
	}
	p3.SetSpacing(spacingConfig)
	p3.SetAlignment(document.AlignJustify) // 两端对齐

	// 4. 添加混合格式的段落
	p4 := doc.AddParagraph("这个段落包含多种格式：")

	// 添加粗体文本
	boldFormat := &document.TextFormat{
		Bold:      true,
		FontColor: "0000FF", // 蓝色
	}
	p4.AddFormattedText("粗体蓝色文本", boldFormat)

	// 添加普通文本
	p4.AddFormattedText("，普通文本，", nil)

	// 添加斜体文本
	italicFormat := &document.TextFormat{
		Italic:    true,
		FontColor: "00FF00", // 绿色
		FontSize:  14,
	}
	p4.AddFormattedText("斜体绿色大字", italicFormat)

	// 5. 添加右对齐段落
	p5 := doc.AddFormattedParagraph("这个段落右对齐", &document.TextFormat{
		FontName: "Times New Roman",
		FontSize: 12,
	})
	p5.SetAlignment(document.AlignRight)

	// 保存文档
	err := doc.Save("../output/formatted_document.docx")
	if err != nil {
		log.Fatalf("保存文档失败: %v", err)
	}

	log.Println("格式化文档创建成功: ../output/formatted_document.docx")
}

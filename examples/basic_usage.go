package main

import (
	"fmt"
	"log"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
)

func main() {
	// 创建新文档
	doc := document.New()

	// 添加标题
	doc.AddParagraph("WordZero 示例文档")

	// 添加内容
	doc.AddParagraph("这是一个使用 WordZero 库创建的示例文档。")
	doc.AddParagraph("WordZero 提供了简单易用的 API 来创建和操作 Word 文档。")

	// 添加列表内容
	doc.AddParagraph("主要功能：")
	doc.AddParagraph("• 创建新文档")
	doc.AddParagraph("• 添加段落")
	doc.AddParagraph("• 保存文档")
	doc.AddParagraph("• 打开现有文档")

	// 保存文档
	outputFile := "output/example_document.docx"
	err := doc.Save(outputFile)
	if err != nil {
		log.Fatalf("保存文档失败: %v", err)
	}

	fmt.Printf("文档已成功保存到: %s\n", outputFile)

	// 演示打开文档
	fmt.Println("\n正在打开刚创建的文档...")
	openedDoc, err := document.Open(outputFile)
	if err != nil {
		log.Fatalf("打开文档失败: %v", err)
	}

	fmt.Printf("文档包含 %d 个段落\n", len(openedDoc.Body.Paragraphs))

	// 打印所有段落内容
	fmt.Println("\n文档内容：")
	for i, para := range openedDoc.Body.Paragraphs {
		if len(para.Runs) > 0 {
			fmt.Printf("段落 %d: %s\n", i+1, para.Runs[0].Text.Content)
		}
	}
}

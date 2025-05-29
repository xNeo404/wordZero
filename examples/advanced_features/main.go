// Package main 演示WordZero高级功能
package main

import (
	"fmt"
	"log"
	"time"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
)

func main() {
	fmt.Println("WordZero 高级功能演示")
	fmt.Println("================")

	// 创建新文档
	doc := document.New()

	// 1. 设置文档属性
	fmt.Println("1. 设置文档属性...")
	if err := doc.SetTitle("WordZero高级功能演示文档"); err != nil {
		log.Printf("设置标题失败: %v", err)
	}
	if err := doc.SetAuthor("WordZero开发团队"); err != nil {
		log.Printf("设置作者失败: %v", err)
	}
	if err := doc.SetSubject("演示WordZero的高级功能"); err != nil {
		log.Printf("设置主题失败: %v", err)
	}
	if err := doc.SetKeywords("WordZero, Go, 文档处理, 高级功能"); err != nil {
		log.Printf("设置关键字失败: %v", err)
	}
	if err := doc.SetDescription("本文档演示了WordZero库的各种高级功能，包括页眉页脚、列表、目录和脚注等。"); err != nil {
		log.Printf("设置描述失败: %v", err)
	}

	// 2. 设置页眉页脚
	fmt.Println("2. 设置页眉页脚...")
	if err := doc.AddHeader(document.HeaderFooterTypeDefault, "WordZero高级功能演示"); err != nil {
		log.Printf("添加页眉失败: %v", err)
	}
	if err := doc.AddFooterWithPageNumber(document.HeaderFooterTypeDefault, "WordZero开发团队", true); err != nil {
		log.Printf("添加页脚失败: %v", err)
	}

	// 3. 添加文档标题
	fmt.Println("3. 添加文档内容...")
	doc.AddHeadingParagraph("WordZero高级功能演示", 1)
	doc.AddParagraph("本文档演示了WordZero库的各种高级功能，展示如何使用Go语言创建复杂的Word文档。")

	// 4. 添加各级标题和内容（先添加内容，后面再生成目录）
	fmt.Println("4. 添加章节内容...")

	// 第一章
	doc.AddHeadingParagraph("第一章 基础功能", 2)
	doc.AddParagraph("WordZero提供了丰富的基础功能，包括文本格式化、段落设置等。")

	// 添加脚注
	if err := doc.AddFootnote("这是一个脚注示例", "脚注内容：WordZero是一个强大的Go语言Word文档处理库。"); err != nil {
		log.Printf("添加脚注失败: %v", err)
	}

	// 第二章 - 列表功能
	doc.AddHeadingParagraph("第二章 列表功能", 2)
	doc.AddParagraph("WordZero支持多种类型的列表：")

	// 5. 演示列表功能
	fmt.Println("5. 演示列表功能...")

	// 无序列表
	doc.AddHeadingParagraph("2.1 无序列表", 3)
	doc.AddBulletList("项目符号列表项1", 0, document.BulletTypeDot)
	doc.AddBulletList("项目符号列表项2", 0, document.BulletTypeDot)
	doc.AddBulletList("二级项目1", 1, document.BulletTypeCircle)
	doc.AddBulletList("二级项目2", 1, document.BulletTypeCircle)
	doc.AddBulletList("项目符号列表项3", 0, document.BulletTypeDot)

	// 有序列表
	doc.AddHeadingParagraph("2.2 有序列表", 3)
	doc.AddNumberedList("编号列表项1", 0, document.ListTypeDecimal)
	doc.AddNumberedList("编号列表项2", 0, document.ListTypeDecimal)
	doc.AddNumberedList("子项目a", 1, document.ListTypeLowerLetter)
	doc.AddNumberedList("子项目b", 1, document.ListTypeLowerLetter)
	doc.AddNumberedList("编号列表项3", 0, document.ListTypeDecimal)

	// 多级列表
	doc.AddHeadingParagraph("2.3 多级列表", 3)
	multiLevelItems := []document.ListItem{
		{Text: "一级项目1", Level: 0, Type: document.ListTypeDecimal},
		{Text: "二级项目1.1", Level: 1, Type: document.ListTypeLowerLetter},
		{Text: "三级项目1.1.1", Level: 2, Type: document.ListTypeLowerRoman},
		{Text: "三级项目1.1.2", Level: 2, Type: document.ListTypeLowerRoman},
		{Text: "二级项目1.2", Level: 1, Type: document.ListTypeLowerLetter},
		{Text: "一级项目2", Level: 0, Type: document.ListTypeDecimal},
	}
	if err := doc.CreateMultiLevelList(multiLevelItems); err != nil {
		log.Printf("创建多级列表失败: %v", err)
	}

	// 第三章 - 高级格式
	doc.AddHeadingParagraph("第三章 高级格式", 2)
	doc.AddParagraph("WordZero还支持各种高级格式功能。")

	// 添加尾注
	if err := doc.AddEndnote("这是尾注示例", "尾注内容：更多信息请访问WordZero项目主页。"); err != nil {
		log.Printf("添加尾注失败: %v", err)
	}

	// 第四章 - 文档属性
	doc.AddHeadingParagraph("第四章 文档属性管理", 2)
	doc.AddParagraph("WordZero允许设置和管理文档的各种属性，包括标题、作者、创建时间等元数据。")

	// 结论
	doc.AddHeadingParagraph("结论", 2)
	doc.AddParagraph("通过以上演示，我们可以看到WordZero提供了全面的Word文档处理能力，" +
		"包括基础的文本处理、高级的格式设置、以及专业的文档结构功能。")

	// 6. 自动生成目录（新功能！）
	fmt.Println("6. 自动生成目录...")

	// 调试：显示检测到的标题
	headings := doc.ListHeadings()
	fmt.Printf("   检测到 %d 个标题:\n", len(headings))
	for i, heading := range headings {
		fmt.Printf("     %d. 级别%d: %s\n", i+1, heading.Level, heading.Text)
	}

	// 显示标题级别统计
	counts := doc.GetHeadingCount()
	fmt.Printf("   标题级别统计: %+v\n", counts)

	// 使用新的AutoGenerateTOC方法自动生成目录
	tocConfig := document.DefaultTOCConfig()
	tocConfig.Title = "目录"
	tocConfig.MaxLevel = 3

	if err := doc.AutoGenerateTOC(tocConfig); err != nil {
		log.Printf("自动生成目录失败: %v", err)
		fmt.Println("   ❌ 目录生成失败，可能是因为未检测到标题")
	} else {
		fmt.Println("   ✅ 自动生成目录成功")
	}

	// 7. 更新文档统计信息
	fmt.Println("7. 更新文档统计信息...")
	if err := doc.UpdateStatistics(); err != nil {
		log.Printf("更新统计信息失败: %v", err)
	}

	// 8. 保存文档
	fmt.Println("8. 保存文档...")
	outputFile := "examples/output/advanced_features_demo.docx"
	if err := doc.Save(outputFile); err != nil {
		log.Fatalf("保存文档失败: %v", err)
	}

	fmt.Printf("✅ 高级功能演示文档已保存至: %s\n", outputFile)

	// 9. 显示文档统计信息
	fmt.Println("9. 文档统计信息:")
	if properties, err := doc.GetDocumentProperties(); err == nil {
		fmt.Printf("   标题: %s\n", properties.Title)
		fmt.Printf("   作者: %s\n", properties.Creator)
		fmt.Printf("   段落数: %d\n", properties.Paragraphs)
		fmt.Printf("   字数: %d\n", properties.Words)
		fmt.Printf("   字符数: %d\n", properties.Characters)
		fmt.Printf("   创建时间: %s\n", properties.Created.Format(time.RFC3339))
	}

	fmt.Printf("   脚注数量: %d\n", doc.GetFootnoteCount())
	fmt.Printf("   尾注数量: %d\n", doc.GetEndnoteCount())

	fmt.Println("\n🎉 演示完成！")
	fmt.Println("\n📝 新增功能说明:")
	fmt.Println("   - 使用 AutoGenerateTOC() 方法自动检测文档中的标题")
	fmt.Println("   - 支持显示检测到的标题列表和级别统计")
	fmt.Println("   - 自动将目录插入到文档开头")
	fmt.Println("   - 修复了样式ID映射问题，现在能正确识别Heading1-9样式")
}

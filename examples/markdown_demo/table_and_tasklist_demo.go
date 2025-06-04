package main

import (
	"fmt"
	"log"

	"github.com/ZeroHawkeye/wordZero/pkg/markdown"
)

func main() {
	// 示例Markdown内容，包含表格和任务列表
	markdownContent := `# 表格和任务列表示例

## 表格示例

下面是一个简单的表格：

| 姓名   | 年龄 | 城市   |
|--------|------|--------|
| 张三   | 25   | 北京   |
| 李四   | 30   | 上海   |
| 王五   | 28   | 广州   |

## 任务列表示例

待办事项：

- [x] 完成项目需求分析
- [ ] 设计系统架构
- [ ] 实现核心功能
  - [x] 用户管理
  - [ ] 权限控制
  - [ ] 数据存储
- [x] 编写测试用例
- [ ] 部署到生产环境

## 混合内容

这是一个包含**粗体**和*斜体*的段落。

### 对齐表格

| 左对齐 | 居中对齐 | 右对齐 |
|:-------|:--------:|-------:|
| 内容1  |   内容2  |  内容3 |
| 较长内容 | 短内容   |    数字 |
`

	// 创建转换器
	opts := markdown.HighQualityOptions()
	opts.EnableTables = true
	opts.EnableTaskList = true
	converter := markdown.NewConverter(opts)

	// 转换为Word文档
	doc, err := converter.ConvertString(markdownContent, opts)
	if err != nil {
		log.Fatalf("转换失败: %v", err)
	}

	// 保存文档
	outputPath := "examples/output/table_and_tasklist_demo.docx"
	err = doc.Save(outputPath)
	if err != nil {
		log.Fatalf("保存文档失败: %v", err)
	}

	fmt.Printf("✅ 表格和任务列表示例已保存到: %s\n", outputPath)
	fmt.Println("📝 示例包含以下功能:")
	fmt.Println("   • GFM表格转换为Word表格")
	fmt.Println("   • 任务列表复选框显示")
	fmt.Println("   • 表格对齐方式保持")
	fmt.Println("   • 混合格式文本支持")
}

# WordZero - Golang Word操作库

[![Go Version](https://img.shields.io/badge/Go-1.19+-00ADD8?style=flat&logo=go)](https://golang.org)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Tests](https://img.shields.io/badge/Tests-Passing-green.svg)](#测试)
[![Benchmark](https://img.shields.io/badge/Benchmark-Go%202.62ms%20%7C%20JS%209.63ms%20%7C%20Python%2055.98ms-success.svg)](https://github.com/ZeroHawkeye/wordZero/wiki/13-%E6%80%A7%E8%83%BD%E5%9F%BA%E5%87%86%E6%B5%8B%E8%AF%95)
[![Performance](https://img.shields.io/badge/Performance-Golang%20优胜-brightgreen.svg)](https://github.com/ZeroHawkeye/wordZero/wiki/13-%E6%80%A7%E8%83%BD%E5%9F%BA%E5%87%86%E6%B5%8B%E8%AF%95)
[![Ask DeepWiki](https://deepwiki.com/badge.svg)](https://deepwiki.com/ZeroHawkeye/wordZero)

## 项目介绍

WordZero 是一个使用 Golang 实现的 Word 文档操作库，提供基础的文档创建、修改等操作功能。该库遵循最新的 Office Open XML (OOXML) 规范，专注于现代 Word 文档格式（.docx）的支持。

### 核心特性

- 🚀 **完整的文档操作**: 创建、读取、修改 Word 文档
- 🎨 **丰富的样式系统**: 18种预定义样式，支持自定义样式和样式继承
- 📝 **文本格式化**: 字体、大小、颜色、粗体、斜体等完整支持
- 📐 **段落格式**: 对齐、间距、缩进等段落属性设置
- 🏷️ **标题导航**: 完整支持Heading1-9样式，可被Word导航窗格识别
- 📊 **表格功能**: 完整的表格创建、编辑、样式设置和迭代器支持
- 📄 **页面设置**: 页面尺寸、边距、页眉页脚等专业排版功能
- 🔧 **高级功能**: 目录生成、脚注尾注、列表编号、模板引擎等
- 🎯 **模板继承**: 支持基础模板和块重写机制，实现模板复用和扩展
- ⚡ **卓越性能**: 零依赖的纯Go实现，平均2.62ms处理速度，比JavaScript快3.7倍，比Python快21倍
- 🔧 **易于使用**: 简洁的API设计，链式调用支持

## 安装

```bash
go get github.com/ZeroHawkeye/wordZero
```

### 版本说明

推荐使用带版本号的安装方式：

```bash
# 安装最新版本
go get github.com/ZeroHawkeye/wordZero@latest

# 安装指定版本
go get github.com/ZeroHawkeye/wordZero@v0.4.0
```

## 快速开始

```go
package main

import (
    "log"
    "github.com/ZeroHawkeye/wordZero/pkg/document"
    "github.com/ZeroHawkeye/wordZero/pkg/style"
)

func main() {
    // 创建新文档
    doc := document.New()
    
    // 添加标题
    titlePara := doc.AddParagraph("WordZero 使用示例")
    titlePara.SetStyle(style.StyleHeading1)
    
    // 添加正文段落
    para := doc.AddParagraph("这是一个使用 WordZero 创建的文档示例。")
    para.SetFontFamily("宋体")
    para.SetFontSize(12)
    para.SetColor("333333")
    
    // 创建表格
    tableConfig := &document.TableConfig{
        Rows:    3,
        Columns: 3,
    }
    table := doc.AddTable(tableConfig)
    table.SetCellText(0, 0, "表头1")
    table.SetCellText(0, 1, "表头2")
    table.SetCellText(0, 2, "表头3")
    
    // 保存文档
    if err := doc.Save("example.docx"); err != nil {
        log.Fatal(err)
    }
}
```

### 模板继承功能示例

```go
// 创建基础模板
engine := document.NewTemplateEngine()
baseTemplate := `{{companyName}} 工作报告

{{#block "summary"}}
默认摘要内容
{{/block}}

{{#block "content"}}
默认主要内容
{{/block}}`

engine.LoadTemplate("base_report", baseTemplate)

// 创建扩展模板，重写特定块
salesTemplate := `{{extends "base_report"}}

{{#block "summary"}}
销售业绩摘要：本月达成 {{achievement}}%
{{/block}}

{{#block "content"}}
销售详情：
- 总销售额：{{totalSales}}
- 新增客户：{{newCustomers}}
{{/block}}`

engine.LoadTemplate("sales_report", salesTemplate)

// 渲染模板
data := document.NewTemplateData()
data.SetVariable("companyName", "WordZero科技")
data.SetVariable("achievement", "125")
data.SetVariable("totalSales", "1,850,000")
data.SetVariable("newCustomers", "45")

doc, _ := engine.RenderToDocument("sales_report", data)
doc.Save("sales_report.docx")
```

## 文档和示例

### 📚 完整文档
- [**📖 Wiki文档**](https://github.com/ZeroHawkeye/wordZero/wiki) - 完整的使用文档和API参考
- [**🚀 快速开始**](https://github.com/ZeroHawkeye/wordZero/wiki/01-快速开始) - 新手入门指南
- [**⚡ 功能特性详览**](https://github.com/ZeroHawkeye/wordZero/wiki/14-功能特性详览) - 所有功能的详细说明
- [**📊 性能基准测试**](https://github.com/ZeroHawkeye/wordZero/wiki/13-性能基准测试) - 跨语言性能对比分析
- [**🏗️ 项目结构详解**](https://github.com/ZeroHawkeye/wordZero/wiki/15-项目结构详解) - 项目架构和代码组织

### 💡 使用示例
查看 `examples/` 目录下的示例代码：

- `examples/basic/` - 基础功能演示
- `examples/style_demo/` - 样式系统演示  
- `examples/table/` - 表格功能演示
- `examples/formatting/` - 格式化演示
- `examples/page_settings/` - 页面设置演示
- `examples/advanced_features/` - 高级功能综合演示
- `examples/template_demo/` - 模板功能演示
- `examples/template_inheritance_demo/` - 模板继承功能演示 ✨ **新增**

运行示例：
```bash
# 运行基础功能演示
go run ./examples/basic/

# 运行样式演示
go run ./examples/style_demo/

# 运行表格演示
go run ./examples/table/

# 运行模板继承演示
go run ./examples/template_inheritance_demo/
```

## 主要功能

### ✅ 已实现功能
- **文档操作**: 创建、读取、保存、解析DOCX文档
- **文本格式化**: 字体、大小、颜色、粗体、斜体等
- **样式系统**: 18种预定义样式 + 自定义样式支持
- **段落格式**: 对齐、间距、缩进等完整支持
- **表格功能**: 完整的表格操作、样式设置、单元格迭代器
- **页面设置**: 页面尺寸、边距、页眉页脚等
- **高级功能**: 目录生成、脚注尾注、列表编号、模板引擎（含模板继承）
- **图片功能**: 图片插入、大小调整、位置设置

### 🚧 规划中功能
- 表格排序和高级操作
- 书签和交叉引用
- 文档批注和修订
- 图形绘制功能
- 多语言和国际化支持

👉 **查看完整功能列表**: [功能特性详览](https://github.com/ZeroHawkeye/wordZero/wiki/14-功能特性详览)

## 性能表现

WordZero 在性能方面表现卓越，通过完整的基准测试验证：

| 语言 | 平均执行时间 | 相对性能 |
|------|-------------|----------|
| **Golang** | **2.62ms** | **1.00×** |
| JavaScript | 9.63ms | 3.67× |
| Python | 55.98ms | 21.37× |

👉 **查看详细性能分析**: [性能基准测试](https://github.com/ZeroHawkeye/wordZero/wiki/13-性能基准测试)

## 项目结构

```
wordZero/
├── pkg/                    # 核心库代码
│   ├── document/          # 文档操作功能
│   └── style/             # 样式管理系统
├── examples/              # 使用示例
├── test/                  # 集成测试
├── benchmark/             # 性能基准测试
└── wordZero.wiki/         # 完整文档
```

👉 **查看详细结构说明**: [项目结构详解](https://github.com/ZeroHawkeye/wordZero/wiki/15-项目结构详解)

## 贡献指南

欢迎提交 Issue 和 Pull Request！在提交代码前请确保：

1. 代码符合 Go 代码规范
2. 添加必要的测试用例
3. 更新相关文档
4. 确保所有测试通过

## 许可证

本项目采用 MIT 许可证。详见 [LICENSE](LICENSE) 文件。

---

**更多资源**:
- 📖 [完整文档](https://github.com/ZeroHawkeye/wordZero/wiki)
- 🔧 [API参考](https://github.com/ZeroHawkeye/wordZero/wiki/10-API参考)
- 💡 [最佳实践](https://github.com/ZeroHawkeye/wordZero/wiki/09-最佳实践)
- 📝 [更新日志](CHANGELOG.md)
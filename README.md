# WordZero - Golang Word操作库

## 项目介绍

WordZero 是一个使用 Golang 实现的 Word 文档操作库，提供基础的文档创建、修改等操作功能。该库遵循最新的 Office Open XML (OOXML) 规范，专注于现代 Word 文档格式（.docx）的支持。

## 功能特性

- 创建新的 Word 文档
- 读取和解析现有文档
- 文本内容的添加和修改
- 段落格式化
- 表格操作
- 图片插入
- 样式管理

## 项目结构

```
wordZero/
├── cmd/                    # 命令行工具
├── pkg/                    # 公共包
│   ├── document/          # 文档核心操作
│   ├── paragraph/         # 段落处理
│   ├── table/            # 表格处理
│   ├── image/            # 图片处理
│   └── style/            # 样式管理
├── internal/              # 内部包
│   ├── xml/              # XML处理
│   └── zip/              # ZIP文件处理
├── examples/              # 使用示例
├── test/                  # 测试文件
├── go.mod
├── go.sum
└── README.md
```

## 待办任务列表

### 基础架构 
- [x] 创建项目基础目录结构
- [x] 初始化 go.mod 依赖管理
- [x] 设置基础的错误处理机制
- [x] 实现日志系统

### 核心功能
- [x] 实现 OOXML 基础结构解析
- [x] 实现 .docx 文件的 ZIP 解压和压缩
- [x] 创建 Document 核心结构体
- [x] 实现文档创建功能
- [x] 实现文档打开和保存功能

### 文本操作
- [x] 实现段落创建和管理
- [x] 实现文本添加功能
- [x] 实现文本格式化（字体、大小、颜色等）
- [x] 实现文本对齐功能
- [x] 实现行间距和段间距设置

### 表格功能
- [ ] 实现表格创建
- [ ] 实现单元格合并
- [ ] 实现表格样式设置
- [ ] 实现表格边框设置

### 图片功能
- [ ] 实现图片插入
- [ ] 实现图片大小调整
- [ ] 实现图片位置设置
- [ ] 支持多种图片格式（JPG、PNG、GIF）

### 样式管理
- [ ] 实现预定义样式
- [ ] 实现自定义样式创建
- [ ] 实现样式应用和修改

### 高级功能
- [ ] 实现页眉页脚
- [ ] 实现目录生成
- [ ] 实现页码设置
- [ ] 实现文档属性设置（作者、标题等）

### 测试和文档
- [x] 编写单元测试
- [ ] 编写集成测试
- [x] 编写使用示例
- [ ] 编写 API 文档

## 快速开始

```go
package main

import (
    "log"
    "github.com/ZeroHawkeye/wordZero/pkg/document"
)

func main() {
    // 创建新文档
    doc := document.New()
    
    // 添加段落
    doc.AddParagraph("Hello, World!")
    doc.AddParagraph("这是使用 WordZero 创建的文档。")
    
    // 保存文档
    err := doc.Save("output.docx")
    if err != nil {
        log.Fatal(err)
    }
}
```

## 安装

```bash
go get github.com/ZeroHawkeye/wordZero
```

## 使用示例

### 创建文档

```go
doc := document.New()
doc.AddParagraph("标题")
doc.AddParagraph("正文内容")
doc.Save("document.docx")
```

### 文本格式化

```go
// 创建格式化文档
doc := document.New()

// 添加格式化标题
titleFormat := &document.TextFormat{
    Bold:      true,
    FontSize:  18,
    FontColor: "FF0000", // 红色
    FontName:  "微软雅黑",
}
title := doc.AddFormattedParagraph("这是标题", titleFormat)
title.SetAlignment(document.AlignCenter) // 居中对齐

// 添加带间距的段落
para := doc.AddParagraph("这个段落有特定的间距设置")
spacingConfig := &document.SpacingConfig{
    LineSpacing:     1.5, // 1.5倍行距
    BeforePara:      12,  // 段前12磅
    AfterPara:       6,   // 段后6磅
    FirstLineIndent: 24,  // 首行缩进24磅
}
para.SetSpacing(spacingConfig)
para.SetAlignment(document.AlignJustify) // 两端对齐

// 添加混合格式的段落
mixed := doc.AddParagraph("这个段落包含多种格式：")
mixed.AddFormattedText("粗体蓝色", &document.TextFormat{
    Bold: true, FontColor: "0000FF"})
mixed.AddFormattedText("，普通文本，", nil)
mixed.AddFormattedText("斜体绿色", &document.TextFormat{
    Italic: true, FontColor: "00FF00"})

doc.Save("formatted.docx")
```

### 打开文档

```go
doc, err := document.Open("existing.docx")
if err != nil {
    log.Fatal(err)
}

// 读取段落内容
for _, para := range doc.Body.Paragraphs {
    if len(para.Runs) > 0 {
        fmt.Println(para.Runs[0].Text.Content)
    }
}
```

### 命令行使用

```bash
# 创建文档
go run main.go -action=create -file=output.docx -text="Hello World"

# 打开并读取文档
go run main.go -action=open -file=output.docx
```

## 贡献指南

欢迎提交 Issue 和 Pull Request！

## 许可证

MIT License 
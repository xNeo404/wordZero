# 段落格式自定义功能完善 - 技术总结

## 问题描述

用户反馈："为啥总有一些功能函数用不了，段落格式那里不能自定义啊"

经过分析，发现WordZero库虽然已经实现了基础的段落格式功能（对齐、间距、缩进、边框），但缺少一些Word文档中常用的高级段落属性设置，导致用户无法完全自定义段落格式。

## 解决方案

### 1. 新增段落属性结构

在`ParagraphProperties`中添加了以下高级属性：

```go
type ParagraphProperties struct {
    // ... 现有字段
    KeepNext        *KeepNext        // 与下一段落保持在一起
    KeepLines       *KeepLines       // 段落中的行保持在一起
    PageBreakBefore *PageBreakBefore // 段前分页
    WidowControl    *WidowControl    // 孤行控制
    OutlineLevel    *OutlineLevel    // 大纲级别
}
```

### 2. 实现的新方法

#### 单独设置方法

```go
// SetKeepWithNext 设置与下一段落保持在同一页
func (p *Paragraph) SetKeepWithNext(keep bool)

// SetKeepLines 设置段落所有行保持在同一页
func (p *Paragraph) SetKeepLines(keep bool)

// SetPageBreakBefore 设置段前分页
func (p *Paragraph) SetPageBreakBefore(pageBreak bool)

// SetWidowControl 设置孤行控制
func (p *Paragraph) SetWidowControl(control bool)

// SetOutlineLevel 设置大纲级别（0-8）
func (p *Paragraph) SetOutlineLevel(level int)
```

#### 综合配置方法

```go
// ParagraphFormatConfig 段落格式配置
type ParagraphFormatConfig struct {
    // 基础格式
    Alignment AlignmentType
    Style     string
    
    // 间距设置
    LineSpacing     float64
    BeforePara      int
    AfterPara       int
    FirstLineIndent int
    
    // 缩进设置
    FirstLineCm float64
    LeftCm      float64
    RightCm     float64
    
    // 分页与控制
    KeepWithNext    bool
    KeepLines       bool
    PageBreakBefore bool
    WidowControl    bool
    
    // 大纲级别
    OutlineLevel int
}

// SetParagraphFormat 一次性设置所有段落格式属性
func (p *Paragraph) SetParagraphFormat(config *ParagraphFormatConfig)
```

### 3. 使用示例

#### 方法1：逐个设置属性

```go
title := doc.AddParagraph("第一章 概述")
title.SetAlignment(document.AlignCenter)
title.SetKeepWithNext(true)
title.SetPageBreakBefore(true)
title.SetOutlineLevel(0)
```

#### 方法2：使用配置结构一次性设置

```go
para := doc.AddParagraph("重要内容")
para.SetParagraphFormat(&document.ParagraphFormatConfig{
    Alignment:       document.AlignJustify,
    Style:           "Normal",
    LineSpacing:     1.5,
    BeforePara:      12,
    AfterPara:       6,
    FirstLineCm:     0.5,
    KeepWithNext:    true,
    KeepLines:       true,
    WidowControl:    true,
    OutlineLevel:    0,
})
```

## 技术细节

### 1. OpenXML 兼容性

所有新增的属性都严格遵循Office Open XML (OOXML)标准：
- `w:keepNext` - 与下一段保持
- `w:keepLines` - 行保持
- `w:pageBreakBefore` - 段前分页
- `w:widowControl` - 孤行控制
- `w:outlineLvl` - 大纲级别

### 2. 类型复用

为避免重复定义，删除了`table.go`中的`KeepNext`和`KeepLines`类型定义，改为在`document.go`中统一定义，供段落和表格共同使用。

### 3. 边界值处理

- `OutlineLevel`设置时会自动验证范围（0-8），超出范围会被调整并记录警告日志
- 所有布尔类型属性使用"1"/"0"字符串表示，符合OpenXML规范

### 4. XML序列化

使用`omitempty`标签确保：
- 未设置的属性不会出现在XML中
- 生成的文档更简洁
- 符合Word默认行为

## 测试覆盖

### 单元测试

创建了8个新的测试用例：

1. `TestParagraphSetKeepWithNext` - 测试与下段保持功能
2. `TestParagraphSetKeepLines` - 测试行保持功能
3. `TestParagraphSetPageBreakBefore` - 测试段前分页
4. `TestParagraphSetWidowControl` - 测试孤行控制
5. `TestParagraphSetOutlineLevel` - 测试大纲级别设置
6. `TestParagraphSetParagraphFormat` - 测试综合格式配置
7. `TestParagraphSetParagraphFormatNil` - 测试nil配置处理
8. `TestParagraphSetParagraphFormatPartial` - 测试部分配置

### 测试结果

```
=== RUN   TestParagraphSetKeepWithNext
--- PASS: TestParagraphSetKeepWithNext (0.00s)
=== RUN   TestParagraphSetKeepLines
--- PASS: TestParagraphSetKeepLines (0.00s)
=== RUN   TestParagraphSetPageBreakBefore
--- PASS: TestParagraphSetPageBreakBefore (0.00s)
=== RUN   TestParagraphSetWidowControl
--- PASS: TestParagraphSetWidowControl (0.00s)
=== RUN   TestParagraphSetOutlineLevel
--- PASS: TestParagraphSetOutlineLevel (0.00s)
=== RUN   TestParagraphSetParagraphFormat
--- PASS: TestParagraphSetParagraphFormat (0.00s)
=== RUN   TestParagraphSetParagraphFormatNil
--- PASS: TestParagraphSetParagraphFormatNil (0.00s)
=== RUN   TestParagraphSetParagraphFormatPartial
--- PASS: TestParagraphSetParagraphFormatPartial (0.00s)
PASS
ok      github.com/ZeroHawkeye/wordZero/pkg/document    0.003s
```

所有现有测试也继续通过，确保了向后兼容性。

## 示例程序

创建了`examples/paragraph_format_demo`，演示所有新功能：

1. 章节标题格式（居中、段前分页、与下段保持、大纲级别）
2. 正文段落格式（两端对齐、行间距、首行缩进、孤行控制）
3. 二级标题格式（段前段后间距、与下段保持、大纲级别）
4. 悬挂缩进列表（负首行缩进和左缩进组合）
5. 引用块格式（左右缩进、行保持、居中对齐）
6. 多级标题（不同大纲级别设置）

## 文档更新

### API文档

更新了`pkg/document/README.md`：
- 添加了所有新方法的详细说明
- 提供了使用示例和应用场景
- 说明了各个功能的适用情况

### 用户文档

更新了`README_zh.md`：
- 在核心特性中更新了段落格式描述
- 在已实现功能列表中更新了段落格式说明
- 添加了新示例到示例列表
- 添加了运行新示例的命令

### 示例文档

创建了`examples/paragraph_format_demo/README.md`：
- 详细说明了示例的功能
- 提供了运行方法
- 展示了两种API使用方式
- 说明了应用场景

## 质量保证

### 代码审查

- 通过代码审查，修复了`Justification`结构的XML标签问题
- 所有方法都有完整的GoDoc注释
- 符合Go代码规范

### 安全扫描

- 使用CodeQL进行安全扫描
- 结果：0个安全漏洞
- 代码安全可靠

### 兼容性

- 所有现有测试通过
- 不影响现有功能
- 向后兼容

## 应用场景

### 1. 文档结构化

使用大纲级别功能创建清晰的文档层次结构，便于Word导航窗格显示。

```go
chapter := doc.AddParagraph("第一章")
chapter.SetOutlineLevel(0)

section := doc.AddParagraph("1.1 小节")
section.SetOutlineLevel(1)
```

### 2. 专业排版

使用分页控制确保标题和内容的关联性，避免标题孤零零地出现在页面底部。

```go
title := doc.AddParagraph("重要章节")
title.SetKeepWithNext(true)  // 确保标题和下一段在同一页
```

### 3. 内容保护

使用行保持功能防止重要段落被分页，保持内容完整性。

```go
important := doc.AddParagraph("这是一段重要内容，需要保持完整")
important.SetKeepLines(true)  // 所有行保持在同一页
```

### 4. 章节管理

使用段前分页实现章节的页面独立性。

```go
newChapter := doc.AddParagraph("第二章")
newChapter.SetPageBreakBefore(true)  // 从新页开始
```

## 总结

本次更新完善了WordZero的段落格式自定义功能，提供了与Microsoft Word相同的高级段落控制能力。通过提供单独设置方法和综合配置方法，满足了不同用户的使用习惯。完整的测试覆盖、详细的文档和示例程序确保了功能的可用性和可维护性。

## 改进文件清单

1. **核心代码**
   - `pkg/document/document.go` - 新增类型定义和方法实现
   - `pkg/document/table.go` - 删除重复类型定义

2. **测试代码**
   - `pkg/document/document_test.go` - 新增8个测试用例

3. **示例代码**
   - `examples/paragraph_format_demo/main.go` - 新增示例程序
   - `examples/paragraph_format_demo/README.md` - 示例文档

4. **文档**
   - `pkg/document/README.md` - 更新API文档
   - `README_zh.md` - 更新功能列表和示例列表

## 提交记录

1. `Add comprehensive paragraph formatting features` - 核心功能实现
2. `Update documentation for paragraph formatting features` - 文档更新
3. `Fix Justification struct to add omitempty tag` - 代码审查修复

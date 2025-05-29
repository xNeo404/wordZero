# WordZero 更新日志

## [v1.3.0] - 2025-01-18

### ✨ 重大功能新增

#### 页眉页脚功能 ✨ **全新实现**
- **完整的页眉页脚支持**: 实现了页眉、页脚的创建和管理
- **多种页眉页脚类型**: 
  - `Default` - 默认页眉页脚
  - `First` - 首页页眉页脚  
  - `Even` - 偶数页页眉页脚
- **页码显示功能**: 支持在页眉页脚中显示页码
- **关键API**:
  - `AddHeader()` - 添加页眉
  - `AddFooter()` - 添加页脚
  - `AddHeaderWithPageNumber()` - 添加带页码的页眉
  - `AddFooterWithPageNumber()` - 添加带页码的页脚
  - `SetDifferentFirstPage()` - 设置首页不同

#### 目录生成功能 ✨ **全新实现**
- **自动目录生成**: 基于标题样式自动创建目录
- **多级目录支持**: 支持1-9级标题的目录条目
- **目录配置选项**:
  - 目录标题自定义
  - 页码显示控制
  - 超链接支持
  - 点状引导线
- **书签集成**: 标题自动生成书签，支持导航
- **关键API**:
  - `GenerateTOC()` - 生成目录
  - `UpdateTOC()` - 更新目录
  - `AddHeadingWithBookmark()` - 添加带书签的标题
  - `AutoGenerateTOC()` - 自动生成目录
  - `GetHeadingCount()` - 获取标题统计

#### 脚注和尾注功能 ✨ **全新实现**
- **脚注管理**: 完整的脚注添加、删除和配置功能
- **尾注支持**: 文档末尾的尾注功能
- **多种编号格式**:
  - 十进制数字 (`decimal`)
  - 小写/大写罗马数字 (`lowerRoman`, `upperRoman`)
  - 小写/大写字母 (`lowerLetter`, `upperLetter`)
  - 符号编号 (`symbol`)
- **脚注位置控制**:
  - 页面底部 (`pageBottom`)
  - 文本下方 (`beneathText`)
  - 节末尾 (`sectEnd`)
  - 文档末尾 (`docEnd`)
- **编号重启规则**:
  - 连续编号 (`continuous`)
  - 每节重启 (`eachSect`)
  - 每页重启 (`eachPage`)
- **关键API**:
  - `AddFootnote()` - 添加脚注
  - `AddEndnote()` - 添加尾注
  - `SetFootnoteConfig()` - 设置脚注配置
  - `GetFootnoteCount()`, `GetEndnoteCount()` - 获取数量统计

#### 列表和编号功能 ✨ **全新实现**
- **无序列表**: 支持多种项目符号
  - 圆点符号 (`•`)
  - 空心圆 (`○`)
  - 方块 (`■`)
  - 短横线 (`–`)
  - 箭头 (`→`)
- **有序列表**: 支持多种编号格式
  - 十进制数字 (`decimal`)
  - 小写/大写字母 (`lowerLetter`, `upperLetter`)
  - 小写/大写罗马数字 (`lowerRoman`, `upperRoman`)
- **多级列表**: 支持最多9级嵌套
- **编号控制**: 支持重新开始编号
- **关键API**:
  - `AddListItem()` - 添加列表项
  - `AddBulletList()` - 添加无序列表
  - `AddNumberedList()` - 添加有序列表
  - `CreateMultiLevelList()` - 创建多级列表
  - `RestartNumbering()` - 重启编号

#### 结构化文档标签（SDT） ✨ **全新实现**
- **目录SDT结构**: 专门用于目录功能的SDT实现
- **SDT属性管理**: 完整的SDT属性和内容控制
- **文档部件支持**: 支持SDT占位符和文档部件
- **关键API**:
  - `CreateTOCSDT()` - 创建目录SDT结构

#### 域字段功能 ✨ **全新实现**
- **超链接域**: 支持文档内部超链接
- **页码引用域**: 支持页码引用和导航
- **域字符控制**: 完整的域开始、分隔、结束标记
- **关键API**:
  - `CreateHyperlinkField()` - 创建超链接域
  - `CreatePageRefField()` - 创建页码引用域

### 🏗️ 架构改进

#### 新增核心文件
- `header_footer.go` - 页眉页脚功能实现
- `toc.go` - 目录生成功能实现
- `footnotes.go` - 脚注尾注功能实现
- `numbering.go` - 列表编号功能实现
- `sdt.go` - 结构化文档标签实现
- `field.go` - 域字段功能实现

#### 文档完善
- 新增 `pkg/document/README.md` 详细API文档更新
- 增加了所有新功能的使用示例和配置说明
- 新增配置结构体文档说明

### 📚 示例程序

#### 新增示例目录
- `examples/page_settings/` - 页面设置演示
- `examples/advanced_features/` - 高级功能综合演示
  - 页眉页脚演示
  - 目录生成演示
  - 脚注尾注演示
  - 列表编号演示

### 🔧 配置结构体

#### 新增配置类型
- `TOCConfig` - 目录配置
- `FootnoteConfig` - 脚注配置
- `ListConfig` - 列表配置
- `HeaderFooterType` - 页眉页脚类型枚举
- `FootnoteNumberFormat` - 脚注编号格式枚举
- `ListType` - 列表类型枚举
- `BulletType` - 项目符号类型枚举

### 📝 使用示例更新

```go
// 页眉页脚示例
doc.AddHeader(document.HeaderFooterTypeDefault, "这是页眉")
doc.AddFooterWithPageNumber(document.HeaderFooterTypeDefault, "第", true)
doc.SetDifferentFirstPage(true)

// 目录示例
doc.AddHeadingWithBookmark("第一章 概述", 1, "chapter1")
tocConfig := document.DefaultTOCConfig()
doc.GenerateTOC(tocConfig)

// 脚注示例
doc.AddFootnote("正文内容", "脚注内容")
doc.AddEndnote("更多说明", "尾注内容")

// 列表示例
doc.AddBulletList("列表项1", 0, document.BulletTypeDot)
doc.AddNumberedList("第一项", 0, document.ListTypeDecimal)
```

### 🎯 兼容性保证

- ✅ **API向下兼容**: 所有现有API保持不变
- ✅ **无破坏性变更**: 现有代码无需修改
- ✅ **渐进增强**: 新功能作为可选功能提供

### 🔍 技术改进

#### 功能模块化
- 每个新功能独立文件实现，降低代码耦合
- 统一的错误处理和日志记录
- 符合Word OOXML标准的实现

#### 代码质量
- 完整的单元测试覆盖
- 详细的API文档和注释
- 规范的Go代码风格

---

## [v1.2.0] - 2025-05-29

### ✨ 新增功能

#### 表格默认样式改进
- **表格默认边框样式**: 新创建的表格现在默认包含单线边框样式，无需手动设置
- **参考标准格式**: 默认样式参考了 Word 标准表格格式（tmp_test 目录中的参考实现）
- **详细规格**:
  - 边框样式：`single`（单线）
  - 边框粗细：`4`（1/8磅单位）
  - 边框颜色：`auto`（自动）
  - 边框间距：`0`
  - 表格布局：`autofit`（自动调整）
  - 单元格边距：左右各 `108 dxa`

#### 功能特性
- ✅ **向下兼容**: 现有代码无需修改，自动享受新的默认样式
- ✅ **样式覆盖**: 仍然支持通过 `SetTableBorders()` 等方法自定义样式
- ✅ **无边框选项**: 可通过 `RemoveTableBorders()` 方法回到原来的无边框效果
- ✅ **标准匹配**: 与 Microsoft Word 创建的表格样式保持一致

### 🔧 改进内容

#### 代码改进
- 修改 `CreateTable()` 函数，在表格属性中增加默认边框配置
- 添加表格布局和单元格边距的默认设置
- 保持原有 API 接口不变，确保兼容性

#### 测试完善
- 新增 `TestTableDefaultStyle` 测试，验证默认样式正确应用
- 新增 `TestDefaultStyleMatchesTmpTest` 测试，确保与参考格式匹配
- 新增 `TestDefaultStyleOverride` 测试，验证样式覆盖功能

#### 示例程序
- 新增 `examples/table_default_style/` 演示程序
- 展示新默认样式、原无边框效果对比、自定义样式覆盖等功能

### 📝 文档更新

#### README.md
- 更新表格功能说明，增加默认样式特性描述
- 标注新增功能和改进点

#### pkg/document/README.md
- 更新 `CreateTable` 方法说明，增加默认样式信息

### 🎯 影响范围

#### 用户体验改进
- **即开即用**: 新创建的表格具有专业的外观，无需额外设置
- **标准化**: 确保表格样式与 Word 标准一致
- **灵活性**: 保持完整的自定义能力

#### 开发者友好
- **API 稳定**: 无破坏性变更，现有代码继续工作
- **渐进增强**: 新功能作为默认行为提供，不影响现有逻辑

### 🔍 技术细节

#### 参考实现
基于 `tmp_test/word/document.xml` 中的表格定义：
```xml
<w:tblBorders>
  <w:top w:val="single" w:color="auto" w:sz="4" w:space="0"/>
  <w:left w:val="single" w:color="auto" w:sz="4" w:space="0"/>
  <w:bottom w:val="single" w:color="auto" w:sz="4" w:space="0"/>
  <w:right w:val="single" w:color="auto" w:sz="4" w:space="0"/>
  <w:insideH w:val="single" w:color="auto" w:sz="4" w:space="0"/>
  <w:insideV w:val="single" w:color="auto" w:sz="4" w:space="0"/>
</w:tblBorders>
```

#### 实现位置
- 文件：`pkg/document/table.go`
- 函数：`CreateTable()`
- 影响：所有通过 `AddTable()` 创建的新表格

---

## [v1.1.0] - 2025-05-28

### 🎨 表格样式系统
- 完整的表格边框设置功能
- 表格和单元格背景颜色支持
- 多种边框样式（单线、双线、虚线、点线等）
- 奇偶行颜色交替功能

### 📐 表格布局功能
- 表格尺寸控制（宽度、高度、列宽）
- 表格对齐和定位
- 单元格合并功能
- 行高设置和分页控制

### 🎯 样式管理系统
- 18种预定义样式支持
- 样式继承机制
- 自定义样式创建
- 样式查询和批量操作 API

---

## [v1.0.0] - 2025-05-27

### 🚀 初始版本
- 基础文档创建和操作功能
- 文本格式化支持
- 段落管理和样式设置
- 基础表格创建和单元格操作 
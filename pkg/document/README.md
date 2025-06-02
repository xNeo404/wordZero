# Document 包 API 文档

本文档记录了 `pkg/document` 包中所有可用的公开方法和功能。

## 核心类型

### Document 文档
- [Document](document.go) - Word文档的核心结构
- [Body](document.go) - 文档主体
- [Paragraph](document.go) - 段落结构
- [Table](table.go) - 表格结构

## 文档操作方法

### 文档创建与加载
- [`New()`](document.go#L232) - 创建新的Word文档
- [`Open(filename string)`](document.go#L269) - 打开现有Word文档 ✨ **重大改进**
  
#### 文档解析功能重大升级 ✨
`Open` 方法现在支持完整的文档结构解析，包括：

**动态元素解析支持**：
- **段落解析** (`<w:p>`): 完整解析段落内容、属性、运行和格式
- **表格解析** (`<w:tbl>`): 支持表格结构、网格、行列、单元格内容
- **节属性解析** (`<w:sectPr>`): 页面设置、边距、分栏等属性
- **扩展性设计**: 新的解析架构可轻松添加更多元素类型

**解析器特性**：
- **流式解析**: 使用XML流式解析器，内存效率高，适用于大型文档
- **结构保持**: 完整保留文档元素的原始顺序和层次结构
- **错误恢复**: 智能跳过未知或损坏的元素，确保解析过程稳定
- **深度解析**: 支持嵌套结构（如表格中的段落、段落中的运行等）

**解析的内容包括**：
- 段落文本内容和所有格式属性（字体、大小、颜色、样式等）
- 表格完整结构（行列定义、单元格内容、表格属性）
- 页面设置信息（页面尺寸、方向、边距等）
- 样式引用和属性继承关系

### 文档保存与导出
- [`Save(filename string)`](document.go#L337) - 保存文档到文件
- [`ToBytes()`](document.go#L1107) - 将文档转换为字节数组

### 文档内容操作
- [`AddParagraph(text string)`](document.go#L420) - 添加简单段落
- [`AddFormattedParagraph(text string, format *TextFormat)`](document.go#L459) - 添加格式化段落
- [`AddHeadingParagraph(text string, level int)`](document.go#L682) - 添加标题段落
- [`AddHeadingParagraphWithBookmark(text string, level int, bookmarkName string)`](document.go#L747) - 添加带书签的标题段落 ✨ **新增功能**

#### 标题段落书签功能 ✨
`AddHeadingParagraphWithBookmark` 方法现在支持为标题段落添加书签：

**书签功能特性**：
- **自动书签生成**: 为标题段落创建唯一的书签标识
- **灵活命名**: 支持自定义书签名称或留空不添加书签
- **目录兼容**: 生成的书签与目录功能完美兼容，支持导航和超链接
- **Word标准**: 符合Microsoft Word的书签格式规范

**书签生成规则**：
- 书签ID自动生成为 `bookmark_{元素索引}_{书签名称}` 格式
- 书签开始标记插入在段落之前
- 书签结束标记插入在段落之后
- 支持空书签名称以跳过书签创建

### 样式管理
- [`GetStyleManager()`](document.go#L791) - 获取样式管理器

### 页面设置 ✨ 新增功能
- [`SetPageSettings(settings *PageSettings)`](page.go) - 设置完整页面属性
- [`GetPageSettings()`](page.go) - 获取当前页面设置
- [`SetPageSize(size PageSize)`](page.go) - 设置页面尺寸
- [`SetCustomPageSize(width, height float64)`](page.go) - 设置自定义页面尺寸（毫米）
- [`SetPageOrientation(orientation PageOrientation)`](page.go) - 设置页面方向
- [`SetPageMargins(top, right, bottom, left float64)`](page.go) - 设置页面边距（毫米）
- [`SetHeaderFooterDistance(header, footer float64)`](page.go) - 设置页眉页脚距离（毫米）
- [`SetGutterWidth(width float64)`](page.go) - 设置装订线宽度（毫米）
- [`DefaultPageSettings()`](page.go) - 获取默认页面设置（A4纵向）

### 页眉页脚操作 ✨ 新增功能
- [`AddHeader(headerType HeaderFooterType, text string)`](header_footer.go) - 添加页眉
- [`AddFooter(footerType HeaderFooterType, text string)`](header_footer.go) - 添加页脚
- [`AddHeaderWithPageNumber(headerType HeaderFooterType, text string, showPageNum bool)`](header_footer.go) - 添加带页码的页眉
- [`AddFooterWithPageNumber(footerType HeaderFooterType, text string, showPageNum bool)`](header_footer.go) - 添加带页码的页脚
- [`SetDifferentFirstPage(different bool)`](header_footer.go) - 设置首页不同

### 目录功能 ✨ 新增功能
- [`GenerateTOC(config *TOCConfig)`](toc.go) - 生成目录
- [`UpdateTOC()`](toc.go) - 更新目录
- [`AddHeadingWithBookmark(text string, level int, bookmarkName string)`](toc.go) - 添加带书签的标题
- [`AutoGenerateTOC(config *TOCConfig)`](toc.go) - 自动生成目录
- [`GetHeadingCount()`](toc.go) - 获取标题统计
- [`ListHeadings()`](toc.go) - 列出所有标题
- [`SetTOCStyle(level int, style *TextFormat)`](toc.go) - 设置目录样式

### 脚注与尾注功能 ✨ 新增功能
- [`AddFootnote(text string, footnoteText string)`](footnotes.go) - 添加脚注
- [`AddEndnote(text string, endnoteText string)`](footnotes.go) - 添加尾注
- [`AddFootnoteToRun(run *Run, footnoteText string)`](footnotes.go) - 为运行添加脚注
- [`SetFootnoteConfig(config *FootnoteConfig)`](footnotes.go) - 设置脚注配置
- [`GetFootnoteCount()`](footnotes.go) - 获取脚注数量
- [`GetEndnoteCount()`](footnotes.go) - 获取尾注数量
- [`RemoveFootnote(footnoteID string)`](footnotes.go) - 移除脚注
- [`RemoveEndnote(endnoteID string)`](footnotes.go) - 移除尾注

### 列表与编号功能 ✨ 新增功能
- [`AddListItem(text string, config *ListConfig)`](numbering.go) - 添加列表项
- [`AddBulletList(text string, level int, bulletType BulletType)`](numbering.go) - 添加无序列表
- [`AddNumberedList(text string, level int, numType ListType)`](numbering.go) - 添加有序列表
- [`CreateMultiLevelList(items []ListItem)`](numbering.go) - 创建多级列表
- [`RestartNumbering(numID string)`](numbering.go) - 重启编号

### 结构化文档标签 ✨ 新增功能
- [`CreateTOCSDT(title string, maxLevel int)`](sdt.go) - 创建目录SDT结构

### 模板功能 ✨ 新增功能
- [`NewTemplateEngine()`](template.go) - 创建新的模板引擎
- [`LoadTemplate(name, content string)`](template.go) - 从字符串加载模板
- [`LoadTemplateFromDocument(name string, doc *Document)`](template.go) - 从现有文档创建模板
- [`GetTemplate(name string)`](template.go) - 获取缓存的模板
- [`RenderToDocument(templateName string, data *TemplateData)`](template.go) - 渲染模板到新文档
- [`ValidateTemplate(template *Template)`](template.go) - 验证模板语法
- [`ClearCache()`](template.go) - 清空模板缓存
- [`RemoveTemplate(name string)`](template.go) - 移除指定模板

#### 模板引擎功能特性 ✨
**变量替换**: 支持 `{{变量名}}` 语法进行动态内容替换
**条件语句**: 支持 `{{#if 条件}}...{{/if}}` 语法进行条件渲染
**循环语句**: 支持 `{{#each 列表}}...{{/each}}` 语法进行列表渲染
**模板继承**: 支持 `{{extends "基础模板"}}` 语法进行模板继承
**循环内条件**: 完美支持循环内部的条件表达式，如 `{{#each items}}{{#if isActive}}...{{/if}}{{/each}}`
**数据类型支持**: 支持字符串、数字、布尔值、对象等多种数据类型
**结构体绑定**: 支持从Go结构体自动生成模板数据

### 模板数据操作
- [`NewTemplateData()`](template.go) - 创建新的模板数据
- [`SetVariable(name string, value interface{})`](template.go) - 设置变量
- [`SetList(name string, list []interface{})`](template.go) - 设置列表
- [`SetCondition(name string, value bool)`](template.go) - 设置条件
- [`SetVariables(variables map[string]interface{})`](template.go) - 批量设置变量
- [`GetVariable(name string)`](template.go) - 获取变量
- [`GetList(name string)`](template.go) - 获取列表
- [`GetCondition(name string)`](template.go) - 获取条件
- [`Merge(other *TemplateData)`](template.go) - 合并模板数据
- [`Clear()`](template.go) - 清空模板数据
- [`FromStruct(data interface{})`](template.go) - 从结构体生成模板数据

### 图片操作功能 ✨ 新增功能
- [`AddImageFromFile(filePath string, config *ImageConfig)`](image.go) - 从文件添加图片
- [`AddImageFromData(imageData []byte, fileName string, format ImageFormat, width, height int, config *ImageConfig)`](image.go) - 从数据添加图片
- [`ResizeImage(imageInfo *ImageInfo, size *ImageSize)`](image.go) - 调整图片大小
- [`SetImagePosition(imageInfo *ImageInfo, position ImagePosition, offsetX, offsetY float64)`](image.go) - 设置图片位置
- [`SetImageWrapText(imageInfo *ImageInfo, wrapText ImageWrapText)`](image.go) - 设置图片文字环绕
- [`SetImageAltText(imageInfo *ImageInfo, altText string)`](image.go) - 设置图片替代文字
- [`SetImageTitle(imageInfo *ImageInfo, title string)`](image.go) - 设置图片标题

## 段落操作方法

### 段落格式设置
- [`SetAlignment(alignment AlignmentType)`](document.go#L521) - 设置段落对齐方式
- [`SetSpacing(config *SpacingConfig)`](document.go#L558) - 设置段落间距
- [`SetStyle(styleID string)`](document.go#L773) - 设置段落样式

### 段落内容操作
- [`AddFormattedText(text string, format *TextFormat)`](document.go#L623) - 添加格式化文本
- [`ElementType()`](document.go#L61) - 获取段落元素类型

## 文档主体操作方法

### 元素查询
- [`GetParagraphs()`](document.go#L1149) - 获取所有段落
- [`GetTables()`](document.go#L1160) - 获取所有表格

### 元素添加
- [`AddElement(element interface{})`](document.go#L1171) - 添加元素到文档主体

## 表格操作方法

### 表格创建
- [`CreateTable(config *TableConfig)`](table.go#L161) - 创建新表格（✨ 新增：默认包含单线边框样式）
- [`AddTable(config *TableConfig)`](table.go#L257) - 添加表格到文档

### 行操作
- [`InsertRow(position int, data []string)`](table.go#L271) - 在指定位置插入行
- [`AppendRow(data []string)`](table.go#L329) - 在表格末尾添加行
- [`DeleteRow(rowIndex int)`](table.go#L334) - 删除指定行
- [`DeleteRows(startIndex, endIndex int)`](table.go#L351) - 删除多行
- [`GetRowCount()`](table.go#L562) - 获取行数

### 列操作
- [`InsertColumn(position int, data []string, width int)`](table.go#L369) - 在指定位置插入列
- [`AppendColumn(data []string, width int)`](table.go#L438) - 在表格末尾添加列
- [`DeleteColumn(colIndex int)`](table.go#L447) - 删除指定列
- [`DeleteColumns(startIndex, endIndex int)`](table.go#L474) - 删除多列
- [`GetColumnCount()`](table.go#L567) - 获取列数

### 单元格操作
- [`GetCell(row, col int)`](table.go#L502) - 获取指定单元格
- [`SetCellText(row, col int, text string)`](table.go#L515) - 设置单元格文本
- [`GetCellText(row, col int)`](table.go#L548) - 获取单元格文本
- [`SetCellFormat(row, col int, format *CellFormat)`](table.go#L691) - 设置单元格格式
- [`GetCellFormat(row, col int)`](table.go#L1238) - 获取单元格格式

### 单元格文本格式化
- [`SetCellFormattedText(row, col int, text string, format *TextFormat)`](table.go#L780) - 设置格式化文本
- [`AddCellFormattedText(row, col int, text string, format *TextFormat)`](table.go#L833) - 添加格式化文本

### 单元格合并
- [`MergeCellsHorizontal(row, startCol, endCol int)`](table.go#L887) - 水平合并单元格
- [`MergeCellsVertical(startRow, endRow, col int)`](table.go#L924) - 垂直合并单元格
- [`MergeCellsRange(startRow, endRow, startCol, endCol int)`](table.go#L971) - 范围合并单元格
- [`UnmergeCells(row, col int)`](table.go#L1004) - 取消合并单元格
- [`IsCellMerged(row, col int)`](table.go#L1074) - 检查单元格是否已合并
- [`GetMergedCellInfo(row, col int)`](table.go#L1098) - 获取合并单元格信息

### 单元格特殊属性
- [`SetCellPadding(row, col int, padding int)`](table.go#L1189) - 设置单元格内边距
- [`SetCellTextDirection(row, col int, direction CellTextDirection)`](table.go#L1202) - 设置文字方向
- [`GetCellTextDirection(row, col int)`](table.go#L1223) - 获取文字方向
- [`ClearCellContent(row, col int)`](table.go#L1138) - 清除单元格内容
- [`ClearCellFormat(row, col int)`](table.go#L1156) - 清除单元格格式

### 表格整体操作
- [`ClearTable()`](table.go#L575) - 清空表格内容
- [`CopyTable()`](table.go#L593) - 复制表格
- [`ElementType()`](table.go#L66) - 获取表格元素类型

### 行高设置
- [`SetRowHeight(rowIndex int, config *RowHeightConfig)`](table.go#L1318) - 设置行高
- [`GetRowHeight(rowIndex int)`](table.go#L1339) - 获取行高
- [`SetRowHeightRange(startRow, endRow int, config *RowHeightConfig)`](table.go#L1371) - 设置多行行高

### 表格布局与对齐
- [`SetTableLayout(config *TableLayoutConfig)`](table.go#L1447) - 设置表格布局
- [`GetTableLayout()`](table.go#L1473) - 获取表格布局
- [`SetTableAlignment(alignment TableAlignment)`](table.go#L1488) - 设置表格对齐

### 行属性设置
- [`SetRowKeepTogether(rowIndex int, keepTogether bool)`](table.go#L1529) - 设置行保持完整
- [`SetRowAsHeader(rowIndex int, isHeader bool)`](table.go#L1552) - 设置行为标题行
- [`SetHeaderRows(startRow, endRow int)`](table.go#L1575) - 设置多行为标题行
- [`IsRowHeader(rowIndex int)`](table.go#L1600) - 检查是否为标题行
- [`IsRowKeepTogether(rowIndex int)`](table.go#L1614) - 检查行是否保持完整
- [`SetRowKeepWithNext(rowIndex int, keepWithNext bool)`](table.go#L1645) - 设置与下一行保持在一起

### 表格分页设置
- [`SetTablePageBreak(config *TablePageBreakConfig)`](table.go#L1636) - 设置表格分页
- [`GetTableBreakInfo()`](table.go#L1657) - 获取分页信息

### 表格样式
- [`ApplyTableStyle(config *TableStyleConfig)`](table.go#L1956) - 应用表格样式
- [`CreateCustomTableStyle(styleID, styleName string, borderConfig *TableBorderConfig, shadingConfig *ShadingConfig, firstRowBold bool)`](table.go#L2213) - 创建自定义表格样式

### 边框设置
- [`SetTableBorders(config *TableBorderConfig)`](table.go#L2038) - 设置表格边框
- [`SetCellBorders(row, col int, config *CellBorderConfig)`](table.go#L2085) - 设置单元格边框
- [`RemoveTableBorders()`](table.go#L2168) - 移除表格边框
- [`RemoveCellBorders(row, col int)`](table.go#L2194) - 移除单元格边框

### 背景与阴影
- [`SetTableShading(config *ShadingConfig)`](table.go#L2069) - 设置表格底纹
- [`SetCellShading(row, col int, config *ShadingConfig)`](table.go#L2121) - 设置单元格底纹
- [`SetAlternatingRowColors(evenRowColor, oddRowColor string)`](table.go#L2142) - 设置交替行颜色

### 单元格遍历迭代器 ✨ **新功能**

提供强大的单元格遍历和查找功能：

##### CellIterator - 单元格迭代器
```go
// 创建迭代器
iterator := table.NewCellIterator()

// 遍历所有单元格
for iterator.HasNext() {
    cellInfo, err := iterator.Next()
    if err != nil {
        break
    }
    fmt.Printf("单元格[%d,%d]: %s\n", cellInfo.Row, cellInfo.Col, cellInfo.Text)
}

// 获取进度
progress := iterator.Progress() // 0.0 - 1.0

// 重置迭代器
iterator.Reset()
```

##### ForEach 批量处理
```go
// 遍历所有单元格
err := table.ForEach(func(row, col int, cell *TableCell, text string) error {
    // 处理每个单元格
    return nil
})

// 按行遍历
err := table.ForEachInRow(rowIndex, func(col int, cell *TableCell, text string) error {
    // 处理行中的每个单元格
    return nil
})

// 按列遍历
err := table.ForEachInColumn(colIndex, func(row int, cell *TableCell, text string) error {
    // 处理列中的每个单元格
    return nil
})
```

##### 范围操作
```go
// 获取指定范围的单元格
cells, err := table.GetCellRange(startRow, startCol, endRow, endCol)
for _, cellInfo := range cells {
    fmt.Printf("单元格[%d,%d]: %s\n", cellInfo.Row, cellInfo.Col, cellInfo.Text)
}
```

##### 查找功能
```go
// 自定义条件查找
cells, err := table.FindCells(func(row, col int, cell *TableCell, text string) bool {
    return strings.Contains(text, "关键词")
})

// 按文本查找
exactCells, err := table.FindCellsByText("精确匹配", true)
fuzzyCells, err := table.FindCellsByText("模糊", false)
```

##### CellInfo 结构
```go
type CellInfo struct {
    Row    int        // 行索引
    Col    int        // 列索引
    Cell   *TableCell // 单元格引用
    Text   string     // 单元格文本
    IsLast bool       // 是否为最后一个单元格
}
```

## 工具函数

### 日志系统
- [`NewLogger(level LogLevel, output io.Writer)`](logger.go#L56) - 创建新的日志记录器
- [`SetGlobalLevel(level LogLevel)`](logger.go#L129) - 设置全局日志级别
- [`SetGlobalOutput(output io.Writer)`](logger.go#L134) - 设置全局日志输出
- [`Debug(msg string)`](logger.go#L159) - 输出调试信息
- [`Info(msg string)`](logger.go#L164) - 输出信息
- [`Warn(msg string)`](logger.go#L169) - 输出警告
- [`Error(msg string)`](logger.go#L174) - 输出错误

### 错误处理
- [`NewDocumentError(operation string, cause error, context string)`](errors.go#L47) - 创建文档错误
- [`WrapError(operation string, err error)`](errors.go#L56) - 包装错误
- [`WrapErrorWithContext(operation string, err error, context string)`](errors.go#L64) - 带上下文包装错误
- [`NewValidationError(field, value, message string)`](errors.go#L84) - 创建验证错误

### 域字段工具 ✨ 新增功能
- [`CreateHyperlinkField(anchor string)`](field.go) - 创建超链接域
- [`CreatePageRefField(anchor string)`](field.go) - 创建页码引用域

## 常用配置结构

### 文本格式
- `TextFormat` - 文本格式配置
- `AlignmentType` - 对齐类型
- `SpacingConfig` - 间距配置

### 表格配置
- `TableConfig` - 表格基础配置
- `CellFormat` - 单元格格式
- `RowHeightConfig` - 行高配置
- `TableLayoutConfig` - 表格布局配置
- `TableStyleConfig` - 表格样式配置
- `BorderConfig` - 边框配置
- `ShadingConfig` - 底纹配置

### 页面设置配置 ✨ 新增
- `PageSettings` - 页面设置配置
- `PageSize` - 页面尺寸类型（A4、Letter、Legal、A3、A5、Custom）
- `PageOrientation` - 页面方向（Portrait纵向、Landscape横向）
- `SectionProperties` - 节属性（包含页面设置信息）

### 页眉页脚配置 ✨ 新增
- `HeaderFooterType` - 页眉页脚类型（Default、First、Even）
- `Header` - 页眉结构
- `Footer` - 页脚结构
- `HeaderFooterReference` - 页眉页脚引用
- `PageNumber` - 页码字段

### 目录配置 ✨ 新增
- `TOCConfig` - 目录配置
- `TOCEntry` - 目录条目
- `Bookmark` - 书签结构
- `BookmarkEnd` - 书签结束标记

### 脚注尾注配置 ✨ 新增
- `FootnoteConfig` - 脚注配置
- `FootnoteType` - 脚注类型（Footnote脚注、Endnote尾注）
- `FootnoteNumberFormat` - 脚注编号格式
- `FootnoteRestart` - 脚注重新开始规则
- `FootnotePosition` - 脚注位置
- `Footnote` - 脚注结构
- `Endnote` - 尾注结构

### 列表编号配置 ✨ 新增
- `ListConfig` - 列表配置
- `ListType` - 列表类型（Bullet无序、Number有序等）
- `BulletType` - 项目符号类型
- `ListItem` - 列表项结构
- `Numbering` - 编号定义
- `AbstractNum` - 抽象编号定义
- `Level` - 编号级别

### 结构化文档标签配置 ✨ 新增
- `SDT` - 结构化文档标签
- `SDTProperties` - SDT属性
- `SDTContent` - SDT内容

### 域字段配置 ✨ 新增
- `FieldChar` - 域字符
- `InstrText` - 域指令文本
- `HyperlinkField` - 超链接域
- `PageRefField` - 页码引用域

### 图片配置 ✨ 新增
- `ImageConfig` - 图片配置
- `ImageSize` - 图片尺寸配置
- `ImageFormat` - 图片格式（PNG、JPEG、GIF）
- `ImagePosition` - 图片位置（inline、floatLeft、floatRight）
- `ImageWrapText` - 文字环绕类型（none、square、tight、topAndBottom）
- `ImageInfo` - 图片信息结构
- `AlignmentType` - 对齐方式（left、center、right、justify）

## 使用示例

```go
// 创建新文档
doc := document.New()

// ✨ 新增：页面设置示例
// 设置页面为A4横向
doc.SetPageOrientation(document.OrientationLandscape)

// 设置自定义边距（上下左右：25mm）
doc.SetPageMargins(25, 25, 25, 25)

// 设置自定义页面尺寸（200mm x 300mm）
doc.SetCustomPageSize(200, 300)

// 或者使用完整页面设置
pageSettings := &document.PageSettings{
    Size:           document.PageSizeLetter,
    Orientation:    document.OrientationPortrait,
    MarginTop:      30,
    MarginRight:    20,
    MarginBottom:   30,
    MarginLeft:     20,
    HeaderDistance: 15,
    FooterDistance: 15,
    GutterWidth:    0,
}
doc.SetPageSettings(pageSettings)

// ✨ 新增：页眉页脚示例
// 添加页眉
doc.AddHeader(document.HeaderFooterTypeDefault, "这是页眉")

// 添加带页码的页脚
doc.AddFooterWithPageNumber(document.HeaderFooterTypeDefault, "第", true)

// 设置首页不同
doc.SetDifferentFirstPage(true)

// ✨ 新增：目录示例
// 添加带书签的标题
doc.AddHeadingWithBookmark("第一章 概述", 1, "chapter1")
doc.AddHeadingWithBookmark("1.1 背景", 2, "section1_1")

// 生成目录
tocConfig := document.DefaultTOCConfig()
tocConfig.Title = "目录"
tocConfig.MaxLevel = 3
doc.GenerateTOC(tocConfig)

// ✨ 新增：脚注示例
// 添加脚注
doc.AddFootnote("这是正文内容", "这是脚注内容")

// 添加尾注
doc.AddEndnote("更多说明", "这是尾注内容")

// ✨ 新增：列表示例
// 添加无序列表
doc.AddBulletList("列表项1", 0, document.BulletTypeDot)
doc.AddBulletList("列表项2", 1, document.BulletTypeCircle)

// 添加有序列表
doc.AddNumberedList("编号项1", 0, document.ListTypeNumber)

// ✨ 新增：图片示例
// 从文件添加图片
imageInfo, err := doc.AddImageFromFile("path/to/image.png", &document.ImageConfig{
    Size: &document.ImageSize{
        Width:  100.0, // 100毫米宽度
        Height: 75.0,  // 75毫米高度
    },
    Position: document.ImagePositionInline,
    WrapText: document.ImageWrapNone,
    AltText:  "示例图片",
    Title:    "这是一个示例图片",
})

// 从数据添加图片
imageData := []byte{...} // 图片二进制数据
imageInfo2, err := doc.AddImageFromData(
    imageData,
    "example.png",
    document.ImageFormatPNG,
    200, 150, // 原始像素尺寸
    &document.ImageConfig{
        Size: &document.ImageSize{
            Width:           60.0, // 只设置宽度
            KeepAspectRatio: true, // 保持长宽比
        },
        AltText: "数据图片",
    },
)

// 调整图片大小
err = doc.ResizeImage(imageInfo, &document.ImageSize{
    Width:  80.0,
    Height: 60.0,
})

// 设置图片属性
err = doc.SetImagePosition(imageInfo, document.ImagePositionFloatLeft, 5.0, 0.0)
err = doc.SetImageWrapText(imageInfo, document.ImageWrapSquare)
err = doc.SetImageAltText(imageInfo, "更新的替代文字")
err = doc.SetImageTitle(imageInfo, "更新的标题")

// ✨ 新增：设置图片对齐方式（仅适用于嵌入式图片）
err = doc.SetImageAlignment(imageInfo, document.AlignCenter)  // 居中对齐
err = doc.SetImageAlignment(imageInfo, document.AlignLeft)    // 左对齐
err = doc.SetImageAlignment(imageInfo, document.AlignRight)   // 右对齐
doc.AddNumberedList("第一项", 0, document.ListTypeDecimal)
doc.AddNumberedList("第二项", 0, document.ListTypeDecimal)

// 添加段落
para := doc.AddParagraph("这是一个段落")
para.SetAlignment(document.AlignCenter)

// 创建表格
table := doc.CreateTable(&document.TableConfig{
    Rows:  3,
    Cols:  3,
    Width: 5000,
})

// 设置单元格内容
table.SetCellText(0, 0, "标题")

// 保存文档
doc.Save("example.docx")
```

## 注意事项

1. 所有位置索引都是从0开始
2. 宽度单位使用磅（pt），1磅 = 20twips
3. 颜色使用十六进制格式，如 "FF0000" 表示红色
4. 在操作表格前请确保行列索引有效，否则可能返回错误
5. 页眉页脚类型包括：Default（默认）、First（首页）、Even（偶数页）
6. 目录功能需要先添加带书签的标题，然后调用生成目录方法
7. 脚注和尾注会自动编号，支持多种编号格式和重启规则
8. 列表支持多级嵌套，最多支持9级缩进
9. 结构化文档标签主要用于目录等特殊功能的实现
10. 图片支持PNG、JPEG、GIF格式，会自动嵌入到文档中
11. 图片尺寸可以用毫米或像素指定，支持保持长宽比的缩放
12. 图片位置支持嵌入式、左浮动、右浮动等多种布局方式
13. 图片对齐功能仅适用于嵌入式图片（ImagePositionInline），浮动图片请使用位置控制 
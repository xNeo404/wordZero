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
- [`Open(filename string)`](document.go#L269) - 打开现有Word文档

### 文档保存与导出
- [`Save(filename string)`](document.go#L337) - 保存文档到文件
- [`ToBytes()`](document.go#L1107) - 将文档转换为字节数组

### 文档内容操作
- [`AddParagraph(text string)`](document.go#L420) - 添加简单段落
- [`AddFormattedParagraph(text string, format *TextFormat)`](document.go#L459) - 添加格式化段落
- [`AddHeadingParagraph(text string, level int)`](document.go#L682) - 添加标题段落

### 样式管理
- [`GetStyleManager()`](document.go#L791) - 获取样式管理器

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

## 使用示例

```go
// 创建新文档
doc := document.New()

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
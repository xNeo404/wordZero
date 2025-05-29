# WordZero 更新日志

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
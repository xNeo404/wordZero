// Package main 演示从现有DOCX模板文件生成新文档
package main

import (
	"fmt"
	"log"
	"os"
	"time"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
)

func main() {
	fmt.Println("=== WordZero 从现有DOCX模板生成文档演示 ===")

	// 确保输出目录存在
	os.MkdirAll("examples/output", 0755)

	// 演示从现有模板文件生成文档
	demonstrateTemplateFromExistingDocx()

	// 演示使用复杂模板（包含条件和循环）
	demonstrateComplexTemplateFromDocx()

	fmt.Println("\n✅ 所有演示完成！")
}

// demonstrateTemplateFromExistingDocx 演示从现有DOCX模板文件生成新文档
func demonstrateTemplateFromExistingDocx() {
	fmt.Println("\n📄 演示1：从现有DOCX模板生成发票文档")

	// 先创建一个模板文件作为示例
	createInvoiceTemplate()

	// 创建模板引擎
	engine := document.NewTemplateEngine()

	// 1. 打开现有的DOCX模板文件
	templateDoc, err := document.Open("examples/output/invoice_template.docx")
	if err != nil {
		log.Fatalf("无法打开模板文件: %v", err)
	}
	fmt.Println("✓ 成功打开模板文件: invoice_template.docx")

	// 2. 从文档创建模板
	template, err := engine.LoadTemplateFromDocument("invoice_template", templateDoc)
	if err != nil {
		log.Fatalf("从文档创建模板失败: %v", err)
	}
	fmt.Printf("✓ 从文档解析到 %d 个模板变量\n", len(template.Variables))

	// 3. 准备发票数据
	data := document.NewTemplateData()

	// 基本信息
	data.SetVariable("invoiceNumber", "INV-2024-12001")
	data.SetVariable("issueDate", time.Now().Format("2006年01月02日"))
	data.SetVariable("dueDate", time.Now().AddDate(0, 0, 30).Format("2006年01月02日"))

	// 出票方信息
	data.SetVariable("sellerName", "WordZero科技有限公司")
	data.SetVariable("sellerAddress", "上海市浦东新区科技园区1号楼")
	data.SetVariable("sellerPhone", "021-12345678")
	data.SetVariable("sellerEmail", "billing@wordzero.com")
	data.SetVariable("sellerTaxId", "91310000123456789X")

	// 收票方信息
	data.SetVariable("buyerName", "某某企业有限公司")
	data.SetVariable("buyerAddress", "北京市朝阳区商务楼A座20层")
	data.SetVariable("buyerPhone", "010-87654321")
	data.SetVariable("buyerTaxId", "91110000987654321Y")

	// 商品信息
	items := []interface{}{
		map[string]interface{}{
			"description":  "WordZero企业版许可证",
			"quantity":     1,
			"unit":         "套",
			"unitPrice":    9999.00,
			"subtotal":     9999.00,
			"isDiscounted": false,
		},
		map[string]interface{}{
			"description":  "技术支持服务（12个月）",
			"quantity":     12,
			"unit":         "月",
			"unitPrice":    500.00,
			"subtotal":     6000.00,
			"isDiscounted": true,
			"discount":     300.00,
		},
		map[string]interface{}{
			"description":  "在线培训课程",
			"quantity":     3,
			"unit":         "次",
			"unitPrice":    800.00,
			"subtotal":     2400.00,
			"isDiscounted": false,
		},
	}
	data.SetList("items", items)

	// 费用计算
	data.SetVariable("subtotalAmount", "18399.00")
	data.SetVariable("totalDiscount", "300.00")
	data.SetVariable("taxRate", "13")
	data.SetVariable("taxAmount", "2352.87")
	data.SetVariable("shippingFee", "50.00")
	data.SetVariable("totalAmount", "20501.87")

	// 条件设置
	data.SetCondition("hasSubtotal", true)
	data.SetCondition("hasDiscount", true)
	data.SetCondition("hasTax", true)
	data.SetCondition("hasShipping", true)
	data.SetCondition("isPaid", false)
	data.SetCondition("isOverdue", false)

	// 其他信息
	data.SetVariable("notes", "请在30天内付款，逾期将收取滞纳金。")
	data.SetVariable("issuer", "张会计")
	data.SetVariable("reviewer", "李经理")

	// 4. 渲染生成新文档
	invoiceDoc, err := engine.RenderToDocument("invoice_template", data)
	if err != nil {
		log.Fatalf("渲染发票失败: %v", err)
	}

	// 5. 保存生成的发票
	outputFile := "examples/output/generated_invoice_" + time.Now().Format("20060102_150405") + ".docx"
	err = invoiceDoc.Save(outputFile)
	if err != nil {
		log.Fatalf("保存发票失败: %v", err)
	}

	fmt.Printf("✓ 成功生成发票文档: %s\n", outputFile)
}

// demonstrateComplexTemplateFromDocx 演示使用包含条件和循环的复杂模板
func demonstrateComplexTemplateFromDocx() {
	fmt.Println("\n📊 演示2：从复杂DOCX模板生成项目报告")

	// 先创建一个复杂模板文件
	createProjectReportTemplate()

	// 创建模板引擎
	engine := document.NewTemplateEngine()

	// 1. 打开复杂模板文件
	templateDoc, err := document.Open("examples/output/project_report_template.docx")
	if err != nil {
		log.Fatalf("无法打开项目报告模板: %v", err)
	}
	fmt.Println("✓ 成功打开模板文件: project_report_template.docx")

	// 2. 从文档创建模板
	template, err := engine.LoadTemplateFromDocument("project_report_template", templateDoc)
	if err != nil {
		log.Fatalf("从项目报告文档创建模板失败: %v", err)
	}
	fmt.Printf("✓ 从项目报告解析到 %d 个模板变量\n", len(template.Variables))

	// 3. 准备项目数据
	data := document.NewTemplateData()

	// 基本信息
	data.SetVariable("projectName", "WordZero企业文档管理系统")
	data.SetVariable("projectManager", "李项目经理")
	data.SetVariable("reportDate", time.Now().Format("2006年01月02日"))
	data.SetVariable("projectStatus", "进行中")
	data.SetVariable("completionRate", "85")

	// 团队成员
	teamMembers := []interface{}{
		map[string]interface{}{
			"name":       "张开发",
			"role":       "高级开发工程师",
			"workload":   "核心功能开发",
			"isTeamLead": true,
		},
		map[string]interface{}{
			"name":       "王测试",
			"role":       "测试工程师",
			"workload":   "功能测试与质量保证",
			"isTeamLead": false,
		},
		map[string]interface{}{
			"name":       "刘设计",
			"role":       "UI/UX设计师",
			"workload":   "界面设计与用户体验",
			"isTeamLead": false,
		},
		map[string]interface{}{
			"name":       "陈运维",
			"role":       "运维工程师",
			"workload":   "系统部署与维护",
			"isTeamLead": false,
		},
	}
	data.SetList("teamMembers", teamMembers)

	// 项目里程碑
	milestones := []interface{}{
		map[string]interface{}{
			"title":       "需求分析完成",
			"date":        "2024年10月15日",
			"isCompleted": true,
			"isCurrent":   false,
		},
		map[string]interface{}{
			"title":       "系统设计完成",
			"date":        "2024年11月1日",
			"isCompleted": true,
			"isCurrent":   false,
		},
		map[string]interface{}{
			"title":       "核心开发阶段",
			"date":        "2024年12月1日",
			"isCompleted": false,
			"isCurrent":   true,
		},
		map[string]interface{}{
			"title":       "系统测试",
			"date":        "2024年12月15日",
			"isCompleted": false,
			"isCurrent":   false,
		},
	}
	data.SetList("milestones", milestones)

	// 成就列表
	achievements := []interface{}{
		"完成了核心模板引擎的开发",
		"实现了完整的样式管理系统",
		"建立了自动化测试流程",
		"完成了API文档编写",
	}
	data.SetList("achievements", achievements)

	// 风险列表
	risks := []interface{}{
		map[string]interface{}{
			"description": "第三方库兼容性问题",
			"level":       "中等",
			"mitigation":  "提前进行兼容性测试，准备备选方案",
		},
		map[string]interface{}{
			"description": "项目进度可能延期",
			"level":       "低",
			"mitigation":  "合理安排任务优先级，增加缓冲时间",
		},
	}
	data.SetList("risks", risks)

	// 条件设置
	data.SetCondition("showTeamMembers", true)
	data.SetCondition("showMilestones", true)
	data.SetCondition("showAchievements", true)
	data.SetCondition("showRisks", true)
	data.SetCondition("isOnTrack", true)
	data.SetCondition("needsAttention", false)

	// 其他信息
	data.SetVariable("nextReviewDate", time.Now().AddDate(0, 0, 7).Format("2006年01月02日"))
	data.SetVariable("reporter", "李项目经理")
	data.SetVariable("reviewer", "王总监")

	// 4. 渲染生成新文档
	reportDoc, err := engine.RenderToDocument("project_report_template", data)
	if err != nil {
		log.Fatalf("渲染项目报告失败: %v", err)
	}

	// 5. 保存生成的报告
	outputFile := "examples/output/generated_project_report_" + time.Now().Format("20060102_150405") + ".docx"
	err = reportDoc.Save(outputFile)
	if err != nil {
		log.Fatalf("保存项目报告失败: %v", err)
	}

	fmt.Printf("✓ 成功生成项目报告: %s\n", outputFile)
}

// createInvoiceTemplate 创建发票模板文件
func createInvoiceTemplate() {
	doc := document.New()

	// 标题
	title := doc.AddParagraph("商业发票")
	title.SetAlignment(document.AlignCenter)

	doc.AddParagraph("")

	// 发票基本信息
	doc.AddParagraph("发票编号：{{invoiceNumber}}")
	doc.AddParagraph("开票日期：{{issueDate}}")
	doc.AddParagraph("付款期限：{{dueDate}}")

	doc.AddParagraph("")
	doc.AddParagraph("═══════════════════════════════════════")
	doc.AddParagraph("")

	// 出票方信息
	doc.AddParagraph("出票方信息：")
	doc.AddParagraph("{{sellerName}}")
	doc.AddParagraph("地址：{{sellerAddress}}")
	doc.AddParagraph("电话：{{sellerPhone}}")
	doc.AddParagraph("邮箱：{{sellerEmail}}")
	doc.AddParagraph("{{#if sellerTaxId}}")
	doc.AddParagraph("税号：{{sellerTaxId}}")
	doc.AddParagraph("{{/if}}")

	doc.AddParagraph("")

	// 收票方信息
	doc.AddParagraph("收票方信息：")
	doc.AddParagraph("{{buyerName}}")
	doc.AddParagraph("地址：{{buyerAddress}}")
	doc.AddParagraph("电话：{{buyerPhone}}")
	doc.AddParagraph("{{#if buyerTaxId}}")
	doc.AddParagraph("税号：{{buyerTaxId}}")
	doc.AddParagraph("{{/if}}")

	doc.AddParagraph("")
	doc.AddParagraph("═══════════════════════════════════════")
	doc.AddParagraph("")

	// 商品明细
	doc.AddParagraph("商品明细：")
	doc.AddParagraph("{{#each items}}")
	doc.AddParagraph("{{@index}}. {{description}}")
	doc.AddParagraph("   数量：{{quantity}} {{unit}}")
	doc.AddParagraph("   单价：{{unitPrice}} 元")
	doc.AddParagraph("   小计：{{subtotal}} 元")
	doc.AddParagraph("   {{#if isDiscounted}}")
	doc.AddParagraph("   折扣：-{{discount}} 元")
	doc.AddParagraph("   {{/if}}")
	doc.AddParagraph("")
	doc.AddParagraph("{{/each}}")

	doc.AddParagraph("───────────────────────────────────────")
	doc.AddParagraph("{{#if hasSubtotal}}")
	doc.AddParagraph("商品小计：{{subtotalAmount}} 元")
	doc.AddParagraph("{{/if}}")
	doc.AddParagraph("")
	doc.AddParagraph("{{#if hasDiscount}}")
	doc.AddParagraph("总折扣：-{{totalDiscount}} 元")
	doc.AddParagraph("{{/if}}")
	doc.AddParagraph("")
	doc.AddParagraph("{{#if hasTax}}")
	doc.AddParagraph("税费（{{taxRate}}%）：{{taxAmount}} 元")
	doc.AddParagraph("{{/if}}")
	doc.AddParagraph("")
	doc.AddParagraph("{{#if hasShipping}}")
	doc.AddParagraph("运费：{{shippingFee}} 元")
	doc.AddParagraph("{{/if}}")
	doc.AddParagraph("")
	doc.AddParagraph("总计：{{totalAmount}} 元")
	doc.AddParagraph("───────────────────────────────────────")

	doc.AddParagraph("")
	doc.AddParagraph("{{#if isPaid}}")
	doc.AddParagraph("✅ 付款状态：已付款")
	doc.AddParagraph("付款日期：{{paymentDate}}")
	doc.AddParagraph("付款方式：{{paymentMethod}}")
	doc.AddParagraph("{{/if}}")
	doc.AddParagraph("")
	doc.AddParagraph("{{#if isOverdue}}")
	doc.AddParagraph("⚠️ 状态：已逾期")
	doc.AddParagraph("逾期天数：{{overdueDays}} 天")
	doc.AddParagraph("{{/if}}")
	doc.AddParagraph("")
	doc.AddParagraph("{{#if notes}}")
	doc.AddParagraph("备注：")
	doc.AddParagraph("{{notes}}")
	doc.AddParagraph("{{/if}}")
	doc.AddParagraph("")
	doc.AddParagraph("感谢您的合作！")
	doc.AddParagraph("")
	doc.AddParagraph("开票人：{{issuer}}")
	doc.AddParagraph("审核人：{{reviewer}}")

	// 保存模板文件
	err := doc.Save("examples/output/invoice_template.docx")
	if err != nil {
		log.Fatalf("保存发票模板失败: %v", err)
	}

	fmt.Println("✓ 创建发票模板文件: invoice_template.docx")
}

// createProjectReportTemplate 创建项目报告模板文件
func createProjectReportTemplate() {
	doc := document.New()

	// 标题
	title := doc.AddParagraph("项目进度报告")
	title.SetAlignment(document.AlignCenter)

	doc.AddParagraph("")

	// 基本信息
	doc.AddParagraph("项目名称：{{projectName}}")
	doc.AddParagraph("项目经理：{{projectManager}}")
	doc.AddParagraph("报告日期：{{reportDate}}")
	doc.AddParagraph("项目状态：{{projectStatus}}")
	doc.AddParagraph("完成度：{{completionRate}}%")

	doc.AddParagraph("")
	doc.AddParagraph("═══════════════════════════════════════")
	doc.AddParagraph("")

	// 团队成员
	doc.AddParagraph("{{#if showTeamMembers}}")
	doc.AddParagraph("团队成员：")
	doc.AddParagraph("{{#each teamMembers}}")
	doc.AddParagraph("• {{name}} - {{role}}")
	doc.AddParagraph("  工作内容：{{workload}}")
	doc.AddParagraph("  {{#if isTeamLead}}")
	doc.AddParagraph("  👨‍💼 团队负责人")
	doc.AddParagraph("  {{/if}}")
	doc.AddParagraph("")
	doc.AddParagraph("{{/each}}")
	doc.AddParagraph("{{/if}}")

	doc.AddParagraph("")

	// 项目里程碑
	doc.AddParagraph("{{#if showMilestones}}")
	doc.AddParagraph("项目里程碑：")
	doc.AddParagraph("{{#each milestones}}")
	doc.AddParagraph("{{#if isCompleted}}")
	doc.AddParagraph("✅ {{title}} - {{date}}")
	doc.AddParagraph("{{/if}}")
	doc.AddParagraph("{{#if isCurrent}}")
	doc.AddParagraph("🔄 {{title}} - {{date}} (进行中)")
	doc.AddParagraph("{{/if}}")
	doc.AddParagraph("{{/each}}")
	doc.AddParagraph("{{/if}}")

	doc.AddParagraph("")

	// 主要成就
	doc.AddParagraph("{{#if showAchievements}}")
	doc.AddParagraph("主要成就：")
	doc.AddParagraph("{{#each achievements}}")
	doc.AddParagraph("✓ {{this}}")
	doc.AddParagraph("{{/each}}")
	doc.AddParagraph("{{/if}}")

	doc.AddParagraph("")

	// 风险管理
	doc.AddParagraph("{{#if showRisks}}")
	doc.AddParagraph("风险管理：")
	doc.AddParagraph("{{#each risks}}")
	doc.AddParagraph("⚠️ {{description}}")
	doc.AddParagraph("   风险等级：{{level}}")
	doc.AddParagraph("   缓解措施：{{mitigation}}")
	doc.AddParagraph("")
	doc.AddParagraph("{{/each}}")
	doc.AddParagraph("{{/if}}")

	doc.AddParagraph("")
	doc.AddParagraph("───────────────────────────────────────")
	doc.AddParagraph("")

	// 项目状态
	doc.AddParagraph("{{#if isOnTrack}}")
	doc.AddParagraph("✅ 项目进展顺利，按计划推进")
	doc.AddParagraph("{{/if}}")
	doc.AddParagraph("")
	doc.AddParagraph("{{#if needsAttention}}")
	doc.AddParagraph("⚠️ 项目需要特别关注")
	doc.AddParagraph("{{/if}}")
	doc.AddParagraph("")

	doc.AddParagraph("下次汇报日期：{{nextReviewDate}}")
	doc.AddParagraph("")
	doc.AddParagraph("报告人：{{reporter}}")
	doc.AddParagraph("审核人：{{reviewer}}")

	// 保存模板文件
	err := doc.Save("examples/output/project_report_template.docx")
	if err != nil {
		log.Fatalf("保存项目报告模板失败: %v", err)
	}

	fmt.Println("✓ 创建项目报告模板文件: project_report_template.docx")
}

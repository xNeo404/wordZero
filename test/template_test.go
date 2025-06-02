// Package test 模板功能集成测试
package test

import (
	"os"
	"testing"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
)

// TestTemplateIntegration 模板功能集成测试
func TestTemplateIntegration(t *testing.T) {
	// 确保输出目录存在
	outputDir := "output"
	if _, err := os.Stat(outputDir); os.IsNotExist(err) {
		err = os.Mkdir(outputDir, 0755)
		if err != nil {
			t.Fatalf("Failed to create output directory: %v", err)
		}
	}

	t.Run("变量替换集成测试", testVariableReplacementIntegration)
	t.Run("条件语句集成测试", testConditionalStatementsIntegration)
	t.Run("循环语句集成测试", testLoopStatementsIntegration)
	t.Run("模板继承集成测试", testTemplateInheritanceIntegration)
	t.Run("复杂模板集成测试", testComplexTemplateIntegration)
	t.Run("文档模板转换集成测试", testDocumentToTemplateIntegration)
	t.Run("结构体绑定集成测试", testStructDataBindingIntegration)
}

// testVariableReplacementIntegration 测试变量替换集成功能
func testVariableReplacementIntegration(t *testing.T) {
	engine := document.NewTemplateEngine()

	// 创建包含多种变量类型的模板
	templateContent := `产品信息单

产品名称：{{productName}}
产品价格：{{price}} 元
产品数量：{{quantity}} 个
是否库存充足：{{inStock}}
产品描述：{{description}}
更新时间：{{updateTime}}`

	template, err := engine.LoadTemplate("product_info", templateContent)
	if err != nil {
		t.Fatalf("Failed to load template: %v", err)
	}

	// 验证解析的变量数量
	expectedVars := 6
	if len(template.Variables) != expectedVars {
		t.Errorf("Expected %d variables, got %d", expectedVars, len(template.Variables))
	}

	// 创建多种类型的数据
	data := document.NewTemplateData()
	data.SetVariable("productName", "WordZero处理器")
	data.SetVariable("price", 299.99)
	data.SetVariable("quantity", 100)
	data.SetVariable("inStock", true)
	data.SetVariable("description", "高效的Word文档处理工具")
	data.SetVariable("updateTime", "2024-12-01 15:30:00")

	// 渲染并保存文档
	doc, err := engine.RenderToDocument("product_info", data)
	if err != nil {
		t.Fatalf("Failed to render template: %v", err)
	}

	err = doc.Save("output/test_variable_replacement_integration.docx")
	if err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// 验证文档内容
	if len(doc.Body.Elements) == 0 {
		t.Error("Expected document to have content")
	}
}

// testConditionalStatementsIntegration 测试条件语句集成功能
func testConditionalStatementsIntegration(t *testing.T) {
	engine := document.NewTemplateEngine()

	// 创建包含嵌套条件的模板
	templateContent := `用户权限报告

用户名：{{username}}

{{#if isAdmin}}
管理员权限：
- 系统配置访问权限
- 用户管理权限
- 数据备份权限
{{/if}}

{{#if isEditor}}
编辑权限：
- 内容编辑权限
- 文档管理权限
{{/if}}

{{#if isViewer}}
查看权限：
- 只读访问权限
{{/if}}

{{#if hasSpecialAccess}}
特殊权限：
- API访问权限
- 高级功能权限
{{/if}}`

	_, err := engine.LoadTemplate("user_permissions", templateContent)
	if err != nil {
		t.Fatalf("Failed to load template: %v", err)
	}

	// 测试不同权限组合
	testCases := []struct {
		name         string
		username     string
		isAdmin      bool
		isEditor     bool
		isViewer     bool
		hasSpecial   bool
		expectedFile string
	}{
		{
			name:         "管理员权限",
			username:     "admin_user",
			isAdmin:      true,
			isEditor:     false,
			isViewer:     false,
			hasSpecial:   true,
			expectedFile: "test_conditional_admin.docx",
		},
		{
			name:         "编辑员权限",
			username:     "editor_user",
			isAdmin:      false,
			isEditor:     true,
			isViewer:     false,
			hasSpecial:   false,
			expectedFile: "test_conditional_editor.docx",
		},
		{
			name:         "查看者权限",
			username:     "viewer_user",
			isAdmin:      false,
			isEditor:     false,
			isViewer:     true,
			hasSpecial:   false,
			expectedFile: "test_conditional_viewer.docx",
		},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			data := document.NewTemplateData()
			data.SetVariable("username", tc.username)
			data.SetCondition("isAdmin", tc.isAdmin)
			data.SetCondition("isEditor", tc.isEditor)
			data.SetCondition("isViewer", tc.isViewer)
			data.SetCondition("hasSpecialAccess", tc.hasSpecial)

			doc, err := engine.RenderToDocument("user_permissions", data)
			if err != nil {
				t.Fatalf("Failed to render template for %s: %v", tc.name, err)
			}

			err = doc.Save("output/" + tc.expectedFile)
			if err != nil {
				t.Fatalf("Failed to save document for %s: %v", tc.name, err)
			}

			// 验证文档有内容
			if len(doc.Body.Elements) == 0 {
				t.Errorf("Expected document for %s to have content", tc.name)
			}
		})
	}
}

// testLoopStatementsIntegration 测试循环语句集成功能
func testLoopStatementsIntegration(t *testing.T) {
	engine := document.NewTemplateEngine()

	// 创建包含多种循环的模板
	templateContent := `库存管理报告

报告日期：{{reportDate}}

商品清单：
{{#each products}}
{{@index}}. 商品名称：{{name}}
   分类：{{category}}
   价格：{{price}} 元
   库存：{{stock}} 件
   {{#if lowStock}}⚠️ 库存不足{{/if}}
   {{#if popular}}🔥 热销商品{{/if}}

{{/each}}

供应商信息：
{{#each suppliers}}
供应商：{{name}}
联系电话：{{phone}}
地址：{{address}}
合作产品：
{{#each products}}
  - {{this}}
{{/each}}

{{/each}}

统计信息：
{{#each statistics}}
- {{key}}：{{value}}
{{/each}}`

	_, err := engine.LoadTemplate("inventory_report", templateContent)
	if err != nil {
		t.Fatalf("Failed to load template: %v", err)
	}

	// 创建测试数据
	data := document.NewTemplateData()
	data.SetVariable("reportDate", "2024年12月1日")

	// 商品列表
	products := []interface{}{
		map[string]interface{}{
			"name":     "笔记本电脑",
			"category": "电子产品",
			"price":    5999,
			"stock":    15,
			"lowStock": false,
			"popular":  true,
		},
		map[string]interface{}{
			"name":     "无线鼠标",
			"category": "电脑配件",
			"price":    199,
			"stock":    3,
			"lowStock": true,
			"popular":  false,
		},
		map[string]interface{}{
			"name":     "机械键盘",
			"category": "电脑配件",
			"price":    599,
			"stock":    25,
			"lowStock": false,
			"popular":  true,
		},
	}
	data.SetList("products", products)

	// 供应商列表
	suppliers := []interface{}{
		map[string]interface{}{
			"name":     "华硕科技",
			"phone":    "400-100-2000",
			"address":  "台北市信义区",
			"products": []interface{}{"笔记本电脑", "主板", "显卡"},
		},
		map[string]interface{}{
			"name":     "罗技公司",
			"phone":    "400-200-3000",
			"address":  "瑞士洛桑",
			"products": []interface{}{"无线鼠标", "键盘", "摄像头"},
		},
	}
	data.SetList("suppliers", suppliers)

	// 统计信息
	statistics := []interface{}{
		map[string]interface{}{
			"key":   "总商品数量",
			"value": "43件",
		},
		map[string]interface{}{
			"key":   "库存总价值",
			"value": "168,425元",
		},
		map[string]interface{}{
			"key":   "低库存商品",
			"value": "1种",
		},
	}
	data.SetList("statistics", statistics)

	// 渲染模板
	doc, err := engine.RenderToDocument("inventory_report", data)
	if err != nil {
		t.Fatalf("Failed to render template: %v", err)
	}

	// 保存文档
	err = doc.Save("output/test_loop_statements_integration.docx")
	if err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// 验证文档内容
	if len(doc.Body.Elements) == 0 {
		t.Error("Expected document to have content")
	}
}

// testTemplateInheritanceIntegration 测试模板继承集成功能
func testTemplateInheritanceIntegration(t *testing.T) {
	engine := document.NewTemplateEngine()

	// 创建基础报告模板
	baseTemplate := `{{companyName}} 业务报告

报告类型：{{reportType}}
生成时间：{{generateTime}}
报告期间：{{reportPeriod}}

================================`

	_, err := engine.LoadTemplate("base_report", baseTemplate)
	if err != nil {
		t.Fatalf("Failed to load base template: %v", err)
	}

	// 创建销售报告模板（继承基础模板）
	salesTemplate := `{{extends "base_report"}}

销售业绩：
总销售额：{{totalSales}} 元
订单数量：{{orderCount}} 个
平均客单价：{{averageOrder}} 元

销售团队表现：
{{#each salesTeam}}
销售员：{{name}}
销售额：{{sales}} 元
完成率：{{completion}}%

{{/each}}`

	_, err = engine.LoadTemplate("sales_report", salesTemplate)
	if err != nil {
		t.Fatalf("Failed to load sales template: %v", err)
	}

	// 创建财务报告模板（继承基础模板）
	financeTemplate := `{{extends "base_report"}}

财务状况：
营业收入：{{revenue}} 元
营业成本：{{cost}} 元
净利润：{{profit}} 元
利润率：{{profitRate}}%

现金流：
{{#each cashFlow}}
项目：{{item}}
金额：{{amount}} 元
类型：{{type}}

{{/each}}`

	_, err = engine.LoadTemplate("finance_report", financeTemplate)
	if err != nil {
		t.Fatalf("Failed to load finance template: %v", err)
	}

	// 创建通用数据
	commonData := document.NewTemplateData()
	commonData.SetVariable("companyName", "WordZero科技有限公司")
	commonData.SetVariable("generateTime", "2024年12月1日 16:00")
	commonData.SetVariable("reportPeriod", "2024年11月")

	// 测试销售报告
	salesData := document.NewTemplateData()
	salesData.Merge(commonData)
	salesData.SetVariable("reportType", "销售业绩报告")
	salesData.SetVariable("totalSales", "856,750")
	salesData.SetVariable("orderCount", 245)
	salesData.SetVariable("averageOrder", "3,497")

	salesTeam := []interface{}{
		map[string]interface{}{
			"name":       "张销售",
			"sales":      285600,
			"completion": 142,
		},
		map[string]interface{}{
			"name":       "李销售",
			"sales":      234800,
			"completion": 117,
		},
		map[string]interface{}{
			"name":       "王销售",
			"sales":      336350,
			"completion": 168,
		},
	}
	salesData.SetList("salesTeam", salesTeam)

	salesDoc, err := engine.RenderToDocument("sales_report", salesData)
	if err != nil {
		t.Fatalf("Failed to render sales report: %v", err)
	}

	err = salesDoc.Save("output/test_inheritance_sales_report.docx")
	if err != nil {
		t.Fatalf("Failed to save sales report: %v", err)
	}

	// 测试财务报告
	financeData := document.NewTemplateData()
	financeData.Merge(commonData)
	financeData.SetVariable("reportType", "财务状况报告")
	financeData.SetVariable("revenue", "1,245,600")
	financeData.SetVariable("cost", "723,400")
	financeData.SetVariable("profit", "522,200")
	financeData.SetVariable("profitRate", "41.9")

	cashFlow := []interface{}{
		map[string]interface{}{
			"item":   "销售收入",
			"amount": 1245600,
			"type":   "收入",
		},
		map[string]interface{}{
			"item":   "原料采购",
			"amount": -456000,
			"type":   "支出",
		},
		map[string]interface{}{
			"item":   "人员工资",
			"amount": -267400,
			"type":   "支出",
		},
	}
	financeData.SetList("cashFlow", cashFlow)

	financeDoc, err := engine.RenderToDocument("finance_report", financeData)
	if err != nil {
		t.Fatalf("Failed to render finance report: %v", err)
	}

	err = financeDoc.Save("output/test_inheritance_finance_report.docx")
	if err != nil {
		t.Fatalf("Failed to save finance report: %v", err)
	}

	// 验证文档内容
	if len(salesDoc.Body.Elements) == 0 {
		t.Error("Expected sales document to have content")
	}
	if len(financeDoc.Body.Elements) == 0 {
		t.Error("Expected finance document to have content")
	}
}

// testComplexTemplateIntegration 测试复杂模板集成功能
func testComplexTemplateIntegration(t *testing.T) {
	engine := document.NewTemplateEngine()

	// 创建复杂的年度报告模板
	complexTemplate := `{{companyName}} {{year}}年度报告

报告编号：{{reportNumber}}
发布日期：{{publishDate}}
审计机构：{{auditFirm}}

===============================

{{#if showExecutiveSummary}}
执行摘要：
{{executiveSummary}}

{{#if showKeyMetrics}}
关键指标：
{{#each keyMetrics}}
{{name}}：{{value}} {{unit}}
{{#if hasGrowth}}增长率：{{growth}}%{{/if}}

{{/each}}
{{/if}}
{{/if}}

业务部门报告：
{{#each departments}}
部门：{{name}}
负责人：{{manager}}
员工人数：{{employeeCount}} 人

{{#if showPerformance}}
业绩表现：
营收：{{revenue}} 万元
{{#if profitable}}✅ 盈利部门{{/if}}
{{#if needImprovement}}⚠️ 需要改进{{/if}}

主要成就：
{{#each achievements}}
- {{this}}
{{/each}}

{{#if showChallenges}}
面临挑战：
{{#each challenges}}
- {{challenge}}
  应对措施：{{solution}}
{{/each}}
{{/if}}
{{/if}}

{{/each}}

{{#if showFinancialData}}
财务数据：
总营收：{{totalRevenue}} 万元
总成本：{{totalCost}} 万元
净利润：{{netProfit}} 万元
{{#if profitGrowth}}利润同比增长：{{profitGrowthRate}}%{{/if}}

{{#if showInvestments}}
投资项目：
{{#each investments}}
项目：{{project}}
投资金额：{{amount}} 万元
预期回报：{{expectedReturn}}%
风险等级：{{riskLevel}}

{{/each}}
{{/if}}
{{/if}}

{{#if showFutureOutlook}}
未来展望：
{{futureOutlook}}

发展计划：
{{#each developmentPlans}}
- 时间：{{timeline}}
  目标：{{goal}}
  预算：{{budget}} 万元
{{/each}}
{{/if}}

===============================

报告编制：{{preparedBy}}
审核：{{reviewedBy}}
批准：{{approvedBy}}`

	_, err := engine.LoadTemplate("annual_report", complexTemplate)
	if err != nil {
		t.Fatalf("Failed to load complex template: %v", err)
	}

	// 创建复杂数据
	data := document.NewTemplateData()

	// 基础信息
	data.SetVariable("companyName", "WordZero科技有限公司")
	data.SetVariable("year", 2024)
	data.SetVariable("reportNumber", "AR-2024-001")
	data.SetVariable("publishDate", "2024年12月1日")
	data.SetVariable("auditFirm", "德勤会计师事务所")
	data.SetVariable("preparedBy", "财务部")
	data.SetVariable("reviewedBy", "CFO")
	data.SetVariable("approvedBy", "CEO")

	// 条件控制
	data.SetCondition("showExecutiveSummary", true)
	data.SetCondition("showKeyMetrics", true)
	data.SetCondition("showPerformance", true)
	data.SetCondition("showChallenges", true)
	data.SetCondition("showFinancialData", true)
	data.SetCondition("showInvestments", true)
	data.SetCondition("showFutureOutlook", true)

	// 执行摘要
	data.SetVariable("executiveSummary", "2024年是公司发展的重要一年，我们在技术创新、市场拓展和团队建设方面都取得了显著成就。")

	// 关键指标
	keyMetrics := []interface{}{
		map[string]interface{}{
			"name":      "年度营收",
			"value":     "2,450",
			"unit":      "万元",
			"hasGrowth": true,
			"growth":    "35.6",
		},
		map[string]interface{}{
			"name":      "客户数量",
			"value":     "1,250",
			"unit":      "家",
			"hasGrowth": true,
			"growth":    "28.9",
		},
		map[string]interface{}{
			"name":      "员工数量",
			"value":     "85",
			"unit":      "人",
			"hasGrowth": true,
			"growth":    "18.2",
		},
	}
	data.SetList("keyMetrics", keyMetrics)

	// 部门报告
	departments := []interface{}{
		map[string]interface{}{
			"name":            "研发部",
			"manager":         "张技术",
			"employeeCount":   35,
			"revenue":         850,
			"profitable":      true,
			"needImprovement": false,
			"achievements": []interface{}{
				"完成核心产品重构",
				"上线3个新功能模块",
				"技术专利申请5项",
			},
			"challenges": []interface{}{
				map[string]interface{}{
					"challenge": "人才招聘困难",
					"solution":  "提高薪酬待遇，完善培训体系",
				},
				map[string]interface{}{
					"challenge": "技术债务积累",
					"solution":  "制定重构计划，分阶段实施",
				},
			},
		},
		map[string]interface{}{
			"name":            "销售部",
			"manager":         "李销售",
			"employeeCount":   25,
			"revenue":         1200,
			"profitable":      true,
			"needImprovement": false,
			"achievements": []interface{}{
				"超额完成销售目标",
				"开拓5个新行业客户",
				"建立完善的CRM系统",
			},
			"challenges": []interface{}{
				map[string]interface{}{
					"challenge": "市场竞争激烈",
					"solution":  "差异化产品策略，提升服务质量",
				},
			},
		},
		map[string]interface{}{
			"name":            "运营部",
			"manager":         "王运营",
			"employeeCount":   15,
			"revenue":         400,
			"profitable":      false,
			"needImprovement": true,
			"achievements": []interface{}{
				"优化运营流程",
				"降低运营成本15%",
			},
			"challenges": []interface{}{
				map[string]interface{}{
					"challenge": "自动化程度不高",
					"solution":  "引入自动化工具，提升效率",
				},
			},
		},
	}
	data.SetList("departments", departments)

	// 财务数据
	data.SetVariable("totalRevenue", "2,450")
	data.SetVariable("totalCost", "1,680")
	data.SetVariable("netProfit", "770")
	data.SetCondition("profitGrowth", true)
	data.SetVariable("profitGrowthRate", "42.3")

	// 投资项目
	investments := []interface{}{
		map[string]interface{}{
			"project":        "AI智能分析系统",
			"amount":         300,
			"expectedReturn": 25.5,
			"riskLevel":      "中等",
		},
		map[string]interface{}{
			"project":        "云服务平台升级",
			"amount":         150,
			"expectedReturn": 18.2,
			"riskLevel":      "低",
		},
		map[string]interface{}{
			"project":        "海外市场拓展",
			"amount":         500,
			"expectedReturn": 35.8,
			"riskLevel":      "高",
		},
	}
	data.SetList("investments", investments)

	// 未来展望
	data.SetVariable("futureOutlook", "展望2025年，我们将继续专注于技术创新和市场拓展，预计营收将达到4000万元，成为行业领先企业。")

	// 发展计划
	developmentPlans := []interface{}{
		map[string]interface{}{
			"timeline": "2025年Q1",
			"goal":     "完成B轮融资",
			"budget":   200,
		},
		map[string]interface{}{
			"timeline": "2025年Q2",
			"goal":     "国际市场进入",
			"budget":   800,
		},
		map[string]interface{}{
			"timeline": "2025年Q3",
			"goal":     "团队扩展至150人",
			"budget":   500,
		},
		map[string]interface{}{
			"timeline": "2025年Q4",
			"goal":     "推出企业级产品",
			"budget":   1200,
		},
	}
	data.SetList("developmentPlans", developmentPlans)

	// 渲染复杂模板
	doc, err := engine.RenderToDocument("annual_report", data)
	if err != nil {
		t.Fatalf("Failed to render complex template: %v", err)
	}

	// 保存文档
	err = doc.Save("output/test_complex_template_integration.docx")
	if err != nil {
		t.Fatalf("Failed to save complex document: %v", err)
	}

	// 验证文档内容
	if len(doc.Body.Elements) == 0 {
		t.Error("Expected complex document to have content")
	}
}

// testDocumentToTemplateIntegration 测试文档转模板集成功能
func testDocumentToTemplateIntegration(t *testing.T) {
	// 创建包含模板变量的源文档
	sourceDoc := document.New()
	sourceDoc.AddParagraph("合同编号：{{contractNumber}}")
	sourceDoc.AddParagraph("甲方：{{partyA}}")
	sourceDoc.AddParagraph("乙方：{{partyB}}")
	sourceDoc.AddParagraph("")
	sourceDoc.AddParagraph("合同内容：")
	sourceDoc.AddParagraph("项目名称：{{projectName}}")
	sourceDoc.AddParagraph("项目金额：{{amount}} 元")
	sourceDoc.AddParagraph("开始日期：{{startDate}}")
	sourceDoc.AddParagraph("结束日期：{{endDate}}")
	sourceDoc.AddParagraph("")
	sourceDoc.AddParagraph("特别条款：")
	sourceDoc.AddParagraph("{{specialTerms}}")

	// 创建模板引擎
	engine := document.NewTemplateEngine()

	// 从文档创建模板
	template, err := engine.LoadTemplateFromDocument("contract_template", sourceDoc)
	if err != nil {
		t.Fatalf("Failed to create template from document: %v", err)
	}

	// 验证模板变量解析
	expectedVars := 8
	if len(template.Variables) != expectedVars {
		t.Errorf("Expected %d variables, got %d", expectedVars, len(template.Variables))
	}

	// 创建合同数据
	contractData := document.NewTemplateData()
	contractData.SetVariable("contractNumber", "WZ-2024-001")
	contractData.SetVariable("partyA", "WordZero科技有限公司")
	contractData.SetVariable("partyB", "客户公司A")
	contractData.SetVariable("projectName", "企业文档管理系统开发")
	contractData.SetVariable("amount", "500,000")
	contractData.SetVariable("startDate", "2024年12月1日")
	contractData.SetVariable("endDate", "2025年6月30日")
	contractData.SetVariable("specialTerms", "本项目包含完整的技术支持和培训服务，质保期为一年。")

	// 渲染合同
	contractDoc, err := engine.RenderToDocument("contract_template", contractData)
	if err != nil {
		t.Fatalf("Failed to render contract: %v", err)
	}

	// 保存合同文档
	err = contractDoc.Save("output/test_document_to_template_integration.docx")
	if err != nil {
		t.Fatalf("Failed to save contract document: %v", err)
	}

	// 验证文档内容
	if len(contractDoc.Body.Elements) == 0 {
		t.Error("Expected contract document to have content")
	}
}

// testStructDataBindingIntegration 测试结构体数据绑定集成功能
func testStructDataBindingIntegration(t *testing.T) {
	// 定义复杂的数据结构
	type Address struct {
		Street   string
		City     string
		Province string
		PostCode string
	}

	type Contact struct {
		Phone string
		Email string
		Fax   string
	}

	type Employee struct {
		ID         int
		Name       string
		Position   string
		Department string
		Salary     float64
		IsManager  bool
		HireDate   string
		Address    Address
		Contact    Contact
	}

	type Company struct {
		Name        string
		Industry    string
		Founded     int
		Employees   int
		Revenue     float64
		Address     Address
		Contact     Contact
		IsPublic    bool
		StockSymbol string
	}

	// 创建测试数据
	employee := Employee{
		ID:         1001,
		Name:       "张工程师",
		Position:   "高级软件工程师",
		Department: "技术部",
		Salary:     25000.00,
		IsManager:  false,
		HireDate:   "2023年5月15日",
		Address: Address{
			Street:   "科技园区1号楼A座",
			City:     "上海",
			Province: "上海市",
			PostCode: "200120",
		},
		Contact: Contact{
			Phone: "138-0013-8888",
			Email: "zhang.engineer@wordzero.com",
			Fax:   "021-6888-9999",
		},
	}

	company := Company{
		Name:      "WordZero科技有限公司",
		Industry:  "软件开发",
		Founded:   2023,
		Employees: 85,
		Revenue:   2450.0,
		Address: Address{
			Street:   "浦东新区科技园区",
			City:     "上海",
			Province: "上海市",
			PostCode: "200122",
		},
		Contact: Contact{
			Phone: "021-6666-8888",
			Email: "info@wordzero.com",
			Fax:   "021-6666-9999",
		},
		IsPublic:    false,
		StockSymbol: "",
	}

	// 创建模板引擎
	engine := document.NewTemplateEngine()

	// 创建员工档案模板
	templateContent := `员工档案详细信息

公司信息：
公司名称：{{name}}
所属行业：{{industry}}
成立年份：{{founded}}
员工总数：{{employees}}
年营收：{{revenue}} 万元
{{#if ispublic}}
股票代码：{{stocksymbol}}
{{/if}}

公司地址：
{{street}}
{{city}}, {{province}} {{postcode}}

联系方式：
电话：{{phone}}
邮箱：{{email}}
传真：{{fax}}

员工基本信息：
员工编号：{{id}}
姓名：{{name}}
职位：{{position}}
部门：{{department}}
月薪：{{salary}} 元
入职日期：{{hiredate}}
{{#if ismanager}}
职级：部门经理
{{/if}}

员工地址：
{{street}}
{{city}}, {{province}} {{postcode}}

员工联系方式：
电话：{{phone}}
邮箱：{{email}}
传真：{{fax}}`

	// 加载模板
	_, err := engine.LoadTemplate("employee_detail", templateContent)
	if err != nil {
		t.Fatalf("Failed to load employee detail template: %v", err)
	}

	// 创建模板数据
	data := document.NewTemplateData()

	// 从公司结构体创建数据
	err = data.FromStruct(company)
	if err != nil {
		t.Fatalf("Failed to create data from company struct: %v", err)
	}

	// 创建员工数据（手动设置以避免字段冲突）
	employeeData := document.NewTemplateData()
	err = employeeData.FromStruct(employee)
	if err != nil {
		t.Fatalf("Failed to create data from employee struct: %v", err)
	}

	// 手动设置员工相关变量
	data.SetVariable("id", employee.ID)
	data.SetVariable("name", employee.Name)
	data.SetVariable("position", employee.Position)
	data.SetVariable("department", employee.Department)
	data.SetVariable("salary", employee.Salary)
	data.SetVariable("hiredate", employee.HireDate)
	data.SetCondition("ismanager", employee.IsManager)

	// 设置员工地址和联系方式
	data.SetVariable("street", employee.Address.Street)
	data.SetVariable("city", employee.Address.City)
	data.SetVariable("province", employee.Address.Province)
	data.SetVariable("postcode", employee.Address.PostCode)
	data.SetVariable("phone", employee.Contact.Phone)
	data.SetVariable("email", employee.Contact.Email)
	data.SetVariable("fax", employee.Contact.Fax)

	// 设置公司相关变量（覆盖冲突字段）
	data.SetVariable("name", company.Name)
	data.SetVariable("industry", company.Industry)
	data.SetVariable("founded", company.Founded)
	data.SetVariable("employees", company.Employees)
	data.SetVariable("revenue", company.Revenue)
	data.SetCondition("ispublic", company.IsPublic)
	data.SetVariable("stocksymbol", company.StockSymbol)

	// 设置公司地址和联系方式（不同的变量名以避免冲突）
	// 在实际应用中，可以使用更好的方式处理这种冲突

	// 渲染模板
	doc, err := engine.RenderToDocument("employee_detail", data)
	if err != nil {
		t.Fatalf("Failed to render employee detail: %v", err)
	}

	// 保存文档
	err = doc.Save("output/test_struct_data_binding_integration.docx")
	if err != nil {
		t.Fatalf("Failed to save employee detail document: %v", err)
	}

	// 验证文档内容
	if len(doc.Body.Elements) == 0 {
		t.Error("Expected employee detail document to have content")
	}
}

package main

import (
	"log"
	"os"
	"path/filepath"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
)

func renderTemplate(data map[string]interface{}) error {
	// 创建新的模板渲染器
	renderer := document.NewTemplateRenderer()
	renderer.SetLogging(true) // 启用详细日志

	// 检查模板文件是否存在
	templatePath := "default_template.docx"
	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		log.Printf("模板文件 %s 不存在", templatePath)
		return err
	}

	// 从文件加载模板
	_, err := renderer.LoadTemplateFromFile("debug_template", templatePath)
	if err != nil {
		log.Printf("加载模板失败: %v", err)
		return err
	}

	// 分析模板结构
	analysis, err := renderer.AnalyzeTemplate("debug_template")
	if err != nil {
		log.Printf("分析模板失败: %v", err)
		return err
	}

	log.Printf("=== 模板分析结果 ===")
	log.Printf("模板名称: %s", analysis.TemplateName)
	log.Printf("变量数量: %d", len(analysis.Variables))
	log.Printf("列表数量: %d", len(analysis.Lists))
	log.Printf("条件数量: %d", len(analysis.Conditions))
	log.Printf("表格数量: %d", len(analysis.Tables))

	for i, table := range analysis.Tables {
		log.Printf("表格 %d: 行=%d, 列=%d, 模板=%v",
			i, table.RowCount, table.ColCount, table.HasTemplate)
		if table.HasTemplate {
			log.Printf("  - 模板行索引: %d", table.TemplateRowIndex)
			log.Printf("  - 循环变量: %v", table.LoopVariables)
		}
	}

	// 创建模板数据
	templateData := document.NewTemplateData()

	// 将数据转换为模板数据格式
	for key, value := range data {
		switch v := value.(type) {
		case []interface{}:
			templateData.SetList(key, v)
		case bool:
			templateData.SetCondition(key, v)
		default:
			templateData.SetVariable(key, v)
		}
	}

	// 渲染模板
	resultDoc, err := renderer.RenderTemplate("debug_template", templateData)
	if err != nil {
		log.Printf("渲染模板失败: %v", err)
		return err
	}

	// 确保输出目录存在
	outputDir := "./"
	err = os.MkdirAll(outputDir, 0755)
	if err != nil {
		log.Printf("创建输出目录失败: %v", err)
		return err
	}

	// 保存渲染结果
	outputPath := filepath.Join(outputDir, "debug_rendered_result.docx")
	err = resultDoc.Save(outputPath)
	if err != nil {
		log.Printf("保存渲染结果失败: %v", err)
		return err
	}

	log.Printf("渲染完成！输出文件: %s", outputPath)
	return nil
}

func main() {
	data := map[string]interface{}{
		"data_1": "普通变量值",
		"data_2": "字体变量值",
		"data_3": []interface{}{
			map[string]interface{}{
				"data_4": "表格项1列1",
				"data_5": "表格项1列2",
				"data_6": "表格项1列3",
			},
			map[string]interface{}{
				"data_4": "表格项2列1",
				"data_5": "表格项2列2",
				"data_6": "表格项2列3",
			},
			map[string]interface{}{
				"data_4": "表格项3列1",
				"data_5": "表格项3列2",
				"data_6": "表格项3列3",
			},
		},
		"data_7": []interface{}{
			map[string]interface{}{
				"data_8":  "非表格循环项1值",
				"data_9":  "非表格循环项1描述",
				"data_10": "非表格循环项1备注",
			},
			map[string]interface{}{
				"data_8":  "非表格循环项2值",
				"data_9":  "非表格循环项2描述",
				"data_10": "非表格循环项2备注",
			},
		},
	}

	if err := renderTemplate(data); err != nil {
		log.Fatalf("渲染模板时出错: %v", err)
	}
}

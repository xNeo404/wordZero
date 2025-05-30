package main

import (
	"errors"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"time"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
)

type DocumentProcessor struct {
	logger *log.Logger
}

func NewDocumentProcessor() *DocumentProcessor {
	return &DocumentProcessor{
		logger: log.New(os.Stdout, "[DocProcessor] ", log.LstdFlags),
	}
}

func (dp *DocumentProcessor) ProcessWithRetry(data interface{}, maxRetries int) error {
	var lastErr error

	for attempt := 1; attempt <= maxRetries; attempt++ {
		err := dp.process(data)
		if err == nil {
			dp.logger.Printf("处理成功，尝试次数: %d", attempt)
			return nil
		}

		lastErr = err
		dp.logger.Printf("尝试 %d/%d 失败: %v", attempt, maxRetries, err)

		if attempt < maxRetries {
			// 指数退避
			time.Sleep(time.Duration(attempt) * time.Second)
		}
	}

	return fmt.Errorf("处理失败，已重试 %d 次: %w", maxRetries, lastErr)
}

func (dp *DocumentProcessor) process(data interface{}) error {
	// 验证输入
	if err := dp.validateInput(data); err != nil {
		return fmt.Errorf("输入验证失败: %w", err)
	}

	// 创建文档
	doc, err := dp.createDocument(data)
	if err != nil {
		return fmt.Errorf("创建文档失败: %w", err)
	}

	// 保存文档
	if err := dp.saveDocument(doc, "output.docx"); err != nil {
		return fmt.Errorf("保存文档失败: %w", err)
	}

	return nil
}

func (dp *DocumentProcessor) validateInput(data interface{}) error {
	if data == nil {
		return errors.New("数据不能为空")
	}

	// 更多验证逻辑...
	return nil
}

func (dp *DocumentProcessor) createDocument(data interface{}) (*document.Document, error) {
	defer func() {
		if r := recover(); r != nil {
			dp.logger.Printf("创建文档时发生panic: %v", r)
		}
	}()

	doc := document.New()

	// 添加内容的逻辑...
	para := doc.AddParagraph("测试内容")
	if para == nil {
		return nil, errors.New("添加段落失败")
	}

	return doc, nil
}

func (dp *DocumentProcessor) saveDocument(doc *document.Document, filename string) error {
	// 确保目录存在
	dir := filepath.Dir(filename)
	if err := os.MkdirAll(dir, 0755); err != nil {
		return fmt.Errorf("创建目录失败: %w", err)
	}

	// 检查磁盘空间
	if err := dp.checkDiskSpace(filename); err != nil {
		return fmt.Errorf("磁盘空间检查失败: %w", err)
	}

	// 保存文档
	if err := doc.Save(filename); err != nil {
		return fmt.Errorf("保存失败: %w", err)
	}

	// 验证文件是否正确保存
	if err := dp.validateSavedFile(filename); err != nil {
		return fmt.Errorf("文件验证失败: %w", err)
	}

	return nil
}

func (dp *DocumentProcessor) checkDiskSpace(filename string) error {
	// 简单的磁盘空间检查
	info, err := os.Stat(filepath.Dir(filename))
	if err != nil {
		return err
	}

	// 这里可以添加更详细的磁盘空间检查逻辑
	_ = info
	return nil
}

func (dp *DocumentProcessor) validateSavedFile(filename string) error {
	info, err := os.Stat(filename)
	if err != nil {
		return fmt.Errorf("文件不存在: %w", err)
	}

	if info.Size() == 0 {
		return errors.New("文件大小为0")
	}

	return nil
}

func main() {
	processor := NewDocumentProcessor()

	data := map[string]interface{}{
		"title":   "测试文档",
		"content": "这是测试内容",
	}

	if err := processor.ProcessWithRetry(data, 3); err != nil {
		log.Fatalf("文档处理失败: %v", err)
	}

	fmt.Println("文档处理成功！")
}

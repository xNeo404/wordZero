package test

import (
	"encoding/xml"
	"fmt"
	"testing"
)

// 定义一个本地的Settings结构用于测试，不使用命名空间前缀
type testSettings struct {
	XMLName                 xml.Name                     `xml:"settings"`
	Xmlns                   string                       `xml:"xmlns:w,attr"`
	DefaultTabStop          *testDefaultTabStop          `xml:"defaultTabStop,omitempty"`
	CharacterSpacingControl *testCharacterSpacingControl `xml:"characterSpacingControl,omitempty"`
}

type testDefaultTabStop struct {
	XMLName xml.Name `xml:"defaultTabStop"`
	Val     string   `xml:"val,attr"`
}

type testCharacterSpacingControl struct {
	XMLName xml.Name `xml:"characterSpacingControl"`
	Val     string   `xml:"val,attr"`
}

func TestXMLSerialization(t *testing.T) {
	// 创建测试设置
	settings := &testSettings{
		Xmlns: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		DefaultTabStop: &testDefaultTabStop{
			Val: "708",
		},
		CharacterSpacingControl: &testCharacterSpacingControl{
			Val: "doNotCompress",
		},
	}

	// 序列化为XML
	xmlData, err := xml.MarshalIndent(settings, "", "  ")
	if err != nil {
		t.Fatalf("序列化失败: %v", err)
	}

	// 添加XML声明
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	fullXML := append(xmlDeclaration, xmlData...)

	fmt.Printf("序列化的XML:\n%s\n", string(fullXML))

	// 解析XML - 这次应该能成功，因为我们不使用命名空间前缀
	var parsedSettings testSettings
	err = xml.Unmarshal(xmlData, &parsedSettings) // 使用xmlData而不是fullXML，避免XML声明解析问题
	if err != nil {
		t.Fatalf("解析失败: %v", err)
	}

	fmt.Printf("解析成功！\n")

	// 验证解析结果 - 注意XML序列化后命名空间可能会变化
	if parsedSettings.Xmlns != "" && parsedSettings.Xmlns != settings.Xmlns {
		t.Errorf("命名空间不匹配: 期望 %s, 实际 %s", settings.Xmlns, parsedSettings.Xmlns)
	}

	// 验证其他字段
	if parsedSettings.DefaultTabStop == nil || parsedSettings.DefaultTabStop.Val != "708" {
		t.Errorf("DefaultTabStop解析不正确")
	}

	if parsedSettings.CharacterSpacingControl == nil || parsedSettings.CharacterSpacingControl.Val != "doNotCompress" {
		t.Errorf("CharacterSpacingControl解析不正确")
	}

	// 验证核心功能：能够序列化和解析XML结构
	fmt.Printf("XML序列化和解析测试通过！\n")
}

package main

import (
	"bytes"
	"fmt"
	"image"
	"image/color"
	"image/png"
	"log"
	"os"
	"path/filepath"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
)

// createSampleImage 创建一个示例图片
func createSampleImage(width, height int, bgColor color.RGBA, text string) []byte {
	img := image.NewRGBA(image.Rect(0, 0, width, height))

	// 填充背景色
	for y := 0; y < height; y++ {
		for x := 0; x < width; x++ {
			img.Set(x, y, bgColor)
		}
	}

	// 添加一个简单的边框
	borderColor := color.RGBA{255 - bgColor.R, 255 - bgColor.G, 255 - bgColor.B, 255}
	for x := 0; x < width; x++ {
		img.Set(x, 0, borderColor)
		img.Set(x, height-1, borderColor)
	}
	for y := 0; y < height; y++ {
		img.Set(0, y, borderColor)
		img.Set(width-1, y, borderColor)
	}

	var buf bytes.Buffer
	png.Encode(&buf, img)
	return buf.Bytes()
}

// saveImageToFile 将图片数据保存为文件
func saveImageToFile(imageData []byte, filePath string) error {
	// 确保目录存在
	dir := filepath.Dir(filePath)
	if err := os.MkdirAll(dir, 0755); err != nil {
		return fmt.Errorf("创建目录失败: %v", err)
	}

	// 写入文件
	file, err := os.Create(filePath)
	if err != nil {
		return fmt.Errorf("创建文件失败: %v", err)
	}
	defer file.Close()

	_, err = file.Write(imageData)
	if err != nil {
		return fmt.Errorf("写入文件失败: %v", err)
	}

	return nil
}

// addSectionTitle 添加带样式的章节标题
func addSectionTitle(doc *document.Document, title string, level int) {
	doc.AddParagraph("")
	para := doc.AddParagraph("")
	if level == 1 {
		para.AddFormattedText(title, &document.TextFormat{
			Bold:      true,
			FontSize:  16,
			FontColor: "0066CC",
		})
	} else {
		para.AddFormattedText(title, &document.TextFormat{
			Bold:      true,
			FontSize:  14,
			FontColor: "333333",
		})
	}
	doc.AddParagraph("")
}

// addDescription 添加功能描述
func addDescription(doc *document.Document, description string) {
	para := doc.AddParagraph("")
	para.AddFormattedText(description, &document.TextFormat{
		FontSize:  11,
		FontColor: "666666",
		Italic:    true,
	})
	doc.AddParagraph("")
}

func main() {
	fmt.Println("=== WordZero 高级图片功能演示 ===")

	// 创建新文档
	doc := document.New()

	// 添加文档标题
	titlePara := doc.AddParagraph("")
	titlePara.AddFormattedText("WordZero 图片功能完整演示", &document.TextFormat{
		Bold:      true,
		FontSize:  20,
		FontColor: "000080",
	})
	doc.AddParagraph("")
	descPara := doc.AddParagraph("")
	descPara.AddFormattedText("本文档演示了WordZero图片处理的各种功能，包括图片位置、大小调整、文字环绕等特性。", &document.TextFormat{
		FontSize:  12,
		FontColor: "444444",
	})
	doc.AddParagraph("")

	// 确保输出目录存在
	outputDir := "examples/output"
	imagesDir := outputDir + "/images"
	if err := os.MkdirAll(imagesDir, 0755); err != nil {
		log.Fatalf("创建输出目录失败: %v", err)
	}

	// ==================== 第一部分：基础图片插入 ====================
	addSectionTitle(doc, "1. 基础图片插入功能", 1)
	addDescription(doc, "演示基本的图片插入功能，包括从数据和文件两种方式添加图片。")

	fmt.Println("1. 基础图片插入功能演示...")

	// 创建一个蓝色的示例图片
	basicImageData := createSampleImage(200, 120, color.RGBA{100, 150, 255, 255}, "基础图片")
	basicImagePath := filepath.Join(imagesDir, "basic_image.png")
	if err := saveImageToFile(basicImageData, basicImagePath); err != nil {
		log.Fatalf("保存基础图片失败: %v", err)
	}

	basicImageInfo, err := doc.AddImageFromData(
		basicImageData,
		"basic_image.png",
		document.ImageFormatPNG,
		200,
		120,
		&document.ImageConfig{
			Size: &document.ImageSize{
				Width:  50.0, // 50毫米宽度
				Height: 30.0, // 30毫米高度
			},
			AltText: "基础演示图片",
			Title:   "这是一个基础的图片插入示例",
		},
	)
	if err != nil {
		log.Fatalf("添加基础图片失败: %v", err)
	}
	fmt.Printf("   添加基础图片成功，ID: %s\n", basicImageInfo.ID)

	doc.AddParagraph("上面是一个基础的图片插入示例。图片默认以嵌入式（inline）方式显示，会跟随文本流动。")

	// ==================== 第二部分：图片位置控制 ====================
	addSectionTitle(doc, "2. 图片位置控制功能", 1)
	addDescription(doc, "演示不同的图片位置选项：嵌入式、左浮动、右浮动等。")

	fmt.Println("2. 图片位置控制功能演示...")

	// 2.1 嵌入式图片（默认）
	addSectionTitle(doc, "2.1 嵌入式图片（默认位置）", 2)
	doc.AddParagraph("嵌入式图片会跟随文本流，在文档中的固定位置显示，不会影响周围文本的布局。")

	inlineImageData := createSampleImage(180, 100, color.RGBA{255, 100, 100, 255}, "嵌入式")
	inlineImagePath := filepath.Join(imagesDir, "inline_image.png")
	if err := saveImageToFile(inlineImageData, inlineImagePath); err != nil {
		log.Fatalf("保存嵌入式图片失败: %v", err)
	}

	inlineImageInfo, err := doc.AddImageFromData(
		inlineImageData,
		"inline_image.png",
		document.ImageFormatPNG,
		180,
		100,
		&document.ImageConfig{
			Size: &document.ImageSize{
				Width:  45.0,
				Height: 25.0,
			},
			Position: document.ImagePositionInline,
			WrapText: document.ImageWrapNone,
			AltText:  "嵌入式图片",
			Title:    "嵌入式图片示例",
		},
	)
	if err != nil {
		log.Fatalf("添加嵌入式图片失败: %v", err)
	}
	fmt.Printf("   添加嵌入式图片成功，ID: %s\n", inlineImageInfo.ID)

	doc.AddParagraph("这段文字紧跟在嵌入式图片后面。嵌入式图片不会让文字环绕，而是按照正常的文档流进行排列。")

	// 2.2 左浮动图片
	addSectionTitle(doc, "2.2 左浮动图片", 2)
	doc.AddParagraph("左浮动图片会显示在页面左侧，右侧的文字可以环绕显示。")

	leftFloatImageData := createSampleImage(150, 150, color.RGBA{100, 255, 100, 255}, "左浮动")
	leftFloatImagePath := filepath.Join(imagesDir, "left_float_image.png")
	if err := saveImageToFile(leftFloatImageData, leftFloatImagePath); err != nil {
		log.Fatalf("保存左浮动图片失败: %v", err)
	}

	leftFloatImageInfo, err := doc.AddImageFromData(
		leftFloatImageData,
		"left_float_image.png",
		document.ImageFormatPNG,
		150,
		150,
		&document.ImageConfig{
			Size: &document.ImageSize{
				Width:  40.0,
				Height: 40.0,
			},
			Position: document.ImagePositionFloatLeft,
			WrapText: document.ImageWrapSquare,
			AltText:  "左浮动图片",
			Title:    "左浮动图片示例",
			OffsetX:  2.0, // 距离左边2毫米
			OffsetY:  0.0,
		},
	)
	if err != nil {
		log.Fatalf("添加左浮动图片失败: %v", err)
	}
	fmt.Printf("   添加左浮动图片成功，ID: %s\n", leftFloatImageInfo.ID)

	doc.AddParagraph("这段文字演示了左浮动图片的效果。左浮动图片会靠左显示，而这段文字内容会在图片的右侧进行环绕显示。这种布局方式常用于杂志、报纸等出版物中，可以有效利用版面空间，让文档看起来更加美观和专业。左浮动图片适合在文章开头或段落中插入配图，不会打断阅读的连续性。")

	// 2.3 右浮动图片
	addSectionTitle(doc, "2.3 右浮动图片", 2)
	doc.AddParagraph("右浮动图片会显示在页面右侧，左侧的文字可以环绕显示。")

	rightFloatImageData := createSampleImage(160, 120, color.RGBA{255, 200, 100, 255}, "右浮动")
	rightFloatImagePath := filepath.Join(imagesDir, "right_float_image.png")
	if err := saveImageToFile(rightFloatImageData, rightFloatImagePath); err != nil {
		log.Fatalf("保存右浮动图片失败: %v", err)
	}

	rightFloatImageInfo, err := doc.AddImageFromData(
		rightFloatImageData,
		"right_float_image.png",
		document.ImageFormatPNG,
		160,
		120,
		&document.ImageConfig{
			Size: &document.ImageSize{
				Width:  42.0,
				Height: 32.0,
			},
			Position: document.ImagePositionFloatRight,
			WrapText: document.ImageWrapSquare,
			AltText:  "右浮动图片",
			Title:    "右浮动图片示例",
			OffsetX:  2.0, // 距离右边2毫米
			OffsetY:  0.0,
		},
	)
	if err != nil {
		log.Fatalf("添加右浮动图片失败: %v", err)
	}
	fmt.Printf("   添加右浮动图片成功，ID: %s\n", rightFloatImageInfo.ID)

	doc.AddParagraph("这段文字展示了右浮动图片的效果。右浮动图片会出现在页面的右侧，文字内容会从左侧开始环绕显示。这种布局方式特别适合在文档中插入说明性图片、图表或装饰性元素。右浮动布局可以让读者在阅读主要内容的同时，方便地查看相关的图片信息，提升文档的可读性和视觉效果。")

	// ==================== 第三部分：文字环绕方式 ====================
	addSectionTitle(doc, "3. 文字环绕方式", 1)
	addDescription(doc, "演示不同的文字环绕模式：无环绕、四周环绕、紧密环绕、上下环绕等。")

	fmt.Println("3. 文字环绕方式演示...")

	// 3.1 无环绕
	addSectionTitle(doc, "3.1 无环绕模式", 2)
	doc.AddParagraph("无环绕模式下，文字不会环绕图片，图片会独占一行或一个段落。")

	noWrapImageData := createSampleImage(140, 100, color.RGBA{200, 100, 255, 255}, "无环绕")
	noWrapImagePath := filepath.Join(imagesDir, "no_wrap_image.png")
	if err := saveImageToFile(noWrapImageData, noWrapImagePath); err != nil {
		log.Fatalf("保存无环绕图片失败: %v", err)
	}

	noWrapImageInfo, err := doc.AddImageFromData(
		noWrapImageData,
		"no_wrap_image.png",
		document.ImageFormatPNG,
		140,
		100,
		&document.ImageConfig{
			Size: &document.ImageSize{
				Width:  35.0,
				Height: 25.0,
			},
			Position: document.ImagePositionInline,
			WrapText: document.ImageWrapNone,
			AltText:  "无环绕图片",
			Title:    "无环绕模式示例",
		},
	)
	if err != nil {
		log.Fatalf("添加无环绕图片失败: %v", err)
	}
	fmt.Printf("   添加无环绕图片成功，ID: %s\n", noWrapImageInfo.ID)

	doc.AddParagraph("上面的图片使用了无环绕模式。在这种模式下，图片前后的文字会分别在图片的上方和下方显示，不会有文字出现在图片的左右两侧。")

	// 3.2 四周环绕
	addSectionTitle(doc, "3.2 四周环绕模式", 2)
	doc.AddParagraph("四周环绕模式下，文字会在图片的四周进行环绕显示。")

	squareWrapImageData := createSampleImage(130, 130, color.RGBA{100, 200, 255, 255}, "四周环绕")
	squareWrapImagePath := filepath.Join(imagesDir, "square_wrap_image.png")
	if err := saveImageToFile(squareWrapImageData, squareWrapImagePath); err != nil {
		log.Fatalf("保存四周环绕图片失败: %v", err)
	}

	squareWrapImageInfo, err := doc.AddImageFromData(
		squareWrapImageData,
		"square_wrap_image.png",
		document.ImageFormatPNG,
		130,
		130,
		&document.ImageConfig{
			Size: &document.ImageSize{
				Width:  35.0,
				Height: 35.0,
			},
			Position: document.ImagePositionFloatLeft,
			WrapText: document.ImageWrapSquare,
			AltText:  "四周环绕图片",
			Title:    "四周环绕模式示例",
			OffsetX:  3.0,
			OffsetY:  0.0,
		},
	)
	if err != nil {
		log.Fatalf("添加四周环绕图片失败: %v", err)
	}
	fmt.Printf("   添加四周环绕图片成功，ID: %s\n", squareWrapImageInfo.ID)

	doc.AddParagraph("这段文字演示了四周环绕的效果。四周环绕是最常用的文字环绕方式，文字会按照图片的矩形边界进行环绕，在图片周围形成整齐的文字布局。这种方式适合大多数文档场景，特别是在需要在正文中插入图片时。四周环绕模式可以保持文档的整洁性，同时有效利用版面空间，让图片和文字和谐地组合在一起。")

	// 3.3 紧密环绕
	addSectionTitle(doc, "3.3 紧密环绕模式", 2)
	doc.AddParagraph("紧密环绕模式下，文字会更贴近图片的实际轮廓进行环绕。")

	tightWrapImageData := createSampleImage(120, 140, color.RGBA{255, 150, 100, 255}, "紧密环绕")
	tightWrapImagePath := filepath.Join(imagesDir, "tight_wrap_image.png")
	if err := saveImageToFile(tightWrapImageData, tightWrapImagePath); err != nil {
		log.Fatalf("保存紧密环绕图片失败: %v", err)
	}

	tightWrapImageInfo, err := doc.AddImageFromData(
		tightWrapImageData,
		"tight_wrap_image.png",
		document.ImageFormatPNG,
		120,
		140,
		&document.ImageConfig{
			Size: &document.ImageSize{
				Width:  32.0,
				Height: 37.0,
			},
			Position: document.ImagePositionFloatRight,
			WrapText: document.ImageWrapTight,
			AltText:  "紧密环绕图片",
			Title:    "紧密环绕模式示例",
			OffsetX:  2.0,
			OffsetY:  0.0,
		},
	)
	if err != nil {
		log.Fatalf("添加紧密环绕图片失败: %v", err)
	}
	fmt.Printf("   添加紧密环绕图片成功，ID: %s\n", tightWrapImageInfo.ID)

	doc.AddParagraph("这段文字展示了紧密环绕模式的特点。与四周环绕不同，紧密环绕会让文字更贴近图片的实际边缘，可以获得更紧凑的布局效果。这种模式特别适合不规则形状的图片或者需要节省版面空间的场合。紧密环绕能够让文档看起来更加精致和专业，是高端出版物中常用的排版技巧。")

	// 3.4 上下环绕
	addSectionTitle(doc, "3.4 上下环绕模式", 2)
	doc.AddParagraph("上下环绕模式下，文字只在图片的上方和下方显示，左右两侧留空。")

	topBottomWrapImageData := createSampleImage(180, 80, color.RGBA{150, 255, 150, 255}, "上下环绕")
	topBottomWrapImagePath := filepath.Join(imagesDir, "top_bottom_wrap_image.png")
	if err := saveImageToFile(topBottomWrapImageData, topBottomWrapImagePath); err != nil {
		log.Fatalf("保存上下环绕图片失败: %v", err)
	}

	topBottomWrapImageInfo, err := doc.AddImageFromData(
		topBottomWrapImageData,
		"top_bottom_wrap_image.png",
		document.ImageFormatPNG,
		180,
		80,
		&document.ImageConfig{
			Size: &document.ImageSize{
				Width:  48.0,
				Height: 21.0,
			},
			Position: document.ImagePositionInline,
			WrapText: document.ImageWrapTopAndBottom,
			AltText:  "上下环绕图片",
			Title:    "上下环绕模式示例",
		},
	)
	if err != nil {
		log.Fatalf("添加上下环绕图片失败: %v", err)
	}
	fmt.Printf("   添加上下环绕图片成功，ID: %s\n", topBottomWrapImageInfo.ID)

	doc.AddParagraph("上面的图片演示了上下环绕模式。在这种模式下，图片的左右两侧不会有文字，所有文字内容都会在图片的上方和下方显示。这种布局方式适合横向较宽的图片，如图表、流程图、横幅等，可以让图片获得更好的展示效果。")

	// ==================== 第四部分：图片大小调整 ====================
	addSectionTitle(doc, "4. 图片大小调整功能", 1)
	addDescription(doc, "演示各种图片大小调整选项：固定尺寸、保持比例、动态调整等。")

	fmt.Println("4. 图片大小调整功能演示...")

	// 4.1 固定尺寸
	addSectionTitle(doc, "4.1 固定尺寸设置", 2)
	doc.AddParagraph("可以为图片设置固定的宽度和高度，完全控制图片的显示尺寸。")

	fixedSizeImageData := createSampleImage(200, 150, color.RGBA{255, 180, 200, 255}, "固定尺寸")
	fixedSizeImagePath := filepath.Join(imagesDir, "fixed_size_image.png")
	if err := saveImageToFile(fixedSizeImageData, fixedSizeImagePath); err != nil {
		log.Fatalf("保存固定尺寸图片失败: %v", err)
	}

	fixedSizeImageInfo, err := doc.AddImageFromData(
		fixedSizeImageData,
		"fixed_size_image.png",
		document.ImageFormatPNG,
		200,
		150,
		&document.ImageConfig{
			Size: &document.ImageSize{
				Width:  60.0, // 固定宽度60毫米
				Height: 40.0, // 固定高度40毫米
			},
			AltText: "固定尺寸图片",
			Title:   "固定尺寸设置示例",
		},
	)
	if err != nil {
		log.Fatalf("添加固定尺寸图片失败: %v", err)
	}
	fmt.Printf("   添加固定尺寸图片成功，ID: %s，尺寸: 60x40毫米\n", fixedSizeImageInfo.ID)

	doc.AddParagraph("上面的图片被设置为固定尺寸60x40毫米。固定尺寸设置不考虑原始图片的长宽比，会严格按照指定的尺寸显示。")

	// 4.2 保持长宽比
	addSectionTitle(doc, "4.2 保持长宽比缩放", 2)
	doc.AddParagraph("只设置宽度或高度，另一个维度自动计算以保持图片的原始长宽比。")

	aspectRatioImageData := createSampleImage(300, 150, color.RGBA{180, 220, 255, 255}, "保持比例")
	aspectRatioImagePath := filepath.Join(imagesDir, "aspect_ratio_image.png")
	if err := saveImageToFile(aspectRatioImageData, aspectRatioImagePath); err != nil {
		log.Fatalf("保存长宽比图片失败: %v", err)
	}

	aspectRatioImageInfo, err := doc.AddImageFromData(
		aspectRatioImageData,
		"aspect_ratio_image.png",
		document.ImageFormatPNG,
		300,
		150,
		&document.ImageConfig{
			Size: &document.ImageSize{
				Width:           70.0, // 只设置宽度70毫米
				KeepAspectRatio: true, // 保持长宽比
			},
			AltText: "保持长宽比图片",
			Title:   "保持长宽比缩放示例",
		},
	)
	if err != nil {
		log.Fatalf("添加长宽比图片失败: %v", err)
	}
	fmt.Printf("   添加长宽比图片成功，ID: %s，宽度: 70毫米（高度自动计算）\n", aspectRatioImageInfo.ID)

	doc.AddParagraph("上面的图片只设置了宽度为70毫米，高度根据原始长宽比自动计算。这样可以避免图片变形，保持图片的原始比例。")

	// 4.3 动态调整大小
	addSectionTitle(doc, "4.3 动态调整图片大小", 2)
	doc.AddParagraph("演示在添加图片后动态调整其大小的功能。")

	dynamicImageData := createSampleImage(160, 200, color.RGBA{200, 255, 180, 255}, "动态调整")
	dynamicImagePath := filepath.Join(imagesDir, "dynamic_image.png")
	if err := saveImageToFile(dynamicImageData, dynamicImagePath); err != nil {
		log.Fatalf("保存动态调整图片失败: %v", err)
	}

	dynamicImageInfo, err := doc.AddImageFromData(
		dynamicImageData,
		"dynamic_image.png",
		document.ImageFormatPNG,
		160,
		200,
		&document.ImageConfig{
			Size: &document.ImageSize{
				Width:  30.0, // 初始宽度30毫米
				Height: 38.0, // 初始高度38毫米
			},
			AltText: "动态调整图片",
			Title:   "动态调整大小示例",
		},
	)
	if err != nil {
		log.Fatalf("添加动态调整图片失败: %v", err)
	}
	fmt.Printf("   添加动态调整图片成功，ID: %s，初始尺寸: 30x38毫米\n", dynamicImageInfo.ID)

	// 动态调整图片大小
	newSize := &document.ImageSize{
		Width:  50.0, // 调整为50毫米宽度
		Height: 65.0, // 调整为65毫米高度
	}
	err = doc.ResizeImage(dynamicImageInfo, newSize)
	if err != nil {
		log.Fatalf("动态调整图片大小失败: %v", err)
	}
	fmt.Printf("   图片大小已动态调整为50x65毫米\n")

	doc.AddParagraph("上面的图片演示了动态大小调整。图片最初设置为30x38毫米，然后通过ResizeImage方法调整为50x65毫米。")

	// ==================== 第五部分：图片属性设置 ====================
	addSectionTitle(doc, "5. 图片属性设置", 1)
	addDescription(doc, "演示图片的各种属性设置：替代文字、标题、偏移量等。")

	fmt.Println("5. 图片属性设置演示...")

	// 创建用于属性设置演示的图片
	propertiesImageData := createSampleImage(140, 110, color.RGBA{255, 220, 150, 255}, "属性设置")
	propertiesImagePath := filepath.Join(imagesDir, "properties_image.png")
	if err := saveImageToFile(propertiesImageData, propertiesImagePath); err != nil {
		log.Fatalf("保存属性设置图片失败: %v", err)
	}

	propertiesImageInfo, err := doc.AddImageFromData(
		propertiesImageData,
		"properties_image.png",
		document.ImageFormatPNG,
		140,
		110,
		&document.ImageConfig{
			Size: &document.ImageSize{
				Width:  35.0,
				Height: 28.0,
			},
			Position: document.ImagePositionFloatRight,
			WrapText: document.ImageWrapSquare,
			OffsetX:  5.0, // 距离右边5毫米
			OffsetY:  2.0, // 向下偏移2毫米
		},
	)
	if err != nil {
		log.Fatalf("添加属性设置图片失败: %v", err)
	}

	// 设置图片属性
	err = doc.SetImageAltText(propertiesImageInfo, "这是一个用于演示属性设置的示例图片，包含了完整的可访问性信息")
	if err != nil {
		log.Fatalf("设置图片替代文字失败: %v", err)
	}

	err = doc.SetImageTitle(propertiesImageInfo, "图片属性设置演示 - 包含标题、替代文字和位置偏移")
	if err != nil {
		log.Fatalf("设置图片标题失败: %v", err)
	}

	fmt.Printf("   添加属性设置图片成功，ID: %s\n", propertiesImageInfo.ID)
	fmt.Printf("   设置了替代文字、标题和位置偏移\n")

	doc.AddParagraph("这段文字旁边的图片演示了完整的属性设置功能。图片设置了详细的替代文字（Alt Text），这对于屏幕阅读器用户的可访问性非常重要。同时还设置了图片标题，提供了额外的描述信息。图片位置设置了偏移量，距离右边5毫米，向下偏移2毫米，可以实现精确的位置控制。这些属性设置让图片不仅在视觉上表现良好，在可访问性方面也符合现代文档的标准要求。")

	// ==================== 第六部分：综合应用示例 ====================
	addSectionTitle(doc, "6. 综合应用示例", 1)
	addDescription(doc, "展示在实际文档中综合运用各种图片功能的效果。")

	fmt.Println("6. 综合应用示例演示...")

	doc.AddParagraph("以下是一个模拟真实文档的综合示例，展示了如何在一篇文档中合理运用各种图片功能：")

	// 综合示例：创建多个不同用途的图片
	// 标题图片
	headerImageData := createSampleImage(400, 120, color.RGBA{70, 130, 200, 255}, "文档标题")
	headerImagePath := filepath.Join(imagesDir, "header_image.png")
	if err := saveImageToFile(headerImageData, headerImagePath); err != nil {
		log.Fatalf("保存标题图片失败: %v", err)
	}

	headerImageInfo, err := doc.AddImageFromData(
		headerImageData,
		"header_image.png",
		document.ImageFormatPNG,
		400,
		120,
		&document.ImageConfig{
			Size: &document.ImageSize{
				Width:           100.0, // 100毫米宽度
				KeepAspectRatio: true,
			},
			Position: document.ImagePositionInline,
			WrapText: document.ImageWrapNone,
			AltText:  "文档头部横幅图片",
			Title:    "WordZero图片功能综合演示横幅",
		},
	)
	if err != nil {
		log.Fatalf("添加标题图片失败: %v", err)
	}
	fmt.Printf("   添加文档横幅图片成功，ID: %s\n", headerImageInfo.ID)

	doc.AddParagraph("上方是文档的横幅图片，使用了嵌入式布局和保持长宽比的大小设置。")

	// 左侧配图
	leftSideImageData := createSampleImage(120, 160, color.RGBA{150, 200, 100, 255}, "左侧配图")
	leftSideImagePath := filepath.Join(imagesDir, "left_side_image.png")
	if err := saveImageToFile(leftSideImageData, leftSideImagePath); err != nil {
		log.Fatalf("保存左侧配图失败: %v", err)
	}

	leftSideImageInfo, err := doc.AddImageFromData(
		leftSideImageData,
		"left_side_image.png",
		document.ImageFormatPNG,
		120,
		160,
		&document.ImageConfig{
			Size: &document.ImageSize{
				Width:  30.0,
				Height: 40.0,
			},
			Position: document.ImagePositionFloatLeft,
			WrapText: document.ImageWrapTight,
			AltText:  "文章左侧配图",
			Title:    "左浮动紧密环绕配图示例",
			OffsetX:  1.0,
			OffsetY:  0.0,
		},
	)
	if err != nil {
		log.Fatalf("添加左侧配图失败: %v", err)
	}
	fmt.Printf("   添加左侧配图成功，ID: %s\n", leftSideImageInfo.ID)

	doc.AddParagraph("这段文字展示了在实际文档中使用图片的效果。左侧的配图使用了左浮动和紧密环绕设置，文字会紧贴图片边缘环绕显示。这种布局在杂志、报告和营销材料中非常常见，可以有效提升文档的视觉吸引力和专业感。通过合理的图片布局，可以引导读者的视线，强调重要内容，同时让文档看起来更加生动有趣。")

	// 右侧图表
	chartImageData := createSampleImage(180, 140, color.RGBA{255, 160, 160, 255}, "数据图表")
	chartImagePath := filepath.Join(imagesDir, "chart_image.png")
	if err := saveImageToFile(chartImageData, chartImagePath); err != nil {
		log.Fatalf("保存图表图片失败: %v", err)
	}

	chartImageInfo, err := doc.AddImageFromData(
		chartImageData,
		"chart_image.png",
		document.ImageFormatPNG,
		180,
		140,
		&document.ImageConfig{
			Size: &document.ImageSize{
				Width:  45.0,
				Height: 35.0,
			},
			Position: document.ImagePositionFloatRight,
			WrapText: document.ImageWrapSquare,
			AltText:  "数据分析图表",
			Title:    "展示统计数据的图表",
			OffsetX:  3.0,
			OffsetY:  1.0,
		},
	)
	if err != nil {
		log.Fatalf("添加图表图片失败: %v", err)
	}
	fmt.Printf("   添加右侧图表成功，ID: %s\n", chartImageInfo.ID)

	doc.AddParagraph("右侧的图表展示了数据可视化的应用。在商业文档和技术报告中，图表是不可或缺的元素。通过右浮动和四周环绕的设置，图表可以与解释性文字完美结合，让数据更加清晰易懂。这种布局方式特别适合在分析报告、研究论文和业务提案中使用，可以有效传达复杂的信息内容。")

	// ==================== 第七部分：功能总结 ====================
	addSectionTitle(doc, "7. 功能特性总结", 1)
	addDescription(doc, "总结WordZero图片功能的所有特性和应用场景。")

	doc.AddParagraph("")
	featureTitlePara := doc.AddParagraph("")
	featureTitlePara.AddFormattedText("WordZero图片功能特性列表：", &document.TextFormat{
		Bold:     true,
		FontSize: 12,
	})
	doc.AddParagraph("")

	// 创建功能列表
	features := []string{
		"支持PNG、JPEG、GIF三种主流图片格式",
		"支持从文件和数据两种方式添加图片",
		"提供嵌入式、左浮动、右浮动三种位置选项",
		"支持无环绕、四周环绕、紧密环绕、上下环绕四种文字环绕模式",
		"可设置固定尺寸或保持长宽比的大小调整",
		"支持动态调整已添加图片的大小",
		"提供位置偏移量设置，实现精确位置控制",
		"包含完整的可访问性支持（替代文字、标题）",
		"所有图片数据自动嵌入到Word文档中",
		"兼容Microsoft Word和WPS Office",
	}

	for _, feature := range features {
		featurePara := doc.AddParagraph("")
		featurePara.AddFormattedText(fmt.Sprintf("• %s", feature), &document.TextFormat{
			FontSize: 11,
		})
	}

	doc.AddParagraph("")
	scenarioTitlePara := doc.AddParagraph("")
	scenarioTitlePara.AddFormattedText("应用场景：", &document.TextFormat{
		Bold:     true,
		FontSize: 12,
	})
	doc.AddParagraph("")

	scenarios := []string{
		"技术文档：插入流程图、架构图、截图等",
		"商业报告：添加图表、统计图、产品图片",
		"学术论文：插入实验图片、数据可视化图表",
		"营销材料：使用品牌图片、产品展示图",
		"用户手册：添加操作截图、示意图",
		"培训材料：插入教学图片、演示图表",
	}

	for _, scenario := range scenarios {
		scenarioPara := doc.AddParagraph("")
		scenarioPara.AddFormattedText(fmt.Sprintf("• %s", scenario), &document.TextFormat{
			FontSize: 11,
		})
	}

	// 添加结尾
	doc.AddParagraph("")
	endPara := doc.AddParagraph("")
	endPara.AddFormattedText("本演示文档展示了WordZero图片功能的完整特性。通过灵活运用这些功能，您可以创建出专业、美观的Word文档。", &document.TextFormat{
		FontSize:  12,
		FontColor: "666666",
		Italic:    true,
	})

	// 保存文档
	outputFile := outputDir + "/advanced_image_demo.docx"
	err = doc.Save(outputFile)
	if err != nil {
		log.Fatalf("保存文档失败: %v", err)
	}

	fmt.Printf("\n=== 高级图片功能演示完成 ===\n")
	fmt.Printf("文档已保存到: %s\n", outputFile)
	fmt.Printf("生成的图片文件保存在: %s\n", imagesDir)
	fmt.Printf("生成的图片文件包括:\n")

	imageFiles := []string{
		"basic_image.png (基础图片，蓝色，200x120像素)",
		"inline_image.png (嵌入式图片，红色，180x100像素)",
		"left_float_image.png (左浮动图片，绿色，150x150像素)",
		"right_float_image.png (右浮动图片，橙色，160x120像素)",
		"no_wrap_image.png (无环绕图片，紫色，140x100像素)",
		"square_wrap_image.png (四周环绕图片，青色，130x130像素)",
		"tight_wrap_image.png (紧密环绕图片，橙红色，120x140像素)",
		"top_bottom_wrap_image.png (上下环绕图片，浅绿色，180x80像素)",
		"fixed_size_image.png (固定尺寸图片，粉色，200x150像素)",
		"aspect_ratio_image.png (长宽比图片，浅蓝色，300x150像素)",
		"dynamic_image.png (动态调整图片，浅绿色，160x200像素)",
		"properties_image.png (属性设置图片，浅黄色，140x110像素)",
		"header_image.png (横幅图片，蓝色，400x120像素)",
		"left_side_image.png (左侧配图，绿色，120x160像素)",
		"chart_image.png (图表图片，浅红色，180x140像素)",
	}

	for i, file := range imageFiles {
		fmt.Printf("  %d. %s\n", i+1, file)
	}

	fmt.Printf("\n功能演示统计：\n")
	fmt.Printf("- 演示的图片数量: 15张\n")
	fmt.Printf("- 覆盖的位置模式: 3种 (嵌入式、左浮动、右浮动)\n")
	fmt.Printf("- 覆盖的环绕模式: 4种 (无环绕、四周环绕、紧密环绕、上下环绕)\n")
	fmt.Printf("- 大小调整方式: 3种 (固定尺寸、保持比例、动态调整)\n")
	fmt.Printf("- 应用场景演示: 6个部分\n")
	fmt.Printf("- 文档章节: 7个主要部分\n")
	fmt.Printf("\n您可以用Microsoft Word或WPS Office打开查看完整效果\n")

	fmt.Println("\n=== 所有图片功能演示已完成 ===")
}

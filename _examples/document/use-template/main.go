// Copyright 2017 Baliance. All rights reserved.

package main

import (
	"fmt"
	"github.com/carmel/gooxml/color"
	"github.com/carmel/gooxml/measurement"
	"github.com/carmel/gooxml/schema/soo/wml"
	"log"

	"github.com/qifengzhang007/gooxml/document"
)

var lorem = `
如今，随着公司等企业的不断发展，对于文档处理的需求也愈发增加。而 Word 就是企业中最常见的文档格式之一，更是我们日常工作、学习中必不可少的一部分。
而 Golang 作为逐渐流行的编程语言之一，也具有与 Word 文件交互的能力。下面就让我们一起来看看 Golang 是如何处理 Word 文件操作的。
`

func main() {
	FormatWordBaseTemplate()

	return
	var pathPre = "F:/Project/2023/word-format/使用的库代码参考/gooxml/_examples/document/use-template/"

	// When Word saves a document, it removes all unused styles.  This means to
	// copy the styles from an existing document, you must first create a
	// document that contains text in each style of interest.  As an example,
	// see the template.docx in this directory.  It contains a paragraph set in
	// each style that Word supports by default.
	doc, err := document.OpenTemplate(pathPre + "template.docx")
	if err != nil {
		log.Fatalf("error opening Windows Word 2016 document: %s", err)
	}

	// We can now print out all styles in the document, verifying that they
	// exist.
	//for _, s := range doc.Styles.Styles() {
	//	fmt.Println("style", s.Name(), "has ID of", s.StyleID(), "type is", s.Type())
	//}

	// 遍历所有的段落
	//for index, graph := range doc.Paragraphs() {
	//	fmt.Printf("%d - 段落样式：%s\n", index, graph.Style())
	//	graph.AddRun().Properties().SetColor(color.RGB(220, 12, 12))
	//	graph.AddRun()
	//	//prop.SetFontFamily("微软雅黑") // 设置字体
	//	//prop.SetBold(true)
	//	//prop.SetColor(color.RGB(220, 12, 12))
	//}

	// And create documents setting their style to the style ID (not style name).
	para := doc.AddParagraph()
	para.SetStyle("Title")
	para.AddRun().AddText("My Document Title")

	para = doc.AddParagraph()
	para.SetStyle("Subtitle")
	para.AddRun().AddText("Document Subtitle")

	para = doc.AddParagraph()
	para.SetStyle("Heading1")
	para.AddRun().AddText("Major Section")
	para = doc.AddParagraph()
	para = doc.AddParagraph()
	for i := 0; i < 4; i++ {
		para.AddRun().AddText(lorem)
	}

	para = doc.AddParagraph()
	para.SetStyle("Heading2")
	para.AddRun().AddText("Minor Section")
	para = doc.AddParagraph()
	for i := 0; i < 4; i++ {
		para.AddRun().AddText(lorem)
	}

	// using a pre-defined table style
	//table := doc.AddTable()
	//table.Properties().SetWidthPercent(90)
	//table.Properties().SetStyle("GridTable4-Accent1")
	//look := table.Properties().TableLook()
	//// these have default values in the style, so we manually turn some of them off
	//look.SetFirstColumn(false)
	//look.SetFirstRow(true)
	//look.SetLastColumn(false)
	//look.SetLastRow(true)
	//look.SetHorizontalBanding(true)
	//
	//for r := 0; r < 5; r++ {
	//	row := table.AddRow()
	//	for c := 0; c < 5; c++ {
	//		cell := row.AddCell()
	//		cell.AddParagraph().AddRun().AddText(fmt.Sprintf("row %d col %d", r+1, c+1))
	//	}
	//}
	doc.SaveToFile(pathPre + "use-template-2.docx")
}

// 测试基于已有模板word，进行全量复制、格式更改
func FormatWordBaseTemplate() {
	var pathPre = "F:/Project/2023/word-format/使用的库代码参考/gooxml/_examples/document/use-template/"
	doc, err := document.Open(pathPre + "template-001.docx")
	//doc, err := document.OpenTemplate(pathPre + "template-001.docx")
	if err != nil {
		log.Fatalf("打开模板文件出错 %s", err)
	}

	// 获取全部的段落信息
	paragraphs := []document.Paragraph{}
	for _, p := range doc.Paragraphs() {
		paragraphs = append(paragraphs, p)
	}
	// 获取全部的结构化文档标签
	for _, sdt := range doc.StructuredDocumentTags() {
		for _, p := range sdt.Paragraphs() {
			paragraphs = append(paragraphs, p)
		}
	}

	// 循环段落并进行文字替换
	for _, p := range paragraphs {
		for _, r := range p.Runs() {
			fmt.Printf("输出所有的子段落-run：%+v - 对应的段落样式：%+v\n", r.Text(), p.Style())
			// 所有段落的行间距设置为18磅
			p.Properties().Spacing().SetLineSpacing(28*measurement.Point, wml.ST_LineSpacingRuleExact)
			r.Properties()
			switch r.Text() {
			case "测试标题": //	fmt.Printf("测试标题的样式：%#+v\n", p.Style())
				r.ClearContent() // 清除原有的文字信息和换行符
				r.AddText("替换[测试标题]替换")
				r.AddBreak()
				// 在该段落前添加新段落 p 表示匹配到的段落
				para := doc.InsertParagraphBefore(p)
				para.AddRun().AddText("整个文档Pre.")
				// 在该段落后添加新段落
				para = doc.InsertParagraphAfter(p)
				para.AddRun().AddText("整个文档after")

			case "正文内容", "测试文档 一级标题":
				fmt.Printf("正文内容 的样式：%#+v\n", p.Style())
				r.Properties().SetFontFamily("微软雅黑")
				r.Properties().SetColor(color.RGB(220, 10, 10)) // 大概是红色
				r.Properties().SetSize(16)
				p.Properties().SetAlignment(wml.ST_JcCenter) // 正文内容居中

				//p.AddRun()
			}
		}
	}

	// 这里开始处理表格，
	tbs := doc.Tables()
	for tbIndex, tb := range tbs {
		fmt.Println("表格序号：", tbIndex)
		for rowIndex, row := range tb.Rows() {
			for cellIndex, cel := range row.Cells() {
				// 解析段落
				for _, pg := range cel.Paragraphs() {
					for _, r := range pg.Runs() {
						fmt.Printf("%d 行 - %d 列的数据：%+v\n", rowIndex, cellIndex, r.Text())
						r.Properties().SetColor(color.RGB(220, 10, 10))
					}
				}
			}
		}
	}

	doc.SaveToFile(pathPre + "template-001-save.docx") // 生成替换后的新文档

}

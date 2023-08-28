// Copyright 2017 Baliance. All rights reserved.

package main

import (
	"fmt"
	"github.com/carmel/gooxml/color"
	"log"

	"github.com/carmel/gooxml/document"
)

var lorem = `
如今，随着公司等企业的不断发展，对于文档处理的需求也愈发增加。而 Word 就是企业中最常见的文档格式之一，更是我们日常工作、学习中必不可少的一部分。而 Golang 作为逐渐流行的编程语言之一，也具有与 Word 文件交互的能力。下面就让我们一起来看看 Golang 是如何处理 Word 文件操作的。
`

func main() {

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
	for index, graph := range doc.Paragraphs() {
		fmt.Printf("%d - 段落样式：%s\n", index, graph.Style())
		graph.AddRun().Properties().SetColor(color.RGB(220, 12, 12))
		graph.AddRun()
		//prop.SetFontFamily("微软雅黑") // 设置字体
		//prop.SetBold(true)
		//prop.SetColor(color.RGB(220, 12, 12))
	}

	// And create documents setting their style to the style ID (not style name).
	//para := doc.AddParagraph()
	//para.SetStyle("Title")
	//para.AddRun().AddText("My Document Title")
	//
	//para = doc.AddParagraph()
	//para.SetStyle("Subtitle")
	//para.AddRun().AddText("Document Subtitle")
	//
	//para = doc.AddParagraph()
	//para.SetStyle("Heading1")
	//para.AddRun().AddText("Major Section")
	//para = doc.AddParagraph()
	//para = doc.AddParagraph()
	//for i := 0; i < 4; i++ {
	//	para.AddRun().AddText(lorem)
	//}
	//
	//para = doc.AddParagraph()
	//para.SetStyle("Heading2")
	//para.AddRun().AddText("Minor Section")
	//para = doc.AddParagraph()
	//for i := 0; i < 4; i++ {
	//	para.AddRun().AddText(lorem)
	//}

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

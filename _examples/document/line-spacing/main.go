// Copyright 2017 Baliance. All rights reserved.
package main

import (
	"log"

	"github.com/carmel/gooxml/document"
	"github.com/carmel/gooxml/measurement"
	"github.com/carmel/gooxml/schema/soo/wml"
)

func main() {
	var pathPre = "F:/Project/2023/word-format/使用的库代码参考/gooxml/_examples/document/line-spacing/"

	doc := document.New()
	lorem := `
如今，随着公司等企业的不断发展，对于文档处理的需求也愈发增加。而 Word 就是企业中最常见的文档格式之一，更是我们日常工作、学习中必不可少的一部分。而 Golang 作为逐渐流行的编程语言之一，也具有与 Word 文件交互的能力。下面就让我们一起来看看 Golang 是如何处理 Word 文件操作的。
`
	// single spaced
	para := doc.AddParagraph()
	run := para.AddRun()
	run.AddText(lorem)
	run.AddText(lorem)
	run.AddBreak()

	// double spaced is twice the text height (24 points in this case as the text height is 12 points)
	para = doc.AddParagraph()
	// 正文段落行间距18磅
	para.Properties().Spacing().SetLineSpacing(18*measurement.Point, wml.ST_LineSpacingRuleExact)
	run = para.AddRun()
	run.AddText(lorem)
	run.AddText(lorem)
	run.AddBreak()

	if err := doc.Validate(); err != nil {
		log.Fatalf("error during validation: %s", err)
	}
	doc.SaveToFile(pathPre + "line-spacing-2.docx")
}

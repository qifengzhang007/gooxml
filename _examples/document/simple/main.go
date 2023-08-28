// Copyright 2017 Baliance. All rights reserved.
package main

import (
	"fmt"

	"github.com/carmel/gooxml/color"
	"github.com/carmel/gooxml/document"
	"github.com/carmel/gooxml/measurement"
	"github.com/carmel/gooxml/schema/soo/wml"
)

func main() {
	var pathPre = "F:/Project/2023/word-format/使用的库代码参考/gooxml/_examples/document/simple/"
	doc := document.New()

	para := doc.AddParagraph()
	para.SetStyle("Title") // 1.这里添加的是样式 ，需要先让段落生效 -> 然后再设置样式
	run := para.AddRun()
	run.AddText("标题")

	para = doc.AddParagraph()
	para.SetStyle("Heading1")
	run = para.AddRun()
	run.AddText("Heading1 一级标题")

	para = doc.AddParagraph()
	para.SetStyle("Heading2")
	run = para.AddRun()
	run.AddText("Heading2 - 二级标题")

	para = doc.AddParagraph()
	para.SetStyle("Heading3")
	run = para.AddRun()
	run.AddText("Heading3 - 三级标题")

	para = doc.AddParagraph()
	para.Properties().SetFirstLineIndent(0.5 * measurement.Inch)
	run = para.AddRun()
	run.AddText("缩进文本 A run is a string of characters with the same formatting. ")

	para = doc.AddParagraph()
	run = para.AddRun() // 2.下一步需要设置属性，这里先让段落运行生效，才能设置样式
	run.Properties().SetBold(true)
	run.Properties().SetFontFamily("微软雅黑") // Courier
	run.Properties().SetSize(12)           // 注意字体大小，必须是 偶数
	run.Properties().SetColor(color.Red)   // 可以自定义设置 color.Color{}
	run.AddText("微软雅黑 同一个段落中可以存在多个不同格式的运行run. 微软雅黑 ")

	//run = para.AddRun()
	run.AddText("下面插入换行符号. ")
	run.AddBreak()
	run.AddBreak()

	run = createParaRun(doc, "Runs support styling options:")

	run = createParaRun(doc, "small caps")
	run.Properties().SetSmallCaps(true)

	run = createParaRun(doc, "strike")
	run.Properties().SetStrikeThrough(true)

	run = createParaRun(doc, "double strike")
	run.Properties().SetDoubleStrikeThrough(true)

	run = createParaRun(doc, "outline")
	run.Properties().SetOutline(true)

	run = createParaRun(doc, "emboss")
	run.Properties().SetEmboss(true)

	run = createParaRun(doc, "shadow")
	run.Properties().SetShadow(true)

	run = createParaRun(doc, "imprint")
	run.Properties().SetImprint(true)

	run = createParaRun(doc, "高亮highlighting高亮")
	run.Properties().SetHighlight(wml.ST_HighlightColorYellow)

	run = createParaRun(doc, "下划线underline下划线")
	run.Properties().SetUnderline(wml.ST_UnderlineWavyDouble, color.Red)

	run = createParaRun(doc, "text effects")
	run.Properties().SetEffect(wml.ST_TextEffectAntsRed)

	nd := doc.Numbering.Definitions()[0]

	for i := 1; i < 5; i++ {
		p := doc.AddParagraph()
		p.SetNumberingLevel(i - 1)
		p.SetNumberingDefinition(nd)
		run := p.AddRun()
		run.AddText(fmt.Sprintf("Level %d", i))
	}
	doc.SaveToFile(pathPre + "simple-2.docx")
}

// 创建一个段落并立即生效
func createParaRun(doc *document.Document, s string) document.Run {
	para := doc.AddParagraph()
	run := para.AddRun()
	run.AddText(s)
	return run
}

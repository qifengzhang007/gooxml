package main

import (
	"fmt"

	"github.com/qifengzhang007/gooxml/common"
	"github.com/qifengzhang007/gooxml/document"
)

func main() {
	// 打开 Word 文档
	pathPre := "F:/OpenSource/gooxml/_examples/document/replace-text-or-pic/"
	doc, err := document.Open(pathPre + "input.docx")

	if err != nil {
		fmt.Printf("打开文档失败: %v\n", err)
		return
	}

	Paragraphs := doc.Paragraphs()
	for index1, item := range Paragraphs {
		for _, item2 := range item.Runs() {
			fmt.Println("run:", item2.Text())
			if item2.Text() == "【点击签名】" {
				item2.ClearContent() // 原始的签名标记内容清空

				// 插入签名图片  - 错误判断暂时忽略,正规业务自己加判断
				img1, _ := common.ImageFromFile(pathPre + "签名图片.png")
				img1ref, _ := doc.AddImage(img1)
				tmpRun := Paragraphs[index1].AddRun()

				_, _ = tmpRun.AddDrawingInline(img1ref)
			}
		}
	}

	_ = doc.SaveToFile(pathPre + "input-2.docx")
}

// 替换文本
//func replaceText() {
//	pathPre := "F:/OpenSource/gooxml/_examples/document/replace-text-or-pic/"
//	doc, err := document.Open(pathPre + "input.docx")
//	if err != nil {
//		fmt.Printf("打开文档失败: %v\n", err)
//		return
//	}
//	Paragraphs := doc.Paragraphs()
//	for _, item := range Paragraphs {
//		for _, item2 := range item.Runs() {
//			if item2.Text() == "{{name}}" {
//				item2.ClearContent()
//				item2.AddText("替换后的左上角标题")
//				item2.AddBreak()
//			}
//		}
//	}
//
//	_ = doc.SaveToFile(pathPre + "替换后的文档0826.docx")
//}

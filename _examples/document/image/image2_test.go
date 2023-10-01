// Copyright 2017 Baliance. All rights reserved.
package main

import (
	"fmt"
	"github.com/qifengzhang007/gooxml/document"
	"github.com/qifengzhang007/gooxml/measurement"
	"testing"
)

func TestImages2(t *testing.T) {
	//
	//var text = `开源项目 > WEB应用开发 > Web开发框架`
	var rootPath = "E:/Project/2023/gooxml/_examples/document/image/"
	doc, err := document.Open(rootPath + "不缩放.docx")
	if err != nil {
		fmt.Printf("打开文件出错：" + err.Error())
		return
	}

	//  doc.GetImageByRelID()

	for pIndex, p := range doc.Paragraphs() {
		fmt.Printf("pIndex:%d -- P段落属性：%+v\n", pIndex, p)

		for rIndex, r := range p.Runs() {
			if r.CurRunIsContainerPic() {
				fmt.Printf("图像节点：%d\n", pIndex)
				for _, pic := range r.GetPicInfo() {
					for _, tmp := range doc.GetWordRelationships() {
						if tmp.IdAttr == *pic.BlipFill.Blip.EmbedAttr {
							fmt.Printf("图像 - run - rIndex:%d  -----   pic信息：%+v   -- %+v \n", rIndex, pic, tmp)
						}
					}
				}
			} else {
				fmt.Printf("文本- run - rIndex:%d  -----   文本信息：%+v  \n", rIndex, r.Text())
			}
		}
	}

	// 新增段落， 插入图片
	tmpP := doc.AddParagraph()
	tmpP.AddRun().AddText("rId4 对应黑色代码图片")
	tmpP.AddRun().AddBreak()
	imgRef, _ := doc.GetImageByRelID("rId4")
	inDraw, _ := tmpP.AddRun().AddDrawingInline(imgRef)
	inDraw.SetSize(measurement.Distance(imgRef.Size().X), measurement.Distance(imgRef.Size().Y)*measurement.Point, 0.5)
	tmpP.AddRun().AddBreak()

	doc.SaveToFile("image2-test.docx")
}

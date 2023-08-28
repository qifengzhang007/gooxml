// Copyright 2017 Baliance. All rights reserved.

package main

import (
	"fmt"
	"log"
	"time"

	"github.com/carmel/gooxml/document"
)

func main() {
	var pathPre = "F:/Project/2023/word-format/使用的库代码参考/gooxml/_examples/document/edit-document/"
	doc, err := document.Open(pathPre + "document.docx")
	if err != nil {
		log.Fatalf("error opening document: %s", err)
	}

	paragraphs := []document.Paragraph{}
	for _, p := range doc.Paragraphs() {
		paragraphs = append(paragraphs, p)
	}

	// This sample document uses structured document tags, which are not common
	// except for in document templates.  Normally you can just iterate over the
	// document's paragraphs.
	for _, sdt := range doc.StructuredDocumentTags() {
		for _, p := range sdt.Paragraphs() {
			paragraphs = append(paragraphs, p)
		}
	}

	for _, p := range paragraphs {
		for _, r := range p.Runs() {
			switch r.Text() {
			case "FIRST NAME":
				// ClearContent clears both text and line breaks within a run,
				// so we need to add the line break back
				r.ClearContent()
				r.AddText("张 ")
				r.AddBreak() // 换行

				para := doc.InsertParagraphBefore(p)
				para.AddRun().AddText("1-前面插入一段.")
				para.SetStyle("Name") // Name is a default style in this template file

				para = doc.InsertParagraphAfter(p)
				para.AddRun().AddText("2-前面插入一段")
				para.SetStyle("Name")

			case "LAST NAME":
				r.ClearContent()
				r.AddText("三丰")
			case "Address | Phone | Email":
				r.ClearContent()
				r.AddText("上海 | 松江| 1990850157qq.com | 166")
			case "Date":
				r.ClearContent()
				r.AddText(time.Now().Format("2006-01-02 15:04:05"))
			case "Recipient Name":
				r.ClearContent()
				r.AddText("Mrs. Smith")
				r.AddBreak()
			case "Title":
				// we remove the title content entirely
				p.RemoveRun(r)
			case "Company":
				r.ClearContent()
				r.AddText("Smith Enterprises")
				r.AddBreak()
			case "Address":
				r.ClearContent()
				r.AddText("112 Rustic Rd")
				r.AddBreak()
			case "City, ST ZIP Code":
				r.ClearContent()
				r.AddText("San Francisco, CA 94016")
				r.AddBreak()
			case "Dear Recipient:":
				r.ClearContent()
				r.AddText("Dear Mrs. Smith:")
				r.AddBreak()
			case "Your Name":
				r.ClearContent()
				r.AddText("John Smith")
				r.AddBreak()

				run := p.InsertRunBefore(r)
				run.AddText("---Before----")
				run.AddBreak()
				run = p.InsertRunAfter(r)
				run.AddText("---After----")

			default:
				fmt.Println("not modifying", r.Text())
			}
		}
	}

	doc.SaveToFile(pathPre + "edit-document2.docx")
}

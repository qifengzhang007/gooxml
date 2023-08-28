// Copyright 2017 Baliance. All rights reserved.

package main

import (
	"fmt"
	"log"
	"time"

	"github.com/carmel/gooxml/document"
)

func main() {
	doc, err := document.Open("F:/Project/2023/word-format/使用的库代码参考/gooxml/_examples/document/doc-properties/document.docx")
	if err != nil {
		log.Fatalf("error opening document: %s", err)
	}

	cp := doc.CoreProperties
	// You can read properties from the document
	fmt.Println("Title:", cp.Title())
	fmt.Println("Author:", cp.Author())
	fmt.Println("Description:", cp.Description())
	fmt.Println("Last Modified By:", cp.LastModifiedBy())
	fmt.Println("Category:", cp.Category())
	fmt.Println("Content Status:", cp.ContentStatus())
	fmt.Println("Created:", cp.Created())
	fmt.Println("Modified:", cp.Modified())

	// And change them as well
	cp.SetTitle("CP Invoices")
	cp.SetAuthor("John Doe")
	cp.SetCategory("Invoices")
	cp.SetContentStatus("Draft")
	cp.SetLastModifiedBy("Jane Smith")
	cp.SetCreated(time.Now())
	cp.SetModified(time.Now())
	doc.SaveToFile("document.docx")
}

// Copyright 2017 Baliance. All rights reserved.
package main

import (
	"io/ioutil"
	"log"
	"testing"

	"github.com/qifengzhang007/gooxml/common"
	"github.com/qifengzhang007/gooxml/document"
	"github.com/qifengzhang007/gooxml/measurement"

	"github.com/qifengzhang007/gooxml/schema/soo/wml"
)

var lorem = `Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin lobortis, lectus dictum feugiat tempus, sem neque finibus enim, sed eleifend sem nunc ac diam. Vestibulum tempus sagittis elementum`

func TestImages1(t *testing.T) {
	doc := document.New()

	img1, err := common.ImageFromFile("gophercolor.png")
	if err != nil {
		log.Fatalf("unable to create image: %s", err)
	}
	img2data, err := ioutil.ReadFile("gophercolor.png")
	if err != nil {
		log.Fatalf("unable to read file: %s", err)
	}
	img2, err := common.ImageFromBytes(img2data)
	if err != nil {
		log.Fatalf("unable to create image: %s", err)
	}

	img1ref, err := doc.AddImage(img1)
	if err != nil {
		log.Fatalf("unable to add image to document: %s", err)
	}
	img2ref, err := doc.AddImage(img2)
	if err != nil {
		log.Fatalf("unable to add image to document: %s", err)
	}

	para := doc.AddParagraph()
	anchored, err := para.AddRun().AddDrawingAnchored(img1ref)
	if err != nil {
		log.Fatalf("unable to add anchored image: %s", err)
	}
	anchored.SetName("Gopher")
	anchored.SetSize(2*measurement.Inch, 2*measurement.Inch)
	anchored.SetOrigin(wml.WdST_RelFromHPage, wml.WdST_RelFromVTopMargin)
	anchored.SetHAlignment(wml.WdST_AlignHCenter)
	anchored.SetYOffset(3 * measurement.Inch)
	anchored.SetTextWrapSquare(wml.WdST_WrapTextBothSides)

	run := para.AddRun()
	for i := 0; i < 16; i++ {
		run.AddText(lorem)

		// drop an inline image in
		if i == 13 {
			inl, err := run.AddDrawingInline(img1ref)
			if err != nil {
				log.Fatalf("unable to add inline image: %s", err)
			}
			inl.SetSize(1*measurement.Inch, 1*measurement.Inch, 1)
		}
		if i == 15 {
			inl, err := run.AddDrawingInline(img2ref)
			if err != nil {
				log.Fatalf("unable to add inline image: %s", err)
			}
			inl.SetSize(1*measurement.Inch, 1*measurement.Inch, 1)
		}
	}
	doc.SaveToFile("image.docx")
}

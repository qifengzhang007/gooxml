// Copyright 2017 Baliance. All rights reserved.
//
// Use of this source code is governed by the terms of the Affero GNU General
// Public License version 3.0 as published by the Free Software Foundation and
// appearing in the file LICENSE included in the packaging of this file. A
// commercial license can be purchased by contacting sales@baliance.com.

package document

import (
	"github.com/qifengzhang007/gooxml/common"
	"github.com/qifengzhang007/gooxml/measurement"
	pic "github.com/qifengzhang007/gooxml/schema/soo/dml/picture"
	"github.com/qifengzhang007/gooxml/schema/soo/wml"
)

// InlineDrawing is an inlined image within a run.
type InlineDrawing struct {
	d *Document
	x *wml.WdInline
}

// X returns the inner wrapped XML type.
func (i InlineDrawing) X() *wml.WdInline {
	return i.x
}

// GetImage returns the ImageRef associated with an InlineDrawing.
func (i InlineDrawing) GetImage() (common.ImageRef, bool) {
	any := i.x.Graphic.GraphicData.Any
	if len(any) > 0 {
		p, ok := any[0].(*pic.Pic)
		if ok {
			if p.BlipFill != nil && p.BlipFill.Blip != nil && p.BlipFill.Blip.EmbedAttr != nil {
				return i.d.GetImageByRelID(*p.BlipFill.Blip.EmbedAttr)
			}
		}
	}
	return common.ImageRef{}, false
}

// SetSize sets the size of the displayed image on the page.
// When setting the image size, please set the size in pixels,
// the program will automatically calculate the unit to meet the ooxml regulations.
// width  = image width
// height   = image height
// isZoom  when isZoom=true, we auto  calc image  width
func (i InlineDrawing) SetSize(width, height measurement.Distance, isZoom bool) {
	var zoomScale float64 = 1
	if isZoom && width > 0 && height > 0 {
		zoomScale = measurement.DocDefaultWidth / float64(width)
		if zoomScale > 1 {
			zoomScale = 1
		}
	}
	i.x.Extent.CxAttr = int64(float64(measurement.Distance(zoomScale)*width*measurement.NOFFZZ) / measurement.Ppi)
	i.x.Extent.CyAttr = int64(float64(measurement.Distance(zoomScale)*height*measurement.NOFFZZ) / measurement.Ppi)

}

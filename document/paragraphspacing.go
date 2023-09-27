// Copyright 2018 Baliance. All rights reserved.
//
// Use of this source code is governed by the terms of the Affero GNU General
// Public License version 3.0 as published by the Free Software Foundation and
// appearing in the file LICENSE included in the packaging of this file. A
// commercial license can be purchased by contacting sales@baliance.com.

package document

import (
	"github.com/qifengzhang007/gooxml"
	"github.com/qifengzhang007/gooxml/measurement"
	"github.com/qifengzhang007/gooxml/schema/soo/ofc/sharedTypes"
	"github.com/qifengzhang007/gooxml/schema/soo/wml"
)

// ParagraphSpacing controls the spacing for a paragraph and its lines.
type ParagraphSpacing struct {
	x *wml.CT_Spacing
}

// SetBefore sets the spacing that comes before the paragraph.
// @before When setting paragraph spacing, the unit must be set in points.
func (p ParagraphSpacing) SetBefore(before measurement.Distance) {
	p.x.BeforeAttr = &sharedTypes.ST_TwipsMeasure{}
	p.x.BeforeAttr.ST_UnsignedDecimalNumber = gooxml.Uint64(uint64(before / measurement.Twips))
	p.setBeforeLine(before)
}

// setBeforeLine
// Used with the previous function, if the original word document paragraph spacing is set by behavioral units. Not call setBeforeLine is not achieve the effect.
// @point unit  is  point
func (p ParagraphSpacing) setBeforeLine(point measurement.Distance) {
	p.x.BeforeLinesAttr = gooxml.Int64(int64(point * 100))
}

// SetAfter sets the spacing that comes after the paragraph.
func (p ParagraphSpacing) SetAfter(after measurement.Distance) {
	p.x.AfterAttr = &sharedTypes.ST_TwipsMeasure{}
	p.x.AfterAttr.ST_UnsignedDecimalNumber = gooxml.Uint64(uint64(after / measurement.Twips))
	p.setAfterLine(after)
}

// setAfterLine sets the spacing that comes before the paragraph.
// @point unit  is  point
func (p ParagraphSpacing) setAfterLine(point measurement.Distance) {
	p.x.AfterLinesAttr = gooxml.Int64(int64(point * 100))
}

// SetLineSpacing sets the spacing between lines in a paragraph.
func (p ParagraphSpacing) SetLineSpacing(d measurement.Distance, rule wml.ST_LineSpacingRule) {
	if rule == wml.ST_LineSpacingRuleUnset {
		p.x.LineRuleAttr = wml.ST_LineSpacingRuleUnset
		p.x.LineAttr = nil
	} else {
		p.x.LineRuleAttr = rule
		p.x.LineAttr = &wml.ST_SignedTwipsMeasure{}
		p.x.LineAttr.Int64 = gooxml.Int64(int64(d / measurement.Twips))
	}
}

// SetBeforeAuto controls if spacing before a paragraph is automatically determined.
func (p ParagraphSpacing) SetBeforeAuto(b bool) {
	if b {
		p.x.BeforeAutospacingAttr = &sharedTypes.ST_OnOff{}
		p.x.BeforeAutospacingAttr.Bool = gooxml.Bool(true)
	} else {
		p.x.BeforeAutospacingAttr = nil
	}
}

// SetAfterAuto controls if spacing after a paragraph is automatically determined.
func (p ParagraphSpacing) SetAfterAuto(b bool) {
	if b {
		p.x.AfterAutospacingAttr = &sharedTypes.ST_OnOff{}
		p.x.AfterAutospacingAttr.Bool = gooxml.Bool(true)
	} else {
		p.x.AfterAutospacingAttr = nil
	}
}

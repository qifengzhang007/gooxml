// Copyright 2017 Baliance. All rights reserved.
//
// Use of this source code is governed by the terms of the Affero GNU General
// Public License version 3.0 as published by the Free Software Foundation and
// appearing in the file LICENSE included in the packaging of this file. A
// commercial license can be purchased by contacting sales@baliance.com.

package document

import (
	"fmt"

	"github.com/qifengzhang007/gooxml"
	"github.com/qifengzhang007/gooxml/measurement"
	"github.com/qifengzhang007/gooxml/schema/soo/ofc/sharedTypes"
	"github.com/qifengzhang007/gooxml/schema/soo/wml"
)

// ParagraphProperties are the properties for a paragraph.
type ParagraphProperties struct {
	D     *Document
	CtPPr *wml.CT_PPr
}

// X returns the inner wrapped XML type.
func (p ParagraphProperties) X() *wml.CT_PPr {
	return p.CtPPr
}

// SetSpacing sets the spacing that comes before and after the paragraph.
// Deprecated: See Spacing() instead which allows finer control.
func (p ParagraphProperties) SetSpacing(before, after measurement.Distance) {
	if p.CtPPr.Spacing == nil {
		p.CtPPr.Spacing = wml.NewCT_Spacing()
	}
	p.CtPPr.Spacing.BeforeAttr = &sharedTypes.ST_TwipsMeasure{}
	p.CtPPr.Spacing.BeforeAttr.ST_UnsignedDecimalNumber = gooxml.Uint64(uint64(before / measurement.Twips))
	p.CtPPr.Spacing.AfterAttr = &sharedTypes.ST_TwipsMeasure{}
	p.CtPPr.Spacing.AfterAttr.ST_UnsignedDecimalNumber = gooxml.Uint64(uint64(after / measurement.Twips))
}

// Spacing returns the paragraph spacing settings.
func (p ParagraphProperties) Spacing() ParagraphSpacing {
	if p.CtPPr.Spacing == nil {
		p.CtPPr.Spacing = wml.NewCT_Spacing()
	}
	return ParagraphSpacing{p.CtPPr.Spacing}
}

// SetAlignment controls the paragraph alignment
func (p ParagraphProperties) SetAlignment(align wml.ST_Jc) {
	if align == wml.ST_JcUnset {
		p.CtPPr.Jc = nil
	} else {
		p.CtPPr.Jc = wml.NewCT_Jc()
		p.CtPPr.Jc.ValAttr = align
	}
}

// Style returns the style for a paragraph, or an empty string if it is unset.
func (p ParagraphProperties) Style() string {
	if p.CtPPr.PStyle != nil {
		return p.CtPPr.PStyle.ValAttr
	}
	return ""
}

// SetStyle sets the style of a paragraph.
func (p ParagraphProperties) SetStyle(s string) {
	if s == "" {
		p.CtPPr.PStyle = nil
	} else {
		p.CtPPr.PStyle = wml.NewCT_String()
		p.CtPPr.PStyle.ValAttr = s
	}
}

// AddTabStop adds a tab stop to the paragraph.  It controls the position of text when using Run.AddTab()
func (p ParagraphProperties) AddTabStop(position measurement.Distance, justificaton wml.ST_TabJc, leader wml.ST_TabTlc) {
	if p.CtPPr.Tabs == nil {
		p.CtPPr.Tabs = wml.NewCT_Tabs()
	}
	tab := wml.NewCT_TabStop()
	tab.LeaderAttr = leader
	tab.ValAttr = justificaton
	tab.PosAttr.Int64 = gooxml.Int64(int64(position / measurement.Twips))
	p.CtPPr.Tabs.Tab = append(p.CtPPr.Tabs.Tab, tab)
}

// AddSection adds a new document section with an optional section break.  If t
// is ST_SectionMarkUnset, then no break will be inserted.
func (p ParagraphProperties) AddSection(t wml.ST_SectionMark) Section {
	p.CtPPr.SectPr = wml.NewCT_SectPr()
	if t != wml.ST_SectionMarkUnset {
		p.CtPPr.SectPr.Type = wml.NewCT_SectType()
		p.CtPPr.SectPr.Type.ValAttr = t
	}
	return Section{p.D, p.CtPPr.SectPr}
}

// SetHeadingLevel sets a heading level and style based on the level to a
// paragraph.  The default styles for a new gooxml document support headings
// from level 1 to 8.
func (p ParagraphProperties) SetHeadingLevel(idx int) {
	p.SetStyle(fmt.Sprintf("Heading%d", idx))
	if p.CtPPr.NumPr == nil {
		p.CtPPr.NumPr = wml.NewCT_NumPr()
	}
	p.CtPPr.NumPr.Ilvl = wml.NewCT_DecimalNumber()
	p.CtPPr.NumPr.Ilvl.ValAttr = int64(idx)
}

// SetKeepWithNext controls if this paragraph should be kept with the next.
func (p ParagraphProperties) SetKeepWithNext(b bool) {
	if !b {
		p.CtPPr.KeepNext = nil
	} else {
		p.CtPPr.KeepNext = wml.NewCT_OnOff()
	}
}

// SetKeepOnOnePage controls if all lines in a paragraph are kept on the same
// page.
func (p ParagraphProperties) SetKeepOnOnePage(b bool) {
	if !b {
		p.CtPPr.KeepLines = nil
	} else {
		p.CtPPr.KeepLines = wml.NewCT_OnOff()
	}
}

// SetPageBreakBefore controls if there is a page break before this paragraph.
func (p ParagraphProperties) SetPageBreakBefore(b bool) {
	if !b {
		p.CtPPr.PageBreakBefore = nil
	} else {
		p.CtPPr.PageBreakBefore = wml.NewCT_OnOff()
	}
}

// SetWindowControl controls if the first or last line of the paragraph is
// allowed to dispay on a separate page.
func (p ParagraphProperties) SetWindowControl(b bool) {
	if !b {
		p.CtPPr.WidowControl = nil
	} else {
		p.CtPPr.WidowControl = wml.NewCT_OnOff()
	}
}

// SetFirstLineIndent controls the indentation of the first line in a paragraph.
func (p ParagraphProperties) SetFirstLineIndent(m measurement.Distance) {
	if p.CtPPr.Ind == nil {
		p.CtPPr.Ind = wml.NewCT_Ind()
	}
	if m == measurement.Zero {
		p.CtPPr.Ind.FirstLineAttr = nil
	} else {
		p.CtPPr.Ind.FirstLineAttr = &sharedTypes.ST_TwipsMeasure{}
		p.CtPPr.Ind.FirstLineAttr.ST_UnsignedDecimalNumber = gooxml.Uint64(uint64(m / measurement.Twips))
	}
}

// SetStartIndent controls the start indentation.
func (p ParagraphProperties) SetStartIndent(m measurement.Distance) {
	if p.CtPPr.Ind == nil {
		p.CtPPr.Ind = wml.NewCT_Ind()
	}
	if m == measurement.Zero {
		p.CtPPr.Ind.StartAttr = nil
	} else {
		p.CtPPr.Ind.StartAttr = &wml.ST_SignedTwipsMeasure{}
		p.CtPPr.Ind.StartAttr.Int64 = gooxml.Int64(int64(m / measurement.Twips))
	}
}

// SetEndIndent controls the end indentation.
func (p ParagraphProperties) SetEndIndent(m measurement.Distance) {
	if p.CtPPr.Ind == nil {
		p.CtPPr.Ind = wml.NewCT_Ind()
	}
	if m == measurement.Zero {
		p.CtPPr.Ind.EndAttr = nil
	} else {
		p.CtPPr.Ind.EndAttr = &wml.ST_SignedTwipsMeasure{}
		p.CtPPr.Ind.EndAttr.Int64 = gooxml.Int64(int64(m / measurement.Twips))
	}
}

// SetHangingIndent controls the indentation of the non-first lines in a paragraph.
func (p ParagraphProperties) SetHangingIndent(m measurement.Distance) {
	if p.CtPPr.Ind == nil {
		p.CtPPr.Ind = wml.NewCT_Ind()
	}
	if m == measurement.Zero {
		p.CtPPr.Ind.HangingAttr = nil
	} else {
		p.CtPPr.Ind.HangingAttr = &sharedTypes.ST_TwipsMeasure{}
		p.CtPPr.Ind.HangingAttr.ST_UnsignedDecimalNumber = gooxml.Uint64(uint64(m / measurement.Twips))
	}
}

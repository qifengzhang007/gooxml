// Copyright 2017 Baliance. All rights reserved.
//
// DO NOT EDIT: generated by gooxml ECMA-376 generator
//
// Use of this source code is governed by the terms of the Affero GNU General
// Public License version 3.0 as published by the Free Software Foundation and
// appearing in the file LICENSE included in the packaging of this file. A
// commercial license can be purchased by contacting sales@baliance.com.

package sml

import (
	"encoding/xml"
	"fmt"
	"strconv"

	"github.com/carmel/gooxml"
)

type CT_GroupLevel struct {
	// Unique Name
	UniqueNameAttr string
	// Grouping Level Display Name
	CaptionAttr string
	// User-Defined Group Level
	UserAttr *bool
	// Custom Roll Up
	CustomRollUpAttr *bool
	// OLAP Level Groups
	Groups *CT_Groups
	// Future Feature Data Storage Area
	ExtLst *CT_ExtensionList
}

func NewCT_GroupLevel() *CT_GroupLevel {
	ret := &CT_GroupLevel{}
	return ret
}

func (m *CT_GroupLevel) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "uniqueName"},
		Value: fmt.Sprintf("%v", m.UniqueNameAttr)})
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "caption"},
		Value: fmt.Sprintf("%v", m.CaptionAttr)})
	if m.UserAttr != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "user"},
			Value: fmt.Sprintf("%d", b2i(*m.UserAttr))})
	}
	if m.CustomRollUpAttr != nil {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "customRollUp"},
			Value: fmt.Sprintf("%d", b2i(*m.CustomRollUpAttr))})
	}
	e.EncodeToken(start)
	if m.Groups != nil {
		segroups := xml.StartElement{Name: xml.Name{Local: "ma:groups"}}
		e.EncodeElement(m.Groups, segroups)
	}
	if m.ExtLst != nil {
		seextLst := xml.StartElement{Name: xml.Name{Local: "ma:extLst"}}
		e.EncodeElement(m.ExtLst, seextLst)
	}
	e.EncodeToken(xml.EndElement{Name: start.Name})
	return nil
}

func (m *CT_GroupLevel) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// initialize to default
	for _, attr := range start.Attr {
		if attr.Name.Local == "uniqueName" {
			parsed, err := attr.Value, error(nil)
			if err != nil {
				return err
			}
			m.UniqueNameAttr = parsed
			continue
		}
		if attr.Name.Local == "caption" {
			parsed, err := attr.Value, error(nil)
			if err != nil {
				return err
			}
			m.CaptionAttr = parsed
			continue
		}
		if attr.Name.Local == "user" {
			parsed, err := strconv.ParseBool(attr.Value)
			if err != nil {
				return err
			}
			m.UserAttr = &parsed
			continue
		}
		if attr.Name.Local == "customRollUp" {
			parsed, err := strconv.ParseBool(attr.Value)
			if err != nil {
				return err
			}
			m.CustomRollUpAttr = &parsed
			continue
		}
	}
lCT_GroupLevel:
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch el := tok.(type) {
		case xml.StartElement:
			switch el.Name {
			case xml.Name{Space: "http://schemas.openxmlformats.org/spreadsheetml/2006/main", Local: "groups"}:
				m.Groups = NewCT_Groups()
				if err := d.DecodeElement(m.Groups, &el); err != nil {
					return err
				}
			case xml.Name{Space: "http://schemas.openxmlformats.org/spreadsheetml/2006/main", Local: "extLst"}:
				m.ExtLst = NewCT_ExtensionList()
				if err := d.DecodeElement(m.ExtLst, &el); err != nil {
					return err
				}
			default:
				gooxml.Log("skipping unsupported element on CT_GroupLevel %v", el.Name)
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			break lCT_GroupLevel
		case xml.CharData:
		}
	}
	return nil
}

// Validate validates the CT_GroupLevel and its children
func (m *CT_GroupLevel) Validate() error {
	return m.ValidateWithPath("CT_GroupLevel")
}

// ValidateWithPath validates the CT_GroupLevel and its children, prefixing error messages with path
func (m *CT_GroupLevel) ValidateWithPath(path string) error {
	if m.Groups != nil {
		if err := m.Groups.ValidateWithPath(path + "/Groups"); err != nil {
			return err
		}
	}
	if m.ExtLst != nil {
		if err := m.ExtLst.ValidateWithPath(path + "/ExtLst"); err != nil {
			return err
		}
	}
	return nil
}
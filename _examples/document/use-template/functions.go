package main

import "github.com/carmel/gooxml/schema/soo/wml"

// 处理word正文段落格式

type paragraph struct {
	FontName    string // 字体名称
	FontSize    int    //字号大小
	Align       byte   // 段落对齐
	LineSpacing int    // 行间距
	Indent      int    // 缩进
	Bold        bool   // 是否加粗
}

var ParagraphConfig = map[string]paragraph{
	// 文当标题
	"title": paragraph{
		FontName:    "方正小标宋简体",
		FontSize:    22, // 数字 22 字号，表示 二号
		Align:       byte(wml.ST_JcCenter),
		LineSpacing: 28,
		Indent:      0,
		Bold:        false,
	},
	// 可以理解为：第一层级
	"heading1": paragraph{
		FontName:    "黑体",
		FontSize:    16, // 数字 16 字号，表示 三号
		Align:       byte(wml.ST_JcLeft),
		LineSpacing: 28,
		Indent:      0,
		Bold:        false,
	},
	// 可以理解为：第二层级
	"heading2": paragraph{
		FontName:    "楷体",
		FontSize:    16, // 数字 16 字号，表示 三号
		Align:       byte(wml.ST_JcLeft),
		LineSpacing: 28,
		Indent:      0,
		Bold:        true,
	},
	// 三级标题与正文通用设置段落样式和属性即可
	"heading3": paragraph{},
	"heading4": paragraph{},
	"heading5": paragraph{},
	// word 正文样式
	"content": paragraph{
		FontName:    "微软雅黑",
		FontSize:    16, // 数字 16 字号，表示 三号
		Align:       byte(wml.ST_JcCenter),
		LineSpacing: 28,
		Indent:      2,
		Bold:        false,
	},
	// word 嵌套的表格样式
	"table": paragraph{
		FontName:    "黑体",
		FontSize:    12,                    // 数字 12 字号，表示 小四号
		Align:       byte(wml.ST_JcCenter), // 表格正文居中
		LineSpacing: 18,
		Indent:      0,
		Bold:        false,
	},
}

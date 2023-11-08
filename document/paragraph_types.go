package document

// ParagraphType 段落类型
type ParagraphType int

const (
	PTypeText      ParagraphType = iota // 段落中只有文本
	PTypeImage                          // 段落中只有图片
	PTypeTextImage                      // 段落中有 文本 + 图片
)

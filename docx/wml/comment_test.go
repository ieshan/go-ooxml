package wml

import (
	"encoding/xml"
	"testing"

	"github.com/ieshan/go-ooxml/xmlutil"
)

func TestCT_Comments_Unmarshal(t *testing.T) {
	input := `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:comment w:id="0" w:author="Joe" w:date="2024-01-01T00:00:00Z" w:initials="JS">
			<w:p><w:r><w:t>This is a comment</w:t></w:r></w:p>
		</w:comment>
		<w:comment w:id="1" w:author="Jane">
			<w:p><w:r><w:t>Another comment</w:t></w:r></w:p>
		</w:comment>
	</w:comments>`
	var comments CT_Comments
	if err := xmlutil.Unmarshal([]byte(input), &comments); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(comments.Comments) != 2 {
		t.Fatalf("count = %d, want 2", len(comments.Comments))
	}
	c := comments.Comments[0]
	if c.ID != 0 || c.Author != "Joe" || c.Initials != "JS" {
		t.Errorf("comment 0: %+v", c)
	}
	if len(c.Content) != 1 || c.Content[0].Paragraph == nil {
		t.Error("comment body missing paragraph")
	}
}

func TestCT_Comments_RoundTrip(t *testing.T) {
	c := CT_Comments{XMLName: xml.Name{Space: Ns, Local: "comments"}}
	comment := &CT_Comment{
		XMLName:  xml.Name{Space: Ns, Local: "comment"},
		ID:       0,
		Author:   "Test",
		Initials: "T",
	}
	p := &CT_P{XMLName: xml.Name{Space: Ns, Local: "p"}}
	p.AddRun("comment text")
	comment.Content = append(comment.Content, BlockLevelContent{Paragraph: p})
	c.Comments = append(c.Comments, comment)

	data, err := xmlutil.Marshal(&c, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var c2 CT_Comments
	if err := xmlutil.Unmarshal(data, &c2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(c2.Comments) != 1 {
		t.Error("lost comment")
	}
	if c2.Comments[0].Author != "Test" {
		t.Error("lost author")
	}
}

func TestCT_Comment_Text(t *testing.T) {
	input := `<w:comment xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="0" w:author="X">
		<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>World</w:t></w:r></w:p>
	</w:comment>`
	var c CT_Comment
	xmlutil.Unmarshal([]byte(input), &c)
	// Verify we can access text through content
	if len(c.Content) == 0 || c.Content[0].Paragraph == nil {
		t.Fatal("no content")
	}
	if c.Content[0].Paragraph.Text() != "Hello World" {
		t.Error("text")
	}
}

func TestCT_Comment_MultipleParagraphs(t *testing.T) {
	input := `<w:comment xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="5" w:author="Alice" w:date="2024-06-01T00:00:00Z" w:initials="A">
		<w:p><w:r><w:t>First paragraph.</w:t></w:r></w:p>
		<w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p>
		<w:p><w:r><w:t>Third paragraph.</w:t></w:r></w:p>
	</w:comment>`
	var c CT_Comment
	if err := xmlutil.Unmarshal([]byte(input), &c); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if c.ID != 5 || c.Author != "Alice" || c.Initials != "A" {
		t.Errorf("attrs: id=%d author=%q initials=%q", c.ID, c.Author, c.Initials)
	}
	if len(c.Content) != 3 {
		t.Fatalf("content = %d, want 3", len(c.Content))
	}
	for i, bc := range c.Content {
		if bc.Paragraph == nil {
			t.Errorf("content[%d] should be paragraph", i)
		}
	}
}

func TestCT_Comments_Empty(t *testing.T) {
	input := `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:comments>`
	var comments CT_Comments
	if err := xmlutil.Unmarshal([]byte(input), &comments); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(comments.Comments) != 0 {
		t.Errorf("count = %d, want 0", len(comments.Comments))
	}
	// Marshal an empty collection
	data, err := xmlutil.Marshal(&comments, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var comments2 CT_Comments
	if err := xmlutil.Unmarshal(data, &comments2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(comments2.Comments) != 0 {
		t.Errorf("count after round-trip = %d, want 0", len(comments2.Comments))
	}
}

func TestCT_CommentReference_MarshalUnmarshal(t *testing.T) {
	// Test CT_CommentReference inside a run round-trip
	r := CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r.Content = append(r.Content, RunContent{CommentReference: &CT_CommentReference{ID: 42}})
	data, err := xmlutil.Marshal(&r, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var r2 CT_R
	if err := xmlutil.Unmarshal(data, &r2); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(r2.Content) != 1 || r2.Content[0].CommentReference == nil {
		t.Fatal("CommentReference not preserved")
	}
	if r2.Content[0].CommentReference.ID != 42 {
		t.Errorf("ID = %d, want 42", r2.Content[0].CommentReference.ID)
	}
}

func TestCT_Comment_RoundTrip_WithDate(t *testing.T) {
	comments := CT_Comments{XMLName: xml.Name{Space: Ns, Local: "comments"}}
	comment := &CT_Comment{
		XMLName:  xml.Name{Space: Ns, Local: "comment"},
		ID:       7,
		Author:   "Bob",
		Date:     "2024-01-15T12:00:00Z",
		Initials: "B",
	}
	p := &CT_P{XMLName: xml.Name{Space: Ns, Local: "p"}}
	p.AddRun("note")
	comment.Content = append(comment.Content, BlockLevelContent{Paragraph: p})
	comments.Comments = append(comments.Comments, comment)

	data, err := xmlutil.Marshal(&comments, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var comments2 CT_Comments
	if err := xmlutil.Unmarshal(data, &comments2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(comments2.Comments) != 1 {
		t.Fatal("comment lost")
	}
	c2 := comments2.Comments[0]
	if c2.ID != 7 || c2.Author != "Bob" || c2.Date != "2024-01-15T12:00:00Z" {
		t.Errorf("attrs: id=%d author=%q date=%q", c2.ID, c2.Author, c2.Date)
	}
}

func TestCT_Comment_WithTable(t *testing.T) {
	// Comment containing a table — exercises table branch in CT_Comment.UnmarshalXML/MarshalXML
	input := `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:comment w:id="2" w:author="Alice">
			<w:p><w:r><w:t>See table:</w:t></w:r></w:p>
			<w:tbl>
				<w:tr><w:tc><w:p><w:r><w:t>data</w:t></w:r></w:p></w:tc></w:tr>
			</w:tbl>
		</w:comment>
	</w:comments>`
	var comments CT_Comments
	if err := xmlutil.Unmarshal([]byte(input), &comments); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(comments.Comments) != 1 {
		t.Fatal("comment count")
	}
	c := comments.Comments[0]
	if len(c.Content) != 2 {
		t.Fatalf("content = %d, want 2", len(c.Content))
	}
	if c.Content[1].Table == nil {
		t.Error("table not parsed in comment")
	}
	// Marshal to exercise table branch in CT_Comment.MarshalXML
	data, err := xmlutil.Marshal(&comments, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var comments2 CT_Comments
	if err := xmlutil.Unmarshal(data, &comments2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(comments2.Comments[0].Content) != 2 || comments2.Comments[0].Content[1].Table == nil {
		t.Error("table lost in comment round-trip")
	}
}

func TestCT_Comment_WithRaw(t *testing.T) {
	// Comment with unknown content preserved as Raw
	input := `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:comment w:id="3" w:author="Bob">
			<w:p><w:r><w:t>note</w:t></w:r></w:p>
			<w:customBlock/>
		</w:comment>
	</w:comments>`
	var comments CT_Comments
	if err := xmlutil.Unmarshal([]byte(input), &comments); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	c := comments.Comments[0]
	if len(c.Content) != 2 || c.Content[1].Raw == nil {
		t.Fatal("Raw content not parsed")
	}
	// Marshal to exercise raw branch in CT_Comment.MarshalXML
	data, err := xmlutil.Marshal(&comments, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var comments2 CT_Comments
	if err := xmlutil.Unmarshal(data, &comments2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(comments2.Comments[0].Content) != 2 || comments2.Comments[0].Content[1].Raw == nil {
		t.Error("raw content lost in comment round-trip")
	}
}

func TestCT_Comments_UnknownChild_Skipped(t *testing.T) {
	// Unknown children in comments root should be skipped
	input := `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:comment w:id="0" w:author="A"><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:comment>
		<w:unknownChild/>
		<w:comment w:id="1" w:author="B"><w:p><w:r><w:t>bye</w:t></w:r></w:p></w:comment>
	</w:comments>`
	var comments CT_Comments
	if err := xmlutil.Unmarshal([]byte(input), &comments); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(comments.Comments) != 2 {
		t.Errorf("count = %d, want 2", len(comments.Comments))
	}
}

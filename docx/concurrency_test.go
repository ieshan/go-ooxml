package docx

import (
	"sync"
	"testing"
)

func TestConcurrent_Reads(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("test content")
	doc.Body().AddTable(2, 2)

	var wg sync.WaitGroup
	for range 100 {
		wg.Add(1)
		go func() {
			defer wg.Done()
			_ = doc.Text(nil)
			for _ = range doc.Body().Paragraphs() {
			}
			for _ = range doc.Body().Tables() {
			}
			for _ = range doc.Comments() {
			}
		}()
	}
	wg.Wait()
}

func TestConcurrent_FindOperations(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	for range 10 {
		doc.Body().AddParagraph().AddRun("searchable text here")
	}

	var wg sync.WaitGroup
	for range 50 {
		wg.Add(1)
		go func() {
			defer wg.Done()
			results := doc.Find("searchable")
			_ = len(results)
		}()
	}
	wg.Wait()
}

func TestConcurrent_MarkdownExtraction(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.SetStyle("Heading1")
	p.AddRun("Title")
	doc.Body().AddParagraph().AddRun("Content")

	var wg sync.WaitGroup
	for range 50 {
		wg.Add(1)
		go func() {
			defer wg.Done()
			_ = doc.Markdown(nil)
			_ = doc.Text(nil)
		}()
	}
	wg.Wait()
}

func TestConcurrent_NoDeadlock_SearchComment(t *testing.T) {
	// Verify the fluent API pattern doesn't deadlock
	doc, _ := New(&Config{Author: "Bot"})
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("target text")

	// This should not deadlock — Search collects under read lock,
	// then AddComment acquires write lock
	comments := doc.Search("target").AddComment("found it")
	if len(comments) != 1 {
		t.Errorf("comments = %d", len(comments))
	}
}

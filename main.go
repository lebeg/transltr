package main

import (
	"bufio"
	"context"
	"flag"
	"fmt"
	"log"
	"os"

	"github.com/unidoc/unioffice/color"
	"github.com/unidoc/unioffice/document"
	"golang.org/x/text/language"

	"github.com/unidoc/unioffice/schema/soo/wml"

	"cloud.google.com/go/translate"
)

type docRun struct {
	properties  document.RunProperties
	original    string
	translation string
}

type docParagraph struct {
	properties document.ParagraphProperties
	style      string
	runs       []docRun
}

func run() error {
	wordPtr := flag.String("doc", "", "Path to a string")
	flag.Parse()

	if *wordPtr == "" {
		log.Fatal("Document path is required")
	}

	fmt.Printf("Processing %s...\n", *wordPtr)

	inputDoc, err := document.Open(*wordPtr)
	if err != nil {
		return fmt.Errorf("error opening document: %s", err)
	}

	target, err := language.Parse("en")
	if err != nil {
		return fmt.Errorf("Failed to parse target language: %v", err)
	}

	paragraphs := []document.Paragraph{}
	for _, p := range inputDoc.Paragraphs() {
		paragraphs = append(paragraphs, p)
	}
	for _, sdt := range inputDoc.StructuredDocumentTags() {
		for _, p := range sdt.Paragraphs() {
			paragraphs = append(paragraphs, p)
		}
	}

	docParagraphs := []docParagraph{}

	ctx := context.Background()

	client, err := translate.NewClient(ctx)
	if err != nil {
		return fmt.Errorf("Failed to create client: %v", err)
	}

	paragraphCounter := 0
	runCounter := 0

	for _, paragraph := range paragraphs {
		docP := docParagraph{
			style:      paragraph.Style(),
			properties: paragraph.Properties(),
		}

		texts := []string{}

		for _, inputRun := range paragraph.Runs() {
			text := inputRun.Text()
			if text != "" {
				texts = append(texts, text)
			}
		}

		if len(texts) == 0 {
			continue
		}

		translations, err := client.Translate(ctx, texts, target, nil)
		if err != nil {
			return fmt.Errorf("Failed to translate text: %v", err)
		}

		for i, tr := range translations {
			docP.runs = append(docP.runs, docRun{
				properties:  paragraph.Runs()[i].Properties(),
				original:    texts[i],
				translation: tr.Text,
			})
		}

		docParagraphs = append(docParagraphs, docP)
	}

	outputDoc := document.New()

	table := outputDoc.AddTable()
	// width of the page
	table.Properties().SetWidthPercent(100)
	// with thick borers
	borders := table.Properties().Borders()
	borders.SetAll(wml.ST_BorderNone, color.Auto, 0)

	for _, p := range docParagraphs {
		paragraphCounter++
		row := table.AddRow()
		inputCell := row.AddCell()
		inputParagraph := inputCell.AddParagraph()
		inputParagraph.SetStyle(p.style)
		outputCell := row.AddCell()
		outputParagraph := outputCell.AddParagraph()
		outputParagraph.SetStyle(p.style)
		for _, docR := range p.runs {
			runCounter++

			inputRun := inputParagraph.AddRun()
			inputRun.AddText(docR.original)
			inputRun.Properties().SetBold(docR.properties.IsBold() || p.properties.IsBold())
			outputRun := outputParagraph.AddRun()
			outputRun.AddText(docR.translation)
			outputRun.Properties().SetBold(docR.properties.IsBold())
		}
	}

	fmt.Printf("Processed %v paragraphs and %v runs\n", paragraphCounter, runCounter)

	f, err := os.Create("./document-output.docx")
	if err != nil {
		return fmt.Errorf("Failed to open file document-output.docx: %v", err)
	}

	defer f.Close()

	writer := bufio.NewWriter(f)
	err = outputDoc.Save(writer)
	if err != nil {
		return fmt.Errorf("Failed to save document: %v", err)
	}

	return nil
}

func main() {
	if err := run(); err != nil {
		log.Fatalf("Error %v", err)
		os.Exit(1)
	}

}

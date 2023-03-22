package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"net/http"
	"os"
	"regexp"
)

type Context struct {
	replacements ReplacementMap
}
type ReplacementMap map[string]string

func XMLCopyHandler(file *zip.File, w *zip.Writer) {
	fw, err := w.Create(file.Name)
	if err != nil {
		panic(err)
	}

	reader, err := file.Open()
	if err != nil {
		panic(err)
	}
	defer reader.Close()

	// Read document

	content, err := ioutil.ReadAll(reader)
	if err != nil {
		panic(err)
	}

	fw.Write(content)
}

func XMLDocumentHandler(file *zip.File, w *zip.Writer, context *Context) {
	reader, err := file.Open()
	if err != nil {
		panic(err)
	}
	defer reader.Close()

	// Read document

	content, err := ioutil.ReadAll(reader)
	if err != nil {
		panic(err)
	}

	updatedContent := processFile(content, context)

	fmt.Println(string(updatedContent))

	fw, err := w.Create(file.Name)
	if err != nil {
		panic(err)
	}

	fw.Write(updatedContent)

}

func handleXMLFile(file *zip.File, w *zip.Writer, context *Context) {
	switch file.Name {
	case "word/document.xml":
		XMLDocumentHandler(file, w, context)
		break
	default:
		XMLCopyHandler(file, w)
	}
}

func processFile(templateData []byte, context *Context) []byte {
	re := regexp.MustCompile("{{([a-zA-Z]\\w+)}}")
	updatedContent := re.ReplaceAllFunc(templateData, func(match []byte) []byte {
		stringMatch := string(match)
		withoutCurly := stringMatch[2 : len(stringMatch)-2]
		_, ok := context.replacements[withoutCurly]
		fmt.Println(withoutCurly)
		if !ok {
			return []byte(fmt.Sprintf("INVALID [%s]", string(match)))
		}
		return []byte(context.replacements[withoutCurly])
	})
	return updatedContent
}

// ========================= File Reader ===================================

func processFiles(inputPath string, outputPath string) {
	context := Context{
		replacements: ReplacementMap{
			"Title":          "something...",
			"more_text":      "Once upon a time there was a Go program that could template docx documents...",
			"even_more_text": "This is an item",
		},
	}

	file, err := zip.OpenReader(inputPath)
	if err != nil {
		panic(err)
	}
	defer file.Close()

	// Writer

	var newFileBuf bytes.Buffer
	newFileWriter := zip.NewWriter(&newFileBuf)

	// Handle

	for _, f := range file.File {
		handleXMLFile(f, newFileWriter, &context)
	}

	newFileWriter.Close()

	ioutil.WriteFile(outputPath, newFileBuf.Bytes(), 0644)
}

// ========================= Web Server ===================================

func processWebData() {
	fmt.Println("Listening on port 4222...")
	http.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		fmt.Fprintf(w, "Hello, there! %d\n", 3)
	})
	http.HandleFunc("/data", func(w http.ResponseWriter, r *http.Request) {

		if r.Method == "POST" {

			err := r.ParseMultipartForm(5 << 20)
			if err != nil {
				http.Error(w, err.Error(), 400)
			}

			jsonDataFile, _, err := r.FormFile("replacements")
			if err != nil {
				http.Error(w, err.Error(), 400)
			}

			jsonData, err := ioutil.ReadAll(jsonDataFile)
			if err != nil {
				http.Error(w, err.Error(), 400)
			}

			rmap := make(ReplacementMap)
			err = json.Unmarshal(jsonData, &rmap)

			// Template file

			templateDataFile, _, err := r.FormFile("template")
			if err != nil {
				http.Error(w, err.Error(), 400)
			}

			templateData, err := ioutil.ReadAll(templateDataFile)
			if err != nil {
				http.Error(w, err.Error(), 400)
			}

			// Make context

			context := &Context{
				replacements: rmap,
			}

			processedFile := processFile(templateData, context)

			// Write zip

			var newFileBuf bytes.Buffer
			newFileWriter := zip.NewWriter(&newFileBuf)

			fw, err := newFileWriter.Create("output.docx")
			if err != nil {
				panic(err)
			}
			fw.Write(processedFile)

			w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
			w.Write(newFileBuf.Bytes())

			return
		}
		fmt.Fprintf(w, "That was NOT a post\n")
	})

	http.ListenAndServe("localhost:4222", nil)

}

// ========================= Entry Point ===================================

func main() {

	if len(os.Args) < 2 {
		fmt.Print("USAGE:\n")
		fmt.Print("gocx [files|serve]\n")
		os.Exit(1)
	}

	switch os.Args[1] {
	case "serve":
		processWebData()
		break
	case "files":
		if len(os.Args) < 4 {
			fmt.Print("USAGE:\n")
			fmt.Print("gocx files [inputPath] [outputPath]\n")
			os.Exit(1)
		}
		processFiles(os.Args[2], os.Args[3])
		break
	default:
		fmt.Print("USAGE:\n")
		fmt.Print("gocx [files|serve]\n")
		os.Exit(1)
	}

}

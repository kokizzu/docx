package docx

import (
	`archive/zip`
	`bufio`
	`bytes`
	`encoding/xml`
	`errors`
	`io`
	`io/ioutil`
	`os`
	`strings`
)

type ReplaceDocx struct {
	zipReader *zip.ReadCloser
	content   string
	header    string
	footer    string
}

func (r *ReplaceDocx) Editable() *Docx {
	return &Docx{
		files:   r.zipReader.File,
		content: r.content,
		header: r.header,
		footer: r.footer,
	}
}

func (r *ReplaceDocx) Close() error {
	return r.zipReader.Close()
}

type Docx struct {
	files   []*zip.File
	content string
	header  string
	footer  string
}

func (d *Docx) ReplaceContentRaw(oldString string, newString string, num int) (err error) {
	oldString, err = encode(oldString)
	if err != nil {
		return err
	}
	newString, err = encode(newString)
	if err != nil {
		return err
	}
	newString = strings.Replace(newString,`&#xA;`,`</w:t><w:br/><w:t>`,-1)
	d.content = strings.Replace(d.content, oldString, newString, num)
	return nil
}

func (d *Docx) ReplaceContent(oldString string, newString string, num int) (err error) {
	oldString, err = encode(oldString)
	if err != nil {
		return err
	}
	newString, err = encode(newString)
	if err != nil {
		return err
	}
	d.content = strings.Replace(d.content, oldString, newString, num)
	return nil
}

func (d *Docx) ReplaceHeader(oldString string, newString string, num int) (err error) {
	oldString, err = encode(oldString)
	if err != nil {
		return err
	}
	newString, err = encode(newString)
	if err != nil {
		return err
	}
	d.header = strings.Replace(d.header, oldString, newString, num)

	return nil
}

func (d *Docx) ReplaceFooter(oldString string, newString string, num int) (err error) {
	oldString, err = encode(oldString)
	if err != nil {
		return err
	}
	newString, err = encode(newString)
	if err != nil {
		return err
	}
	d.footer = strings.Replace(d.footer, oldString, newString, num)

	return nil
}

func (d *Docx) WriteToFile(path string) (err error) {
	var target *os.File
	target, err = os.Create(path)
	if err != nil {
		return
	}
	defer target.Close()
	err = d.Write(target)
	return
}

func (d *Docx) Write(ioWriter io.Writer) (err error) {
	w := zip.NewWriter(ioWriter)
	for _, file := range d.files {
		var writer io.Writer
		var readCloser io.ReadCloser

		writer, err = w.Create(file.Name)
		if err != nil {
			return err
		}
		readCloser, err = file.Open()
		if err != nil {
			return err
		}
		switch file.Name {
		case `word/document.xml`:
			writer.Write([]byte(d.content))
		case `word/header1.xml`:
			writer.Write([]byte(d.header))
		case `word/footer1.xml`:
			writer.Write([]byte(d.footer))
		default:
			writer.Write(streamToByte(readCloser))
		}
	}
	w.Close()
	return
}

func ReadDocxFile(cpath string) (*ReplaceDocx, error) {
	reader, err := zip.OpenReader(cpath)
	if err != nil {
		return nil, err
	}
	content, header, footer, err := readText(reader.File)
	if err != nil {
		return nil, err
	}

	return &ReplaceDocx{zipReader: reader, content: content, header: header, footer: footer}, nil
}

func readText(files []*zip.File) (content, header, footer string, err error) {
	content, err = retrieveWordXml(files, `document.xml`)
	header, err = retrieveWordXml(files, `header1.xml`)
	footer, err = retrieveWordXml(files,`footer1.xml`)
	return
}


func retrieveWordXml(files []*zip.File, name string) (string, error) {
	var file *zip.File
	for _, f := range files {
		if f.Name == `word/` + name {
			file = f
		}
	}
	if file == nil {
		return ``, errors.New(`document.xml file not found`)
	}
	documentReader, err := file.Open()
	if err != nil {
		return ``, err
	}
	b, err := ioutil.ReadAll(documentReader)
	if err != nil {
		return ``, err
	}
	return string(b), nil
}

func streamToByte(stream io.Reader) []byte {
	buf := new(bytes.Buffer)
	buf.ReadFrom(stream)
	return buf.Bytes()
}

func encode(s string) (string, error) {
	var b bytes.Buffer
	enc := xml.NewEncoder(bufio.NewWriter(&b))
	if err := enc.Encode(s); err != nil {
		return s, err
	}
	return strings.Replace(strings.Replace(b.String(), `<string>`, ``, 1), `</string>`, ``, 1), nil
}

## Simple Google Go (golang) library for replace text in microsoft word (.docx) file

The following constitutes the bare minimum required to replace text in DOCX document.
``` go 

import (
	"github.com/kokizzu/docx"
)

func main() {
	r, err := docx.ReadDocxFile("./template.docx")
	if err != nil {
		panic(err)
	}
	docx1 := r.Editable()
	docx1.ReplaceContent("old_1_1", "new_1_1", -1)
	docx1.ReplaceContentRaw("old_1_2", "new_1_2", -1) // replace newline with `</w:t><w:br/><w:t>``
	docx1.WriteToFile("./new_result_1.docx")

	docx2 := r.Editable()
	docx2.ReplaceHeader("old_2_1", "new_2_1", -1)
	docx2.ReplaceFooter("old_2_2", "new_2_2", -1)
	docx2.WriteToFile("./new_result_2.docx")

	r.Close()
}

```
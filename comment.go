package excelize

import (
	"encoding/json"
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
)

// parseFormatCommentsSet provides a function to parse the format settings of
// the comment with default value.
func parseFormatCommentsSet(formatSet string) (*formatComment, error) {
	format := formatComment{
		Author: "Author:",
		Text:   " ",
	}
	err := json.Unmarshal([]byte(formatSet), &format)
	return &format, err
}

// GetComments retrieves all comments and returns a map of worksheet name to
// the worksheet comments.
func (f *File) GetComments() (comments map[string]*xlsxComments) {
	comments = map[string]*xlsxComments{}
	for n := range f.sheetMap {
		commentID := f.GetSheetIndex(n)
		commentsXML := "xl/comments" + strconv.Itoa(commentID) + ".xml"
		c, ok := f.XLSX[commentsXML]
		if ok {
			d := xlsxComments{}
			xml.Unmarshal([]byte(c), &d)
			comments[n] = &d
		}
	}
	return
}

// AddComment provides the method to add comment in a sheet by given worksheet
// index, cell and format set (such as author and text). Note that the max
// author length is 255 and the max text length is 32512. For example, add a
// comment in Sheet1!$A$30:
//
//    xlsx.AddComment("Sheet1", "A30", `{"author":"Excelize: ","text":"This is a comment."}`)
//
func (f *File) AddComment(sheet, cell, format string) error {
	yAxis, xAxis, err := CellCoords(cell)
	if err != nil {
		return err
	}
	return f.AddCommentAt(sheet, cell, format, xAxis+1, yAxis+1)
}

// AddComment provides the method to add comment in a sheet by given worksheet
// index, cell and format set (such as author and text). Note that the max
// author length is 255 and the max text length is 32512. For example, add a
// comment in Sheet1!$A$30:
//
//    xlsx.AddComment("Sheet1", "A30", `{"author":"Excelize: ","text":"This is a comment."}`)
//
func (f *File) AddCommentAt(sheet, cell, format string, col, row int) error {
	formatSet, err := parseFormatCommentsSet(format)
	if err != nil {
		return err
	}
	col1Width := f.GetColWidth(sheet, ColIndexToLetters(col+1))
	col2Width := f.GetColWidth(sheet, ColIndexToLetters(col+2))
	// maxChar := ((col1Width+col2Width)/10.452839 - 0.007) / 0.17390901
	// maxChar := (col1Width/5.2264195 - 0.007) / 0.17390901
	maxChar := ((col1Width+col2Width)/5.2264195 - 0.007) / 0.15
	// maxChar := ((col1Width+col2Width)/5.2264195 - 0.007) / 0.17390901
	// Read sheet data.
	xlsx := f.workSheetReader(sheet)
	commentID := f.countComments() + 1
	drawingVML := "xl/drawings/vmlDrawing" + strconv.Itoa(commentID) + ".vml"
	sheetRelationshipsComments := "../comments" + strconv.Itoa(commentID) + ".xml"
	sheetRelationshipsDrawingVML := "../drawings/vmlDrawing" + strconv.Itoa(commentID) + ".vml"
	if xlsx.LegacyDrawing != nil {
		// The worksheet already has a comments relationships, use the relationships drawing ../drawings/vmlDrawing%d.vml.
		sheetRelationshipsDrawingVML = f.getSheetRelationshipsTargetByID(sheet, xlsx.LegacyDrawing.RID)
		commentID, _ = strconv.Atoi(strings.TrimSuffix(strings.TrimPrefix(sheetRelationshipsDrawingVML, "../drawings/vmlDrawing"), ".vml"))
		drawingVML = strings.Replace(sheetRelationshipsDrawingVML, "..", "xl", -1)
	} else {
		// Add first comment for given sheet.
		rID := f.addSheetRelationships(sheet, SourceRelationshipDrawingVML, sheetRelationshipsDrawingVML, "")
		f.addSheetRelationships(sheet, SourceRelationshipComments, sheetRelationshipsComments, "")
		f.addSheetLegacyDrawing(sheet, rID)
	}
	commentsXML := "xl/comments" + strconv.Itoa(commentID) + ".xml"
	f.addComment(commentsXML, cell, formatSet)
	var hight float64
	for i, l := range strings.Split(formatSet.Text, "\n") {
		ll := len(l)
		if i == 0 {
			ll += len(formatSet.Author)
		}
		hight += float64(ll) / maxChar
	}
	// f.addDrawingVML(commentID, drawingVML, cell, strings.Count(formatSet.Text, "\n")+1, colCount)
	if hight < 0.5 {
		hight = 1.0
	}
	f.addBoxDrawingVML(
		commentID, drawingVML, cell,
		col, 23, row, 5, col+2, 23, row+int(hight+0.5), 5)

	//1+xAxis, 23, 1+yAxis, 0, 2+xAxis+lineCount, colCount+xAxis, 2+yAxis+lineCount, 5)
	// leftColumn, leftOffset, topRow, topOffset, rightColumn, rightOffset, bottomRow, bottomOffset int) {
	f.addContentTypePart(commentID, "comments")
	return err
}

// addDrawingVML provides a function to create comment as
// xl/drawings/vmlDrawing%d.vml by given commit ID and cell.
func (f *File) addDrawingVML(commentID int, drawingVML, cell string, lineCount, colCount int) {
	col := string(strings.Map(letterOnlyMapF, cell))
	row, _ := strconv.Atoi(strings.Map(intOnlyMapF, cell))
	xAxis := TitleToNumber(col)
	yAxis := row - 1

	f.addBoxDrawingVML(commentID, drawingVML, cell,
		1+xAxis, 23, 1+yAxis, 0, 2+xAxis+lineCount, colCount+xAxis, 2+yAxis+lineCount, 5)
}

// addDrawingVML provides function to create comment as
// xl/drawings/vmlDrawing%d.vml by given commit ID and cell.
func (f *File) addBoxDrawingVML(commentID int, drawingVML, cell string,
	leftColumn, leftOffset, topRow, topOffset, rightColumn, rightOffset, bottomRow, bottomOffset int) {
	col := string(strings.Map(letterOnlyMapF, cell))
	row, _ := strconv.Atoi(strings.Map(intOnlyMapF, cell))
	xAxis := TitleToNumber(col)
	yAxis := row - 1
	vml := vmlDrawing{
		XMLNSv:  "urn:schemas-microsoft-com:vml",
		XMLNSo:  "urn:schemas-microsoft-com:office:office",
		XMLNSx:  "urn:schemas-microsoft-com:office:excel",
		XMLNSmv: "http://macVmlSchemaUri",
		Shapelayout: &xlsxShapelayout{
			Ext: "edit",
			IDmap: &xlsxIDmap{
				Ext:  "edit",
				Data: commentID,
			},
		},
		Shapetype: &xlsxShapetype{
			ID:        "_x0000_t202",
			Coordsize: "21600,21600",
			Spt:       202,
			Path:      "m0,0l0,21600,21600,21600,21600,0xe",
			Stroke: &xlsxStroke{
				Joinstyle: "miter",
			},
			VPath: &vPath{
				Gradientshapeok: "t",
				Connecttype:     "miter",
			},
		},
	}
	sp := encodeShape{
		Fill: &vFill{
			Color2: "#fbfe82",
			Angle:  -180,
			Type:   "gradient",
			Fill: &oFill{
				Ext:  "view",
				Type: "gradientUnscaled",
			},
		},
		Shadow: &vShadow{
			On:       "t",
			Color:    "black",
			Obscured: "t",
		},
		Path: &vPath{
			Connecttype: "none",
		},
		Textbox: &vTextbox{
			Style: "mso-direction-alt:auto",
			Div: &xlsxDiv{
				Style: "text-align:left",
			},
		},
		ClientData: &xClientData{
			ObjectType: "Note",
			Anchor: fmt.Sprintf(
				"%d, %d, %d, %d, %d, %d, %d, %d",
				leftColumn, leftOffset, topRow, topOffset, rightColumn, rightOffset, bottomRow, bottomOffset),
			AutoFill: "True",
			Row:      yAxis,
			Column:   xAxis,
		},
	}
	s, _ := xml.Marshal(sp)
	shape := xlsxShape{
		ID:          "_x0000_s1025",
		Type:        "#_x0000_t202",
		Style:       "position:absolute;73.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden",
		Fillcolor:   "#fbf6d6",
		Strokecolor: "#edeaa1",
		Val:         string(s[13 : len(s)-14]),
	}
	c, ok := f.XLSX[drawingVML]
	if ok {
		d := decodeVmlDrawing{}
		_ = xml.Unmarshal([]byte(c), &d)
		for _, v := range d.Shape {
			s := xlsxShape{
				ID:          "_x0000_s1025",
				Type:        "#_x0000_t202",
				Style:       "position:absolute;73.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden",
				Fillcolor:   "#fbf6d6",
				Strokecolor: "#edeaa1",
				Val:         v.Val,
			}
			vml.Shape = append(vml.Shape, s)
		}
	}
	vml.Shape = append(vml.Shape, shape)
	v, _ := xml.Marshal(vml)
	f.XLSX[drawingVML] = v
}

// addComment provides a function to create chart as xl/comments%d.xml by
// given cell and format sets.
func (f *File) addComment(commentsXML, cell string, formatSet *formatComment) {
	a := formatSet.Author
	t := formatSet.Text
	if len(a) > 255 {
		a = a[0:255]
	}
	if len(t) > 32512 {
		t = t[0:32512]
	}
	comments := xlsxComments{
		Authors: []xlsxAuthor{
			{
				Author: formatSet.Author,
			},
		},
	}
	cmt := xlsxComment{
		Ref:      cell,
		AuthorID: 0,
		Text: xlsxText{
			R: []xlsxR{
				{
					RPr: &xlsxRPr{
						B:  " ",
						Sz: &attrValFloat{Val: 9},
						Color: &xlsxColor{
							Indexed: 81,
						},
						RFont:  &attrValString{Val: "Calibri"},
						Family: &attrValInt{Val: 2},
					},
					T: a,
				},
				{
					RPr: &xlsxRPr{
						Sz: &attrValFloat{Val: 9},
						Color: &xlsxColor{
							Indexed: 81,
						},
						RFont:  &attrValString{Val: "Calibri"},
						Family: &attrValInt{Val: 2},
					},
					T: t,
				},
			},
		},
	}
	c, ok := f.XLSX[commentsXML]
	if ok {
		d := xlsxComments{}
		_ = xml.Unmarshal([]byte(c), &d)
		comments.CommentList.Comment = append(comments.CommentList.Comment, d.CommentList.Comment...)
	}
	comments.CommentList.Comment = append(comments.CommentList.Comment, cmt)
	v, _ := xml.Marshal(comments)
	f.saveFileList(commentsXML, v)
}

// countComments provides a function to get comments files count storage in
// the folder xl.
func (f *File) countComments() int {
	count := 0
	for k := range f.XLSX {
		if strings.Contains(k, "xl/comments") {
			count++
		}
	}
	return count
}

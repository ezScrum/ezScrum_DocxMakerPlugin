package plugin.docxMaker;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import org.apache.commons.io.IOUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.codehaus.jettison.json.JSONArray;
import org.codehaus.jettison.json.JSONException;
import org.codehaus.jettison.json.JSONObject;
import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.Br;
import org.docx4j.wml.CTVerticalJc;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase.Spacing;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STBrType;
import org.docx4j.wml.STVerticalJc;
import org.docx4j.wml.Style;
import org.docx4j.wml.Styles;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblGrid;
import org.docx4j.wml.TblGridCol;
import org.docx4j.wml.TblPr;
import org.docx4j.wml.TblWidth;
import org.docx4j.wml.Tc;
import org.docx4j.wml.TcPr;
import org.docx4j.wml.Tr;

public class DocxMaker {
	private static Log log = LogFactory.getLog(DocxMaker.class);
	private ObjectFactory mFactory;
	private WordprocessingMLPackage wordMLPackage;
	final String[] title = {"ID", "Name", "Est.", "Handler", "Partners", "Notes"};
	final int FONT_SIZE = 12 * 2;	// 轉成word需要*2才是正確的大小

	public DocxMaker() {}

	/**
	 * get release plan docx file
	 * 
	 * @param releaseJSON - a JSON contain all release data
	 */
	public File getReleasePlanDocx(JSONObject releaseJSON) throws JSONException {
		try {
			wordMLPackage = WordprocessingMLPackage.createPackage();	// Create the package
			MainDocumentPart mainDoc = wordMLPackage.getMainDocumentPart();
			mFactory = Context.getWmlObjectFactory();
			alterStyleSheet();					// change the style of this docx
			addReleaseInfo(mainDoc, releaseJSON);	// the title of docx
			JSONArray sprintJSONArray = releaseJSON.getJSONArray("sprints");
			for (int i = 0, sprintSize = sprintJSONArray.length(); i < sprintSize; i++) {
				JSONObject sprintJSON = sprintJSONArray.getJSONObject(i);
				int totalStoryPoints = sprintJSON.getInt("totalStoryPoint");
				addSprintInfo(mainDoc, sprintJSON, totalStoryPoints);
				JSONArray storyJSONArray = sprintJSON.getJSONArray("stories");
				addStoryInfo(mainDoc, storyJSONArray);
				if (i != sprintSize) addPageBreak(mainDoc);	// add new page
			}
			String filePath = releaseJSON.get("name") + "_ReleasePlan.docx";
			wordMLPackage.save(new File(filePath));	// Save it
			return new File(filePath);
		} catch (Docx4JException e) {
			e.printStackTrace();
		} catch (JAXBException e) {
			e.printStackTrace();
		}
		return null;
	}

	/**
	 * add release info as title
	 * 
	 * @param mainDoc - the Doc content
	 * @param releaseJSON - a JSON contain all release data
	 * @throws JSONException 
	 */
	private void addReleaseInfo(MainDocumentPart mainDoc, JSONObject releaseJSON) throws JSONException {
		mainDoc.addStyledParagraphOfText("Title", "Release Plan #" + releaseJSON.get("serial_id") + "： " + releaseJSON.get("name"));
		mainDoc.addStyledParagraphOfText("Subtitle", "Start Date： " + releaseJSON.get("start_date"));
		mainDoc.addStyledParagraphOfText("Subtitle", "End Date： " + releaseJSON.get("end_date"));
		mainDoc.addStyledParagraphOfText("Subtitle", "Description： " + releaseJSON.get("description"));
	}

	/**
	 * add sprint info above its stories
	 * 
	 * @param mainDoc - the Doc content
	 * @param sprintJSON - a JSON contain all sprint data
	 * @param tatolStoryPoints - sprint stories estimate
	 * @throws JSONException 
	 */
	private void addSprintInfo(MainDocumentPart mainDoc, JSONObject sprintJSON, int tatolStoryPoints) throws JSONException {
		mainDoc.addStyledParagraphOfText("Heading1", "Sprint #" + sprintJSON.get("serial_id") + "： " + sprintJSON.get("goal"));
		mainDoc.addStyledParagraphOfText("Subtitle", "Start Date： " + sprintJSON.get("start_date"));
		mainDoc.addStyledParagraphOfText("Subtitle", "End Date： " + sprintJSON.get("end_date"));
		mainDoc.addStyledParagraphOfText("Subtitle", "Total Story Points： " + tatolStoryPoints);
	}

	/**
	 * add story info below its sprint
	 * 
	 * @param mainDoc - the Doc content
	 * @param storyJSONArray - a JSON array contain all stories in the same release
	 * @throws JAXBException
	 * @throws JSONException 
	 */
	private void addStoryInfo(MainDocumentPart mainDoc, JSONArray storyJSONArray) throws JAXBException, JSONException {
		try {
			String tblXML = IOUtils.toString(new FileReader("StoryCard.xml"));	// 使用預設的Story Card 格式
			for (int i = 0; i < storyJSONArray.length(); i++) {
				JSONObject storyJSON = storyJSONArray.getJSONObject(i);
				mainDoc.addObject(getStoryTable(mainDoc, storyJSON, tblXML));
				mainDoc.addParagraphOfText("Task List: ");
				JSONArray taskJSONArray = storyJSON.getJSONArray("tasks");
				mainDoc.addObject(addTaskInfo(mainDoc, taskJSONArray));
				mainDoc.addParagraphOfText("");
				mainDoc.addParagraphOfText(" - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -");
				mainDoc.addParagraphOfText("");
			}
		} catch (JAXBException e) {
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * add data to story card as docx table
	 * 
	 * @param mainDoc - the Doc content
	 * @param storyJSON - a JSON contain story data
	 * @param tblXML - the table xml for story card
	 * @return the story table
	 * @throws JAXBException
	 * @throws FileNotFoundException
	 * @throws IOException
	 * @throws JSONException 
	 */
	private Tbl getStoryTable(MainDocumentPart mainDoc, JSONObject storyJSON, String tblXML) throws JAXBException, FileNotFoundException, IOException, JSONException {
		Tbl table = (Tbl) XmlUtils.unmarshalString(tblXML);	// row 2, 4, 5, 7, 9  是 table 的空白格
		Tr row = null;
		row = (Tr) table.getContent().get(0);	// get ID
		JAXBElement<?> element = (JAXBElement<?>) row.getContent().get(0);
		Tc tc = (Tc) element.getValue();
		tc.getContent().set(0, mainDoc.createParagraphOfText("Sprint Backlog Item #" + storyJSON.get("serial_id")));							// add ID
		row = (Tr) table.getContent().get(1);	// get Name
		element = (JAXBElement<?>) row.getContent().get(0);
		tc = (Tc) element.getValue();
		tc.getContent().set(0, mainDoc.createParagraphOfText(storyJSON.getString("name")));													// add Name
		setStoryDataToTableColumn((JAXBElement<?>) row.getContent().get(1), mainDoc.createStyledParagraphOfText("Bold_Number", String.valueOf(storyJSON.get("importance"))), true);	// add Importance
		row = (Tr) table.getContent().get(3);	// get Notes
		setStoryDataToTableColumn((JAXBElement<?>) row.getContent().get(0), mainDoc.createParagraphOfText(storyJSON.getString("notes")));	// add Notes
		setStoryDataToTableColumn((JAXBElement<?>) row.getContent().get(1), mainDoc.createStyledParagraphOfText("Bold_Number", String.valueOf(storyJSON.get("estimate"))), true);	// add Estimate
		row = (Tr) table.getContent().get(6);	// get Tags
		String tags = "";
		JSONArray tagJSONArray = storyJSON.getJSONArray("tags");
		for (int i = 0; i < tagJSONArray.length(); i++) {
			if (!tags.isEmpty()) tags = tags + ", ";
			JSONObject tagJSON = tagJSONArray.getJSONObject(i);
			tags = tags + tagJSON.getString("name");
		}
		setStoryDataToTableColumn((JAXBElement<?>) row.getContent().get(0), mainDoc.createParagraphOfText(tags));			// add Tags
		row = (Tr) table.getContent().get(8);	// get How to Demo
		setStoryDataToTableColumn((JAXBElement<?>) row.getContent().get(0), mainDoc.createParagraphOfText(storyJSON.getString("how_to_demo")));// add How to Demo
		return table;
	}

	/**
	 * analysis the table content and set story data
	 * 
	 * @param columnElement - the column element, like name, notes...
	 * @param paragraph - p that you want to add to table
	 * @param isSetCenter - if you like to set table cell content to center, set true
	 */
	private void setStoryDataToTableColumn(JAXBElement<?> columnElement, P paragraph, boolean isSetCenter) {
		Tc tc = (Tc) columnElement.getValue();
		JAXBElement<?> element = (JAXBElement<?>) tc.getContent().get(1);
		Tbl howToDemoTable = (Tbl) element.getValue();
		Tr tr = (Tr) howToDemoTable.getContent().get(0);
		element = (JAXBElement<?>) tr.getContent().get(0);
		tc = (Tc) element.getValue();
		if (isSetCenter) {
			TcPr tcPr = tc.getTcPr();
			if (tcPr == null) tcPr = mFactory.createTcPr();
			setTextVerticalCenter(tcPr);
		}
		tc.getContent().set(0, paragraph);
	}

	/**
	 * analysis the table content and set story data
	 * 
	 * @param columnElement - the column element, like name, notes...
	 * @param paragraph - p that you want to add to table
	 */
	private void setStoryDataToTableColumn(JAXBElement<?> columnElement, P paragraph) {
		setStoryDataToTableColumn(columnElement, paragraph, false);
	}

	/**
	 * add task info below its story
	 * 
	 * @param mainDoc - the Doc content
	 * @param taskJSONArray - a JSON array contain all tasks in the same story
	 * @return table
	 * @throws JAXBException
	 * @throws JSONException 
	 */
	private Tbl addTaskInfo(MainDocumentPart mainDoc, JSONArray taskJSONArray) throws JAXBException, JSONException {
		if (taskJSONArray.length() == 0) {
			return null;
		}
		int writableWidthTwips = wordMLPackage.getDocumentModel().getSections().get(0).getPageDimensions().getWritableWidthTwips();
		int rows = taskJSONArray.length();
		int cols = title.length;
		int cellWidthTwips = new Double(Math.floor((writableWidthTwips / cols))).intValue();
		Tbl table = mFactory.createTbl();
		TblGrid tblGrid = mFactory.createTblGrid();

		// 預設table樣式
		String strTblPr = "<w:tblPr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:tblStyle w:val=\"TableGrid\"/><w:tblW w:w=\"0\" w:type=\"auto\"/><w:tblLook w:val=\"04A0\"/></w:tblPr>";
		TblPr tblPr = (TblPr) XmlUtils.unmarshalString(strTblPr);
		table.setTblPr(tblPr);
		table.setTblGrid(tblGrid);
		for (int i = 0; i < cols; i++) {
			TblGridCol gridCol = mFactory.createTblGridCol();
			gridCol.setW(BigInteger.valueOf(cellWidthTwips));
			tblGrid.getGridCol().add(gridCol);
		}
		setTaskInfoToDocx(mainDoc, table, taskJSONArray, rows, cols, cellWidthTwips);
		return table;
	}

	/**
	 * set the task info to table of Doc
	 * 
	 * @param mainDoc - the Doc content
	 * @param table - the table of Doc
	 * @param taskJSONArray - a JSON array contain all tasks in the same story
	 * @param rows - table rows
	 * @param cols - table columns
	 * @param cellWidthTwips - table cell width
	 * @throws JSONException 
	 */
	private void setTaskInfoToDocx(MainDocumentPart mainDoc, Tbl table, JSONArray taskJSONArray, int rows, int cols, int cellWidthTwips) throws JSONException {
		for (int row = -1; row < rows; row++) {
			Tr tr = mFactory.createTr();
			table.getContent().add(tr);
			for (int col = 0; col < cols; col++) {
				Tc tc = mFactory.createTc();
				TcPr tcPr = mFactory.createTcPr();
				if (row == -1) {		// title row
					tc.getContent().add(mainDoc.createStyledParagraphOfText("Table_Cell_Center", title[col]));
				} else if (col == 0) {	// task Id
					tc.getContent().add(mainDoc.createStyledParagraphOfText("Table_Cell_Center", taskJSONArray.getJSONObject(row).getString("serial_id")));
				} else if (col == 1) {	// task Name
					tc.getContent().add(mainDoc.createStyledParagraphOfText("Table_Cell", taskJSONArray.getJSONObject(row).getString("name")));
				} else if (col == 2) {	// task Est.
					tc.getContent().add(mainDoc.createStyledParagraphOfText("Table_Cell_Center", taskJSONArray.getJSONObject(row).getString("estimate")));
				} else if (col == 3) {	// task Handler
					JSONObject handlerJSON = taskJSONArray.getJSONObject(row).getJSONObject("handler");
					String handlerUsername = "";
					try {
						handlerUsername = handlerJSON.getString("username");
					} catch (JSONException e) {
						e.printStackTrace();
					}
					tc.getContent().add(mainDoc.createStyledParagraphOfText("Table_Cell", handlerUsername));
				} else if (col == 4) {	// task Partners
					String partnerUsernames = "";
					JSONArray partnerJSONArray = taskJSONArray.getJSONObject(row).getJSONArray("partners");
					for (int i = 0; i < partnerJSONArray.length(); i++) {
						if (!partnerUsernames.isEmpty()) partnerUsernames = partnerUsernames + ", ";
						JSONObject partnerJSON = partnerJSONArray.getJSONObject(i);
						partnerUsernames = partnerUsernames + partnerJSON.getString("username");
					}
					tc.getContent().add(mainDoc.createStyledParagraphOfText("Table_Cell", partnerUsernames));
				} else if (col == 5) {	// task Notes
					tc.getContent().add(mainDoc.createStyledParagraphOfText("Table_Cell", taskJSONArray.getJSONObject(row).getString("notes")));
				}
				setTextVerticalCenter(tcPr);
				setTableCellWidth(tcPr, cellWidthTwips);
				tc.setTcPr(tcPr);
				tr.getContent().add(tc);
			}
		}
	}

	/**
	 * set table cell width
	 * 
	 * @param tcPr - table property
	 * @param cellWidthTwips - cell width
	 */
	private void setTableCellWidth(TcPr tcPr, int cellWidthTwips) {
		if (tcPr == null) tcPr = mFactory.createTcPr();
		TblWidth cellWidth = mFactory.createTblWidth();
		cellWidth.setType(TblWidth.TYPE_AUTO);
		cellWidth.setW(BigInteger.valueOf(cellWidthTwips));
		tcPr.setTcW(cellWidth);
	}

	/**
	 * Adds a page break to the document.
	 * 
	 * @param mainDoc - the Doc content
	 */
	private void addPageBreak(MainDocumentPart mainDoc) {
		Br breakObj = new Br();
		breakObj.setType(STBrType.PAGE);
		P paragraph = mFactory.createP();
		paragraph.getContent().add(breakObj);
		mainDoc.getJaxbElement().getBody().getContent().add(paragraph);
	}

	/**
	 * The document style setting
	 */
	public void alterStyleSheet() {
		StyleDefinitionsPart styleDefinitionsPart = wordMLPackage.getMainDocumentPart().getStyleDefinitionsPart();
		Styles styles = styleDefinitionsPart.getJaxbElement();
		List<Style> stylesList = styles.getStyle();
		addOwnStyle(stylesList);
	}

	/**
	 * add custom style
	 * 
	 * @param stylesList - the list of default styles
	 */
	private void addOwnStyle(List<Style> stylesList) {
		Style s = new Style();
		s.setStyleId("Bold");
		stylesList.add(s);
		s = new Style();
		s.setStyleId("Bold_Number");
		stylesList.add(s);
		s = new Style();
		s.setStyleId("Table_Cell_Center");
		stylesList.add(s);
		s = new Style();
		s.setStyleId("Table_Cell");
		stylesList.add(s);

		// set custom style property
		setStyleProperty(stylesList);
	}

	/**
	 * set custom style property
	 * 
	 * @param stylesList - the list of default styles
	 */
	private void setStyleProperty(List<Style> stylesList) {
		RPr rpr;
		for (Style style : stylesList) {
			rpr = style.getRPr();
			if (rpr == null) rpr = new RPr();
			if (style.getStyleId().equals("Bold")) {
				setBoldStyle(rpr, true);
			} else if (style.getStyleId().equals("Bold_Number")) {
				setFontSize(rpr, 56);
				setBoldStyle(rpr, true);
				setTableCellCenter(style);
			} else if (style.getStyleId().equals("Table_Cell_Center")) {
				setTableCellCenter(style);
			} else if (style.getStyleId().equals("Table_Cell")) {
				PPr ppr = style.getPPr();
				if (ppr == null) ppr = mFactory.createPPr();
				setTextAfterSpace(ppr);
			}
			setFontStyle(rpr);
			style.setRPr(rpr);
		}
	}

	/**
	 * table cell to horizontal center
	 * 
	 * @param style
	 */
	private void setTableCellCenter(Style style) {
		PPr ppr = style.getPPr();
		if (ppr == null) ppr = mFactory.createPPr();
		setTextCenter(ppr);
		setTextAfterSpace(ppr);
		style.setPPr(ppr);
	}

	/**
	 * 設定該段落下面的空白
	 * 
	 * @param ppr - the paragraph property
	 */
	private void setTextAfterSpace(PPr ppr) {
		Spacing spacing = Context.getWmlObjectFactory().createPPrBaseSpacing();
		spacing.setAfterAutospacing(true);
		ppr.setSpacing(spacing);
	}

	/**
	 * set text in center for horizontal
	 * 
	 * @param ppr - the paragraph property
	 */
	private void setTextCenter(PPr ppr) {
		Jc jc = new Jc();
		jc.setVal(JcEnumeration.CENTER);
		ppr.setJc(jc);
	}

	/**
	 * set text in center for vertical
	 * 
	 * @param ppr - the paragraph property
	 */
	private void setTextVerticalCenter(TcPr tcPr) {
		CTVerticalJc value = new CTVerticalJc();
		value = new CTVerticalJc();
		value.setVal(STVerticalJc.CENTER);
		tcPr.setVAlign(value);
	}

	/**
	 * 設定字型
	 * 
	 * @param ppr - the run property
	 */
	private static void setFontStyle(RPr runProperties) {
		RFonts runFont = new RFonts();
		runFont.setAscii("Times New Roman");	// for English and Number
		runFont.setEastAsia("標楷體");			// for Chinese
		runProperties.setRFonts(runFont);
	}

	/**
	 * 設定字的大小
	 * 
	 * @param ppr - the run property
	 */
	private static void setFontSize(RPr runProperties, int fontSize) {
		HpsMeasure size = new HpsMeasure();
		size.setVal(BigInteger.valueOf(fontSize));
		runProperties.setSz(size);
		runProperties.setSzCs(size);
	}

	/**
	 * 設定粗體
	 * 
	 * @param ppr - the run property
	 */
	private static void setBoldStyle(RPr runProperties, boolean bool) {
		BooleanDefaultTrue b = new BooleanDefaultTrue();
		b.setVal(bool);
		runProperties.setB(b);
	}
}

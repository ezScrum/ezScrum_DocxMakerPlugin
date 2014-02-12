package plugin.docxMaker;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import org.apache.commons.io.IOUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
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

import plugin.docxMaker.dataModel.ReleasePlanObject;
import plugin.docxMaker.dataModel.SprintPlanObject;
import plugin.docxMaker.dataModel.StoryObject;
import plugin.docxMaker.dataModel.TagObject;
import plugin.docxMaker.dataModel.TaskObject;

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
	 * @param releases - release info object
	 * @param sprints - release info objects
	 * @param storyMap - use sprint id as key to get its story objects
	 * @param taskMap - use story id as key to get its task objects
	 * @param tatolStoryPoints - sprint stories count
	 * @return
	 */
	public File getReleasePlanDocx(ReleasePlanObject releases, List<SprintPlanObject> sprints, HashMap<String, List<StoryObject>> storyMap, LinkedHashMap<String, List<TaskObject>> taskMap,
	        HashMap<String, Float> tatolStoryPoints) {
		try {
			wordMLPackage = WordprocessingMLPackage.createPackage();	// Create the package
			MainDocumentPart mainDoc = wordMLPackage.getMainDocumentPart();
			mFactory = Context.getWmlObjectFactory();
			alterStyleSheet();					// change the style of this docx
			addReleaseInfo(mainDoc, releases);	// the title of docx
			SprintPlanObject sprint;
			for (int i = 0, sprintSize = sprints.size(); i < sprintSize; i++) {
				sprint = sprints.get(i);
				addSprintInfo(mainDoc, sprints.get(i), tatolStoryPoints);
				addStoryInfo(mainDoc, storyMap.get(sprint.getId()), taskMap);
				if (i != sprintSize) addPageBreak(mainDoc);	// add new page
			}
			String filePath = releases.getName() + "_ReleasePlan.docx";
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
	 * @param releases - release info object
	 */
	private void addReleaseInfo(MainDocumentPart mainDoc, ReleasePlanObject releases) {
		mainDoc.addStyledParagraphOfText("Title", "Release Plan #" + releases.getId() + "： " + releases.getName());
		mainDoc.addStyledParagraphOfText("Subtitle", "Start Date： " + releases.getStartDate());
		mainDoc.addStyledParagraphOfText("Subtitle", "End Date： " + releases.getEndDate());
		mainDoc.addStyledParagraphOfText("Subtitle", "Description： " + releases.getDescription());
	}

	/**
	 * add sprint info above its stories
	 * 
	 * @param mainDoc - the Doc content
	 * @param sprint - sprint info object
	 * @param tatolStoryPoints - sprint stories count
	 */
	private void addSprintInfo(MainDocumentPart mainDoc, SprintPlanObject sprint, HashMap<String, Float> tatolStoryPoints) {
		mainDoc.addStyledParagraphOfText("Heading1", "Sprint #" + sprint.getId() + "： " + sprint.getGoal());
		mainDoc.addStyledParagraphOfText("Subtitle", "Start Date： " + sprint.getStartDate());
		mainDoc.addStyledParagraphOfText("Subtitle", "End Date： " + sprint.getEndDate());
		mainDoc.addStyledParagraphOfText("Subtitle", "Total Story Points： " + tatolStoryPoints.get(sprint.getId()));
	}

	/**
	 * add story info below its sprint
	 * 
	 * @param mainDoc - the Doc content
	 * @param storyList - story objects
	 * @param taskMap - task objects
	 * @throws JAXBException
	 */
	private void addStoryInfo(MainDocumentPart mainDoc, List<StoryObject> storyList, LinkedHashMap<String, List<TaskObject>> taskMap) throws JAXBException {
		try {
			String tblXML = IOUtils.toString(new FileReader("StoryCard.xml"));	// 使用預設的Story Card 格式
			StoryObject story = null;
			int storySize = storyList.size();
			for (int i = 0; i < storySize; i++) {
				story = storyList.get(i);
				mainDoc.addObject(getStoryTable(mainDoc, story, tblXML));
				mainDoc.addParagraphOfText("Task List: ");
				mainDoc.addObject(addTaskInfo(mainDoc, taskMap.get(story.id)));
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
	 * @param story - the story object for table
	 * @param tblXML - the table xml for story card
	 * @return the story table
	 * @throws JAXBException
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	private Tbl getStoryTable(MainDocumentPart mainDoc, StoryObject story, String tblXML) throws JAXBException, FileNotFoundException, IOException {
		Tbl table = (Tbl) XmlUtils.unmarshalString(tblXML);	// row 2, 4, 5, 7, 9  是 table 的空白格
		Tr row = null;
		row = (Tr) table.getContent().get(0);	// get ID
		JAXBElement<?> element = (JAXBElement<?>) row.getContent().get(0);
		Tc tc = (Tc) element.getValue();
		tc.getContent().set(0, mainDoc.createParagraphOfText("Sprint Backlog Item #" + story.id));							// add ID
		row = (Tr) table.getContent().get(1);	// get Name
		element = (JAXBElement<?>) row.getContent().get(0);
		tc = (Tc) element.getValue();
		tc.getContent().set(0, mainDoc.createParagraphOfText(story.name));													// add Name
		setStoryDataToTableColumn((JAXBElement<?>) row.getContent().get(1), mainDoc.createStyledParagraphOfText("Bold_Number", story.importance), true);	// add Importance
		row = (Tr) table.getContent().get(3);	// get Notes
		setStoryDataToTableColumn((JAXBElement<?>) row.getContent().get(0), mainDoc.createParagraphOfText(story.notes));	// add Notes
		setStoryDataToTableColumn((JAXBElement<?>) row.getContent().get(1), mainDoc.createStyledParagraphOfText("Bold_Number", story.estimation), true);	// add Estimate
		row = (Tr) table.getContent().get(6);	// get Tags
		String tags = "";
		for (TagObject tag : story.tagList) {
			if (!tags.isEmpty()) tags = tags + ", ";
			tags = tags + tag.getTagName();
		}
		setStoryDataToTableColumn((JAXBElement<?>) row.getContent().get(0), mainDoc.createParagraphOfText(tags));			// add Tags
		row = (Tr) table.getContent().get(8);	// get How to Demo
		setStoryDataToTableColumn((JAXBElement<?>) row.getContent().get(0), mainDoc.createParagraphOfText(story.howToDemo));// add How to Demo
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
	 * @param taskList - task info objects
	 * @return table
	 * @throws JAXBException
	 */
	private Tbl addTaskInfo(MainDocumentPart mainDoc, List<TaskObject> taskList) throws JAXBException {
		if (taskList == null) return null;
		int writableWidthTwips = wordMLPackage.getDocumentModel().getSections().get(0).getPageDimensions().getWritableWidthTwips();
		int rows = taskList.size();
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
		setTaskInfoToDocx(mainDoc, table, taskList, rows, cols, cellWidthTwips);
		return table;
	}

	/**
	 * set the task info to table of Doc
	 * 
	 * @param mainDoc - the Doc content
	 * @param table - the table of Doc
	 * @param taskList - task objects
	 * @param rows - table rows
	 * @param cols - table columns
	 * @param cellWidthTwips - table cell width
	 */
	private void setTaskInfoToDocx(MainDocumentPart mainDoc, Tbl table, List<TaskObject> taskList, int rows, int cols, int cellWidthTwips) {
		for (int row = -1; row < rows; row++) {
			Tr tr = mFactory.createTr();
			table.getContent().add(tr);
			for (int col = 0; col < cols; col++) {
				Tc tc = mFactory.createTc();
				TcPr tcPr = mFactory.createTcPr();
				if (row == -1) {		// title row
					tc.getContent().add(mainDoc.createStyledParagraphOfText("Table_Cell_Center", title[col]));
				} else if (col == 0) {	// task Id
					tc.getContent().add(mainDoc.createStyledParagraphOfText("Table_Cell_Center", taskList.get(row).id));
				} else if (col == 1) {	// task Name
					tc.getContent().add(mainDoc.createStyledParagraphOfText("Table_Cell", taskList.get(row).name));
				} else if (col == 2) {	// task Est.
					tc.getContent().add(mainDoc.createStyledParagraphOfText("Table_Cell_Center", taskList.get(row).estimation));
				} else if (col == 3) {	// task Handler
					tc.getContent().add(mainDoc.createStyledParagraphOfText("Table_Cell", taskList.get(row).handler));
				} else if (col == 4) {	// task Partners
					tc.getContent().add(mainDoc.createStyledParagraphOfText("Table_Cell", taskList.get(row).partners));
				} else if (col == 5) {	// task Notes
					tc.getContent().add(mainDoc.createStyledParagraphOfText("Table_Cell", taskList.get(row).notes));
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

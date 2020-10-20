package com.dev4k.birt.excelcomparator.comparator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.ListIterator;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.model.api.CellHandle;
import org.eclipse.birt.report.model.api.ElementFactory;
import org.eclipse.birt.report.model.api.GridHandle;
import org.eclipse.birt.report.model.api.ReportDesignHandle;
import org.eclipse.birt.report.model.api.RowOperationParameters;
import org.eclipse.birt.report.model.api.TextItemHandle;
import org.eclipse.birt.report.model.api.activity.SemanticException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;

import com.dev4k.birt.excelcomparator.designer.ReportDesigner;

public class ExcelComparator {

	@Autowired
	private ReportDesigner reportDesigner;

	@Value("${excel1.path}")
	private String sourcePath1;
	@Value("${excel2.path}")
	private String sourcePath2;

	private ReportDesignHandle design;
	private ElementFactory factory;
	int summaryGridRowCount = 1;

	public ReportDesignHandle compareExcel()
			throws BirtException, EncryptedDocumentException, InvalidFormatException, IOException {

		design = reportDesigner.buildReport();
		factory = design.getElementFactory();

		int[] fileNotFound = { 0, 0 };

		// begin excel comparison
		
		// load souce file
		FileInputStream file1 = null;
		try {
			file1 = new FileInputStream(new File(sourcePath1));
		} catch (FileNotFoundException e) {
			fileNotFound[0] = 1;
		}
		FileInputStream file2 = null;
		try {
			file2 = new FileInputStream(new File(sourcePath2));
		} catch (FileNotFoundException e) {
			fileNotFound[1] = 1;
		}

		if (fileNotFound[0] == 1 || fileNotFound[1] == 1) {
			fileNotFoundMismatch(fileNotFound);
			return design;
		}

		// create workbook objects using source files
		Workbook workbook1 = WorkbookFactory.create(file1);
		Workbook workbook2 = WorkbookFactory.create(file2);

		// compare excels by their sheet name
		// save matched sheet names to common list
		List<String> sheetNames1 = new ArrayList<>();
		List<String> sheetNames2 = new ArrayList<>();
		List<String> matchedSheetNames = new ArrayList<>();

		for (int i = 0; i < workbook1.getNumberOfSheets(); i++) {
			sheetNames1.add(workbook1.getSheetName(i));
		}
		for (int i = 0; i < workbook2.getNumberOfSheets(); i++) {
			sheetNames2.add(workbook2.getSheetName(i));
		}

		ListIterator<String> sheetNames1Iterator = sheetNames1.listIterator();
		while (sheetNames1Iterator.hasNext()) {
			String currentSheet = sheetNames1Iterator.next();

			if (sheetNames2.indexOf(currentSheet) != -1) {
				matchedSheetNames.add(currentSheet);
			} else {
				/**
				 * sheet not found mismatch
				 */
				sheetNames1Iterator.remove();
				sheetNotFoundException(currentSheet, 1);
			}

		}

		ListIterator<String> sheetNames2Iterator = sheetNames2.listIterator();
		while (sheetNames2Iterator.hasNext()) {
			String currentSheet = sheetNames2Iterator.next();

			if (sheetNames1.indexOf(currentSheet) != -1) {
				matchedSheetNames.add(currentSheet);
			} else {
				/**
				 * sheet not found mismatch
				 */
				sheetNames2Iterator.remove();
				sheetNotFoundException(currentSheet, 2);
			}

		}

		// further matching the sheets which are in both workbook
		for (int i = 0; i < matchedSheetNames.size(); i++) {
			String sheetName = matchedSheetNames.get(i);
			Sheet sheet1 = workbook1.getSheet(sheetName);
			Sheet sheet2 = workbook2.getSheet(sheetName);

			compare(sheet1, sheet2);
		}

		if (summaryGridRowCount == 1) {
			noDiscrepancyFound();
		}

		return design;
	}

	private void compare(Sheet sheet1, Sheet sheet2) throws SemanticException {

		DataFormatter formatter = new DataFormatter();

		// check if number of rows are equal
		if (sheet1.getPhysicalNumberOfRows() != sheet2.getPhysicalNumberOfRows()) {
			rowCountMismatch(sheet1.getSheetName(), sheet1.getPhysicalNumberOfRows(), sheet2.getPhysicalNumberOfRows());
		}

		// create data structure to store header metadata
		Map<String, List<String>> columns1 = new HashMap<>();
		Row row = sheet1.getRow(0);

		int firstCellIndex1 = row.getFirstCellNum();
		int lastCellIndex1 = row.getLastCellNum();
		int columnLength1 = lastCellIndex1 - firstCellIndex1;

		for (int i = firstCellIndex1; i < columnLength1; i++) {
			Row keyRow = sheet1.getRow(0);
			String columnKey = formatter.formatCellValue(keyRow.getCell(i));
			List<String> columnValues = new ArrayList<>();
			for (int j = firstCellIndex1 + 1; j < sheet1.getPhysicalNumberOfRows(); j++) {
				Row valueRow = sheet1.getRow(j);
				columnValues.add(formatter.formatCellValue(valueRow.getCell(i)));
			}
			columns1.put(columnKey, columnValues);
		}

		Map<String, List<String>> columns2 = new HashMap<>();
		row = sheet2.getRow(0);

		int firstCellIndex2 = row.getFirstCellNum();
		int lastCellIndex2 = row.getLastCellNum();
		int columnLength2 = lastCellIndex2 - firstCellIndex2;

		for (int i = firstCellIndex2; i < columnLength2; i++) {
			Row keyRow = sheet2.getRow(0);
			String columnKey = formatter.formatCellValue(keyRow.getCell(i));
			List<String> columnValues = new ArrayList<>();
			for (int j = firstCellIndex2 + 1; j < sheet2.getPhysicalNumberOfRows(); j++) {
				Row valueRow = sheet2.getRow(j);
				columnValues.add(formatter.formatCellValue(valueRow.getCell(i)));
			}
			columns2.put(columnKey, columnValues);
		}

		// compare columns in each sheets
		Map<String, List<String>> matchedColumns1 = new HashMap<>();

		Iterator<Entry<String, List<String>>> columns1Iterator = columns1.entrySet().iterator();
		while (columns1Iterator.hasNext()) {
			Entry<String, List<String>> mapElement = columns1Iterator.next();
			String key = mapElement.getKey();

			if (columns2.containsKey(key)) {
				matchedColumns1.put(key, mapElement.getValue());
			} else {
				missingColumn(sheet1.getSheetName(), key, 1);
			}
		}

		Map<String, List<String>> matchedColumns2 = new HashMap<>();

		Iterator<Entry<String, List<String>>> columns2Iterator = columns2.entrySet().iterator();
		while (columns2Iterator.hasNext()) {
			Entry<String, List<String>> mapElement = columns2Iterator.next();
			String key = mapElement.getKey();

			if (columns2.containsKey(key)) {
				matchedColumns2.put(key, mapElement.getValue());
			} else {
				missingColumn(sheet1.getSheetName(), key, 2);
			}
		}

		// compare data of each cell for matched/same column headers
		Iterator<Entry<String, List<String>>> matchedColumnsIterator = matchedColumns2.entrySet().iterator();
		while (matchedColumnsIterator.hasNext()) {
			Map.Entry<String, List<String>> next = (Map.Entry<String, List<String>>) matchedColumnsIterator.next();

			String key = next.getKey();

			// since key is present in both map
			List<String> values1 = matchedColumns1.get(key);
			List<String> values2 = matchedColumns2.get(key);

			for (int i = 0; i < values1.size(); i++) {
				if (!values1.get(i).equals(values2.get(i))) {
					valueMismatch(sheet1.getSheetName(), key, (i + 1), values1.get(i), values2.get(i));
				}
			}

		}

	}

	private void noDiscrepancyFound() throws SemanticException {

		GridHandle paramGrid = (GridHandle) design.findElement("SummaryGrid");
		paramGrid.drop();
		TextItemHandle text = factory.newTextItem(null);
		text.setProperty("contentType", "HTML");
		text.setContent("<b>No Discrepancy found between the Excel Sources.<b>");

		design.getBody().add(text);
	}

	private void valueMismatch(String sheetName, String column, int i, String value1, String value2)
			throws SemanticException {
		summaryGridRowCount++;

		GridHandle grid = (GridHandle) design.findElement("SummaryGrid");
		RowOperationParameters rowParam = new RowOperationParameters(1, 0, summaryGridRowCount - 1);
		grid.insertRow(rowParam);

		CellHandle cell = grid.getCell(summaryGridRowCount, 1);
		TextItemHandle sheetNameText = factory.newTextItem(null);
		sheetNameText.setContent(sheetName);
		cell.getContent().add(sheetNameText);

		cell = grid.getCell(summaryGridRowCount, 2);
		TextItemHandle mismatchType = factory.newTextItem(null);
		mismatchType.setContent("Value Mismatch");
		cell.getContent().add(mismatchType);

		cell = grid.getCell(summaryGridRowCount, 3);
		TextItemHandle columnName = factory.newTextItem(null);
		columnName.setContent(column);
		cell.getContent().add(columnName);

		cell = grid.getCell(summaryGridRowCount, 4);
		TextItemHandle rowNumber = factory.newTextItem(null);
		rowNumber.setContent(Integer.toString(i));
		cell.getContent().add(rowNumber);

		cell = grid.getCell(summaryGridRowCount, 5);
		TextItemHandle source1Text = factory.newTextItem(null);
		source1Text.setContent(value1);
		cell.getContent().add(source1Text);

		cell = grid.getCell(summaryGridRowCount, 6);
		TextItemHandle source2Text = factory.newTextItem(null);
		source2Text.setContent(value2);
		cell.getContent().add(source2Text);

	}

	private void missingColumn(String sheetName, String column, int posn) throws SemanticException {
		summaryGridRowCount++;

		GridHandle grid = (GridHandle) design.findElement("SummaryGrid");
		RowOperationParameters rowParam = new RowOperationParameters(1, 0, summaryGridRowCount - 1);
		grid.insertRow(rowParam);

		CellHandle cell = grid.getCell(summaryGridRowCount, 1);
		TextItemHandle sheetNameText = factory.newTextItem(null);
		sheetNameText.setContent(sheetName);
		cell.getContent().add(sheetNameText);

		cell = grid.getCell(summaryGridRowCount, 2);
		TextItemHandle mismatchType = factory.newTextItem(null);
		mismatchType.setContent("Missing Column");
		cell.getContent().add(mismatchType);

		cell = grid.getCell(summaryGridRowCount, posn + 4);

		TextItemHandle exception = factory.newTextItem(null);
		exception.setContent(column);
		cell.getContent().add(exception);

	}

	private void rowCountMismatch(String sheetName, int physicalNumberOfRows1, int physicalNumberOfRows2)
			throws SemanticException {
		summaryGridRowCount++;

		GridHandle grid = (GridHandle) design.findElement("SummaryGrid");
		RowOperationParameters rowParam = new RowOperationParameters(1, 0, summaryGridRowCount - 1);
		grid.insertRow(rowParam);

		CellHandle cell = grid.getCell(summaryGridRowCount, 1);
		TextItemHandle sheetNameText = factory.newTextItem(null);
		sheetNameText.setContent(sheetName);
		cell.getContent().add(sheetNameText);

		cell = grid.getCell(summaryGridRowCount, 2);
		TextItemHandle mismatchType = factory.newTextItem(null);
		mismatchType.setContent("Row Count Mismatch");
		cell.getContent().add(mismatchType);

		cell = grid.getCell(summaryGridRowCount, 5);

		TextItemHandle exception1 = factory.newTextItem(null);
		exception1.setContent(Integer.toString(physicalNumberOfRows1));
		cell.getContent().add(exception1);

		cell = grid.getCell(summaryGridRowCount, 6);

		TextItemHandle exception2 = factory.newTextItem(null);
		exception1.setContent(Integer.toString(physicalNumberOfRows2));
		cell.getContent().add(exception2);

	}

	private void fileNotFoundMismatch(int[] fileNotFound) throws SemanticException {

		GridHandle paramGrid = (GridHandle) design.findElement("SummaryGrid");
		paramGrid.drop();
		TextItemHandle text = factory.newTextItem(null);
		text.setProperty("contentType", "HTML");

		if (fileNotFound[0] == 1 && fileNotFound[1] == 1) {
			text.setContent("Source File 1 and 2 was not found on the path specified.");
		} else if (fileNotFound[0] == 1) {
			text.setContent("Source File 1 was not found on the path specified.");
		} else if (fileNotFound[1] == 1) {
			text.setContent("Source File 2 was not found on the path specified.");
		}

		design.getBody().add(text);

	}

	private void sheetNotFoundException(String currentSheet, int posn) throws SemanticException {
		summaryGridRowCount++;

		GridHandle grid = (GridHandle) design.findElement("SummaryGrid");
		RowOperationParameters rowParam = new RowOperationParameters(1, 0, summaryGridRowCount - 1);
		grid.insertRow(rowParam);

		CellHandle cell = grid.getCell(summaryGridRowCount, 1);
		TextItemHandle sheetName = factory.newTextItem(null);
		sheetName.setContent(currentSheet);
		cell.getContent().add(sheetName);

		cell = grid.getCell(summaryGridRowCount, 2);
		TextItemHandle mismatchType = factory.newTextItem(null);
		mismatchType.setContent("Missing Sheet");
		cell.getContent().add(mismatchType);

		cell = grid.getCell(summaryGridRowCount, posn + 4);

		TextItemHandle exception = factory.newTextItem(null);
		exception.setContent(currentSheet);
		cell.getContent().add(exception);

	}

}

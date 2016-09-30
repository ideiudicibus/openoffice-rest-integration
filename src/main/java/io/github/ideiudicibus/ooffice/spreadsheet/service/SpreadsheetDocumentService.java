package io.github.ideiudicibus.ooffice.spreadsheet.service;

import java.io.File;

import org.artofsolving.jodconverter.office.ExternalOfficeManager;
import org.artofsolving.jodconverter.office.OfficeContext;
import org.artofsolving.jodconverter.office.OfficeException;
import org.artofsolving.jodconverter.office.OfficeUtils;
import org.springframework.stereotype.Service;

import com.sun.star.beans.PropertyValue;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XModel;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.sheet.CellFlags;
import com.sun.star.sheet.XCellRangeData;
import com.sun.star.sheet.XCellRangesQuery;
import com.sun.star.sheet.XSheetCellRanges;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XSpreadsheets;
import com.sun.star.table.CellRangeAddress;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;
import com.sun.star.text.XText;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.XCloseable;
import com.sun.star.util.XModifiable;

@Service
public class SpreadsheetDocumentService {

	public SpreadsheetDocumentService() throws BootstrapException, Exception {

		super();
	}

	public void loadDocument() {

		String unoUrl = "8100";
		manager = new ExternalOfficeManager(unoUrl, true);
		manager.start();
		OfficeContext context = manager.getConnection();

		XComponentLoader loader = (XComponentLoader) UnoRuntime.queryInterface(XComponentLoader.class,
				context.getService(OfficeUtils.SERVICE_DESKTOP));

		assert loader != null : "desktop object is null";
		try {
			PropertyValue[] arguments = new PropertyValue[] { OfficeUtils.property("Hidden", true) };
			// XComponent document =
			// loader.loadComponentFromURL("private:factory/swriter", "_blank",
			// 0, arguments);
			File file = new File(filename);
			xComponent = loader.loadComponentFromURL(OfficeUtils.toUrl(file), "_blank", 0, arguments);
			xSpreadsheetDoc = (XSpreadsheetDocument) UnoRuntime.queryInterface(XSpreadsheetDocument.class, xComponent);

		} catch (Exception exception) {
			throw new OfficeException("failed to create document", exception);
		}
	}

	public Object[][] getSheetContentData(XSpreadsheet xSheet) throws Exception {
		XCellRangesQuery xRangesQuery = (XCellRangesQuery) UnoRuntime.queryInterface(XCellRangesQuery.class, xSheet);
		XSheetCellRanges xCellRanges = xRangesQuery.queryContentCells(
				(short) (CellFlags.DATETIME | CellFlags.FORMULA | CellFlags.STRING | CellFlags.VALUE));
		CellRangeAddress[] addrs = xCellRanges.getRangeAddresses();
		int startR = 0, startC = 0, endR = 0, endC = 0;
		for (CellRangeAddress addr : addrs) {
			if (addr.StartRow < startR)
				startR = addr.StartRow;
			if (addr.StartColumn < startC)
				startC = addr.StartColumn;
			if (addr.EndRow > endR)
				endR = addr.EndRow;
			if (addr.EndColumn > endC)
				endC = addr.EndColumn;
		}
		XCellRange xRange = xSheet.getCellRangeByPosition(startC, startR, endC, endR);
		return ((XCellRangeData) UnoRuntime.queryInterface(XCellRangeData.class, xRange)).getDataArray();
	}

	public void setText(XSpreadsheet xSheet, int nRow, int nColumn, String text) throws Exception {
		XCell cell = xSheet.getCellByPosition(nColumn, nRow);
		XText xText = (XText) UnoRuntime.queryInterface(XText.class, cell);
		xText.setString(text);
	}

	public void save() throws Exception {
		XModel xModel = (XModel) UnoRuntime.queryInterface(XModel.class, xSpreadsheetDoc);
		XModifiable xModified = (XModifiable) UnoRuntime.queryInterface(XModifiable.class, xModel);
		if (xModified.isModified()) {
			XStorable xStore = (XStorable) UnoRuntime.queryInterface(XStorable.class, xModel);
			xStore.store();
		}
	}

	public void close() throws Exception {
		XModel xModel = (XModel) UnoRuntime.queryInterface(XModel.class, xComponent);
		XCloseable xCloseable = (XCloseable) UnoRuntime.queryInterface(XCloseable.class, xModel);
		xCloseable.close(true);
	}

	public XSpreadsheet getSpreadsheet(int nIndex) throws Exception {
		XSpreadsheets xSheets = xSpreadsheetDoc.getSheets();
		return (XSpreadsheet) UnoRuntime.queryInterface(XSpreadsheet.class,
				xSheets.getByName(xSheets.getElementNames()[nIndex]));
	}

	public XComponent getxComponent() {
		return xComponent;
	}

	public XSpreadsheetDocument getxSpreadsheetDoc() {
		return xSpreadsheetDoc;
	}

	public ExternalOfficeManager getManager() {
		return manager;
	}

	private ExternalOfficeManager manager;
	private XComponent xComponent;
	private XSpreadsheetDocument xSpreadsheetDoc;

	private String filename;

	public void setFilename(String filename) {
		this.filename = filename;
	}

	public void loadDocument(SpreadsheetXComponentLoader xComponentLoader) {
		

		try {
			PropertyValue[] arguments = new PropertyValue[] { OfficeUtils.property("Hidden", true) };
			
			File file = new File(filename);
			xComponent = (xComponentLoader.getLoader()).loadComponentFromURL(OfficeUtils.toUrl(file), "_blank", 0, arguments);
			xSpreadsheetDoc = (XSpreadsheetDocument) UnoRuntime.queryInterface(XSpreadsheetDocument.class, xComponent);

		} catch (Exception exception) {
			throw new OfficeException("failed to create document", exception);
		}
		
		
	}

}

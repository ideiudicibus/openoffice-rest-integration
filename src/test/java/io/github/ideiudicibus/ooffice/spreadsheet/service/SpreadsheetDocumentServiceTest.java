package io.github.ideiudicibus.ooffice.spreadsheet.service;

import static org.junit.Assert.*;

import java.io.File;

import org.artofsolving.jodconverter.office.ExternalOfficeManager;
import org.artofsolving.jodconverter.office.OfficeContext;
import org.artofsolving.jodconverter.office.OfficeUtils;
import org.junit.Before;
import org.junit.Test;
import org.springframework.beans.factory.annotation.Autowired;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.sun.star.beans.PropertyValue;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.io.IOException;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.XComponent;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;

public class SpreadsheetDocumentServiceTest {
	@Autowired
	SpreadsheetDocumentService sds;

	@Before
	public void before() {

	}

	@Test
	public void testLoadDocument() throws JsonProcessingException, BootstrapException, Exception {
		String filename = "C:/Users/ignazio/projects/pcc-etl/export_lista_fatture-ares-118-test.xlsx";
		sds =new SpreadsheetDocumentService();
		sds.setFilename(filename);
		XSpreadsheet xSheet = sds.getSpreadsheet(0);
		Object[][] retValue=sds.getSheetContentData(xSheet);
		assertNotNull(retValue);
		ObjectMapper mapper = new ObjectMapper();
	    System.out.println(mapper.writeValueAsString(retValue));
	     xSheet = sds.getSpreadsheet(0);
		 retValue=sds.getSheetContentData(xSheet);
		System.out.println(mapper.writeValueAsString(retValue));

	}

}

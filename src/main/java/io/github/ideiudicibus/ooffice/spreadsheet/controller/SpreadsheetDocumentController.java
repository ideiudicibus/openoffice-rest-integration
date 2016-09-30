package io.github.ideiudicibus.ooffice.spreadsheet.controller;

import java.util.HashMap;
import java.util.Map;

import org.springframework.aop.TargetSource;
import org.springframework.aop.target.CommonsPoolTargetSource;
import org.springframework.beans.BeansException;
import org.springframework.beans.factory.BeanFactory;
import org.springframework.beans.factory.BeanFactoryAware;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.uno.Exception;

import io.github.ideiudicibus.ooffice.spreadsheet.service.SpreadsheetDocumentService;
import io.github.ideiudicibus.ooffice.spreadsheet.service.SpreadsheetXComponentLoader;



@RestController

public class SpreadsheetDocumentController implements BeanFactoryAware  {

	Map<String,String> hm=new HashMap<String,String>();
	
	public SpreadsheetDocumentController(){
		
		hm.put("1", "C:/Users/ignazio/projects/pcc-etl/export_lista_fatture-ares-118-test.xlsx");
	}
	private SpreadsheetXComponentLoader spreadsheetXComponentLoader;
	@Autowired
	SpreadsheetDocumentService spreadsheetDocSvc;
	  @RequestMapping("/spreadsheets")
	    public String greeting(@RequestParam(value="id") String spreadsheetId) throws java.lang.Exception {
		  Object[][] retValue=null;
		  
		  try {
			spreadsheetDocSvc=new SpreadsheetDocumentService();
			spreadsheetDocSvc.setFilename(hm.get(spreadsheetId));
			//spreadsheetDocSvc.loadDocument();
			spreadsheetDocSvc.loadDocument(spreadsheetXComponentLoader);
			XSpreadsheet xSheet = spreadsheetDocSvc.getSpreadsheet(0);
			retValue=spreadsheetDocSvc.getSheetContentData(xSheet);
		} catch (Exception | BootstrapException e) {
			// TODO Auto-generated catch block
			return e.toString();
		}
		  
	         ObjectMapper mapper = new ObjectMapper();
	         try {
				return mapper.writeValueAsString(retValue);
			} catch (JsonProcessingException e) {
				// TODO Auto-generated catch block
				return e.toString();
			}
	    }
	@Override
	public void setBeanFactory(BeanFactory ctx) throws BeansException {
		spreadsheetXComponentLoader =(SpreadsheetXComponentLoader)ctx.getBean("spreadsheetXComponentLoader");
		
	}
	
}

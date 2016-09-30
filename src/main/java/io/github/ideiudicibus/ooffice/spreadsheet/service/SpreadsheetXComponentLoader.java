package io.github.ideiudicibus.ooffice.spreadsheet.service;

import org.artofsolving.jodconverter.office.ExternalOfficeManager;
import org.artofsolving.jodconverter.office.OfficeContext;
import org.artofsolving.jodconverter.office.OfficeUtils;
import org.springframework.stereotype.Component;

import com.sun.star.frame.XComponentLoader;
import com.sun.star.uno.UnoRuntime;

@Component
public class SpreadsheetXComponentLoader {

	private XComponentLoader loader;
	
	public SpreadsheetXComponentLoader(){
		
		String unoUrl = "8100";
		ExternalOfficeManager manager = new ExternalOfficeManager(unoUrl, true);
		manager.start();
		OfficeContext context = manager.getConnection();

		loader = (XComponentLoader) UnoRuntime.queryInterface(XComponentLoader.class,context.getService(OfficeUtils.SERVICE_DESKTOP));

	}
	public XComponentLoader getLoader(){
		return this.loader;
	}
}

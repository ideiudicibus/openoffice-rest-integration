package io.github.ideiudicibus.ooffice.spreadsheet;



import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.Scope;

import io.github.ideiudicibus.ooffice.spreadsheet.service.SpreadsheetXComponentLoader;

@SuppressWarnings("deprecation")
@Configuration
public class ObjectPoolConfig {

	@Bean(name = "spreadsheetXComponentLoader")
	@Scope("prototype")
	public SpreadsheetXComponentLoader spreadsheetDocumentService() {
		return new SpreadsheetXComponentLoader();
		
	}

	
}
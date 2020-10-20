package com.dev4k.birt.excelcomparator.designer;

import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.core.framework.Platform;
import org.eclipse.birt.report.model.api.CellHandle;
import org.eclipse.birt.report.model.api.DesignConfig;
import org.eclipse.birt.report.model.api.ElementFactory;
import org.eclipse.birt.report.model.api.GridHandle;
import org.eclipse.birt.report.model.api.IDesignEngine;
import org.eclipse.birt.report.model.api.IDesignEngineFactory;
import org.eclipse.birt.report.model.api.ReportDesignHandle;
import org.eclipse.birt.report.model.api.SessionHandle;
import org.eclipse.birt.report.model.api.TextItemHandle;

import com.ibm.icu.util.ULocale;


public class ReportDesigner {
	
	public ReportDesignHandle buildReport() throws BirtException {
		
		final DesignConfig config = new DesignConfig();
		
		final IDesignEngine engine;
		
		try {
			Platform.startup(config);
			IDesignEngineFactory factory = (IDesignEngineFactory) Platform.createFactoryObject(IDesignEngineFactory.EXTENSION_DESIGN_ENGINE_FACTORY);
			engine = factory.createDesignEngine(config);
		}catch( Exception ex ) {
			throw ex;
		}
		
		SessionHandle session = engine.newSessionHandle(ULocale.ENGLISH);
		ReportDesignHandle design = session.createDesign();
		
		ElementFactory elementFactory = design.getElementFactory();
		
		design.setTitle("Birt Excel Comparison Report");
		
		// create report title
		TextItemHandle title = elementFactory.newTextItem("Title");
		title.setProperty("contentType", "HTML");
		title.setContent("<h1 style=\"color:#004080; font-size:150%; text-align:center;\"> Birt Excel Comparison Report </h1>");
		design.getBody().add(title);
		
		// parameter grid
		GridHandle paramGrid = elementFactory.newGridItem("ParameterGrid", 2, 2);
		TextItemHandle source1 = elementFactory.newTextItem(null);
		source1.setProperty("contentType", "HTML");
		source1.setContent("Source 1:");
		TextItemHandle source2 = elementFactory.newTextItem(null);
		source2.setProperty("contentType", "HTML");
		source2.setContent("Source 2:");
		CellHandle cell = paramGrid.getCell(1, 1);
		cell.getContent().add(source1);
		cell = paramGrid.getCell(2, 1);
		cell.getContent().add(source2);
		design.getBody().add(paramGrid);

		
		// preparing summary grid
		String[] headers = {"Sheet Name", "Mismatch Type", "Mismatch on Column", "Mismatch on Row", "Source 1", "Source 2"};
		TextItemHandle text = elementFactory.newTextItem(null);
		
		GridHandle summaryGrid = elementFactory.newGridItem("SummaryGrid", headers.length, 1);
		for(int i=0; i<headers.length; i++) {
			cell = summaryGrid.getCell(i, 1);
			text = elementFactory.newTextItem(null);
			text.setProperty("contentType", "HTML");
			text.setContent(headers[i]);
		}
		design.getBody().add(summaryGrid);
		
		
		return design;
	}
	
	
}

package com.dev4k.birt.excelcomparator.engine;

import java.util.logging.Level;

import org.eclipse.birt.core.framework.Platform;
import org.eclipse.birt.report.engine.api.EngineConfig;
import org.eclipse.birt.report.engine.api.EngineException;
import org.eclipse.birt.report.engine.api.HTMLRenderOption;
import org.eclipse.birt.report.engine.api.IReportEngine;
import org.eclipse.birt.report.engine.api.IReportEngineFactory;
import org.eclipse.birt.report.engine.api.IReportRunnable;
import org.eclipse.birt.report.engine.api.IRunAndRenderTask;
import org.eclipse.birt.report.engine.api.PDFRenderOption;
import org.eclipse.birt.report.model.api.ReportDesignHandle;
import org.springframework.beans.factory.annotation.Value;

public class BirtReportEngine {

	private IReportEngine engine = null;
	private EngineConfig config = null;

	@Value("${report.output.format}")
	private String outputType;

	@Value("${report.output.path}")
	private String outputPath;

	public void runReport(ReportDesignHandle design) throws EngineException {

		try {

			config = new EngineConfig();
			config.setLogConfig(null, Level.WARNING);
			Platform.startup(config);
			IReportEngineFactory factory = (IReportEngineFactory) Platform
					.createFactoryObject(IReportEngineFactory.EXTENSION_REPORT_ENGINE_FACTORY);
			engine = factory.createReportEngine(config);
		} catch (Exception ex) {

		}

		IReportRunnable runnableDesign = engine.openReportDesign(design);

		IRunAndRenderTask runAndRenderTask = engine.createRunAndRenderTask(runnableDesign);

		if (outputType.toLowerCase().equals("html")) {
			HTMLRenderOption htmlOptions = new HTMLRenderOption();
			htmlOptions.setOutputFormat("html");
			htmlOptions.setOutputFileName(outputPath + "Birt_Excel_Comparison_Report" + ".html");
			runAndRenderTask.setRenderOption(htmlOptions);
		}

		if (outputType.toLowerCase().equals("pdf")) {
			PDFRenderOption pdfOptions = new PDFRenderOption();
			pdfOptions.setOutputFormat("pdf");
			pdfOptions.setOutputFileName(outputPath + "Birt_Excel_Comparison_Report" + ".pdf");
			runAndRenderTask.setRenderOption(pdfOptions);
		}

		runAndRenderTask.run();

		runAndRenderTask.close();
		engine.destroy();
		Platform.shutdown();
		design.close();
	}

}

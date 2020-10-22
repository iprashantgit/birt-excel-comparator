package com.dev4k.birt.excelcomparator;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.model.api.ReportDesignHandle;
import org.springframework.context.annotation.AnnotationConfigApplicationContext;

import com.dev4k.birt.excelcomparator.comparator.ExcelComparator;
import com.dev4k.birt.excelcomparator.engine.BirtReportEngine;

public class Application {

	public static void main(String[] args)
			throws EncryptedDocumentException, InvalidFormatException, BirtException, IOException {
		AnnotationConfigApplicationContext context = new AnnotationConfigApplicationContext(ApplicationConfig.class);

		try {
			ExcelComparator excelComparator = context.getBean(ExcelComparator.class);

			ReportDesignHandle design = excelComparator.compareExcel();

			BirtReportEngine reportEngine = context.getBean(BirtReportEngine.class);

			reportEngine.runReport(design);

			System.out.println("Excel Comparison Complete.");
		} finally {
			context.close();
		}
	}

}

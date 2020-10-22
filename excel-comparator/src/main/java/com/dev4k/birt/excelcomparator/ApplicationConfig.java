package com.dev4k.birt.excelcomparator;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.PropertySource;

import com.dev4k.birt.excelcomparator.comparator.ExcelComparator;
import com.dev4k.birt.excelcomparator.designer.ReportDesigner;
import com.dev4k.birt.excelcomparator.engine.BirtReportEngine;

@Configuration
@PropertySource("file:application.properties")
public class ApplicationConfig {
	
	@Bean
	public ExcelComparator excelComparator() {
		return new ExcelComparator();
	}
	
	@Bean
	public ReportDesigner reportDesigner() {
		return new ReportDesigner();
	}
	
	@Bean
	public BirtReportEngine reportEngine() {
		return new BirtReportEngine();
	}
}

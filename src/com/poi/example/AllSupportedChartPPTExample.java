package com.poi.example;
import java.awt.geom.Rectangle2D;
import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.XDDFChart;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFSlide;

/**
 * Build charts without reading template file
 * @author sandeep tiwari
 */
public class AllSupportedChartPPTExample {

    public static void main(String[] args) throws Exception {
        try (BufferedReader modelReader = new BufferedReader(new FileReader("BarData.txt"))) {

        	String chartTitle = modelReader.readLine();  // first line is chart title
        	String[] series = modelReader.readLine().split(",");

        	// Category Axis Data
        	List<String> listLanguages = new ArrayList<>(10);

        	// Values
        	List<Double> listCountries = new ArrayList<>(10);
        	List<Double> listSpeakers = new ArrayList<>(10);
        	
        	// set model
        	String ln;
        	while((ln = modelReader.readLine()) != null) {
        		String[] vals = ln.split(",");
        		listCountries.add(Double.valueOf(vals[0]));
        		listSpeakers.add(Double.valueOf(vals[1]));
        		listLanguages.add(vals[2]);
        	}
        	
        	String[] categories = listLanguages.toArray(new String[listLanguages.size()]);
        	Double[] values1 = listCountries.toArray(new Double[listCountries.size()]);
        	Double[] values2 = listSpeakers.toArray(new Double[listSpeakers.size()]);
        	
        	try(XMLSlideShow ppt = new XMLSlideShow()) {
        		ChartData data = new ChartData();
        		XSLFSlide slide = ppt.createSlide();
        		XSLFChart chart = ppt.createChart();
        		Rectangle2D rect2D = new java.awt.Rectangle(XDDFChart.DEFAULT_X, XDDFChart.DEFAULT_Y, 
        				XDDFChart.DEFAULT_WIDTH*10, XDDFChart.DEFAULT_HEIGHT*10);
        		slide.addChart(chart, rect2D);
        		data.setPieChartData(chart, chartTitle, series, categories, values1, values2,ChartTypes.PIE);
        		
        		
        		slide = ppt.createSlide();
        		XSLFChart chart1 = ppt.createChart();
        		slide.addChart(chart1, rect2D);
        		data.setPieChartData(chart1, chartTitle, series, categories, values1, values2,ChartTypes.PIE3D);
        		
        		slide = ppt.createSlide();
        		XSLFChart chart2 = ppt.createChart();
        		slide.addChart(chart2, rect2D);
        		data.setChartData(chart2, chartTitle, series, categories, values1, values2,ChartTypes.AREA);
        		
        		slide = ppt.createSlide();
        		XSLFChart chart3 = ppt.createChart();
        		slide.addChart(chart3, rect2D);
        		data.setChartData(chart3, chartTitle, series, categories, values1, values2,ChartTypes.AREA3D);
        		
        		slide = ppt.createSlide();
        		XSLFChart chart4 = ppt.createChart();
        		slide.addChart(chart4, rect2D);
        		data.setChartData(chart4, chartTitle, series, categories, values1, values2,ChartTypes.BAR);
        		
        		slide = ppt.createSlide();
        		XSLFChart chart5 = ppt.createChart();
        		slide.addChart(chart5, rect2D);
        		data.setChartData(chart5, chartTitle, series, categories, values1, values2,ChartTypes.BAR3D);
        		
        		slide = ppt.createSlide();
        		XSLFChart chart6 = ppt.createChart();
        		slide.addChart(chart6, rect2D);
        		data.setChartData(chart6, chartTitle, series, categories, values1, values2,ChartTypes.LINE);
        		
        		slide = ppt.createSlide();
        		XSLFChart chart7 = ppt.createChart();
        		slide.addChart(chart7, rect2D);
        		data.setChartData(chart7, chartTitle, series, categories, values1, values2,ChartTypes.LINE3D);
        		
        		slide = ppt.createSlide();
        		XSLFChart chart8 = ppt.createChart();
        		slide.addChart(chart8, rect2D);
        		data.setChartData(chart8, chartTitle, series, categories, values1, values2,ChartTypes.RADAR);
        		
        		slide = ppt.createSlide();
        		XSLFChart chart9 = ppt.createChart();
        		slide.addChart(chart9, rect2D);
        		data.setChartData(chart9, chartTitle, series, categories, values1, values2,ChartTypes.SCATTER);
        		
        		slide = ppt.createSlide();
        		XSLFChart chart10 = ppt.createChart();
        		slide.addChart(chart10, rect2D);
        		data.setChartData(chart10, chartTitle, series, categories, values1, values2,ChartTypes.SURFACE);
        		
        		slide = ppt.createSlide();
        		XSLFChart chart11 = ppt.createChart();
        		slide.addChart(chart11, rect2D);
        		data.setChartData(chart11, chartTitle, series, categories, values1, values2,ChartTypes.SURFACE3D);
        		// save the result
        		try (OutputStream out = new FileOutputStream("All-chart-demo-output.pptx")) {
        			ppt.write(out);
        		}
        	}
        	catch(Exception e)
        	{
        		e.printStackTrace();
        	}
        }
        System.out.println("Done");
    }
}


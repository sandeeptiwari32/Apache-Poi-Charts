package com.poi.example;
import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.XDDFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * Build chart without reading template file
 * @author sandeep tiwari
 */
public class AllSupportedChartDocExample {

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
        	
        	try (XWPFDocument doc = new XWPFDocument()) {
        		ChartData data = new ChartData();
        		XDDFChart chart = doc.createChart(XDDFChart.DEFAULT_WIDTH*10, XDDFChart.DEFAULT_HEIGHT*10);
        		data.setPieChartData(chart, chartTitle, series, categories, values1, values2,ChartTypes.PIE);
        		
        		XDDFChart chart1 = doc.createChart(XDDFChart.DEFAULT_WIDTH*10, XDDFChart.DEFAULT_HEIGHT*10);
        		data.setPieChartData(chart1, chartTitle, series, categories, values1, values2,ChartTypes.PIE3D);
        		
        		XDDFChart chart2 = doc.createChart(XDDFChart.DEFAULT_WIDTH*10, XDDFChart.DEFAULT_HEIGHT*10);
        		data.setChartData(chart2, chartTitle, series, categories, values1, values2,ChartTypes.AREA);
        		
        		XDDFChart chart3 = doc.createChart(XDDFChart.DEFAULT_WIDTH*10, XDDFChart.DEFAULT_HEIGHT*10);
        		data.setChartData(chart3, chartTitle, series, categories, values1, values2,ChartTypes.AREA3D);
        		
        		XDDFChart chart4 = doc.createChart(XDDFChart.DEFAULT_WIDTH*10, XDDFChart.DEFAULT_HEIGHT*10);
        		data.setChartData(chart4, chartTitle, series, categories, values1, values2,ChartTypes.BAR);
        		
        		XDDFChart chart5 = doc.createChart(XDDFChart.DEFAULT_WIDTH*10, XDDFChart.DEFAULT_HEIGHT*10);
        		data.setChartData(chart5, chartTitle, series, categories, values1, values2,ChartTypes.BAR3D);
        		
        		XDDFChart chart6 = doc.createChart(XDDFChart.DEFAULT_WIDTH*10, XDDFChart.DEFAULT_HEIGHT*10);
        		data.setChartData(chart6, chartTitle, series, categories, values1, values2,ChartTypes.LINE);
        		
        		XDDFChart chart7 = doc.createChart(XDDFChart.DEFAULT_WIDTH*10, XDDFChart.DEFAULT_HEIGHT*10);
        		data.setChartData(chart7, chartTitle, series, categories, values1, values2,ChartTypes.LINE3D);
        		
        		XDDFChart chart8 = doc.createChart(XDDFChart.DEFAULT_WIDTH*10, XDDFChart.DEFAULT_HEIGHT*10);
        		data.setChartData(chart8, chartTitle, series, categories, values1, values2,ChartTypes.RADAR);
        		
        		XDDFChart chart9 = doc.createChart(XDDFChart.DEFAULT_WIDTH*10, XDDFChart.DEFAULT_HEIGHT*10);
        		data.setChartData(chart9, chartTitle, series, categories, values1, values2,ChartTypes.SCATTER);
        		
        		XDDFChart chart10 = doc.createChart(XDDFChart.DEFAULT_WIDTH*10, XDDFChart.DEFAULT_HEIGHT*10);
        		data.setChartData(chart10, chartTitle, series, categories, values1, values2,ChartTypes.SURFACE);
        		
        		XDDFChart chart11 = doc.createChart(XDDFChart.DEFAULT_WIDTH*10, XDDFChart.DEFAULT_HEIGHT*10);
        		data.setChartData(chart11, chartTitle, series, categories, values1, values2,ChartTypes.SURFACE3D);
        		// save the result
        		try (OutputStream out = new FileOutputStream("All-chart-demo-output.docx")) {
        			doc.write(out);
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


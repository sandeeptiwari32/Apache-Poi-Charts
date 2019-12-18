package com.poi.example;
import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.XDDFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Build charts without reading template file
 * @author sandeep tiwari
 */
public class AllSupportedChartInPOIExcel {
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

        	try (XSSFWorkbook excel = new XSSFWorkbook()) {
        		XSSFSheet sheet = excel.createSheet("Sheet1");
        		XSSFDrawing drawing = sheet.createDrawingPatriarch();
        		ChartData data = new ChartData();
        		int col1 = 4;
        		int row1 = 0;
        		int col2 = 11;
        		int row2 = 15;
                XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, col1, row1, col2, row2);
                XDDFChart chart = drawing.createChart(anchor);
        		data.setPieChartData(chart, chartTitle, series, categories, values1, values2,ChartTypes.PIE);
        		
        		row1 = row2+2;
        		row2 = row1+15;
        		anchor = drawing.createAnchor(0, 0, 0, 0, col1, row1, col2, row2);
        		XDDFChart chart1 = drawing.createChart(anchor);
        		data.setPieChartData(chart1, chartTitle, series, categories, values1, values2,ChartTypes.PIE3D);
        		
        		row1 = row2+2;
        		row2 = row1+15;
        		anchor = drawing.createAnchor(0, 0, 0, 0, col1, row1, col2, row2);
        		XDDFChart chart2 = drawing.createChart(anchor);
        		data.setChartData(chart2, chartTitle, series, categories, values1, values2,ChartTypes.AREA);
        		
        		row1 = row2+2;
        		row2 = row1+15;
        		anchor = drawing.createAnchor(0, 0, 0, 0, col1, row1, col2, row2);
        		XDDFChart chart3 = drawing.createChart(anchor);
        		data.setChartData(chart3, chartTitle, series, categories, values1, values2,ChartTypes.AREA3D);
        		
        		row1 = row2+2;
        		row2 = row1+15;
        		anchor = drawing.createAnchor(0, 0, 0, 0, col1, row1, col2, row2);
        		XDDFChart chart4 = drawing.createChart(anchor);
        		data.setChartData(chart4, chartTitle, series, categories, values1, values2,ChartTypes.BAR);
        		
        		row1 = row2+2;
        		row2 = row1+15;
        		anchor = drawing.createAnchor(0, 0, 0, 0, col1, row1, col2, row2);
        		XDDFChart chart5 = drawing.createChart(anchor);
        		data.setChartData(chart5, chartTitle, series, categories, values1, values2,ChartTypes.BAR3D);
        		
        		row1 = row2+2;
        		row2 = row1+15;
        		anchor = drawing.createAnchor(0, 0, 0, 0, col1, row1, col2, row2);
        		XDDFChart chart6 = drawing.createChart(anchor);
        		data.setChartData(chart6, chartTitle, series, categories, values1, values2,ChartTypes.LINE);
        		
        		row1 = row2+2;
        		row2 = row1+15;
        		anchor = drawing.createAnchor(0, 0, 0, 0, col1, row1, col2, row2);
        		XDDFChart chart7 = drawing.createChart(anchor);
        		data.setChartData(chart7, chartTitle, series, categories, values1, values2,ChartTypes.LINE3D);
        		
        		row1 = row2+2;
        		row2 = row1+15;
        		anchor = drawing.createAnchor(0, 0, 0, 0, col1, row1, col2, row2);
        		XDDFChart chart8 = drawing.createChart(anchor);
        		data.setChartData(chart8, chartTitle, series, categories, values1, values2,ChartTypes.RADAR);
        		
        		row1 = row2+2;
        		row2 = row1+15;
        		anchor = drawing.createAnchor(0, 0, 0, 0, col1, row1, col2, row2);;
        		XDDFChart chart9 = drawing.createChart(anchor);
        		data.setChartData(chart9, chartTitle, series, categories, values1, values2,ChartTypes.SCATTER);
        		
        		row1 = row2+2;
        		row2 = row1+15;
        		anchor = drawing.createAnchor(0, 0, 0, 0, col1, row1, col2, row2);
        		XDDFChart chart10 = drawing.createChart(anchor);
        		data.setChartData(chart10, chartTitle, series, categories, values1, values2,ChartTypes.SURFACE);
        		
        		row1 = row2+2;
        		row2 = row1+15;
        		anchor = drawing.createAnchor(0, 0, 0, 0, col1, row1, col2, row2);
        		XDDFChart chart11 = drawing.createChart(anchor);
        		data.setChartData(chart11, chartTitle, series, categories, values1, values2,ChartTypes.SURFACE3D);
        		// save the result
        		try (OutputStream out = new FileOutputStream("All-chart-demo-output.xlsx")) {
        			excel.write(out);
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


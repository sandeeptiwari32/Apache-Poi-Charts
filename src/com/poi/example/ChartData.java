package com.poi.example;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.PresetColor;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFChart;
import org.apache.poi.xddf.usermodel.chart.XDDFChartAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData.Series;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFSeriesAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFSurface3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFSurfaceChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFView3D;
/**
 * class to add data into chart
 * @author sandeep tiwari
 *
 */
public class ChartData {

	protected void setChartData(XDDFChart chart, String chartTitle, String[] series, String[] categories, Double[] values1, Double[] values2, ChartTypes type)
    {
        // Use a category axis for the bottom axis.
        XDDFChartAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle(series[2]);
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle(series[0]+","+series[1]);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        
        final int numOfPoints = categories.length;
        final String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
        final String valuesDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 1, 1));
        final String valuesDataRange2 = chart.formatRange(new CellRangeAddress(1, numOfPoints, 2, 2));
        final XDDFDataSource<?> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);
        final XDDFNumericalDataSource<? extends Number> valuesData = XDDFDataSourcesFactory.fromArray(values1, valuesDataRange, 1);
        values1[6] = 16.0; // if you ever want to change the underlying data
        final XDDFNumericalDataSource<? extends Number> valuesData2 = XDDFDataSourcesFactory.fromArray(values2, valuesDataRange2, 2);

        
        XDDFChartData chartData = chart.createData(type, bottomAxis, leftAxis);
        Series series1 = chartData.addSeries(categoriesData, valuesData);
        series1.setTitle(series[0], chart.setSheetTitle(series[0], 1));
        Series series2 = chartData.addSeries(categoriesData, valuesData2);
        series2.setTitle(series[1], chart.setSheetTitle(series[1], 2));
        chartData.setVaryColors(true);

        solidFillSeries(chartData, 0, PresetColor.CHARTREUSE);
        solidFillSeries(chartData, 1, PresetColor.TURQUOISE);
        if(ChartTypes.SURFACE3D == type)
        {
        	XDDFSeriesAxis seriesAxis = chart.createSeriesAxis(AxisPosition.LEFT);
        	((XDDFSurface3DChartData)chartData).defineSeriesAxis(seriesAxis);
        	((XDDFSurface3DChartData)chartData).setWireframe(true);
        }
        else if(ChartTypes.SURFACE == type)
        {
        	XDDFView3D view3D = chart.getOrAddView3D();
        	view3D.setXRotationAngle((byte)90);
            view3D.setYRotationAngle(0);
            view3D.setRightAngleAxes(false);
            view3D.setPerspectiveAngle((short)0);
        	
        	XDDFSeriesAxis seriesAxis = chart.createSeriesAxis(AxisPosition.LEFT);
        	((XDDFSurfaceChartData)chartData).defineSeriesAxis(seriesAxis);
        	((XDDFSurfaceChartData)chartData).setWireframe(true);
        }
        chart.plot(chartData);
        
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.LEFT);
        legend.setOverlay(false);
        
        chart.setTitleText(chartTitle);
        chart.setTitleOverlay(false);
    
    }

    protected void setPieChartData(XDDFChart chart, String chartTitle, String[] series, String[] categories, Double[] values1, Double[] values2, ChartTypes type) {
        final XDDFDataSource<?> categoriesData = XDDFDataSourcesFactory.fromArray(categories, null, 0);
        final XDDFNumericalDataSource<? extends Number> valuesData = XDDFDataSourcesFactory.fromArray(values1, null, 1);
        values1[6] = 16.0; // if you ever want to change the underlying data
        final XDDFNumericalDataSource<? extends Number> valuesData2 = XDDFDataSourcesFactory.fromArray(values2, null, 2);

        
        XDDFChartData bar = chart.createData(type, null, null);
        Series series1 = bar.addSeries(categoriesData, valuesData);
        series1.setTitle(series[0], chart.setSheetTitle(series[0], 1));
        Series series2 = bar.addSeries(categoriesData, valuesData2);
        series2.setTitle(series[1], chart.setSheetTitle(series[1], 2));
        bar.setVaryColors(true);
        chart.plot(bar);
        
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.LEFT);
        legend.setOverlay(false);
        
        chart.setTitleText(chartTitle);
        chart.setTitleOverlay(false);
    }
    
    private void solidFillSeries(XDDFChartData data, int index, PresetColor color) {
        XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(color));
        XDDFChartData.Series series = data.getSeries().get(index);
        XDDFShapeProperties properties = series.getShapeProperties();
        if (properties == null) {
            properties = new XDDFShapeProperties();
        }
        properties.setFillProperties(fill);
        series.setShapeProperties(properties);
    }
}

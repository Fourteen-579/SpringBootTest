package com.fourteen.springboottest.util;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.Charts;
import com.deepoove.poi.data.MergeCellRule;
import com.deepoove.poi.data.TableRenderData;
import com.deepoove.poi.data.Tables;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2023/5/23 14:40
 */
public class DocxUtils {

    // 模板文件路径
    private static final String TEMPLATE_PATH = "C:\\Users\\Administrator\\Desktop\\Test.docx";
    // 生成的文件路径
    private static final String OUTPUT_PATH = "C:\\Users\\Administrator\\Desktop\\Test(2).docx";

    public static void test2() {
        try {
            InputStream is = new FileInputStream(new File(TEMPLATE_PATH));
            XWPFDocument doc = new XWPFDocument(is);

            //模拟统计图数据
            //系列
            String[] seriesTitles = {"日处理能力(kg)", "湿垃圾(kg)"};
            //x轴
            String[] categories = {"2020-02-20", "2020-02-21", "2020-02-22", "2020-02-23", "2020-02-24", "2020-02-25", "2020-02-26"};
            List<Number[]> values = new ArrayList<>();
            //日处理能力
            Number[] value1 = {1000, 1000, 1000, 1000, 1000, 1000, 1000};
            //湿垃圾
            Number[] value2 = {450.2, 622.1, 514, 384.7, 486.5, 688.9, 711.1};

            values.add(value1);
            values.add(value2);

            XWPFChart xChart = doc.getCharts().get(0);//获取第1个图表
            generateChart(xChart, seriesTitles, categories, values);

            FileOutputStream fos = new FileOutputStream(OUTPUT_PATH);
            doc.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void test3() {
        String path = "C:\\Users\\Administrator\\Desktop\\CertificateLetterTemplate.docx";
        try {
            InputStream inputStream = Files.newInputStream(new File(TEMPLATE_PATH).toPath());


            String[][] exportFileInfoBOS = {{"张三1","xxxx","yyyy1"},{"张三2","mmmm","yyyy2"},{"张三3","zzzz","yyyy3"}};
            MergeCellRule rule = MergeCellRule.builder().map(MergeCellRule.Grid.of(1, 0), MergeCellRule.Grid.of(1, 2)).build();
            TableRenderData tableRenderData = Tables.of(exportFileInfoBOS).mergeRule(rule).create();

            Map<String, Object> dataMap = new HashMap<>();
            dataMap.put("exportFileInfoBOS", tableRenderData);

            FileOutputStream outputStream = new FileOutputStream(OUTPUT_PATH);
            XWPFTemplate template = XWPFTemplate.compile(inputStream).render(dataMap);
            template.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void generateChart(XWPFChart chart, String[] series, String[] categories, List<Number[]> values) {
        String chartTitle = "收运量统计图";
        final List<XDDFChartData> data = chart.getChartSeries();
        for (int i = 0; i < data.size(); i++) {
            final XDDFLineChartData line = (XDDFLineChartData) data.get(i);//这里一般获取第一个,我们这里是折线图就是XDDFLineChartData
            final int numOfPoints = categories.length;
            final String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));

            final XDDFDataSource<?> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);
            final String valuesDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, i + 1, i + 1));
            Number[] value = values.get(i);
            final XDDFNumericalDataSource<? extends Number> valuesData = XDDFDataSourcesFactory.fromArray(value, valuesDataRange, i + 1);
            XDDFChartData.Series ser;//图表中的系列
            ser = line.getSeries().get(0);
            ser.replaceData(categoriesData, valuesData);
            CellReference cellReference = chart.setSheetTitle(series[i], 1);//修改系列标题
            ser.setTitle(series[i], cellReference);


            chart.plot(line);
            chart.setTitleText(chartTitle);//折线图标题
            chart.setTitleOverlay(false);
        }
    }

    public static void test1() {
        // 准备数据
        Map<String, Object> dataMap = new HashMap<>();

        Charts.ChartCombos chartCombos = Charts.ofComboSeries("ChartTitle", new String[]{"中文", "English"});

        //设置需要展示多少种数据和对应数据
        chartCombos.addLineSeries("countries", new Double[]{15.0, 6.0});
        chartCombos.addLineSeries("speakers", new Double[]{223.0, 119.0});

        dataMap.put("table", chartCombos.create());

        try {
            XWPFTemplate template = XWPFTemplate.compile(TEMPLATE_PATH).render(dataMap);

            template.writeAndClose(Files.newOutputStream(Paths.get(OUTPUT_PATH)));
            System.out.println("Word success");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}

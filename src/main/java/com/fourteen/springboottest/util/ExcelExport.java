package com.fourteen.springboottest.util;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/11/4 9:39
 */
@Slf4j
public class ExcelExport {

    public static void main(String[] args) {
        String templatePath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\market.xlsx";
        String outputPath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\market-export.xlsx";

        try {
            // 1. 读取模板
            XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(Paths.get(templatePath)));

            // 2. 构造数据
            List<List<MarketExportBO>> allSheetData = buildMockData(3, 5);

            workbook.setSheetName(0, "市场总览");
            // 3. 克隆模板 Sheet
            for (int i = 1; i < allSheetData.size(); i++) {
                workbook.cloneSheet(0, allSheetData.get(i).get(0).getAbbrName());
            }

            // 4. 写出新的带多个 sheet 的模板到内存
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            workbook.write(bos);
            workbook.close();

            // 5. 用新的模板流作为 EasyExcel 模板
            try (InputStream inputStream = new ByteArrayInputStream(bos.toByteArray());
                 OutputStream out = Files.newOutputStream(Paths.get(outputPath));
                 ExcelWriter excelWriter = EasyExcel.write(out).withTemplate(inputStream).build()) {

                FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();

                // 6. 循环填充数据
                for (int i = 0; i < allSheetData.size(); i++) {
                    WriteSheet writeSheet = EasyExcel.writerSheet(i).build();
                    excelWriter.fill(allSheetData.get(i), fillConfig, writeSheet);
                }

                excelWriter.finish();
            }

            System.out.println("✅ 导出成功: " + outputPath);

        } catch (Exception e) {
            log.error("❌ 导出失败: {}", e.getMessage(), e);
        }
    }

    /**
     * 构造多个 sheet 的假数据，每个 sheet 含若干行市场数据
     */
    public static List<List<MarketExportBO>> buildMockData(int sheetCount, int rowsPerSheet) {
        List<List<MarketExportBO>> allSheetData = new ArrayList<>();
        Random random = new Random();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");

        for (int i = 1; i <= sheetCount; i++) {
            List<MarketExportBO> sheetData = new ArrayList<>();
            for (int j = 1; j <= rowsPerSheet; j++) {
                MarketExportBO bo = new MarketExportBO();
                bo.setStockCode("600" + (100 + random.nextInt(900)));
                bo.setAbbrName("测试公司" + i);
                bo.setTradeDate(LocalDate.now().minusDays(random.nextInt(30)).format(formatter));
                bo.setClosePrise(String.format("%.2f", 10 + random.nextDouble() * 20));
                bo.setVolume(String.format("%.2f万股", random.nextDouble() * 500));
                bo.setDealAmount(String.format("%.2f万元", random.nextDouble() * 10000));
                sheetData.add(bo);
            }
            allSheetData.add(sheetData);
        }

        return allSheetData;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class MarketExportBO {

        private String stockCode;

        private String abbrName;

        private String tradeDate;

        private String closePrise;

        private String volume;

        private String dealAmount;

    }

}

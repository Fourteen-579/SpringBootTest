package com.fourteen.springboottest;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/3/22 10:04
 */
//@SpringBootTest
public class PdfWriter {

    // 测试用例
    @Test
    public void testPdfWrite() throws Exception {
        // 1. 定义输入和输出文件路径
        String inputPdfPath = "D:/test.pdf";   // 替换为实际输入文件路径
        String outputPdfPath = "D:/output.pdf"; // 替换为实际输出文件路径

        // 2. 读取本地 PDF 文件并处理
        modifyPdf(inputPdfPath, outputPdfPath, "回避表决");
    }

    /**
     * 修改 PDF 并保存到新文件
     *
     * @param inputPath  输入 PDF 路径
     * @param outputPath 输出 PDF 路径
     * @param text       要写入的文本
     */
    public void modifyPdf(String inputPath, String outputPath, String text) throws Exception {
        PdfReader pdfReader = null;
        PdfStamper pdfStamper = null;
        FileOutputStream outputStream = null;

        try {
            // 1. 读取输入 PDF
            pdfReader = new PdfReader(new FileInputStream(inputPath));
            outputStream = new FileOutputStream(outputPath);
            pdfStamper = new PdfStamper(pdfReader, outputStream);

            // 2. 获取 PDF 页面尺寸（假设操作第 1 页）
            Rectangle pageSize = pdfReader.getPageSize(1);

            // 3. 获取绘制层
            PdfContentByte content = pdfStamper.getOverContent(1); // 第 1 页

            // 4. 调用绘制方法
            drawText(content, pageSize, text);

        } finally {
            // 5. 确保资源关闭
            if (pdfStamper != null) {
                pdfStamper.close();
            }
            if (pdfReader != null) {
                pdfReader.close();
            }
            if (outputStream != null) {
                outputStream.close();
            }
        }
    }

    // 绘制文本方法（调整后的版本）
    private void drawText(PdfContentByte content, Rectangle pdfPageSize, String text) throws Exception {
        // 坐标计算（示例值，根据你的需求调整）
        float x = (float) ((0.137755 - (0.27551 / 2f)) * pdfPageSize.getWidth());
        float y = (float) ((0.971468 - (0.057064 / 2f)) * pdfPageSize.getHeight());
        float boxWidth = (float) (pdfPageSize.getWidth() * 0.27551);
        float boxHeight = (float) (pdfPageSize.getHeight() * 0.057064);

        // 使用嵌入字体（确保字体存在）
        BaseFont baseFont = BaseFont.createFont("C:/Windows/Fonts/simsun.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

        // 计算字体大小
        float fontSize = boxHeight * 2 / 3;

        // 计算文本位置
        float textWidth = baseFont.getWidthPoint(text, fontSize);
        float textX = x + (boxWidth - textWidth) / 2;
        float textY = y + (boxHeight - fontSize) / 2 + baseFont.getAscentPoint(text, fontSize) * 0.9f;

        // 绘制背景
        content.saveState();
        content.setColorFill(new BaseColor(236, 236, 236, 205));
        content.rectangle(x, y, boxWidth, boxHeight);
        content.fill();
        content.restoreState();

        // 绘制文本
        content.beginText();
        content.setFontAndSize(baseFont, fontSize);
        content.setColorFill(BaseColor.BLACK);
        content.setTextMatrix(textX, textY);
        content.showText(text);
        content.endText();
    }
}

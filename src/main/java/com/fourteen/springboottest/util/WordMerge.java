package com.fourteen.springboottest.util;

import com.aspose.words.*;
import lombok.extern.slf4j.Slf4j;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/11/5 9:57
 */
@Slf4j
public class WordMerge {

    private static void getLicense() {
        try (InputStream is = WordMerge.class.getClassLoader().getResourceAsStream("License.xml")) {
            License license = new License();
            license.setLicense(is);
        } catch (Exception e) {
            log.error("license 获取失败", e);
        }
    }

    /**
     * 在正文文档中找到指定二级标题（Heading2）后的章节末尾，
     * 插入包含图表的 Word 内容（保留图表样式与数据）
     *
     * @param bodyBytes   正文文档字节数组
     * @param chartBytes  图表文档字节数组
     * @param headingText 要匹配的二级标题文本，例如“股价走势”
     * @return 合并后的 Word 文档字节数组
     */
    public static byte[] insertChartAfterHeading2(byte[] bodyBytes, byte[] chartBytes, String headingText) {
        getLicense();
        try (ByteArrayInputStream bodyIn = new ByteArrayInputStream(bodyBytes);
             ByteArrayInputStream chartIn = new ByteArrayInputStream(chartBytes);
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {

            // 加载主文档与图表文档
            Document bodyDoc = new Document(bodyIn);
            Document chartDoc = new Document(chartIn);

            // 查找指定的二级标题（Heading 2）
            Paragraph targetHeading = null;
            NodeCollection<Paragraph> paras = bodyDoc.getChildNodes(NodeType.PARAGRAPH, true);
            for (Paragraph p : paras) {
                if (p.getParagraphFormat().getStyleIdentifier() == StyleIdentifier.HEADING_2 &&
                        p.getText().contains(headingText)) {
                    targetHeading = p;
                    break;
                }
            }

            if (targetHeading == null) {
                System.out.println("未找到指定的二级标题：" + headingText);
                return bodyBytes;
            }

            // 找到当前 Heading2 所在章节末尾（即下一个 Heading2 前的位置）
            Paragraph insertAfter = targetHeading;
            for (Node node = targetHeading.getNextSibling(); node != null; node = node.getNextSibling()) {
                if (node instanceof Paragraph) {
                    Paragraph nextP = (Paragraph) node;
                    if (nextP.getParagraphFormat().getStyleIdentifier() == StyleIdentifier.HEADING_2) {
                        break;
                    }
                    insertAfter = nextP;
                }
            }

            // 创建节点导入器（保留图表样式和嵌入对象）
            NodeImporter importer = new NodeImporter(chartDoc, bodyDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

            // 在指定位置后插入图表内容
            CompositeNode parent = insertAfter.getParentNode();
            for (Section section : chartDoc.getSections()) {
                NodeCollection<Node> allChildNodes = section.getBody().getChildNodes();
                for (Node node : allChildNodes) {
                    Node imported = importer.importNode(node, true);
                    parent.insertAfter(imported, insertAfter);
                    insertAfter = (Paragraph) imported; // 更新插入位置
                }
            }

            // 保存结果到字节数组
            bodyDoc.save(output, SaveFormat.DOCX);
            return output.toByteArray();

        } catch (Exception e) {
            log.warn("合并 Word 文档失败：", e);
        }
        return null;
    }

    /**
     * 在封面文档的最后一个 Section 末尾添加一个尾页文字，并在正文文档的最后一个 Section 末尾添加两个换行
     *
     * @param coverBytes 封面文档字节数组
     * @param bodyBytes  正文文档字节数组
     * @param footerText 尾页文字
     * @return 合并后的 Word 文档字节数组
     */
    public static byte[] mergeWordWithCoverAndFooter(byte[] coverBytes, byte[] bodyBytes, String footerText) {
        getLicense();
        try {
            // 加载封面文档
            Document coverDoc;
            try (InputStream in = new ByteArrayInputStream(coverBytes)) {
                coverDoc = new Document(in);
            }

            // 加载正文文档
            Document bodyDoc;
            try (InputStream in = new ByteArrayInputStream(bodyBytes)) {
                bodyDoc = new Document(in);
            }

            // 使用 NodeImporter 保留源格式
            NodeImporter importer = new NodeImporter(bodyDoc, coverDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

            // 将正文内容导入封面文档
            for (Section section : bodyDoc.getSections()) {
                Section importedSection = (Section) importer.importNode(section, true);
                coverDoc.appendChild(importedSection);
            }

            // 获取最终的最后一个 Section
            Section lastSection = coverDoc.getLastSection();

            // 在正文后添加两个换行
            Paragraph lineBreaks = new Paragraph(coverDoc);
            Run lineRun = new Run(coverDoc, "\n\n");
            lineBreaks.appendChild(lineRun);
            lastSection.getBody().appendChild(lineBreaks);

            // 添加尾页文字（不分页）
            Paragraph footerPara = new Paragraph(coverDoc);
            Run footerRun = new Run(coverDoc, footerText);
            footerRun.getFont().setSize(12);
            footerPara.appendChild(footerRun);
            lastSection.getBody().appendChild(footerPara);

            // 导出为 byte[]
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                coverDoc.save(out, SaveFormat.DOCX);
                return out.toByteArray();
            }

        } catch (Exception e) {
            log.warn("合并 Word 文档失败：", e);
            return null;
        }
    }

    public static byte[] setTableStyle(byte[] bytes) {
        try (InputStream in = new ByteArrayInputStream(bytes)) {
            Document doc = new Document(in);
            DocumentBuilder builder = new DocumentBuilder(doc);

            // 遍历所有表格
            for (Table table : (Iterable<Table>) doc.getChildNodes(NodeType.TABLE, true)) {

                // 表格整体宽度居中 + 自动适应内容
                table.setAlignment(TableAlignment.CENTER);
                table.autoFit(AutoFitBehavior.AUTO_FIT_TO_CONTENTS);

                // 去除表格之间的空隙（单元格间距）
                table.setAllowAutoFit(true);
                table.setCellSpacing(0);

                // 设置边框样式（等价于 docx4j 中的 CTBorder）
                table.setBorders(LineStyle.SINGLE, 0.75, java.awt.Color.BLACK);

                // 表格宽度为页面宽度的 50%
                double pageWidth = doc.getFirstSection().getPageSetup().getPageWidth()
                        - doc.getFirstSection().getPageSetup().getLeftMargin()
                        - doc.getFirstSection().getPageSetup().getRightMargin();
                table.setPreferredWidth(PreferredWidth.fromPoints(pageWidth * 0.5));

                // 遍历行
                for (int i = 0; i < table.getRows().getCount(); i++) {
                    Row row = table.getRows().get(i);

                    for (int j = 0; j < row.getCells().getCount(); j++) {
                        Cell cell = row.getCells().get(j);

                        // 段落居中，行距设置
                        for (Paragraph para : cell.getParagraphs()) {
                            para.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
                            para.getParagraphFormat().setSpaceBefore(6);
                            para.getParagraphFormat().setSpaceAfter(6);
                        }

                        // 设置首行底色
                        if (i == 0) {
                            cell.getCellFormat().getShading().setBackgroundPatternColor(
                                    new java.awt.Color(0x95, 0xB3, 0xD7)  // RGB: 149, 179, 215
                            );
                        }

                        // 首行首列加粗
                        if (i == 0 || j == 0) {
                            for (Run run : (Iterable<Run>) cell.getChildNodes(NodeType.RUN, true)) {
                                run.getFont().setBold(true);
                            }
                        }
                    }
                }
            }

            // 输出到字节数组
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                doc.save(out, SaveFormat.DOCX);
                return out.toByteArray();
            }
        } catch (Exception e) {
            throw new RuntimeException("设置表格样式失败", e);
        }
    }
}

package com.fourteen.springboottest.util;

import com.aspose.words.*;
import lombok.extern.slf4j.Slf4j;

import java.awt.*;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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
     * 合并封面与正文文档，并在末尾追加“参考文献”部分（含超链接）
     *
     * @param coverBytes 封面文档字节数组
     * @param bodyBytes  正文文档字节数组
     * @param references 参考文献列表（按顺序）
     * @return 合并后的 Word 文档字节数组
     */
    public static byte[] mergeWordWithCoverAndFooter(byte[] coverBytes,
                                                     byte[] bodyBytes,
                                                     LinkedHashMap<String, MarkDownToWord.Reference> references) {
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

            // 导入正文内容到封面文档
            NodeImporter importer = new NodeImporter(bodyDoc, coverDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
            for (Section section : bodyDoc.getSections()) {
                Section importedSection = (Section) importer.importNode(section, true);
                coverDoc.appendChild(importedSection);
            }

            // 获取合并后的最后一个 Section
            Section lastSection = coverDoc.getLastSection();
            Body body = lastSection.getBody();

            // 添加两个换行
            body.appendChild(new Paragraph(coverDoc));
            body.appendChild(new Paragraph(coverDoc));

            // 创建 DocumentBuilder 方便写入
            DocumentBuilder builder = new DocumentBuilder(coverDoc);
            builder.moveToDocumentEnd();

            // 添加标题“参考文献”，样式设为 Heading1
            builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
            builder.writeln("参考文献");

            // 换回正文样式
            builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.NORMAL);

            // 添加参考文献内容
            if (references != null && !references.isEmpty()) {
                for (MarkDownToWord.Reference ref : references.values()) {
                    // 格式示例：[1] 标题 | 来源 | 日期
                    String prefix = "[" + ref.getIndex() + "] ";
                    builder.write(prefix);

                    // 如果有链接，用超链接写入标题
                    if (ref.getUrl() != null && !ref.getUrl().isEmpty()) {
                        builder.insertHyperlink(ref.getTitle(), ref.getUrl(), false);
                    } else {
                        builder.write(ref.getTitle());
                    }

                    // 追加来源与日期信息
                    String tail = "";
                    if (ref.getSource() != null && !ref.getSource().isEmpty()) {
                        tail += " | " + ref.getSource();
                    }
                    if (ref.getDate() != null && !ref.getDate().isEmpty()) {
                        tail += " | " + ref.getDate();
                    }

                    builder.writeln(tail);
                }
            } else {
                builder.writeln("（无参考文献）");
            }

            // 保存为字节数组
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

            // 遍历所有表格
            for (Table table : (Iterable<Table>) doc.getChildNodes(NodeType.TABLE, true)) {

                // 表格整体宽度居中 + 自动适应内容
                table.setAlignment(TableAlignment.CENTER);
                table.autoFit(AutoFitBehavior.AUTO_FIT_TO_CONTENTS);
                table.setAllowAutoFit(true);

                // 设置表格单元格之间的间距为0
                table.setCellSpacing(-1);

                // 设置边框样式（等价于 docx4j 中的 CTBorder）
                // 1. 禁用内部边框重复绘制
                table.setBorders(LineStyle.SINGLE, 0.75, new Color(128, 128, 128));

                // 表格宽度为页面宽度
                double pageWidth = doc.getFirstSection().getPageSetup().getPageWidth()
                        - doc.getFirstSection().getPageSetup().getLeftMargin()
                        - doc.getFirstSection().getPageSetup().getRightMargin();
                table.setPreferredWidth(PreferredWidth.fromPoints(pageWidth));

                // 遍历行
                for (int i = 0; i < table.getRows().getCount(); i++) {
                    Row row = table.getRows().get(i);

                    for (int j = 0; j < row.getCells().getCount(); j++) {
                        Cell cell = row.getCells().get(j);
                        cell.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
                        // 段落居中，行距设置
                        for (Paragraph para : cell.getParagraphs()) {
                            para.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
                            para.getParagraphFormat().setSpaceBefore(5);
                            para.getParagraphFormat().setSpaceAfter(5);
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

    // 替换文档中的指定格式文字为上标样式的方法
    public static byte[] replaceTextWithSuperscript(byte[] text, LinkedHashMap<String, MarkDownToWord.Reference> testReferences) {
        getLicense();
        try {
            // 从字节数组加载 Word 文档
            ByteArrayInputStream inputStream = new ByteArrayInputStream(text);
            Document doc;
            doc = new Document(inputStream);

            // 正则表达式，用于匹配类似 【#4303865#】 这样的格式
            String regex = "【#(.*?)#】";
            Pattern pattern = Pattern.compile(regex);
            Matcher matcher = pattern.matcher(doc.getText());

            // 遍历所有匹配到的内容，并将其替换为上标
            while (matcher.find()) {
                String originalText = matcher.group(0); // 获取完整匹配文本
                String number = matcher.group(1); // 获取编号部分（例如 4303865）

                String index = "";
                if (testReferences.containsKey(number)) {
                    index = String.valueOf(testReferences.get(number).getIndex());
                }

                // 将原文本替换为 [编号] 格式
                String replacement = "[" + index + "]";

                // 遍历文档中的所有 RUN 节点
                for (Node node : (Iterable<Node>) doc.getChildNodes(NodeType.RUN, true)) {
                    Run run = (Run) node;

                    // 如果 RUN 的文本中包含需要替换的部分
                    if (run.getText().contains(originalText)) {
                        String runText = run.getText();
                        int startIndex = runText.indexOf(originalText);
                        int endIndex = startIndex + originalText.length();

                        Run leftRun = (Run) run.deepClone(true);
                        leftRun.setText(runText.substring(0, startIndex));

                        Run centerRun = (Run) run.deepClone(true);
                        centerRun.setText(replacement);
                        centerRun.getFont().setSuperscript(true);

                        Run rightRun = (Run) run.deepClone(true);
                        rightRun.setText(runText.substring(endIndex));

                        // 将三个 Run 节点插入文档中
                        run.getParentNode().insertBefore(leftRun, run);
                        run.getParentNode().insertBefore(centerRun, run);
                        run.getParentNode().insertBefore(rightRun, run);

                        // 删除原始 Run
                        run.remove();
                    }
                }
            }

            // 将修改后的文档保存到字节数组中
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            doc.save(outputStream, SaveFormat.DOCX);
            byte[] modifiedText = outputStream.toByteArray();

            // 返回修改后的字节数组
            return modifiedText;
        } catch (Exception e) {
            log.warn("替换文字为上标样式失败：", e);
            return null;
        }
    }
}

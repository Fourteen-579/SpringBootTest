package com.fourteen.springboottest.util;

import cn.hutool.core.util.ObjectUtil;
import com.deepoove.poi.XWPFTemplate;
import com.vladsch.flexmark.ext.tables.TablesExtension;
import com.vladsch.flexmark.html.HtmlRenderer;
import com.vladsch.flexmark.parser.Parser;
import com.vladsch.flexmark.util.data.MutableDataSet;
import lombok.extern.slf4j.Slf4j;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/10/24 9:28
 */
@Slf4j
public class MarkDownToWord {

    public static void main(String[] args) throws IOException, Docx4JException {
        //封面页
        String coverPath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\frontCover.docx";
        byte[] cover = createCover(coverPath);
        if (ObjectUtil.isNotEmpty(cover)) {
            String outputPath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\coverOutput.docx";
            Files.write(Paths.get(outputPath), cover);
        } else {
            log.warn("createCover-转换失败，coverPath：{}", coverPath);
        }

        //正文
        String mdPath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\test.md";
        byte[] text = createText(mdPath);
        if (ObjectUtil.isNotEmpty(text)) {
//            text = setWordTableStyle(text);
            String outputPath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\textOutput.docx";
            Files.write(Paths.get(outputPath), text);
        } else {
            log.warn("createText-转换失败，mdPath：{}", mdPath);
        }

        byte[] bytes = mergeWordWithCoverAndFooter(cover, text, "参考来源\n" +
                "[1] 决胜“十四五” 打好收官战|“富矿精开”“点石成金”——贵州用好资源优势打造发展新动能 | 2025-09-16 \n" +
                "[2] 工程机械行业稳步迈入新一轮增长周期 | 2025-09-17 \n" +
                "[3] 国家统计局：8月份规上工业原煤产量3.9亿吨 同比下降3.2% | 2025-09-15\n" +
                "[4] 国家统计局：1-8月份全国房地产开发投资60309亿元，同比下降12.9% | 2025-09-15\n" +
                "[5] 外来电规模创历史新高 江苏多举措迎峰度夏 | 2025-09-17\n" +
                "[6] 我国建成全球规模最大的碳排放权交易市场 | 2025-09-19\n" +
                "[7] 国家统计局：8月份，全国规模以上工业增加值同比增长5.2% | 2025-09-15\n" +
                "[8] 河南证监局党委书记、局长李钢赴新乡市调研 | 2025-09-17\n" +
                "[9] 深市监管动态（2025年9月15日-2025年9月19日） | 2025-09-19\n" +
                "[10] 深圳证监局联合央地金融监管部门开展2025年“金融教育宣传周”活动 | 2025-09-19\n" +
                "[11] 吉林证监局联合启动2025年金融教育宣传周系列活动 | 2025-09-19\n" +
                "[12] 矿山地质环境保护与土地复垦方案审查结果公示 | 2025-09-16");
        if (ObjectUtil.isNotEmpty(bytes)) {
            String outputPath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\output.docx";
            Files.write(Paths.get(outputPath), bytes);
        } else {
            log.warn("mergeWordWithCoverAndFooter-合并失败");
        }
    }

    public static byte[] createCover(String coverPath) {
        try {
            Map<String, Object> dataMap = new HashMap<>();
            dataMap.put("enterpriseName", "东方财富");
            dataMap.put("stockCode", "300039.SZ");
            dataMap.put("startDate", "2000-01-01");
            dataMap.put("endDate", "2020-12-31");
            dataMap.put("swIndustry", "这是申万一级行业、这是申万三级行业");
            dataMap.put("cftIndustry", "这是财富通行业");
            dataMap.put("attentionEnterpriseList", "这是关注公司1、关注公司2");

            InputStream inputStream = Files.newInputStream(Paths.get(coverPath));
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            XWPFTemplate template = XWPFTemplate.compile(inputStream).render(dataMap);
            template.writeAndClose(outputStream);
            return outputStream.toByteArray();
        } catch (Exception e) {
            log.error("createCover-转换失败，失败原因：", e);
        }
        return null;
    }

    public static byte[] createText(String textPath) {
        try {
            String markdown = new String(Files.readAllBytes(Paths.get(textPath)), "UTF-8");
            return markdownToWord(markdown);
        } catch (Exception e) {
            log.error("createText-转换失败，失败原因：", e);
        }
        return null;
    }

    public static byte[] setWordTableStyle(byte[] bytes) {
        try (InputStream in = new ByteArrayInputStream(bytes)) {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(in);
            MainDocumentPart mainPart = wordMLPackage.getMainDocumentPart();

            List<Object> tables = mainPart.getJAXBNodesViaXPath("//w:tbl", true);
            for (Object tblObj : tables) {
                Tbl tbl = (Tbl) ((JAXBElement<?>) tblObj).getValue();

                TblPr tblPr = tbl.getTblPr();
                if (tblPr == null) {
                    tblPr = new TblPr();
                    tbl.setTblPr(tblPr);
                }

                // 设置表格边框
                TblBorders borders = new TblBorders();
                CTBorder border = new CTBorder();
                border.setVal(STBorder.SINGLE); // 单线边框
                border.setSz(BigInteger.valueOf(8)); // 0.5pt
                border.setColor("000000"); // 黑色

                borders.setTop(border);
                borders.setBottom(border);
                borders.setLeft(border);
                borders.setRight(border);
                borders.setInsideH(border);
                borders.setInsideV(border);
                tblPr.setTblBorders(borders);

                // 设置表格宽度
                TblWidth width = new TblWidth();
                width.setType("pct");
                width.setW(BigInteger.valueOf(5000)); // 50%
                tblPr.setTblW(width);

                // 设置表格居中
                Jc jc = new Jc();
                jc.setVal(JcEnumeration.CENTER);
                tblPr.setJc(jc);
            }

            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                wordMLPackage.save(out);
                return out.toByteArray();
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static byte[] markdownToWord(String markdown) {
        String html = null;
        try {
            //将所有表格前加一个换行符，避免表格被当成代码块处理
            markdown = markdown.replaceAll(
                    "(?m)(^.+?[:：]?\\s*\\n)(\\|.*\\|\\s*\\n\\|[-| :]+\\|\\s*\\n(?:\\|.*\\|\\s*\\n?)+)",
                    "$1\n$2"
            );

            // markdown 转 html
            MutableDataSet options = new MutableDataSet();
            options.set(Parser.EXTENSIONS, Collections.singletonList(TablesExtension.create()));
            options.set(HtmlRenderer.ESCAPE_HTML, true);

            Parser parser = Parser.builder(options).build();
            HtmlRenderer renderer = HtmlRenderer.builder(options).build();
            com.vladsch.flexmark.util.ast.Node document = parser.parse(markdown);
            html = "<html><body>" + renderer.render(document) + "</body></html>";


            // html 转 docx
            WordprocessingMLPackage wordPackage = WordprocessingMLPackage.createPackage();
            MainDocumentPart mainDocumentPart = wordPackage.getMainDocumentPart();

            XHTMLImporterImpl xhtmlImporter = new XHTMLImporterImpl(wordPackage);
            mainDocumentPart.getContent().addAll(xhtmlImporter.convert(html, null));

            ByteArrayOutputStream outArrayByteStream = new ByteArrayOutputStream();
            wordPackage.save(outArrayByteStream);
            return outArrayByteStream.toByteArray();
        } catch (Exception e) {
            log.warn("markdownToWord-转换失败，markdown：{}", markdown);
            log.warn("markdownToWord-转换失败，html：{}", html);
            log.error("markdownToWord-转换失败，失败原因：", e);

            return null;
        }
    }

    /**
     * 将封面文件与正文文件合并，并在最后追加文字。
     *
     * @param coverBytes 封面 Word 文件内容
     * @param bodyBytes  正文 Word 文件内容
     * @param footerText 末尾追加的文字
     * @return 合并后的 Word 文件字节数组
     */
    public static byte[] mergeWordWithCoverAndFooter(byte[] coverBytes, byte[] bodyBytes, String footerText) {
        try {
            // 1. 加载封面文档
            WordprocessingMLPackage coverPackage;
            try (InputStream in = new ByteArrayInputStream(coverBytes)) {
                coverPackage = WordprocessingMLPackage.load(in);
            }

            // 2. 加载正文文档
            WordprocessingMLPackage bodyPackage;
            try (InputStream in = new ByteArrayInputStream(bodyBytes)) {
                bodyPackage = WordprocessingMLPackage.load(in);
            }

            // 3. 获取主文档内容
            MainDocumentPart mainPart = coverPackage.getMainDocumentPart();

            // 4. 在封面后分页
            mainPart.addObject(createPageBreak());

            // 5. 合并正文内容
            mainPart.getContent().addAll(bodyPackage.getMainDocumentPart().getContent());

            // 6. 再分页 + 添加结尾文字
            mainPart.addObject(createPageBreak());
            mainPart.addParagraphOfText(footerText);

            // 7. 输出为 byte[]
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            coverPackage.save(outputStream);
            return outputStream.toByteArray();
        } catch (Exception e) {
            log.error("mergeWordWithCoverAndFooter-合并失败，失败原因：", e);
        }
        return null;
    }

    /**
     * 创建分页符
     */
    private static P createPageBreak() {
        ObjectFactory factory = Context.getWmlObjectFactory();
        P p = factory.createP();
        R run = factory.createR();
        Br br = new Br();
        br.setType(STBrType.PAGE);
        run.getContent().add(br);
        p.getContent().add(run);
        return p;
    }

}

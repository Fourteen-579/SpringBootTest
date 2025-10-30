package com.fourteen.springboottest.util;

import cn.hutool.core.util.ObjectUtil;
import com.deepoove.poi.XWPFTemplate;
import com.vladsch.flexmark.ext.tables.TablesExtension;
import com.vladsch.flexmark.html.HtmlRenderer;
import com.vladsch.flexmark.parser.Parser;
import com.vladsch.flexmark.util.ast.Node;
import com.vladsch.flexmark.util.data.MutableDataSet;
import lombok.extern.slf4j.Slf4j;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/10/24 9:28
 */
@Slf4j
public class MarkDownToWord {

    public static final String STYLE = "<style>\n" +
            "                body {\n" +
            "                    font-family: SimSun, '宋体', serif;\n" +
            "                    font-size: 12pt;\n" +
            "                    line-height: 1.5;\n" +
            "                    margin: 1cm 2cm;\n" +
            "                }\n" +
            "                h1 {\n" +
            "                    font-size: 24pt;\n" +
            "                    font-weight: bold;\n" +
            "                    text-align: center;\n" +
            "                    margin-top: 24pt;\n" +
            "                    margin-bottom: 12pt;\n" +
            "                }\n" +
            "                h2 {\n" +
            "                    font-size: 20pt;\n" +
            "                    font-weight: bold;\n" +
            "                    margin-top: 20pt;\n" +
            "                    margin-bottom: 10pt;\n" +
            "                }\n" +
            "                h3 {\n" +
            "                    font-size: 16pt;\n" +
            "                    font-weight: bold;\n" +
            "                    margin-top: 16pt;\n" +
            "                    margin-bottom: 8pt;\n" +
            "                }\n" +
            "                p {\n" +
            "                    text-indent: 2em;\n" +
            "                    margin-top: 6pt;\n" +
            "                    margin-bottom: 6pt;\n" +
            "                }\n" +
            "                table {\n" +
            "                    border-collapse: collapse;\n" +
            "                    width: 100%;\n" +
            "                    margin-top: 10pt;\n" +
            "                    margin-bottom: 10pt;\n" +
            "                }\n" +
            "                th, td {\n" +
            "                    border: 1px solid #000000;\n" +
            "                    padding: 5px;\n" +
            "                    text-align: center;\n" +
            "                }\n" +
            "                th {\n" +
            "                    background-color: #eaeaea;\n" +
            "                    font-weight: bold;\n" +
            "                }\n" +
            "                </style>";
    public static final String END_TEXT = "参考来源\n" +
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
            "[12] 矿山地质环境保护与土地复垦方案审查结果公示 | 2025-09-16";

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
            String outputPath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\textOutput.docx";
            Files.write(Paths.get(outputPath), text);
        } else {
            log.warn("createText-转换失败，mdPath：{}", mdPath);
        }

        //合并
        byte[] bytes = mergeWordWithCoverAndFooter(cover, text, END_TEXT);
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

    public static byte[] markdownToWord(String markdown) {
        String html = null;
        try {
            //删除第一行标题
            markdown = markdown.replaceFirst("^.*?\\r?\\n", "");

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
            Node document = parser.parse(markdown);

            html = "<html><head>" + STYLE + "</head><body>" + renderer.render(document) + "</body></html>";

            // 替换 <h2>
            html = html.replaceAll("<h2(.*?)>", "<h2 class=\"Heading1\"$1>");

            // 替换 <h3>
            html = html.replaceAll("<h3(.*?)>", "<h3 class=\"Heading2\"$1>");

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

    public static void testHeading() {
        String html = "<html><body>" + "<h1 class=\"Heading1\">一级标题</h1>\n" +
                "<h2 class=\"Heading2\">二级标题</h2>" + "</body></html>";
        try {
            WordprocessingMLPackage wordPackage = WordprocessingMLPackage.createPackage();
            MainDocumentPart mainDocumentPart = wordPackage.getMainDocumentPart();

            XHTMLImporterImpl xhtmlImporter = new XHTMLImporterImpl(wordPackage);
            mainDocumentPart.getContent().addAll(xhtmlImporter.convert(html, null));

            wordPackage.save(Files.newOutputStream(Paths.get("C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\heading.docx")));
        } catch (Exception e) {
            log.error("testHeading-转换失败，失败原因：", e);
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

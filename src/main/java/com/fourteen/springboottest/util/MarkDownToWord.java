package com.fourteen.springboottest.util;

import com.vladsch.flexmark.ast.Paragraph;
import com.vladsch.flexmark.ext.tables.TableBlock;
import com.vladsch.flexmark.ext.tables.TablesExtension;
import com.vladsch.flexmark.html.HtmlRenderer;
import com.vladsch.flexmark.parser.Parser;
import com.vladsch.flexmark.util.ast.Node;
import com.vladsch.flexmark.util.data.MutableDataSet;
import org.docx4j.convert.in.xhtml.FormattingOption;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Arrays;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/10/24 9:28
 */
public class MarkDownToWord {

    public static void main(String[] args) throws IOException, Docx4JException {
        test2();
    }

    public static void test2()throws IOException, Docx4JException {
        // 1️⃣ 读取 Markdown 文件
        String mdPath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\test.md";
        String markdown = new String(Files.readAllBytes(Paths.get(mdPath)), "UTF-8");
        markdown = markdown.replaceAll(
                "(?m)(^.+?[:：]?\\s*\\n)(\\|.*\\|\\s*\\n\\|[-| :]+\\|\\s*\\n(?:\\|.*\\|\\s*\\n?)+)",
                "$1\n$2"
        );

        MutableDataSet options = new MutableDataSet();
        options.set(Parser.EXTENSIONS, Arrays.asList(TablesExtension.create()));
        options.set(HtmlRenderer.ESCAPE_HTML, true);

        Parser parser = Parser.builder(options).build();
        HtmlRenderer renderer = HtmlRenderer.builder(options).build();
        Node document = parser.parse(markdown);
        String html = "<html><body>" + renderer.render(document) + "</body></html>";

        System.out.println(html);

        // 创建 WordprocessingMLPackage 实例
        WordprocessingMLPackage wordPackage = WordprocessingMLPackage.createPackage();
        MainDocumentPart mainDocumentPart = wordPackage.getMainDocumentPart();

        XHTMLImporterImpl xhtmlImporter = new XHTMLImporterImpl(wordPackage);
        xhtmlImporter.setTableFormatting(FormattingOption.IGNORE_CLASS);
        mainDocumentPart.getContent().addAll(xhtmlImporter.convert(html, null));

        // 保存 DOCX
        String outputPath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\output.docx";
        wordPackage.save(new File(outputPath));
    }

    public static void test() {
        MutableDataSet options = new MutableDataSet();
//        options.set(HtmlRenderer.SOFT_BREAK, "<br/>\n");
        options.set(Parser.EXTENSIONS, Arrays.asList(TablesExtension.create()));
        options.set(HtmlRenderer.ESCAPE_HTML, true);

        Parser parser = Parser.builder(options).build();
        HtmlRenderer renderer = HtmlRenderer.builder(options).build();

        String markdown = "**监管重点领域对比**：" +
                "| 监管领域 | 主要内容 | 相关影响 |" +
                "|---------|---------|---------|" +
                "| 投资者保护 | 防非宣传、金融知识普及 | 提升公司投资者关系管理要求 |\n" +
                "| 信息披露监管 | 纪律处分、问询函件 | 强化年报及重大事项披露质量 |\n" +
                "| 服务实体经济 | 走访企业、融资支持 | 利好公司再融资及并购重组 |";

        markdown = markdown.replaceAll("(?m)(^.*?\\n)(\\|.+\\|\\s*\\n(?:\\|[-:|]+\\|\\s*\\n(?:\\|.*\\|\\s*\\n?)+))", "$1\n$2");
        Node document = parser.parse(markdown);
        String html = renderer.render(document);

        System.out.println(html);
    }

}

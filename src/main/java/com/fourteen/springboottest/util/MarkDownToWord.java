package com.fourteen.springboottest.util;

import cn.hutool.core.util.ObjectUtil;
import com.deepoove.poi.XWPFTemplate;
import com.vladsch.flexmark.ext.tables.TablesExtension;
import com.vladsch.flexmark.html.HtmlRenderer;
import com.vladsch.flexmark.parser.Parser;
import com.vladsch.flexmark.util.ast.Node;
import com.vladsch.flexmark.util.data.MutableDataSet;
import lombok.extern.slf4j.Slf4j;
import org.docx4j.XmlUtils;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
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

    private static final ObjectFactory factory = new ObjectFactory();

    public static void main(String[] args) throws IOException, Docx4JException {
        //封面页
        String coverPath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\frontCover.docx";
        byte[] cover = createCover(coverPath);

        //正文
        String mdPath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\test.md";
        byte[] text = createText(mdPath);

        //合并
        byte[] bytes = mergeWordWithCoverAndFooter(cover, text, END_TEXT);

        //整体样式处理
        bytes = applyTextStyles(bytes);
        bytes = setTableStyle(bytes);

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
            log.warn("createCover-转换失败，失败原因：", e);
        }
        return null;
    }

    public static byte[] createText(String textPath) {
        try {
            String markdown = new String(Files.readAllBytes(Paths.get(textPath)), "UTF-8");
            return markdownToWord(markdown);
        } catch (Exception e) {
            log.warn("createText-转换失败，失败原因：", e);
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

            html = "<html><body>" + renderer.render(document) + "</body></html>";

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

    /**
     * 合并封面、正文、尾页
     *
     * @param coverBytes 封面
     * @param bodyBytes  正文
     * @param footerText 尾页文字
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

            MainDocumentPart coverMainPart = coverPackage.getMainDocumentPart();
            MainDocumentPart bodyMainPart = bodyPackage.getMainDocumentPart();

            // 3. 拷贝正文样式到封面文档
            StyleDefinitionsPart bodyStyles = bodyMainPart.getStyleDefinitionsPart();
            StyleDefinitionsPart coverStyles = coverMainPart.getStyleDefinitionsPart();
            if (bodyStyles != null && coverStyles != null) {
                coverStyles.getContents().getStyle().addAll(bodyStyles.getContents().getStyle());
            }

            // 4. 封面后分页
            coverMainPart.getContent().add(createPageBreak());

            // 5. 合并正文段落内容
            List<Object> bodyContent = bodyMainPart.getContent();
            coverMainPart.getContent().addAll(bodyContent);

            // 6. 尾页文字
            coverMainPart.getContent().add(createPageBreak());
            coverMainPart.addParagraphOfText(footerText);

            // 7. 输出 byte[]
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            coverPackage.save(outputStream);
            return outputStream.toByteArray();

        } catch (Exception e) {
            log.warn("mergeWordWithCoverAndFooter-合并失败，失败原因：", e);
            return null;
        }
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

    public static byte[] setTableStyle(byte[] bytes) {
        try (InputStream in = new ByteArrayInputStream(bytes)) {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(in);
            MainDocumentPart mainPart = wordMLPackage.getMainDocumentPart();

            List<Object> tables = mainPart.getJAXBNodesViaXPath("//w:tbl", true);

            // 遍历每个表格
            for (Object tblObj : tables) {
                Tbl tbl = (Tbl) ((JAXBElement<?>) tblObj).getValue();
                TblPr tblPr = tbl.getTblPr();
                if (tblPr == null) {
                    tblPr = new TblPr();  // 如果表格没有属性，则创建一个新的
                    tbl.setTblPr(tblPr);
                }

                // 创建单线边框样式
                CTBorder border = new CTBorder();
                border.setVal(STBorder.SINGLE);         // 单线边框
                border.setSz(BigInteger.valueOf(8));    // 边框宽度 (单位: 1/8 pt)，8 ≈ 1pt
                border.setSpace(BigInteger.ZERO);       // 无间距
                border.setColor("000000");              // 黑色边框

                // 将边框应用到表格各个方向
                TblBorders borders = new TblBorders();
                borders.setTop(border);
                borders.setBottom(border);
                borders.setLeft(border);
                borders.setRight(border);
                borders.setInsideH(border); // 表格内部的水平线
                borders.setInsideV(border); // 表格内部的垂直线

                // 应用到表格
                tblPr.setTblBorders(borders);

                // 设置表格宽度和居中对齐
                TblWidth width = new TblWidth();
                width.setType("pct");
                width.setW(BigInteger.valueOf(5000)); // 50% 宽度
                tblPr.setTblW(width);

                Jc jc = new Jc();
                jc.setVal(JcEnumeration.CENTER);  // 设置表格居中
                tblPr.setJc(jc);

                // 获取表格内容：List<Object> rowObjects
                List<Object> rowObjects = tbl.getContent();
                for (int i = 0; i < rowObjects.size(); i++) {
                    Object rowObj = XmlUtils.unwrap(rowObjects.get(i));  // 解包行对象
                    if (!(rowObj instanceof Tr)) {
                        continue;  // 如果不是行，跳过
                    }
                    Tr row = (Tr) rowObj;

                    // 获取行内的单元格对象：List<Object> cellObjects
                    List<Object> cellObjects = row.getContent();
                    for (int j = 0; j < cellObjects.size(); j++) {
                        Object cellObj = XmlUtils.unwrap(cellObjects.get(j));  // 解包单元格对象
                        if (!(cellObj instanceof Tc)) {
                            continue;  // 如果不是单元格，跳过
                        }
                        Tc cell = (Tc) cellObj;

                        //表格内容垂直居中
                        TcPr tcPr = cell.getTcPr();
                        if (tcPr == null) {
                            tcPr = factory.createTcPr();
                            cell.setTcPr(tcPr);
                        }
                        CTVerticalJc tcPrValign = factory.createCTVerticalJc();
                        tcPrValign.setVal(STVerticalJc.CENTER);
                        tcPr.setVAlign(tcPrValign);

                        // 表格内容水平居中（针对单元格里的每个段落）
                        for (Object pObj : cell.getContent()) {
                            Object unwrappedP = XmlUtils.unwrap(pObj);
                            if (unwrappedP instanceof P) {
                                P p = (P) unwrappedP;

                                // 获取或创建段落属性
                                PPr pPr = p.getPPr();
                                if (pPr == null) {
                                    pPr = factory.createPPr();
                                    p.setPPr(pPr);
                                }

                                // 设置段落居中
                                pPr.setJc(jc);
                            }
                        }

                        // 首行设置单元格底色
                        if (i == 0) {
                            CTShd shd = factory.createCTShd();
                            shd.setFill("95b3d7");  // 设置底色为 #95b3d7
                            tcPr.setShd(shd);
                        }

                        // 首行首列单元格加粗
                        if (i == 0 || j == 0) {
                            for (Object pObj : cell.getContent()) {
                                P p = (P) XmlUtils.unwrap(pObj);  // 解包段落对象
                                if (p == null) continue;

                                for (Object rObj : p.getContent()) {
                                    R r = (R) XmlUtils.unwrap(rObj);  // 解包文本对象
                                    if (r == null) continue;

                                    RPr rpr = r.getRPr();
                                    if (rpr == null) {
                                        rpr = factory.createRPr();
                                        r.setRPr(rpr);
                                    }
                                    rpr.setB(factory.createBooleanDefaultTrue());  // 设置加粗
                                }
                            }
                        }
                    }
                }
            }

            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                wordMLPackage.save(out);
                return out.toByteArray();
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static byte[] applyTextStyles(byte[] wordBytes) {
        try (ByteArrayInputStream in = new ByteArrayInputStream(wordBytes)) {
            WordprocessingMLPackage wordPackage = WordprocessingMLPackage.load(in);
            List<Object> paragraphs = wordPackage.getMainDocumentPart().getContent();

            for (Object obj : paragraphs) {
                if (obj instanceof P) {
                    P p = (P) obj;
                    PPr pPr = p.getPPr();
                    if (pPr == null) {
                        pPr = factory.createPPr();
                        p.setPPr(pPr);
                    }

                    // 字体和字号
                    List<Object> runs = p.getContent();
                    for (Object runObj : runs) {
                        if (runObj instanceof R) {
                            R run = (R) runObj;
                            RPr rpr = run.getRPr();
                            if (rpr == null) {
                                rpr = factory.createRPr();
                                run.setRPr(rpr);
                            }
                            RFonts rFonts = factory.createRFonts();
                            rFonts.setAscii("Times New Roman");  // 英文
                            rFonts.setHAnsi("Times New Roman");  // 英文
                            rFonts.setEastAsia("SimSun");        // 中文
                            rpr.setRFonts(rFonts);

                            HpsMeasure sz = factory.createHpsMeasure();
                            sz.setVal(BigInteger.valueOf(22)); // 11pt
                            rpr.setSz(sz);
                        }
                    }

                    // 段落格式
                    PPrBase.Spacing spacing = factory.createPPrBaseSpacing();
                    spacing.setLine(BigInteger.valueOf(240)); // 单倍行距
                    spacing.setBefore(BigInteger.valueOf(120)); // 段前6磅
                    spacing.setAfter(BigInteger.valueOf(120));  // 段后6磅
                    pPr.setSpacing(spacing);

                    // 标题1/2
                    if (pPr.getPStyle() != null) {
                        String style = pPr.getPStyle().getVal();
                        if ("Heading1".equals(style)) {
                            spacing.setBefore(BigInteger.valueOf(600)); // 段前30磅
                            spacing.setAfter(BigInteger.valueOf(120));
                            for (Object runObj : runs) {
                                if (runObj instanceof R) {
                                    ((R) runObj).getRPr().getSz().setVal(BigInteger.valueOf(36)); // 18pt
                                }
                            }
                        } else if ("Heading2".equals(style)) {
                            spacing.setBefore(BigInteger.valueOf(120)); // 6磅
                            spacing.setAfter(BigInteger.valueOf(120));
                            for (Object runObj : runs) {
                                if (runObj instanceof R) {
                                    ((R) runObj).getRPr().getSz().setVal(BigInteger.valueOf(28)); // 14pt
                                }
                            }
                        }
                    }
                }
            }

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            wordPackage.save(out);
            return out.toByteArray();

        } catch (Exception e) {
            log.warn("applyStyles-样式处理失败，失败原因：", e);
            return null;
        }
    }
}

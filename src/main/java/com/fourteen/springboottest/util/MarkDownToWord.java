package com.fourteen.springboottest.util;

import cn.hutool.core.util.ObjectUtil;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.ChartMultiSeriesRenderData;
import com.deepoove.poi.data.ChartSingleSeriesRenderData;
import com.deepoove.poi.data.Charts;
import com.vladsch.flexmark.ext.tables.TablesExtension;
import com.vladsch.flexmark.html.HtmlRenderer;
import com.vladsch.flexmark.parser.Parser;
import com.vladsch.flexmark.util.ast.Node;
import com.vladsch.flexmark.util.data.MutableDataSet;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.docx4j.XmlUtils;
import org.docx4j.convert.in.xhtml.FormattingOption;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import java.io.*;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;

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

    public static void main(String[] args) throws IOException, Docx4JException {
        //图表数据
        String chartPath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\tableTemplate.docx";
        byte[] chart = createChart(chartPath);
        if(ObjectUtil.isEmpty(chart)){
            log.warn("createChart-转换失败");
            return;
        }
        Files.write(Paths.get("C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\chart.docx"), chart);

/*        //封面页
        String coverPath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\frontCover.docx";
        byte[] cover = createCover(coverPath);

        //正文
        String mdPath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\test.md";
        byte[] text = createText(mdPath);

        //合并
        byte[] bytes = mergeWordWithCoverAndFooter(cover, text, END_TEXT);

        //整体样式处理
        bytes = applyTextStyles(bytes);
//        bytes = setTableStyle(bytes);
        bytes = setParagraphsStyle(bytes);

        if (ObjectUtil.isNotEmpty(bytes)) {
            String outputPath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\output.docx";
            Files.write(Paths.get(outputPath), bytes);
        } else {
            log.warn("mergeWordWithCoverAndFooter-合并失败");
        }*/
    }

    public static byte[] createChart(String chartPath){
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            Map<String, Object> map = new HashMap<>();

            ChartMultiSeriesRenderData chart = Charts
                    .ofComboSeries("报告期涨跌幅", new String[]{
                            "平煤股份",
                            "潞安环能",
                            "山西焦煤",
                            "永泰能源",
                            "煤炭(申万一级)",
                            "焦煤(申万三级)",
                            "上证指数",
                            "沪深300"
                    })
                    .addBarSeries("涨跌幅", new Double[]{
                            2.02,
                            13.04,
                            6.90,
                            13.42,
                            3.51,
                            7.26,
                            -1.30,
                            -0.44
                    })
                    .create();

            map.put("statisticChart", chart);
            map.put("abbrName","东方财富");

            XWPFTemplate template = XWPFTemplate.compile(chartPath).render(map);
            template.write(outputStream);
            return outputStream.toByteArray();
        } catch (Exception e) {
            log.error("createChart-生成图表失败", e);
            return null;
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
            xhtmlImporter.setTableFormatting(FormattingOption.CLASS_TO_STYLE_ONLY);
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

            // 4. 拷贝正文的编号定义 (防止圆点丢失)
            NumberingDefinitionsPart bodyNumbering = bodyMainPart.getNumberingDefinitionsPart();
            if (bodyNumbering != null) {
                NumberingDefinitionsPart coverNumbering = coverMainPart.getNumberingDefinitionsPart();
                if (coverNumbering == null) {
                    coverNumbering = new NumberingDefinitionsPart();
                    coverMainPart.addTargetPart(coverNumbering);
                }

                // 将正文的 numbering.xml 合并到封面
                org.docx4j.wml.Numbering bodyNum = bodyNumbering.getContents();
                if (bodyNum != null && bodyNum.getAbstractNum() != null) {
                    if (coverNumbering.getContents() == null) {
                        coverNumbering.setJaxbElement(Context.getWmlObjectFactory().createNumbering());
                    }
                    org.docx4j.wml.Numbering coverNum = coverNumbering.getContents();
                    coverNum.getAbstractNum().addAll(bodyNum.getAbstractNum());
                    coverNum.getNum().addAll(bodyNum.getNum());
                }
            }

            // 5. 封面后分页
            coverMainPart.getContent().add(createPageBreak());

            // 6. 合并正文内容
            List<Object> bodyContent = bodyMainPart.getContent();
            coverMainPart.getContent().addAll(bodyContent);

            // 7. 添加尾页
            coverMainPart.getContent().add(createPageBreak());
            coverMainPart.addParagraphOfText(footerText);

            // 8. 输出 byte[]
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

    public static byte[] setParagraphsStyle(byte[] bytes) {
        try (InputStream in = new ByteArrayInputStream(bytes)) {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(in);
            ObjectFactory factory = Context.getWmlObjectFactory();
            List<Object> paragraphs = wordMLPackage.getMainDocumentPart()
                    .getJAXBNodesViaXPath("//w:p", true);

            for (Object obj : paragraphs) {
                P p = (P) XmlUtils.unwrap(obj);
                if (p == null) continue;

                PPr pPr = p.getPPr();
                if (pPr == null) {
                    pPr = factory.createPPr();
                    p.setPPr(pPr);
                }

                if (pPr.getNumPr() != null) { // 带编号段落
                    PPrBase.Ind ind = pPr.getInd();
                    if (ind == null) {
                        ind = factory.createPPrBaseInd();
                        pPr.setInd(ind);
                    }

                    ind.setLeft(BigInteger.valueOf(567));
                }
            }

            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                wordMLPackage.save(out);
                return out.toByteArray();
            }
        } catch (Exception e) {
            log.warn("setParagraphsStyle-样式处理失败，失败原因：", e);
            return null;
        }
    }

    public static byte[] setTableStyle(byte[] bytes) {
        try (InputStream in = new ByteArrayInputStream(bytes)) {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(in);
            MainDocumentPart mainPart = wordMLPackage.getMainDocumentPart();
            ObjectFactory factory = Context.getWmlObjectFactory();

            // 获取所有表格
            List<Object> tables = mainPart.getJAXBNodesViaXPath("//w:tbl", true);
            for (Object tblObj : tables) {
                Tbl tbl = (Tbl) ((JAXBElement<?>) tblObj).getValue();

                // 表格属性
                TblPr tblPr = Optional.ofNullable(tbl.getTblPr()).orElseGet(factory::createTblPr);
                tbl.setTblPr(tblPr);

                tblPr.setTblW(null);
                // 设置表格布局为自动（允许列宽根据内容自动调整）
                CTTblLayoutType layoutType = factory.createCTTblLayoutType();
                layoutType.setType(STTblLayoutType.AUTOFIT);
                tblPr.setTblLayout(layoutType);

                // 单元格间距（0表示无缝隙）
                TblWidth cellSpacing = factory.createTblWidth();
                cellSpacing.setType("dxa");
                cellSpacing.setW(BigInteger.ZERO);
                tblPr.setTblCellSpacing(cellSpacing);

                // 边框
                CTBorder border = createBorder(factory);
                TblBorders tblBorders = factory.createTblBorders();
                tblBorders.setTop(border);
                tblBorders.setBottom(border);
                tblBorders.setLeft(border);
                tblBorders.setRight(border);
                tblBorders.setInsideH(border);
                tblBorders.setInsideV(border);
                tblPr.setTblBorders(tblBorders);

                // 表格宽度与居中
                TblWidth width = factory.createTblWidth();
                width.setType("pct");
                width.setW(BigInteger.valueOf(5000)); // 50%
                tblPr.setTblW(width);
                tblPr.setJc(createJc(factory, JcEnumeration.CENTER));

                // 遍历行
                for (int i = 0; i < tbl.getContent().size(); i++) {
                    Tr row = (Tr) XmlUtils.unwrap(tbl.getContent().get(i));
                    if (row == null) continue;

                    for (int j = 0; j < row.getContent().size(); j++) {
                        Tc cell = (Tc) XmlUtils.unwrap(row.getContent().get(j));
                        if (cell == null) continue;

                        TcPr tcPr = Optional.ofNullable(cell.getTcPr()).orElseGet(factory::createTcPr);
                        cell.setTcPr(tcPr);

                        tcPr.setTcW(null);

                        // 段落居中与间距
                        for (Object pObj : cell.getContent()) {
                            P p = (P) XmlUtils.unwrap(pObj);
                            if (p == null) continue;

                            PPr pPr = Optional.ofNullable(p.getPPr()).orElseGet(factory::createPPr);
                            p.setPPr(pPr);
                            pPr.setJc(createJc(factory, JcEnumeration.CENTER));
                            pPr.setSpacing(createSpacing(factory, 240, 120, 120));
                        }

                        // 首行底色
                        if (i == 0) {
                            CTShd shd = factory.createCTShd();
                            shd.setFill("95b3d7");
                            tcPr.setShd(shd);
                        }

                        // 首行首列加粗
                        if (i == 0 || j == 0) {
                            boldCellText(factory, cell);
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

    /**
     * 创建统一边框
     */
    private static CTBorder createBorder(ObjectFactory factory) {
        CTBorder border = factory.createCTBorder();
        border.setVal(STBorder.SINGLE);
        border.setSz(BigInteger.valueOf(1));  // 1/8 pt 单位，1 ≈ 0.125pt
        border.setSpace(BigInteger.ZERO);
        border.setColor("#808080");
        return border;
    }

    /**
     * 创建居中方式
     */
    private static Jc createJc(ObjectFactory factory, JcEnumeration val) {
        Jc jc = factory.createJc();
        jc.setVal(val);
        return jc;
    }

    /**
     * 创建段落间距
     */
    private static PPrBase.Spacing createSpacing(ObjectFactory factory, int line, int before, int after) {
        PPrBase.Spacing spacing = factory.createPPrBaseSpacing();
        spacing.setLine(BigInteger.valueOf(line));
        spacing.setBefore(BigInteger.valueOf(before));
        spacing.setAfter(BigInteger.valueOf(after));
        return spacing;
    }

    /**
     * 设置单元格内文字加粗
     */
    private static void boldCellText(ObjectFactory factory, Tc cell) {
        for (Object pObj : cell.getContent()) {
            P p = (P) XmlUtils.unwrap(pObj);
            if (p == null) continue;

            for (Object rObj : p.getContent()) {
                R r = (R) XmlUtils.unwrap(rObj);
                if (r == null) continue;

                RPr rPr = Optional.ofNullable(r.getRPr()).orElseGet(factory::createRPr);
                r.setRPr(rPr);
                rPr.setB(factory.createBooleanDefaultTrue());
            }
        }
    }

    public static byte[] applyTextStyles(byte[] wordBytes) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        try (ByteArrayInputStream in = new ByteArrayInputStream(wordBytes)) {
            WordprocessingMLPackage wordPackage = WordprocessingMLPackage.load(in);
            List<Object> contents = wordPackage.getMainDocumentPart().getContent();

            for (Object obj : contents) {
                if (!(obj instanceof P)) continue;

                P p = (P) obj;
                PPr pPr = Optional.ofNullable(p.getPPr()).orElseGet(factory::createPPr);
                p.setPPr(pPr);

                List<Object> runs = p.getContent();
                for (Object runObj : runs) {
                    if (!(runObj instanceof R)) continue;

                    R run = (R) runObj;
                    RPr rpr = Optional.ofNullable(run.getRPr()).orElseGet(factory::createRPr);
                    run.setRPr(rpr);

                    // 字体设置
                    RFonts rFonts = factory.createRFonts();
                    rFonts.setAscii("Times New Roman");
                    rFonts.setHAnsi("Times New Roman");
                    rFonts.setEastAsia("SimSun");
                    rpr.setRFonts(rFonts);

                    // 字号设置
                    HpsMeasure sz = factory.createHpsMeasure();
                    sz.setVal(BigInteger.valueOf(22)); // 11pt
                    rpr.setSz(sz);
                }

                // 段落行距
                PPrBase.Spacing spacing = createSpacing(factory, 240, 120, 120);
                pPr.setSpacing(spacing);

                // 标题样式
                if (pPr.getPStyle() != null) {
                    String style = pPr.getPStyle().getVal();
                    if ("Heading1".equals(style)) {
                        applyHeadingStyle(factory, runs, spacing, 600, 120, 36);
                    } else if ("Heading2".equals(style)) {
                        applyHeadingStyle(factory, runs, spacing, 120, 120, 28);
                    }
                }
            }

            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                wordPackage.save(out);
                return out.toByteArray();
            }
        } catch (Exception e) {
            log.warn("applyStyles-样式处理失败，失败原因：", e);
            return null;
        }
    }

    /**
     * 应用标题样式
     */
    private static void applyHeadingStyle(ObjectFactory factory, List<Object> runs,
                                          PPrBase.Spacing spacing, int before, int after, int fontSize) {
        spacing.setBefore(BigInteger.valueOf(before));
        spacing.setAfter(BigInteger.valueOf(after));
        for (Object runObj : runs) {
            if (runObj instanceof R) {
                RPr rPr = Optional.ofNullable(((R) runObj).getRPr()).orElseGet(factory::createRPr);
                ((R) runObj).setRPr(rPr);
                rPr.getSz().setVal(BigInteger.valueOf(fontSize));
            }
        }
    }

}

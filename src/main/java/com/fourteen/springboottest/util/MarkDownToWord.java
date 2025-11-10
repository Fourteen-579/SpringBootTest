package com.fourteen.springboottest.util;

import cn.hutool.core.util.ObjectUtil;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.ChartMultiSeriesRenderData;
import com.deepoove.poi.data.Charts;
import com.vladsch.flexmark.ext.tables.TablesExtension;
import com.vladsch.flexmark.html.HtmlRenderer;
import com.vladsch.flexmark.parser.Parser;
import com.vladsch.flexmark.util.ast.Node;
import com.vladsch.flexmark.util.data.MutableDataSet;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.docx4j.XmlUtils;
import org.docx4j.convert.in.xhtml.FormattingOption;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.wml.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
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
        if (ObjectUtil.isEmpty(chart)) {
            log.warn("createChart-转换失败");
            return;
        }

        //封面页
        String coverPath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\frontCover.docx";
        byte[] cover = createCover(coverPath);
        if (ObjectUtil.isEmpty(cover)) {
            log.warn("createCover-转换失败");
            return;
        }

        //正文
        String mdPath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\test.md";
        byte[] text = createText(mdPath);
        if (ObjectUtil.isEmpty(text)) {
            log.warn("createText-转换失败");
            return;
        }

        //设置上标
        LinkedHashMap<String, Reference> testReferences = createTestReferences();
        text = WordMerge.replaceTextWithSuperscript(text,testReferences);
        if (ObjectUtil.isEmpty(text)) {
            log.warn("replaceTextWithSuperscript-转换失败");
            return;
        }

        //设置段落格式
        text = setParagraphsStyle(text);

        //合并图表和正文
        text = WordMerge.insertChartAfterHeading2(text, chart, "股价走势");
        if (ObjectUtil.isEmpty(text)) {
            log.warn("insertChartAfterHeading2-转换失败");
            return;
        }

        //合并封面 正文 参考文献
        byte[] bytes = WordMerge.mergeWordWithCoverAndFooter(cover, text, testReferences);
        if (ObjectUtil.isEmpty(bytes)) {
            log.warn("mergeWordWithCoverAndFooter-转换失败");
            return;
        }

        //整体样式处理
        bytes = applyTextStyles(bytes);
        bytes = WordMerge.setTableStyle(bytes);

        if (ObjectUtil.isNotEmpty(bytes)) {
            String outputPath = "C:\\Users\\Administrator\\Desktop\\资本市场智能报告\\output.docx";
            Files.write(Paths.get(outputPath), bytes);
        } else {
            log.warn("mergeWordWithCoverAndFooter-合并失败");
        }
    }

    /**
     * 创建测试数据
     */
    public static LinkedHashMap<String, Reference> createTestReferences() {
        LinkedHashMap<String, Reference> map = new LinkedHashMap<>();
        map.put("4303865", new Reference("4303865", 1, "决胜“十四五” 打好收官战|“富矿精开”“点石成金”——贵州用好资源优势打造发展新动能", "新闻", "2025-09-16", "新华社", "https://news.example.com/article1"));
        map.put("NW202509153514054281", new Reference("NW202509153514054281", 2, "工程机械行业稳步迈入新一轮增长周期", "行业", "2025-09-17", "人民网", "https://news.example.com/article2"));
        map.put("NW202509153514054632", new Reference("NW202509153514054632", 3, "国家统计局：8月份规上工业原煤产量3.9亿吨 同比下降3.2%", "数据", "2025-09-15", "国家统计局", null));
        return map;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class Reference {
        private String id;
        private Integer index;
        private String title;
        private String infoType;
        private String date;
        private String source;
        private String url;
    }

    public static byte[] createChart(String chartPath) {
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            Map<String, Object> map = new HashMap<>();

            map.put("statisticChart", createStatisticChartData());
            map.put("marketChart", createMarketChartData());
            map.put("abbrName", "东方财富");

            XWPFTemplate template = XWPFTemplate.compile(chartPath).render(map);
            template.write(outputStream);
            return outputStream.toByteArray();
        } catch (Exception e) {
            log.error("createChart-生成图表失败", e);
            return null;
        }
    }

    public static ChartMultiSeriesRenderData createStatisticChartData() {
        ChartMultiSeriesRenderData chart = Charts
                .ofComboSeries("", new String[]{
                        "平煤股份",
                        "潞安环能",
                        "山西焦煤",
                        "永泰能源",
                        "煤炭(申万一级)",
                        "焦煤(申万三级)",
                        "上证指数",
                        "沪深300"
                })
                .addBarSeries("报告期涨跌幅(%)", new Double[]{
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

        return chart;
    }

    public static ChartMultiSeriesRenderData createMarketChartData() {
        // 数据：交易时间，收盘价（不复权），成交量（万股），成交金额（万元）
        String[] dates = {
                "2025-09-15", "2025-09-16", "2025-09-17", "2025-09-18", "2025-09-19"
        };

        // 收盘价（不复权）
        Double[] closePrices = {
                8.0300, 8.0600, 8.1000, 7.8900, 8.0700
        };

        // 成交量（万股）
        Double[] volumes = {
                2815.4088, 3674.6904, 3532.8273, 3302.1030, 3239.6603
        };

        // 成交金额（万元）
        Double[] turnover = {
                22413.0359, 29814.4976, 28645.1952, 26282.2999, 25926.0647
        };

        // 创建图表
        ChartMultiSeriesRenderData chart = Charts
                .ofComboSeries("", dates)
                .addLineSeries("收盘价(不复权)", closePrices)
                .addBarSeries("成交量(万股)", volumes)
                .addBarSeries("成交金额(万元)", turnover)
                .create();

        return chart;
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

    public static byte[] setParagraphsStyle(byte[] bytes) {
        try (InputStream in = new ByteArrayInputStream(bytes)) {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(in);
            ObjectFactory factory = Context.getWmlObjectFactory();

            // numbering 部分（可能为 null）
            NumberingDefinitionsPart numberingPart = wordMLPackage.getMainDocumentPart().getNumberingDefinitionsPart();
            org.docx4j.wml.Numbering numbering = numberingPart == null ? null : numberingPart.getContents();

            // 获取所有段落
            List<Object> paragraphs = wordMLPackage.getMainDocumentPart()
                    .getJAXBNodesViaXPath("//w:p", true);

            for (Object obj : paragraphs) {
                // 可能是 JAXBElement 或 已 unwrap 的对象
                Object unwrapped = XmlUtils.unwrap(obj);
                if (!(unwrapped instanceof P)) continue;
                P p = (P) unwrapped;

                PPr pPr = p.getPPr();
                if (pPr == null) continue;

                // 如果没有 numPr 则跳过
                if (pPr.getNumPr() == null || pPr.getNumPr().getNumId() == null) continue;

                // 获取 numId（BigInteger）
                BigInteger numIdVal = pPr.getNumPr().getNumId().getVal();
                if (numIdVal == null || numbering == null) continue;

                // 通过 numId 在 numbering.getNum() 中找到对应的 <w:num>
                Numbering.Num matchedNum = null;
                List<Numbering.Num> nums = numbering.getNum();
                if (nums != null) {
                    for (Numbering.Num n : nums) {
                        if (n.getNumId() != null && n.getNumId().equals(numIdVal)) {
                            matchedNum = n;
                            break;
                        }
                    }
                }
                if (matchedNum == null) continue;

                // 从 matchedNum 找到 abstractNumId，然后在 numbering.getAbstractNum() 中查找对应的 AbstractNum
                BigInteger abstractNumIdVal = null;
                if (matchedNum.getAbstractNumId() != null && matchedNum.getAbstractNumId().getVal() != null) {
                    abstractNumIdVal = matchedNum.getAbstractNumId().getVal();
                }
                if (abstractNumIdVal == null) continue;

                Numbering.AbstractNum matchedAbstract = null;
                List<Numbering.AbstractNum> abstractNums = numbering.getAbstractNum();
                if (abstractNums != null) {
                    for (Numbering.AbstractNum an : abstractNums) {
                        if (an.getAbstractNumId() != null && an.getAbstractNumId().equals(abstractNumIdVal)) {
                            matchedAbstract = an;
                            break;
                        }
                    }
                }
                if (matchedAbstract == null) continue;

                // 通常我们检查第 0 级（lvl index 0），也可以根据 ilvl 决定具体级别
                // 如果段落有 ilvl 那么优先用 ilvl；否则用 0
                int levelIndex = 0;
                if (pPr.getNumPr().getIlvl() != null && pPr.getNumPr().getIlvl().getVal() != null) {
                    try {
                        levelIndex = pPr.getNumPr().getIlvl().getVal().intValue();
                    } catch (Exception ignored) {
                        levelIndex = 0;
                    }
                }

                // 找出对应级别的 Lvl 节点（防 NPE）
                Lvl lvl = null;
                List<Lvl> lvls = matchedAbstract.getLvl();
                if (lvls != null) {
                    for (Lvl l : lvls) {
                        if (l.getIlvl() != null && l.getIlvl().intValue() == levelIndex) {
                            lvl = l;
                            break;
                        }
                    }
                    if (lvl == null && !lvls.isEmpty()) lvl = lvls.get(0); // 回退到第一个级别
                }
                if (lvl == null) continue;

                // numFmt 存在于 lvl.getNumFmt().getVal()
                String numFmt = null;
                if (lvl.getNumFmt() != null && lvl.getNumFmt().getVal() != null) {
                    numFmt = lvl.getNumFmt().getVal().toString(); // e.g. "decimal", "bullet"
                }

                // 获取或创建 ind
                PPrBase.Ind ind = pPr.getInd();
                if (ind == null) {
                    ind = factory.createPPrBaseInd();
                    pPr.setInd(ind);
                }

                // 根据 numFmt 设定不同的 left 值（你可以根据需要调整数值）
                if (numFmt != null) {
                    if ("bullet".equalsIgnoreCase(numFmt)) {
                        // 无序列表（项目符号）
                        ind.setLeft(BigInteger.valueOf(1000));
                    } else {
                        // 其它当作有序处理（decimal, upperRoman, lowerLetter 等）
                        ind.setLeft(BigInteger.valueOf(500));
                    }
                } else {
                    // 未知类型，使用默认
                    ind.setLeft(BigInteger.valueOf(500));
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
                Color color = new Color();
                color.setVal("000000");
                rPr.setColor(color);
            }
        }
    }

}

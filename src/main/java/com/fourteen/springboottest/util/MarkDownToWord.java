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
import org.docx4j.convert.in.xhtml.FormattingOption;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
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
//        bytes = setTableStyle(bytes);
        bytes = setParagraphsStyle(bytes);

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
    /**
     * 在样式为 Heading2 且文字为 titleName 的段落后插入图片+说明文字
     */
    public static byte[] insertImageAndTextAfterHeading2(byte[] wordBytes,
                                                         String titleName,
                                                         List<byte[]> images,
                                                         List<String> captions) {
        try (InputStream in = new ByteArrayInputStream(wordBytes)) {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(in);
            MainDocumentPart mainPart = wordMLPackage.getMainDocumentPart();
            ObjectFactory factory = Context.getWmlObjectFactory();

            // 获取所有段落
            List<Object> paragraphs = mainPart.getJAXBNodesViaXPath("//w:p", true);

            int insertIndex = -1;
            for (int i = 0; i < paragraphs.size(); i++) {
                P p = (P) XmlUtils.unwrap(paragraphs.get(i));
                String styleId = getParagraphStyleId(p);
                String text = getParagraphText(p);

                // 匹配样式为 Heading2 且文字等于指定标题
                if ("Heading2".equalsIgnoreCase(styleId)
                        && titleName.equals(text != null ? text.trim() : "")) {
                    insertIndex = i + 1;
                    break;
                }
            }

            if (insertIndex == -1) {
                System.out.println("未找到二级标题: " + titleName);
                return wordBytes;
            }

            // 插入图片与说明文字
            for (int j = 0; j < images.size(); j++) {
                byte[] imgBytes = images.get(j);
                String caption = captions.size() > j ? captions.get(j) : "";

                P imageParagraph = createImageParagraph(wordMLPackage, imgBytes, 5000, 3000); // 宽高可调
                mainPart.getContent().add(insertIndex++, imageParagraph);

                P captionParagraph = createFormattedTextParagraph(factory, caption);
                mainPart.getContent().add(insertIndex++, captionParagraph);
            }

            // 导出结果
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            wordMLPackage.save(out);
            return out.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return wordBytes;
        }
    }

    /** 提取段落文字 */
    private static String getParagraphText(P paragraph) {
        StringBuilder sb = new StringBuilder();
        for (Object o : paragraph.getContent()) {
            Object unwrapped = XmlUtils.unwrap(o);
            if (unwrapped instanceof R) {
                R run = (R) unwrapped;
                for (Object rObj : run.getContent()) {
                    Object t = XmlUtils.unwrap(rObj);
                    if (t instanceof Text) {
                        sb.append(((Text) t).getValue());
                    }
                }
            }
        }
        return sb.toString();
    }

    /** 获取段落样式ID（如 "Heading1"、"Heading2"） */
    private static String getParagraphStyleId(P p) {
        if (p.getPPr() != null && p.getPPr().getPStyle() != null) {
            return p.getPPr().getPStyle().getVal();
        }
        return null;
    }

    /** 创建图片段落 */
    private static P createImageParagraph(WordprocessingMLPackage wordMLPackage, byte[] imageBytes, int width, int height) throws Exception {
        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, imageBytes);
        Inline inline = imagePart.createImageInline(null, null, 0, 1, width, height, false);

        ObjectFactory factory = Context.getWmlObjectFactory();
        P p = factory.createP();
        R r = factory.createR();
        Drawing drawing = factory.createDrawing();
        drawing.getAnchorOrInline().add(inline);
        r.getContent().add(drawing);
        p.getContent().add(r);
        return p;
    }

    /** 创建格式化文字段落（如加粗、居中） */
    private static P createFormattedTextParagraph(ObjectFactory factory, String text) {
        P p = factory.createP();
        R r = factory.createR();
        Text t = factory.createText();
        t.setValue(text);

        RPr rPr = factory.createRPr();
        // 字号：12pt
        HpsMeasure size = new HpsMeasure();
        size.setVal(BigInteger.valueOf(24)); // 12pt
        rPr.setSz(size);
        rPr.setSzCs(size);
        // 加粗
        BooleanDefaultTrue b = new BooleanDefaultTrue();
        rPr.setB(b);
        r.setRPr(rPr);
        r.getContent().add(t);
        p.getContent().add(r);

        // 段落居中
        PPr pPr = factory.createPPr();
        Jc jc = factory.createJc();
        jc.setVal(JcEnumeration.CENTER);
        pPr.setJc(jc);
        p.setPPr(pPr);

        return p;
    }

}

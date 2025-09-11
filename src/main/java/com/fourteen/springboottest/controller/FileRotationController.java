package com.fourteen.springboottest.controller;

import cn.hutool.core.util.ObjUtil;
import lombok.RequiredArgsConstructor;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.graphics.image.LosslessFactory;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.util.Matrix;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.net.URLEncoder;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/9/3 15:44
 */
@RestController
@RequiredArgsConstructor
public class FileRotationController {

    public static byte[] normalizePdf(MultipartFile multipartFile) throws Exception {
        try (PDDocument doc = PDDocument.load(multipartFile.getInputStream())) {
            PDDocument newDoc = new PDDocument();
            PDFRenderer renderer = new PDFRenderer(doc);

            for (int i = 0; i < doc.getNumberOfPages(); i++) {
                PDPage page = doc.getPage(i);
                int rotation = page.getRotation();
                PDRectangle mediaBox = page.getMediaBox();
                float pageWidth = mediaBox.getWidth();
                float pageHeight = mediaBox.getHeight();

                // 旋转 90/270 时需要交换宽高
                PDRectangle newBox = mediaBox;
                if (rotation == 90 || rotation == 270) {
                    newBox = new PDRectangle(pageHeight, pageWidth);
                }

                PDPage newPage = new PDPage(newBox);
                newDoc.addPage(newPage);

                try (PDPageContentStream contentStream =
                             new PDPageContentStream(newDoc, newPage, PDPageContentStream.AppendMode.APPEND, true, true)) {
                    // 用 renderer 渲染旧页面为图像再画进去（避免 importPage）
                    BufferedImage bim = renderer.renderImageWithDPI(i, 300); // 渲染成 300 DPI
                    PDImageXObject pdImage = LosslessFactory.createFromImage(newDoc, bim);
                    contentStream.drawImage(pdImage, 0, 0, newBox.getWidth(), newBox.getHeight());
                }
            }

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            newDoc.save(baos);
            newDoc.close();
            return baos.toByteArray();
        }
    }

    @PostMapping("/rotationFile")
    public void rotationFile(@RequestParam("file") MultipartFile file,
                             HttpServletResponse response) throws IOException {

        byte[] fileBytes = null;
        try {
            fileBytes = normalizePdf(file);
        } catch (Exception e) {
            e.printStackTrace();
        }
        if (ObjUtil.isEmpty(fileBytes)) {
            System.out.println("失败");
            return;
        }

        // 设置响应头
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode("new.pdf", "UTF-8"));
        response.setContentLength(fileBytes.length);

        // 写入响应输出流
        try (ServletOutputStream outputStream = response.getOutputStream()) {
            outputStream.write(fileBytes);
            outputStream.flush();
        }
    }
}

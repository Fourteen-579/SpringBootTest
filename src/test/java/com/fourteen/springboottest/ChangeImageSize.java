package com.fourteen.springboottest;

import net.coobird.thumbnailator.Thumbnails;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2024/5/11 16:13
 */
@SpringBootTest
public class ChangeImageSize {

    @Test
    void changeImageSize() throws IOException {
        String base64 = "iVBORw0KGgoAAAANSUhEUgAAAGQAAABkAQMAAABKLAcXAAAAAXNSR0IB2cksfwAAAAlwSFlzAAALEwAACxMBAJqcGAAAAANQTFRFAAAAp3o92gAAAAF0Uk5TAEDm2GYAAAAUSURBVHicY2AYBaNgFIyCUUBPAAAFeAABCP1FgwAAAABJRU5ErkJggg==";
        System.out.println(changeSignImageSize(base64));
    }

    private String changeSignImageSize(String signImageBase64) throws IOException {
        byte[] imageBytes = java.util.Base64.getDecoder().decode(signImageBase64);
        ByteArrayInputStream bis = new ByteArrayInputStream(imageBytes);

        ByteArrayOutputStream bos = new ByteArrayOutputStream();

        Thumbnails.of(bis)
                .size(200, 100)
                .imageType(BufferedImage.TYPE_INT_ARGB)
                .keepAspectRatio(false)
                .outputFormat("png")
                .toOutputStream(bos);

        byte[] resizedImageBytes = bos.toByteArray();
        bis.close();
        bos.close();
        return java.util.Base64.getEncoder().encodeToString(resizedImageBytes);
    }


}

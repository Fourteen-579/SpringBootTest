package com.fourteen.springboottest;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.plugin.table.LoopRowTableRenderPolicy;
import lombok.AllArgsConstructor;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@SpringBootTest
class PoiWordTest {

    private static final String OUTPUT_PATH = "C:\\Users\\Administrator\\Desktop\\Test.docx";

//    @Test
    void contextLoads() {
        String path = "C:\\Users\\Administrator\\Desktop\\CertificateLetterTemplate.docx";
        try {
            LoopRowTableRenderPolicy policy = new LoopRowTableRenderPolicy();
            Configure config = Configure.builder()
                    .bind("exportFileInfoBOS", policy)
                    .bind("signMemberInfoBO", policy).build();

            InputStream inputStream = Files.newInputStream(new File(path).toPath());

            List<FileInfo> exportFileInfoBOS = Arrays.asList(
                    new FileInfo("111", "222", "333"),
                    new FileInfo("111", "222", "333"),
                    new FileInfo("111", "222", "333")
            );
            Map<String, Object> dataMap = new HashMap<>();
            dataMap.put("exportFileInfoBOS", exportFileInfoBOS);
            dataMap.put("matterName", "事项名");

            FileOutputStream outputStream = new FileOutputStream(OUTPUT_PATH);
            XWPFTemplate template = XWPFTemplate.compile(inputStream, config).render(dataMap);
            template.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @AllArgsConstructor
    class FileInfo {
        private String name;

        private String positionDetail;

        private String fileHash;
    }
}

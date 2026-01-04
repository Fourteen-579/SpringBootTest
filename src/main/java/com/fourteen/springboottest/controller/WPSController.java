package com.fourteen.springboottest.controller;

import com.fourteen.springboottest.util.ObjectMappers;
import com.fourteen.springboottest.util.WpsClient;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import okhttp3.Response;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.HashMap;
import java.util.Map;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/12/30 13:31
 */
@Slf4j
@RestController
@RequiredArgsConstructor
public class WPSController {

    @PostMapping("/wps/{fileId}")
    public String wordToPdf(@PathVariable String fileId) throws Exception {
        Map<String, String> map = new HashMap<>();
        map.put("url", "https://choiceqyttest.eastmoney.com/ent/temp/downloadFile/" + fileId);
        map.put("filename", fileId + ".docx");

        String bodyJson = ObjectMappers.writeAsJsonStrThrow(map);

        Response result = WpsClient.postJson(
                "/api/developer/v1/office/convert/to/pdf",
                bodyJson
        );

        log.info("result={}", result);
        log.info("result body={}", result.body().string());
        return ObjectMappers.writeAsJsonStrThrow(result);
    }

    @PostMapping("/wps2/{taskId}")
    public String getResult(@PathVariable String taskId) throws Exception {
        String url = "/api/developer/v1/tasks/open:" + taskId;

        Response result = WpsClient.get(url);

        log.info("result={}", result);
        log.info("result body={}", result.body().string());
        return ObjectMappers.writeAsJsonStrThrow(result);
    }
}

package com.fourteen.springboottest.controller;

import cn.hutool.core.util.ObjectUtil;
import cn.hutool.core.util.StrUtil;
import com.fourteen.springboottest.util.AesUtil;
import com.fourteen.springboottest.util.HttpUtils;
import lombok.RequiredArgsConstructor;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import javax.annotation.Resource;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/7/29 15:59
 */
@RestController
@RequiredArgsConstructor
public class FileDownloadController {

    @Resource
    private HttpUtils httpUtils;

    @PostMapping("/downloadFile")
    public String downloadFile(@RequestParam("url") String url, @RequestParam("fileId") String fileId) throws IOException {
        saveToFile(threeMeetingFileDownload(fileId), url);
        return "success";
    }

    public void saveToFile(ByteArrayResource resource, String filePath) throws IOException {
        File file = new File(filePath);

        // 创建父目录（如果不存在）
        File parent = file.getParentFile();
        if (parent != null && !parent.exists()) {
            parent.mkdirs();
        }

        try (InputStream inputStream = resource.getInputStream();
             FileOutputStream outputStream = new FileOutputStream(file)) {

            byte[] buffer = new byte[4096];
            int bytesRead;

            while ((bytesRead = inputStream.read(buffer)) != -1) {
                outputStream.write(buffer, 0, bytesRead);
            }

            outputStream.flush();
        }
    }

    public ByteArrayResource threeMeetingFileDownload(String fileId) {
        byte[] bytes = downloadFile(fileId);
        if (ObjectUtil.isEmpty(bytes)) {
            return null;
        }

        byte[] desBytes = AesUtil.aesDecrypt256(bytes);
        return new ByteArrayResource(desBytes);
    }

    public byte[] downloadFile(String fileId) {
        if (StrUtil.isBlank(fileId)) {
            return null;
        }

        Map<String, String> paramMap = new HashMap<>();
        paramMap.put("fid", fileId);
        Map<String, String> headerMap = new HashMap<>();
        headerMap.put("ossToken", "e465dd65e6ef49768972015ce254f7bc");

        String downloadFileUrl = "http://10.205.37.102:8899/oss-back/qyt-biz-test/shareholder-meeting/download";

        byte[] bytes = httpUtils.postFormGetByte(downloadFileUrl, paramMap, headerMap);
        if (ObjectUtil.isEmpty(bytes)) {
            return null;
        }

        return bytes;
    }

}

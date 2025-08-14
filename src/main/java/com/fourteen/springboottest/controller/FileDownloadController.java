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
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.nio.file.Paths;
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

    public static final String TEST_URL = "http://10.150.224.34:8899";
    public static final String TEST_TOKEN = "kjdhqih91nfkqbfih91n";
    public static final String PROD_URL = "http://10.205.37.102:8899";
    public static final String PROD_TOKEN = "e465dd65e6ef49768972015ce254f7bc";
    @Resource
    private HttpUtils httpUtils;

    @PostMapping("/downloadFile")
    public void downloadFile(@RequestParam("fileId") String fileId,
                             HttpServletResponse response) throws IOException {
        // 获取文件字节流（你原先的逻辑）
        byte[] fileBytes = threeMeetingFileDownload(fileId);

        // 设置响应头
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(fileId+".pdf", "UTF-8"));
        response.setContentLength(fileBytes.length);

        // 写入响应输出流
        try (ServletOutputStream outputStream = response.getOutputStream()) {
            outputStream.write(fileBytes);
            outputStream.flush();
        }
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

    public byte[] threeMeetingFileDownload(String fileId) {
        byte[] bytes = downloadFile(fileId);
        if (ObjectUtil.isEmpty(bytes)) {
            return null;
        }

        return AesUtil.aesDecrypt256(bytes);
    }

    public byte[] downloadFile(String fileId) {
        if (StrUtil.isBlank(fileId)) {
            return null;
        }

        Map<String, String> paramMap = new HashMap<>();
        paramMap.put("fid", fileId);
        Map<String, String> headerMap = new HashMap<>();
        headerMap.put("ossToken", PROD_TOKEN);

        String downloadFileUrl = PROD_URL + "/oss-back/qyt-biz-test/shareholder-meeting/download";

        byte[] bytes = httpUtils.postFormGetByte(downloadFileUrl, paramMap, headerMap);
        if (ObjectUtil.isEmpty(bytes)) {
            return null;
        }

        return bytes;
    }

}

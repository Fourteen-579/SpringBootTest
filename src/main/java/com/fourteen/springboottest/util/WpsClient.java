package com.fourteen.springboottest.util;

import lombok.extern.slf4j.Slf4j;
import okhttp3.*;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.time.ZoneOffset;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Locale;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/12/30 14:37
 */
@Slf4j
public class WpsClient {

    private static final String APP_ID = "SX20251230DAGSWR";
    private static final String APP_SECRET = "SVTfQWEjIqREeMjdmTIRjqGDYFdrKnlb";
    private static final String BASE_URL = "https://solution.wps.cn";
    private static final String CONTENT_TYPE = "application/json;charset=utf-8";

    private static final OkHttpClient CLIENT = new OkHttpClient();

    private static String buildSignature(String contentMd5, String date) {
        String signSource = APP_SECRET + contentMd5 + CONTENT_TYPE + date;
        return sha1(signSource);
    }

    private static Request.Builder baseBuilder(String url, String contentMd5, String date) {
        return new Request.Builder()
                .url(BASE_URL + url)
                .addHeader("Date", date)
                .addHeader("Content-Type", CONTENT_TYPE)
                .addHeader("Content-Md5", contentMd5)
                .addHeader("Authorization",
                        "WPS-2:" + APP_ID + ":" + buildSignature(contentMd5, date));
    }

    /**
     * 获取请求body MD5
     */
    public static String generateMD5(String input) {
        try {
            MessageDigest md = MessageDigest.getInstance("MD5");
            byte[] digest = md.digest(input.getBytes(StandardCharsets.UTF_8));

            StringBuilder hexString = new StringBuilder();
            for (byte b : digest) {
                String hex = Integer.toHexString(0xff & b);
                if (hex.length() == 1) {
                    hexString.append('0');
                }
                hexString.append(hex);
            }
            return hexString.toString();
        } catch (NoSuchAlgorithmException e) {
            throw new RuntimeException("MD5 algorithm not found", e);
        }
    }

    /**
     * 获取日期字符串
     */
    public static String dateStr() {
        ZonedDateTime now = ZonedDateTime.now(ZoneOffset.UTC);
        // 定义 RFC 1123 格式
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(
                "EEE, dd MMM yyyy HH:mm:ss O",
                Locale.US
        );
        return now.format(formatter);
    }

    public static String sha1(String input) {
        try {
            MessageDigest digest = MessageDigest.getInstance("SHA-1");
            byte[] hashBytes = digest.digest(input.getBytes(StandardCharsets.UTF_8));
            // 转换为十六进制字符串
            StringBuilder hexString = new StringBuilder();
            for (byte b : hashBytes) {
                hexString.append(String.format("%02x", b));
            }
            return hexString.toString();
        } catch (NoSuchAlgorithmException e) {
            throw new RuntimeException(e);
        }
    }

    public static Response postJson(String url, String bodyJson) throws Exception {
        String date = dateStr();
        String contentMd5 = generateMD5(bodyJson);

        RequestBody body = RequestBody.create(
                MediaType.parse(CONTENT_TYPE), bodyJson);

        Request request = baseBuilder(url, contentMd5, date)
                .post(body)
                .build();

        log.info("request={}", request);

        return CLIENT.newCall(request).execute();
    }

    public static Response get(String url) throws Exception {
        String date = dateStr();
        String contentMd5 = generateMD5(url);

        Request request = baseBuilder(url, contentMd5, date)
                .get()
                .build();

        log.info("request={}", request);

        return CLIENT.newCall(request).execute();
    }
}

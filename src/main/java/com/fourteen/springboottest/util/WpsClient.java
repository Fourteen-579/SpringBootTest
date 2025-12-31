package com.fourteen.springboottest.util;

import lombok.extern.slf4j.Slf4j;
import okhttp3.*;

import java.time.ZoneOffset;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;

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
    private static final String CONTENT_TYPE = "application/json";

    private static final OkHttpClient CLIENT = new OkHttpClient();

    private static String now() {
        return ZonedDateTime.now(ZoneOffset.UTC)
                .format(DateTimeFormatter.RFC_1123_DATE_TIME);
    }

    private static String buildSignature(String contentMd5, String date) {
        String signSource = APP_SECRET + contentMd5 + CONTENT_TYPE + date;
        return Sha1Util.sha1(signSource);
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

    public static Response postJson(String url, String bodyJson) throws Exception {
        String date = now();
        String contentMd5 = Md5Util.md5HexLower(bodyJson);

        RequestBody body = RequestBody.create(
                MediaType.parse(CONTENT_TYPE), bodyJson);

        Request request = baseBuilder(url, contentMd5, date)
                .post(body)
                .build();

        return CLIENT.newCall(request).execute();
    }

    public static Response get(String url) throws Exception {
        String date = now();
        String contentMd5 = Md5Util.md5HexLower(url);

        Request request = baseBuilder(url, contentMd5, date)
                .get()
                .build();

        return CLIENT.newCall(request).execute();
    }
}

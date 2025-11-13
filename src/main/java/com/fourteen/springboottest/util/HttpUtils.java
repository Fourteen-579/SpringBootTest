package com.fourteen.springboottest.util;

import cn.hutool.core.collection.CollUtil;
import lombok.extern.slf4j.Slf4j;
import okhttp3.*;
import org.apache.commons.lang3.tuple.ImmutablePair;
import org.springframework.stereotype.Component;
import org.springframework.util.CollectionUtils;

import javax.annotation.Resource;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.Dictionary;
import java.util.Enumeration;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.TimeUnit;

/**
 * @Description: HTTP封装类
 * @Author: huangwei
 * @Date: 2022/08/18 19:48
 */
@Slf4j
@Component
public class HttpUtils {

    public static final OkHttpClient langCallHttpClient = newLongCallHttpClient();
    private static final String JSON_TYPE = "application/json; charset=utf-8";

    @Resource
    private OkHttpClient client;

    public static OkHttpClient newLongCallHttpClient() {
        Dispatcher dispatcher = new Dispatcher();
        dispatcher.setMaxRequests(1200);
        dispatcher.setMaxRequestsPerHost(120);

        OkHttpClient.Builder builder = new OkHttpClient.Builder()
                .dispatcher(dispatcher)
                .connectionPool(new ConnectionPool(5, 3, TimeUnit.MINUTES))
                .connectTimeout(10, TimeUnit.SECONDS)
                .writeTimeout(30, TimeUnit.SECONDS)
                .readTimeout(600, TimeUnit.SECONDS)
                .pingInterval(5, TimeUnit.SECONDS);

        return builder.build();
    }

    public String post(String url, String jsonParam, Dictionary<String, String> headers) {
        String result = null;
        //请求参数
        RequestBody requestBody = FormBody.create(MediaType.parse(JSON_TYPE), jsonParam);
        Request.Builder requestBuilder = new Request.Builder().url(url).post(requestBody);
        //添加请求头
        if (headers != null) {
            for (Enumeration<?> keys = headers.keys(); keys.hasMoreElements(); ) {
                Object key = keys.nextElement();
                Object value = headers.get(key);
                requestBuilder.addHeader(key.toString(), value.toString());
            }
        }
        Request request = requestBuilder.build();
        try (Response response = client.newCall(request).execute()) {
            if (response != null && response.code() == 200 && response.body() != null) {
                result = response.body().string();
            } else {
                log.error("HttpUtils -> post error, url={}, param={}, header={}, response={}", url, jsonParam, ObjectMappers.writeAsJsonStrThrow(headers), response == null ? "" : response.toString());
            }
        } catch (Exception e) {
            log.error("HttpUtils -> post error, url={}, json={}, header={}, error={}", url, jsonParam, ObjectMappers.writeAsJsonStrThrow(headers), e.getMessage(), e);
        }
        return result;
    }


    public String postForm(String url, Map<String, String> bodyParam, Map<String, String> headers) {
        String result = null;
        //构建Form格式请求参数
        FormBody.Builder builder = new FormBody.Builder();
        if (bodyParam != null) {
            bodyParam.forEach(builder::add);
        }
        FormBody formBody = builder.build();
        //添加请求头
        Request.Builder requestBuilder = new Request.Builder().url(url);
        if (headers != null) {
            headers.forEach(requestBuilder::addHeader);
        }
        //发送请求
        Request request = requestBuilder.post(formBody).build();
        try (Response response = client.newBuilder().build().newCall(request).execute()) {
            if (response != null && response.code() == 200 && response.body() != null) {
                result = response.body().string();
            } else {
                log.error("HttpUtils -> post error, url={}, param={}, header={}, response={}", url, ObjectMappers.writeAsJsonStrThrow(bodyParam), ObjectMappers.writeAsJsonStrThrow(headers), response == null ? "" : response.toString());
            }
        } catch (Exception e) {
            log.error("HttpUtils -> post error, url={}, from={}, header={}, error={}", url, ObjectMappers.writeAsJsonStrThrow(bodyParam), ObjectMappers.writeAsJsonStrThrow(headers), e.getMessage(), e);
        }
        return result;
    }


    public String post(String url, String jsonParam, Map<String, String> headers) {
        String result = null;
        //请求参数
        RequestBody requestBody = FormBody.create(MediaType.parse(JSON_TYPE), jsonParam);
        Request.Builder requestBuilder = new Request.Builder().url(url).post(requestBody);
        //添加请求头
        if (headers != null) {
            headers.forEach(requestBuilder::addHeader);
        }


        Request request = requestBuilder.build();
        try (Response response = client.newCall(request).execute()) {
            if (response != null && response.code() == 200 && response.body() != null) {
                result = response.body().string();
            } else {
                log.error("HttpUtils -> post error");
            }
            log.info("HttpUtils -> post url={}, param={}, header={}, response={},result={}", url, jsonParam, ObjectMappers.writeAsJsonStrThrow(headers), response == null ? "" : response.toString(), result);
        } catch (Exception e) {
            log.error("HttpUtils -> post error, url={}, json={}, header={}, error={}", url, jsonParam, ObjectMappers.writeAsJsonStrThrow(headers), e.getMessage(), e);
        }
        return result;
    }


    public String postForLongCall(String url, String jsonParam, Map<String, String> headers) {
        String result = null;
        //请求参数
        RequestBody requestBody = FormBody.create(MediaType.parse(JSON_TYPE), jsonParam);
        Request.Builder requestBuilder = new Request.Builder().url(url).post(requestBody);
        //添加请求头
        if (headers != null) {
            headers.forEach(requestBuilder::addHeader);
        }
        Request request = requestBuilder.build();
        try (Response response = langCallHttpClient.newCall(request).execute()) {
            if (response != null && response.code() == 200 && response.body() != null) {
                result = response.body().string();
            } else {
                log.error("HttpUtils -> post error, url={}, param={}, header={}, response={}", url, jsonParam, ObjectMappers.writeAsJsonStrThrow(headers), response == null ? "" : response.toString());
            }
        } catch (Exception e) {
            log.error("HttpUtils -> post error, url={}, json={}, header={}, error={}", url, jsonParam, ObjectMappers.writeAsJsonStrThrow(headers), e.getMessage(), e);
        }
        return result;
    }


    public String postOnlyHeader(String url, Map<String, String> headers) {
        String result = null;
        //请求参数
        RequestBody requestBody = new FormBody.Builder().build();
        Request.Builder requestBuilder = new Request.Builder().url(url).post(requestBody);
        //添加请求头
        if (headers != null) {
            headers.forEach(requestBuilder::addHeader);
        }
        Request request = requestBuilder.build();
        try (Response response = client.newCall(request).execute()) {
            if (response != null && response.code() == 200 && response.body() != null) {
                result = response.body().string();
            } else {
                log.error("HttpUtils -> postOnlyHeader error, url={}, param={}, header={}, response={}", url, "", ObjectMappers.writeAsJsonStrThrow(headers), response == null ? "" : response.toString());
            }
        } catch (Exception e) {
            log.error("HttpUtils -> postOnlyHeader error, url={}, json={}, header={}, error={}", url, "", ObjectMappers.writeAsJsonStrThrow(headers), e.getMessage(), e);
        }
        return result;
    }


    public String post(String url, String jsonParam) {
        String result = null;
        //请求参数
        RequestBody requestBody = FormBody.create(MediaType.parse(JSON_TYPE), jsonParam);
        Request.Builder requestBuilder = new Request.Builder().url(url).post(requestBody);
        Request request = requestBuilder.build();
        try (Response response = client.newCall(request).execute()) {
            if (response != null && response.code() == 200 && response.body() != null) {
                result = response.body().string();
            } else {
                log.error("HttpUtils -> post error, url={}, param={}, response={}", url, jsonParam, response == null ? "" : response.toString());
            }
        } catch (Exception e) {
            log.error("HttpUtils -> post error, url={}, json={}, error={}", url, jsonParam, e.getMessage(), e);
        }
        return result;
    }

    /**
     * post 请求 参数拼接在url后
     *
     * @param url
     * @param paramMap
     * @param headers
     * @return
     */

    public String post(String url, Map<String, String> paramMap, Map<String, String> headers) {
        String result = "";
        StringBuilder urlString = new StringBuilder(url);
        if (!CollectionUtils.isEmpty(paramMap)) {
            urlString.append("?");
            for (Map.Entry<String, String> entry : paramMap.entrySet()) {
                try {
                    urlString.append(URLEncoder.encode(null == entry.getKey() ? "" : entry.getKey(), "utf-8")).append("=").append(URLEncoder.encode(null == entry.getValue() ? "" : entry.getValue(), "utf-8")).append("&");
                } catch (UnsupportedEncodingException e) {
                    log.error("Parse url failed", e);
                }
            }
            urlString.deleteCharAt(urlString.lastIndexOf("&"));
        }
        Request.Builder requestBuilder = new Request.Builder().url(urlString.toString());
        if (headers != null) {
            headers.forEach(requestBuilder::addHeader);
        }
        //post 请求 必须有body
        RequestBody requestBody = FormBody.create(MediaType.parse(JSON_TYPE), "{}");
        try (Response response = client.newCall(requestBuilder.post(requestBody).build()).execute()) {
            if (response != null && response.code() == 200 && response.body() != null) {
                result = response.body().string();
            } else {
                log.error("HttpUtils -> post error, url={}, param={}, header={}, response={}", url, null, ObjectMappers.writeAsJsonStrThrow(headers), response == null ? "" : response.toString());
            }
        } catch (Exception e) {
            log.error("HttpUtils -> post error, url={}, json={}, header={}, error={}", url, null, ObjectMappers.writeAsJsonStrThrow(headers), e.getMessage(), e);
        }
        return result;
    }


    public String get(String url) {
        String result = "";
        try (Response response = client.newCall(new Request.Builder().url(url).build()).execute()) {
            if (response != null && response.code() == 200 && response.body() != null) {
                result = response.body().string();
            } else {
                log.error("HttpUtils -> get error, url={}, response={}", url, response == null ? "" : response.toString());
            }
        } catch (Exception e) {
            log.error("HttpUtils -> get error, url={}, error={}", url, e.getMessage(), e);
        }
        return result;
    }


    public String get(String url, Dictionary<String, String> headers) {
        String result = "";
        Request.Builder build = new Request.Builder().url(url);
        //添加请求头
        if (headers != null) {
            for (Enumeration<?> keys = headers.keys(); keys.hasMoreElements(); ) {
                Object key = keys.nextElement();
                Object value = headers.get(key);
                build.addHeader(key.toString(), value.toString());
            }
        }
        try (Response response = client.newCall(build.build()).execute()) {
            if (response != null && response.code() == 200 && response.body() != null) {
                result = response.body().string();
            } else {
                log.error("HttpUtils -> get error, url={}, headers={}, response={}", url, ObjectMappers.writeAsJsonStrThrow(headers), response == null ? "" : response.toString());
            }
        } catch (Exception e) {
            log.error("HttpUtils -> get error, url={}, headers={}, error={}", url, ObjectMappers.writeAsJsonStrThrow(headers), e.getMessage(), e);
        }
        return result;
    }


    public String get(String url, Map<String, String> paramMap, Map<String, String> headers) {
        String result = "";
        StringBuilder urlString = new StringBuilder(url);
        if (!CollectionUtils.isEmpty(paramMap)) {
            urlString.append("?");
            for (Map.Entry<String, String> entry : paramMap.entrySet()) {
                try {
                    urlString.append(URLEncoder.encode(null == entry.getKey() ? "" : entry.getKey(), "utf-8")).append("=").append(URLEncoder.encode(null == entry.getValue() ? "" : entry.getValue(), "utf-8")).append("&");
                } catch (UnsupportedEncodingException e) {
                    log.error("Parse url failed", e);
                }
            }
            urlString.deleteCharAt(urlString.lastIndexOf("&"));
        }
        Request.Builder requestBuilder = new Request.Builder().url(urlString.toString());
        if (headers != null) {
            headers.forEach(requestBuilder::addHeader);
        }
        try (Response response = client.newCall(requestBuilder.build()).execute()) {
            if (response != null && response.code() == 200 && response.body() != null) {
                result = response.body().string();
            } else {
                log.error("HttpUtils -> get error, url={}, header={}, response={}", urlString, ObjectMappers.writeAsJsonStrThrow(headers), response == null ? "" : response.toString());
            }
        } catch (Exception e) {
            log.error("HttpUtils -> get error, url={}, header={}, error={}", urlString, ObjectMappers.writeAsJsonStrThrow(headers), e.getMessage(), e);
        }
        return result;
    }

    public ImmutablePair<String, byte[]> getBytes(String url, Map<String, String> paramMap, Map<String, String> headers) {
        long start = System.currentTimeMillis();
        StringBuffer urlString = new StringBuffer(url);
        if (!CollectionUtils.isEmpty(paramMap)) {
            urlString.append("?");
            for (Map.Entry<String, String> entry : paramMap.entrySet()) {
                urlString.append(entry.getKey()).append("=").append(entry.getValue()).append("&");
            }
            urlString.deleteCharAt(urlString.lastIndexOf("&"));
        }
        Request.Builder requestBuilder = new Request.Builder().url(urlString.toString());
        if (headers != null) {
            headers.forEach(requestBuilder::addHeader);
        }

        try (Response response = client.newCall(requestBuilder.build()).execute()) {
            if (response != null && response.code() == 200 && response.body() != null) {
                byte[] bytes = response.body().bytes();
                int contentLength = Objects.isNull(bytes) ? -1 : bytes.length;
                log.info("remoteAPI HttpUtils getBytes request:{}, response length:{}, use time {} ms", urlString, contentLength, System.currentTimeMillis() - start);
                return ImmutablePair.of(response.header("attachFilename"), bytes);
            } else {
                log.error("HttpUtils -> getBytes error, url={}, headers={}, response={}", urlString, ObjectMappers.writeAsJsonStrThrow(headers), response == null ? "" : response.toString());
            }
        } catch (Exception e) {
            log.error("HttpUtils -> getBytes error, url={}, headers={}, error={}", urlString, ObjectMappers.writeAsJsonStrThrow(headers), e.getMessage(), e);
        }
        return null;
    }


    public String postByteArray(String url, byte[] data) {
        String result = "";
        Request.Builder requestBuilder = new Request.Builder()
                .url(url)
                .post(RequestBody.create(MediaType.parse("application/octet-stream;;charset=UTF-8"), data));
        try (Response response = client.newCall(requestBuilder.build()).execute()) {
            if (response != null && response.code() == 200 && response.body() != null) {
                result = response.body().string();
            } else {
                log.error("HttpUtils -> post error, url={}, response={}", url, response == null ? "" : response.toString());
            }
        } catch (Exception e) {
            log.error("HttpUtils -> post error, url={}, error={}", url, e.getMessage(), e);
        }
        return result;
    }


    public byte[] getBytes(String url, Map<String, String> param) {
        long time = System.currentTimeMillis();
        StringBuilder builder = new StringBuilder(url);
        if (param != null) {
            builder.append("?");
            for (Map.Entry<String, String> entry : param.entrySet()) {
                builder.append(entry.getKey()).append("=").append(entry.getValue()).append("&");
            }
            builder.deleteCharAt(builder.lastIndexOf("&"));
        }
        Request request = new Request.Builder().url(builder.toString()).build();
        try (Response response = client.newCall(request).execute()) {
            if (response != null && response.code() == 200 & response.body() != null) {
                byte[] result = response.body().bytes();
                int length = result == null ? -1 : result.length;
                log.info("remoteAPI HttpUtils getBytes request:{}, response length:{}, use time {} ms", builder.toString(), length, System.currentTimeMillis() - time);
                return result;
            } else {
                log.error("HttpUtils -> getBytes error, url={}, response={}", builder.toString(), response == null ? "" : response.toString());
            }
        } catch (IOException e) {
            log.error("HttpUtils -> get error, url={}, error={}", url, e.getMessage(), e);
        }
        return new byte[0];
    }


    public String postForm(String url, Map<String, String> params) {
        long time = System.currentTimeMillis();
        if (CollUtil.isEmpty(params)) {
            return "";
        }
        FormBody.Builder builder = new FormBody.Builder();
        params.keySet().forEach(it -> builder.add(it, params.get(it)));
        FormBody formBody = builder.build();

        Request req = new Request.Builder().url(url).post(formBody).build();
        try (Response response = client.newCall(req).execute()) {
            if (response.isSuccessful() && response.body() != null) {
                String res = response.body().string();
                int length = res.length();
                log.info("remoteAPI HttpUtils getBytes request:{}, response length:{}, use time {} ms", builder, length, System.currentTimeMillis() - time);
                return res;
            } else {
                log.error("HttpUtils -> getBytes error, url={}, response={}", builder, response);
            }
        } catch (Exception e) {
            log.error("HttpUtils -> get error, url={}, error={}", url, e.getMessage(), e);
        }
        return "";
    }


    public String postMultipartBody(String url, byte[] bytes, String fileName, Map<String, String> headers) {
        long time = System.currentTimeMillis();
        if (bytes == null) {
            return "";
        }
        try {
            MultipartBody.Builder builder = new MultipartBody.Builder();
            MultipartBody multipartBody = builder.setType(MultipartBody.FORM)
                    .addFormDataPart("file",
                            fileName,
                            RequestBody.create(MediaType.parse("application/octet-stream"), bytes)).build();
            Request.Builder reqBuilder = new Request.Builder();
            if (CollUtil.isNotEmpty(headers)) {
                headers.forEach(reqBuilder::addHeader);
            }
            Request req = reqBuilder.url(url).post(multipartBody).build();
            try (Response response = client.newCall(req).execute()) {
                if (response.isSuccessful() && response.body() != null) {
                    String res = response.body().string();
                    int length = res.length();
                    log.info("remoteAPI HttpUtils getBytes request:{}, response length:{}, use time {} ms", builder, length, System.currentTimeMillis() - time);
                    return res;
                } else {
                    log.error("HttpUtils -> getBytes error, url={}, response={}", builder, response);
                }
            } catch (Exception e) {
                log.error("HttpUtils -> get error, url={}, error={}", url, e.getMessage(), e);
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return "";
    }


    public String postMultipartBody(String url, Map<String, String> formData, byte[] bytes, String fileName, Map<String, String> headers) {
        long time = System.currentTimeMillis();
        if (bytes == null) {
            return "";
        }
        try {
            MultipartBody.Builder builder = new MultipartBody.Builder();
            builder.setType(MultipartBody.FORM);
            builder.addFormDataPart("file",
                    fileName,
                    RequestBody.create(MediaType.parse("application/octet-stream"), bytes));
            if (CollUtil.isNotEmpty(formData)) {
                formData.forEach(builder::addFormDataPart);
            }
            MultipartBody multipartBody = builder.build();
            Request.Builder reqBuilder = new Request.Builder();
            if (CollUtil.isNotEmpty(headers)) {
                headers.forEach(reqBuilder::addHeader);
            }
            Request req = reqBuilder.url(url).post(multipartBody).build();
            try (Response response = client.newCall(req).execute()) {
                if (response.isSuccessful() && response.body() != null) {
                    String res = response.body().string();
                    int length = res.length();
                    log.info("remoteAPI HttpUtils getBytes request:{}, response length:{}, use time {} ms", builder, length, System.currentTimeMillis() - time);
                    return res;
                } else {
                    log.error("HttpUtils -> getBytes error, url={}, response={}", builder, response);
                }
            } catch (Exception e) {
                log.error("HttpUtils -> get error, url={}, error={}", url, e.getMessage(), e);
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return "";
    }

    public byte[] postFormGetByte(String url, Map<String, String> params, Map<String, String> headers) {
        long time = System.currentTimeMillis();
        if (CollUtil.isEmpty(params)) {
            return null;
        }
        FormBody.Builder builder = new FormBody.Builder();
        params.keySet().forEach(it -> builder.add(it, params.get(it)));
        FormBody formBody = builder.build();

        Request.Builder post = new Request.Builder().url(url).post(formBody);
        if (headers != null) {
            headers.forEach(post::addHeader);
        }
        Request request = post.build();
        try (Response response = client.newCall(request).execute()) {
            if (response.isSuccessful() && response.body() != null) {
                byte[] res = response.body().bytes();
                int length = res.length;
                log.info("remoteAPI HttpUtils postFormGetByte request:{}, response length:{}, use time {} ms", builder, length, System.currentTimeMillis() - time);
                return res;
            } else {
                log.error("HttpUtils -> postFormGetByte error, url={}, response={}", builder, response);
            }
        } catch (Exception e) {
            log.error("HttpUtils -> get error, url={}, error={}", url, e.getMessage(), e);
        }
        return null;
    }


    public byte[] postBytes(String url, String jsonParam, Map<String, String> headers) {
        byte[] result = null;
        //请求参数
        RequestBody requestBody = FormBody.create(MediaType.parse(JSON_TYPE), jsonParam);
        Request.Builder requestBuilder = new Request.Builder().url(url).post(requestBody);
        //添加请求头
        if (headers != null) {
            headers.forEach(requestBuilder::addHeader);
        }
        Request request = requestBuilder.build();
        try (Response response = client.newCall(request).execute()) {
            if (response.code() == 200 && response.body() != null) {
                result = response.body().bytes();
            } else {
                log.error("HttpUtils -> post error, url={}, param={}, header={}, response={}", url, jsonParam, ObjectMappers.writeAsJsonStrThrow(headers), response);
            }
        } catch (Exception e) {
            log.error("HttpUtils -> post error, url={}, json={}, header={}, error={}", url, jsonParam, ObjectMappers.writeAsJsonStrThrow(headers), e.getMessage(), e);
        }
        return result;
    }
}









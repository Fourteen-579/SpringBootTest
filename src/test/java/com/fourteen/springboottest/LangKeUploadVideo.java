package com.fourteen.springboottest;

import org.springframework.boot.test.context.SpringBootTest;

import javax.crypto.Mac;
import javax.crypto.spec.SecretKeySpec;
import java.util.Base64;
import java.util.Random;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2024/6/24 10:56
 */

@SpringBootTest
public class LangKeUploadVideo {
/*    public static void main(String[] args) {
        VodUploadClient client = new VodUploadClient("AKIDUjcCYIu57gj4x3BehrLuQJyO1I4nGFsglCCGPTlTuuRLPTIF4fPTnU_mvuBE0bmi",
                "o0Qn/a/Om51OLxg6oqI6s5n9ny4AqisyawB623e8GQg=",
                "maIMBfQS9VWkPl2wUJaecX3opG3SgYia6d0c30c7a36ad92bb8d01db723da12bcvz0ynV7dE8rAEVDwqD9srb8gTbtQRpmH2e5Ik_C15lD2XTtwreErYxGkEmFvfFJG3AvjZo4NUZ5J7HRMHtVjTP0EEYkFNJ-l_uLdhyaZyvZe5FJO-2hSn-BBor5c6yIsXeVUTraG11ResYTiIf5olZRpn4KgFW5illHUOHvrd-6VvKfSr37PfuTZUSJbtOxJaQaLWfxC4ZhGMnrSiPps30uYUIDx8qlRY2jwU6HZwsUCBget-k88EZhuaNmutJGOLdbpZpY2TEJxyUa7ikITkBFQDrFIdUuNw7V3qvJmM9I-y2q9N3lFiF8ja7QyyOYvSOvuOOKY_Rvge6X4HFc-zQ");
        VodUploadRequest request = new VodUploadRequest();
        request.setSubAppId(1500019286L);
        request.setSourceContext("qJh--LVhvFebSUC4QwwGltir7VsP7LC5eM0__xRqdPHu8THcK8yWYfYwROKqxOZbz4Ka3XKyRHYSXvn1lau36FHs0ZA9blwEftiHa4l-ovOX1RtzwzvRzCygAamjlt3kzacQSjrrr2TD8H6NTQyYKybd3UrZYNdEri1lYXGttqSWNJ_ZORzSguZcF0JKw3jbB1XSHTbYTHCQ36NiENPtBggi2ejgW-r59k13bl-rAqsJ0z8q3TKO5WadYJwt0Lm6");
//        request.setCoverFilePath("D:\\Downloads\\导出专用.jpg");
        request.setMediaFilePath("D:\\Downloads\\test.mp4");
        request.setProcedure("素材库测试转码");
        try {
            VodUploadResponse response = client.upload("ap-guangzhou", request);
            System.out.println(response.getFileId());
            System.out.println(response.getCoverUrl());
            System.out.println(response.getMediaUrl());
            System.out.println(response.getRequestId());
        } catch (Exception e) {
            e.printStackTrace();
            // 业务方进行异常处理
        }
    }*/

    public static void main(String[] args) {
        Signature sign = new Signature();
        // 设置 App 的云 API 密钥
        sign.setSecretId("SecretId");
        sign.setSecretKey("SecretKey");
        sign.setCurrentTime(System.currentTimeMillis() / 1000);
        sign.setRandom(new Random().nextInt(java.lang.Integer.MAX_VALUE));
        sign.setSignValidDuration(3600 * 24 * 2); // 签名有效期：2天


        try {
            String signature = sign.getUploadSignature();
            System.out.println("signature : " + signature);
        } catch (Exception e) {
            System.out.print("获取签名失败");
            e.printStackTrace();
        }
    }


}

// 签名工具类
class Signature {
    private static final String HMAC_ALGORITHM = "HmacSHA1"; //签名算法
    private static final String CONTENT_CHARSET = "UTF-8";
    private String secretId;
    private String secretKey;
    private long currentTime;
    private int random;
    private int signValidDuration;

    public static byte[] byteMerger(byte[] byte1, byte[] byte2) {
        byte[] byte3 = new byte[byte1.length + byte2.length];
        System.arraycopy(byte1, 0, byte3, 0, byte1.length);
        System.arraycopy(byte2, 0, byte3, byte1.length, byte2.length);
        return byte3;
    }


    // 获取签名
    public String getUploadSignature() throws Exception {
        String strSign = "";
        String contextStr = "";


        // 生成原始参数字符串
        long endTime = (currentTime + signValidDuration);
        contextStr += "secretId=" + java.net.URLEncoder.encode(secretId, "utf8");
        contextStr += "&currentTimeStamp=" + currentTime;
        contextStr += "&expireTime=" + endTime;
        contextStr += "&random=" + random;


        try {
            Mac mac = Mac.getInstance(HMAC_ALGORITHM);
            SecretKeySpec secretKey = new SecretKeySpec(this.secretKey.getBytes(CONTENT_CHARSET), mac.getAlgorithm());
            mac.init(secretKey);


            byte[] hash = mac.doFinal(contextStr.getBytes(CONTENT_CHARSET));
            byte[] sigBuf = byteMerger(hash, contextStr.getBytes("utf8"));
            strSign = base64Encode(sigBuf);
            strSign = strSign.replace(" ", "").replace("\n", "").replace("\r", "");
        } catch (Exception e) {
            throw e;
        }
        return strSign;
    }


    private String base64Encode(byte[] buffer) {
        return Base64.getEncoder().encodeToString(buffer);
    }


    public void setSecretId(String secretId) {
        this.secretId = secretId;
    }


    public void setSecretKey(String secretKey) {
        this.secretKey = secretKey;
    }


    public void setCurrentTime(long currentTime) {
        this.currentTime = currentTime;
    }


    public void setRandom(int random) {
        this.random = random;
    }


    public void setSignValidDuration(int signValidDuration) {
        this.signValidDuration = signValidDuration;
    }
}




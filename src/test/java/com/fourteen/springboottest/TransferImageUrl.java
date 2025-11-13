package com.fourteen.springboottest;

import org.apache.commons.io.IOUtils;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.net.URL;
import java.util.Base64;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2024/5/16 10:49
 */
@SpringBootTest
public class TransferImageUrl {

    @Test
    public void test() throws Exception {
        String url = "https://mmbiz.qpic.cn/mmbiz_jpg/Ek8t172PVzNKxicNcDHuE2Q6MbQ1haicBLbEXJKmibaRX6HBsmvez5xUQGhb7oJOyiaIFjWVVWWA2kPgkJaTKdu80Q/0?wx_fmt=jpeg";
        System.out.println(Base64.getEncoder().encodeToString(IOUtils.toByteArray(new URL(url))));
    }

}

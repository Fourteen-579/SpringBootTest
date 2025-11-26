package com.fourteen.springboottest.client;

import com.fourteen.springboottest.bo.qyt.MessageReq;
import com.fourteen.springboottest.util.HttpUtils;
import com.fourteen.springboottest.util.ObjectMappers;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.Dictionary;
import java.util.Hashtable;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/11/12 16:24
 */
@Slf4j
@Component
@RequiredArgsConstructor
public class DongMessageClient {

    private final static String url = "http://10.205.214.16/api/message/push";
    private final static String token = "cHQeSKowHJohJ2NgZKSxkYorZ9wzAi";

    private final HttpUtils httpUtils;

    public final void sendDongMessage(String title, String content, String receiver) {
        MessageReq messageReq = new MessageReq(title, receiver, content, "group");

        String jsonParam = ObjectMappers.writeAsJsonStrThrow(messageReq);

        Dictionary<String, String> headers = new Hashtable<>();
        headers.put("token", token);

        httpUtils.post(url, jsonParam, headers);
    }

}

package com.fourteen.springboottest.controller;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.mvc.method.annotation.SseEmitter;

import java.util.concurrent.Executors;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/6/16 10:28
 */
@RestController
public class SseEmitterController {

    @GetMapping("/start")
    public SseEmitter startTask() {
        // 设置超时时间为30秒
        SseEmitter emitter = new SseEmitter(30_000L);

        // 异步执行任务，防止阻塞 Servlet 线程
        Executors.newSingleThreadExecutor().submit(() -> {
            try {
                for (int i = 1; i <= 10; i++) {
                    // 模拟任务进度
                    String progress = "任务进度: " + (i * 10) + "%";

                    // 发送数据到客户端
                    emitter.send(SseEmitter.event()
                            .name("task-progress") // 自定义事件名
                            .data(progress)
                            .id(String.valueOf(i))
                            .reconnectTime(3000)); // 客户端断开后尝试重新连接间隔

                    Thread.sleep(1000); // 模拟耗时操作
                }

                // 任务完成后通知客户端
                emitter.send(SseEmitter.event()
                        .name("task-complete")
                        .data("任务完成！"));

                // 标记完成
                emitter.complete();
            } catch (Exception e) {
                emitter.completeWithError(e);
            }
        });

        return emitter;
    }

}

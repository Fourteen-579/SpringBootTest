package com.fourteen.springboottest.controller;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.read.listener.PageReadListener;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.scheduling.support.CronExpression;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.net.URLEncoder;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/9/11 10:17
 */
@Slf4j
@RestController
@RequiredArgsConstructor
public class CronCheckController {

    /**
     * 判断表达式在时间段内是否会执行
     */
    public static String willRunInRange(String cron, LocalDateTime start, LocalDateTime end) {
        try {
            CronExpression cronExpression = CronExpression.parse(cron);
            LocalDateTime next = cronExpression.next(start);
            return next != null && !next.isAfter(end) ? "1" : "0";
        } catch (Exception e) {
            System.err.println("无效的cron表达式: " + cron);
            log.error("error:", e);
            return "无效";
        }
    }

    @PostMapping("/cron-check")
    public void getSuffix(@RequestParam("file") MultipartFile file, HttpServletResponse response) {

        List<CronCheck> cronList = new ArrayList<>();

        try {
            EasyExcel.read(file.getInputStream(),
                            CronCheck.class,
                            new PageReadListener<CronCheck>(cronList::addAll))
                    .sheet(0)
                    .doRead();

            if (CollUtil.isEmpty(cronList)) {
                log.error("文件内容为空");
            }

            // 检查区间 2024-09-14 16:00 到 18:00
            LocalDateTime start = LocalDateTime.of(2024, 9, 14, 16, 0);
            LocalDateTime end = LocalDateTime.of(2024, 9, 14, 18, 0);

            for (CronCheck cronEntity : cronList) {
                String cron = cronEntity.getCron();
                String willRun = willRunInRange(cron, start, end);
                cronEntity.setEnable(willRun);
            }

            String filename = URLEncoder.encode("cron.xlsx", "utf8");
            response.addHeader("Content-Type", "application/octet-stream");
            response.addHeader("Content-Disposition", "attachment;filename=" + filename);

            EasyExcel.write(response.getOutputStream(), CronCheck.class)
                    .sheet("Sheet1")
                    .doWrite(cronList);

        } catch (Exception e) {
            log.error("error:", e);
        }
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class CronCheck {

        @ExcelProperty(value = "任务名")
        private String name;

        @ExcelProperty(value = "cron")
        private String cron;

        @ExcelProperty(value = "是否会在9月14日16点-18点执行")
        private String enable;

        @ExcelProperty(value = "是否有影响")
        private Boolean hasImpact;

        @ExcelProperty(value = "补全操作")
        private String operation;
    }

}

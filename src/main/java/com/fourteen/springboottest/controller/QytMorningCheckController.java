package com.fourteen.springboottest.controller;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fourteen.springboottest.bo.qyt.DayActiveListResp;
import com.fourteen.springboottest.bo.qyt.PageResVO;
import com.fourteen.springboottest.bo.qyt.Response;
import com.fourteen.springboottest.bo.qyt.UserActiveRequest;
import com.fourteen.springboottest.bo.sw.ServiceLoadResponse;
import com.fourteen.springboottest.bo.sw.SkywalkingMetricRequest;
import com.fourteen.springboottest.bo.sw.TimePercentileResponse;
import com.fourteen.springboottest.util.HttpUtils;
import com.fourteen.springboottest.util.ObjectMappers;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import javax.annotation.Resource;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.HashMap;
import java.util.IntSummaryStatistics;
import java.util.Map;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/11/12 9:47
 */
@Slf4j
@RestController
@RequiredArgsConstructor
public class QytMorningCheckController {

    private final static DateTimeFormatter DATE_TIME_FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd HHmm");
    private final static DateTimeFormatter DATE_TIME_FORMATTER2 = DateTimeFormatter.ofPattern("yyyy-MM-dd");
    private final static String TEMPLATE = "%s企业通系统巡检结果：\n" +
            "1、【正常】接口耗时分布，P90：%sms；P95：%sms；p99：%sms\n" +
            "2、【正常】接口请求量，总数：%s；最高QPS：%s\n" +
            "3、【正常】接口返回结果：正常\n" +
            "4、【正常】应用日志报错查看：正常\n" +
            "5、【正常】应用服务器资源指标，cpu使用率峰值：{}%%；内存使用率峰值：{}%%；磁盘使用率：{}%%\n" +
            "6、【正常】定时任务&数仓数据同步任务：正常执行\n" +
            "7、【正常】日活账号数，企业数：%s；账号数：%s\n" +
            "晨检人：%s";


    @Resource
    private HttpUtils httpUtils;

    @GetMapping("/morning-check")
    public String getQytMorningCheck(@RequestParam("session") String session,
                                     @RequestParam("checkDate") String checkDateStr,
                                     @RequestParam("checkUser") String checkUser) {

        LocalDate checkDate = LocalDate.parse(checkDateStr, DATE_TIME_FORMATTER2);

        //获取接口耗时分布
        TimePercentileResponse timePercentileResponse = getTimePercentileResponse(session, checkDateStr, checkDateStr);
        TimePercentileResponse.ServicePercentile p99 = timePercentileResponse.getData().getService_percentile0().get(4);
        TimePercentileResponse.ServicePercentile p95 = timePercentileResponse.getData().getService_percentile0().get(3);
        TimePercentileResponse.ServicePercentile p90 = timePercentileResponse.getData().getService_percentile0().get(2);
        int p99Ms = (int) p99.getValues().getValues().get(0).getValue();
        int p95Ms = (int) p95.getValues().getValues().get(0).getValue();
        int p90Ms = (int) p90.getValues().getValues().get(0).getValue();

        //获取接口请求量
        ServiceLoadResponse serviceLoad = getServiceLoad(session, checkDateStr+" 12", checkDateStr+" 18");
        IntSummaryStatistics statistics = serviceLoad.getData().getService_cpm0().getValues().getValues().stream()
                .map(ServiceLoadResponse.ValueItem::getValue)
                .mapToInt(Double::intValue)
                .summaryStatistics();
        int totalCount = (int) statistics.getSum();
        int maxQps = statistics.getMax();

        //获取日活账号数
        DayActiveListResp dayActiveList = getDayActiveList(checkDateStr, checkDateStr);
        Integer entCount = dayActiveList.getEntCount();
        Integer userCount = dayActiveList.getUserCount();

        return String.format(TEMPLATE, checkDateStr, p90Ms, p95Ms, p99Ms, totalCount, maxQps, entCount, userCount, checkUser);
    }

    private DayActiveListResp getDayActiveList(String start, String end) {
        String url = "https://qiyetong.eastmoney.com/ent_backend/activeDataBoard/dayActive/list";
        UserActiveRequest request = new UserActiveRequest();
        request.setBeginTime(start);
        request.setEndTime(end);
        request.setStatisticType(2);
        request.setSize(10);
        request.setCurrent(1);

        Map<String, String> params = new HashMap<>();
        params.put("token", "1762304382108GrFO9DSXcqh84mWnx-28ZhJidva5Mmi0B4M9ptvL9K0D39OcjeY");

        String result = httpUtils.post(url, ObjectMappers.writeAsJsonStrThrow(request), params);
        if (StringUtils.isNotBlank(result)) {
            Response<PageResVO<DayActiveListResp>> response = ObjectMappers.readAsObjThrow(result, new TypeReference<Response<PageResVO<DayActiveListResp>>>() {
            });
            if (response.getCode() == 200 && response.getData() != null && response.getData().getRecords() != null && response.getData().getRecords().size() > 0) {
                return response.getData().getRecords().get(0);
            }
        }

        return null;
    }

    private ServiceLoadResponse getServiceLoad(String session, String start, String end) {
        SkywalkingMetricRequest request = buildServiceLoadRequest(start, end);
        String result = getResult(request, session);
        if (StringUtils.isNotBlank(result)) {
            return ObjectMappers.readAsObjThrow(result, ServiceLoadResponse.class);
        }
        return null;
    }

    private SkywalkingMetricRequest buildServiceLoadRequest(String start, String end) {
        SkywalkingMetricRequest request = new SkywalkingMetricRequest();
        request.setQuery("query queryData($duration: Duration!,$condition0: MetricsCondition!) {service_cpm0: readMetricsValues(condition: $condition0, duration: $duration){\n    label\n    values {\n      values {value}\n    }\n  }}");

        // 组装 variables
        SkywalkingMetricRequest.Variables vars = new SkywalkingMetricRequest.Variables();
        SkywalkingMetricRequest.Duration duration = new SkywalkingMetricRequest.Duration(start, end, "HOUR");
        vars.setDuration(duration);

        SkywalkingMetricRequest.Condition condition0 = new SkywalkingMetricRequest.Condition();
        condition0.setName("service_cpm");

        SkywalkingMetricRequest.Entity entity = new SkywalkingMetricRequest.Entity();
        entity.setScope("Service");
        entity.setServiceName("qiyetong_webserver");
        entity.setNormal(true);

        condition0.setEntity(entity);
        vars.setCondition0(condition0);

        request.setVariables(vars);
        return request;
    }

    private TimePercentileResponse getTimePercentileResponse(String session, String start, String end) {
        SkywalkingMetricRequest request = buildTimePercentileRequest(start, end);
        String result = getResult(request, session);
        if (StringUtils.isNotBlank(result)) {
            return ObjectMappers.readAsObjThrow(result, TimePercentileResponse.class);
        }
        return null;
    }

    private String getResult(SkywalkingMetricRequest request, String session) {
        String url = "http://emcd-skywalking.bdeastmoney.net/graphql";
        Map<String, String> map = new HashMap<>();
        map.put("cookie", "SESSION=" + session);

        String jsonParm = ObjectMappers.writeAsJsonStrThrow(request);
        return httpUtils.post(url, jsonParm, map);
    }

    private SkywalkingMetricRequest buildTimePercentileRequest(String start, String end) {
        SkywalkingMetricRequest request = new SkywalkingMetricRequest();
        request.setQuery("query queryData($duration: Duration!,$labels0: [String!]!,$condition0: MetricsCondition!) " +
                "{service_percentile0: readLabeledMetricsValues(condition: $condition0, labels: $labels0, duration: $duration)" +
                "{ label values { values {value} } } }");

        // 组装 variables
        SkywalkingMetricRequest.Variables vars = new SkywalkingMetricRequest.Variables();
        SkywalkingMetricRequest.Duration duration = new SkywalkingMetricRequest.Duration(start, end, "DAY");
        vars.setDuration(duration);
        vars.setLabels0(Arrays.asList("0", "1", "2", "3", "4"));

        SkywalkingMetricRequest.Condition condition0 = new SkywalkingMetricRequest.Condition();
        condition0.setName("service_percentile");

        SkywalkingMetricRequest.Entity entity = new SkywalkingMetricRequest.Entity();
        entity.setScope("Service");
        entity.setServiceName("qiyetong_webserver");
        entity.setNormal(true);

        condition0.setEntity(entity);
        vars.setCondition0(condition0);

        request.setVariables(vars);
        return request;
    }

}

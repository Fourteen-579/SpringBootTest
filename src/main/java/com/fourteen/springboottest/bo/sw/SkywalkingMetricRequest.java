package com.fourteen.springboottest.bo.sw;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/11/12 10:49
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class SkywalkingMetricRequest {

    private String query;
    private Variables variables;

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class Variables {
        private Duration duration;
        private List<String> labels0;
        private Condition condition0;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class Duration {
        private String start;
        private String end;
        private String step;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class Condition {
        private String name;
        private Entity entity;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class Entity {
        private String scope;
        private String serviceName;
        private boolean normal;
    }
}

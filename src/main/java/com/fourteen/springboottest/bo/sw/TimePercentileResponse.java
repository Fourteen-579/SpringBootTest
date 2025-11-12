package com.fourteen.springboottest.bo.sw;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/11/12 10:50
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class TimePercentileResponse {

    private DataBody data;

    @Data
    public static class DataBody {
        private List<ServicePercentile> service_percentile0;
    }

    @Data
    public static class ServicePercentile {
        private String label;
        private MetricValues values;
    }

    @Data
    public static class MetricValues {
        private List<ValueWrapper> values;
    }

    @Data
    public static class ValueWrapper {
        private double value;
    }
}

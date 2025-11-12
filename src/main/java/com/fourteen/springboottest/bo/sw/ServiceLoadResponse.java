package com.fourteen.springboottest.bo.sw;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/11/12 11:21
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ServiceLoadResponse {

    private DataContainer data;

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class DataContainer {
        private ServiceCpm service_cpm0;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class ServiceCpm {
        private String label;
        private ValueContainer values;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class ValueContainer {
        private List<ValueItem> values;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class ValueItem {
        private double value;
    }

}

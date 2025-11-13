package com.fourteen.springboottest.bo.zabbix;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/11/13 10:02
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ZabbixRequest {

    private String jsonrpc;
    private String method;
    private Params params;
    private String auth;
    private int id;

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class Params {
        private String output;

        /**
         * 类型	history值
         * float	0
         * string	1
         * log	    2
         * integer	3
         */
        private int history;
        private List<String> itemids;
        private String sortfield;
        private String sortorder;
        private int limit;

        @JsonProperty("time_from")
        private long timeFrom;

        @JsonProperty("time_till")
        private long timeTill;
    }

}

package com.fourteen.springboottest.bo.zabbix;

import com.fasterxml.jackson.annotation.JsonIgnore;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.List;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/11/13 10:03
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ZabbixResponse {

    private String jsonrpc;
    private List<Result> result;
    private int id;

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class Result {
        private String itemid;
        private String clock;
        private String value;
        private String ns;

        @JsonIgnore
        public LocalDateTime getClockDateTime() {
            if (clock == null) return null;
            try {
                long epochSeconds = Long.parseLong(clock);
                return LocalDateTime.ofInstant(Instant.ofEpochSecond(epochSeconds), ZoneId.systemDefault());
            } catch (NumberFormatException e) {
                return null;
            }
        }
    }

}

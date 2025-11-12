package com.fourteen.springboottest.bo.qyt;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/11/12 11:30
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class UserActiveRequest {

    private String beginTime;
    private String endTime;
    private int current;
    private int size;
    private int statisticType;

}

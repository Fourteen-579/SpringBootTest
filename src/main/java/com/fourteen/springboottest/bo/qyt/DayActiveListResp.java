package com.fourteen.springboottest.bo.qyt;

import lombok.Data;

/**
 * @description:
 * @author: yuzhiying
 * @date: 2023-02-07 10:12
 **/
@Data
public class DayActiveListResp {

    private String id;

    private String dayTime;

    private Integer entCount;

    private Integer userCount;

}

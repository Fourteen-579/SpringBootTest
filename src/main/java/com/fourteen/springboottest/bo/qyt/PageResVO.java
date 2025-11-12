package com.fourteen.springboottest.bo.qyt;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.ArrayList;
import java.util.List;

/**
 * @author : chenyujiang
 * @version: 1.0
 * @date: 2022/08/22
 * @Description:
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class PageResVO<T> {

    private Long total;
    private Long size;
    private Long current;
    private Long pages;
    private List<T> records;

    public static <T> PageResVO<T> emptyPageRes() {
        PageResVO<T> res = new PageResVO<>();
        res.setRecords(new ArrayList<T>());
        return res;
    }

}

package com.fourteen.springboottest;

import cn.hutool.core.collection.CollUtil;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2024/5/23 9:28
 */
@SpringBootTest
public class PageTest {

    public static  <T> List<T> pagination(List<T> list, Integer pageNum, Integer pageSize) {
        if (CollUtil.isEmpty(list)) {
            return null;
        }

        if (pageNum == null || pageNum <= 0) {
            pageNum = 1;
        }

        if (pageSize == null || pageSize <= 0) {
            pageSize = 10;
        }

        int fromIndex = (pageNum - 1) * pageSize;
        if (fromIndex >= list.size()) {
            return new ArrayList<>();
        }

        int toIndex = Math.min(fromIndex + pageSize, list.size());
        return list.subList(fromIndex, toIndex);
    }

    public static void main(String[] args) {
        List<Integer> list = Arrays.asList(0,1,2,3,4,5,6,7,8,9,10,11,12,13,14);

        System.out.println(pagination(list,2,10));
    }

}

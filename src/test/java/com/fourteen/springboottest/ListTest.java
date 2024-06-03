package com.fourteen.springboottest;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2024/5/29 9:44
 */
public class ListTest {

    public static void main(String[] args) {
        List<Integer> list = null;

        List<Integer> collect = list.stream().map(t -> -t).collect(Collectors.toList());
        System.out.println(collect);
    }
}

package com.fourteen.springboottest;

import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2024/5/27 14:18
 */
public class NumberTest {

    public static void main(String[] args) {
        List<Integer> list = Arrays.asList(null,null);

        System.out.println(list.stream().filter(Objects::nonNull).collect(Collectors.averagingDouble(Integer::doubleValue)));
    }

}

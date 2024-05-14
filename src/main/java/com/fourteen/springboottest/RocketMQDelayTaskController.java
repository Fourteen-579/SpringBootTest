package com.fourteen.springboottest;

import apache.rocketmq.v2.Message;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.math3.stat.descriptive.rank.Percentile;
import org.apache.rocketmq.client.producer.DefaultMQProducer;
import org.apache.rocketmq.remoting.common.RemotingHelper;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import javax.annotation.Resource;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Random;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2023/4/13 11:10
 */
@RestController
@Slf4j
public class RocketMQDelayTaskController {

    @GetMapping("/rocketmq/add")
    public static BigDecimal calculatePercentile(List<BigDecimal> numberList, BigDecimal number) {
        /*
         * 计算某个数字在一个数字列表中的分位数
         *
         * 参数：
         * numberList: 数字列表
         * number: 需要计算分位数的数字
         *
         * 返回值：
         * 分位数，介于0到100之间的一个 BigDecimal 类型数值
         */

        // 首先将数字列表进行排序
        List<BigDecimal> sortedList = new ArrayList<>(numberList);
        Collections.sort(sortedList);

        // 计算数字在排序后的列表中的位置
        int position = sortedList.indexOf(number);

        // 计算分位数
        BigDecimal percentile = new BigDecimal(position)
                .divide(new BigDecimal(sortedList.size()), 4, BigDecimal.ROUND_HALF_UP)
                .multiply(new BigDecimal(100));

        return percentile;
    }

    public static double calculateMedian(List<Double> numberList) {
        /*
         * 计算 Double 类型的数字列表的中位数
         *
         * 参数：
         * numberList: 数字列表
         *
         * 返回值：
         * 中位数
         */

        int size = numberList.size();
        if (size == 0) {
            throw new IllegalArgumentException("List is empty");
        }

        int mid = size / 2;
        if (size % 2 == 0) {
            // 如果列表长度为偶数，取中间两个数的平均值
            double median1 = quickSelect(numberList, mid - 1);
            double median2 = quickSelect(numberList, mid);
            return (median1 + median2) / 2.0;
        } else {
            // 如果列表长度为奇数，取中间数
            return quickSelect(numberList, mid);
        }
    }

    private static double quickSelect(List<Double> numberList, int k) {
        /*
         * 使用 Quickselect 算法查找 Double 类型的数字列表中第 k 小的数
         *
         * 参数：
         * numberList: 数字列表
         * k: 目标位置，从 0 开始计数
         *
         * 返回值：
         * 第 k 小的数
         */

        int left = 0;
        int right = numberList.size() - 1;
        Random random = new Random();

        while (left <= right) {
            // 随机选择一个 pivot 元素
            int pivotIndex = random.nextInt(right - left + 1) + left;
            int newPivotIndex = partition(numberList, left, right, pivotIndex);

            if (newPivotIndex == k) {
                return numberList.get(newPivotIndex);
            } else if (newPivotIndex < k) {
                left = newPivotIndex + 1;
            } else {
                right = newPivotIndex - 1;
            }
        }

        throw new RuntimeException("Failed to find k-th element");
    }

    private static int partition(List<Double> numberList, int left, int right, int pivotIndex) {
        /*
         * 将 Double 类型的数字列表按照 pivot 元素分为两个子集，并返回 pivot 元素的新位置
         *
         * 参数：
         * numberList: 数字列表
         * left: 子集的左边界
         * right: 子集的右边界
         * pivotIndex: pivot 元素的位置
         *
         * 返回值：
         * pivot 元素的新位置
         */

        double pivotValue = numberList.get(pivotIndex);
        swap(numberList, pivotIndex, right);
        int storeIndex = left;

        for (int i = left; i < right; i++) {
            if (numberList.get(i) < pivotValue) {
                swap(numberList, i, storeIndex);
                storeIndex++;
            }
        }

        swap(numberList, storeIndex, right);
        return storeIndex;
    }

    private static void swap(List<Double> numberList, int i, int j) {
        /*
         * 交换 Double 类型的数字列表中两个元素的位置
         *
         * 参数：
         * numberList: 数字列表
         * i: 第一个元素的位置
         * j: 第二个元素的位置
         */

        double temp = numberList.get(i);
        numberList.set(i, numberList.get(j));
        numberList.set(j, temp);
    }
}

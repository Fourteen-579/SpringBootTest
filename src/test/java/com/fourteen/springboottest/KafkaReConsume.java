package com.fourteen.springboottest;

import org.apache.kafka.clients.consumer.*;
import org.apache.kafka.common.TopicPartition;

import java.time.Duration;
import java.util.*;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/3/1 14:34
 */
public class KafkaReConsume {

    public static void main(String[] args) {
        String topic = "txz_enterprise_certification_information_audit_results";
        Properties props = new Properties();
        props.put(ConsumerConfig.BOOTSTRAP_SERVERS_CONFIG, "10.205.215.93:9092");
        props.put(ConsumerConfig.GROUP_ID_CONFIG, "qyt_backend_consumer_my_audit_test");
        props.put(ConsumerConfig.KEY_DESERIALIZER_CLASS_CONFIG, "org.apache.kafka.common.serialization.StringDeserializer");
        props.put(ConsumerConfig.VALUE_DESERIALIZER_CLASS_CONFIG, "org.apache.kafka.common.serialization.StringDeserializer");
        props.put(ConsumerConfig.AUTO_OFFSET_RESET_CONFIG, "earliest");
        props.put(ConsumerConfig.ENABLE_AUTO_COMMIT_CONFIG, "false");

        KafkaConsumer<String, String> consumer = new KafkaConsumer<>(props);
        consumer.subscribe(Collections.singletonList(topic));

        // 先 poll 一次，确保分区被分配
        consumer.poll(Duration.ofMillis(1000));

        // 检查分区
        Set<TopicPartition> partitions = consumer.assignment();
        if (partitions.isEmpty()) {
            System.out.println("No partitions assigned, exiting...");
            return;
        }

        // 获取 Kafka 主题的最早和最晚偏移量
        for (TopicPartition partition : partitions) {
            long beginningOffset = consumer.beginningOffsets(Collections.singleton(partition)).get(partition);
            long endOffset = consumer.endOffsets(Collections.singleton(partition)).get(partition);
            System.out.printf("Partition %s: beginningOffset=%d, endOffset=%d%n", partition, beginningOffset, endOffset);
        }

        // 设定需要查找的时间戳（比如 1 小时前）
        long oneHourAgo = System.currentTimeMillis() - 3600 * 1000;
        Map<TopicPartition, Long> timestampsToSearch = new HashMap<>();
        for (TopicPartition partition : partitions) {
            timestampsToSearch.put(partition, oneHourAgo);
        }

        // 查询指定时间点的 offset
        Map<TopicPartition, OffsetAndTimestamp> offsets = consumer.offsetsForTimes(timestampsToSearch);
        for (TopicPartition partition : partitions) {
            OffsetAndTimestamp offsetAndTimestamp = offsets.get(partition);
            if (offsetAndTimestamp != null) {
                consumer.seek(partition, offsetAndTimestamp.offset());
                System.out.printf("Partition %s: seeking to offset %d%n", partition, offsetAndTimestamp.offset());
            } else {
                System.out.printf("Partition %s: No offset found for timestamp, seeking to earliest%n", partition);
                consumer.seekToBeginning(Collections.singleton(partition));
            }
        }

        // 消费数据
        int messageCount = 0;
        while (messageCount < 100) {
            ConsumerRecords<String, String> records = consumer.poll(Duration.ofMillis(100));
            if (records.isEmpty()) {
                System.out.println("No more messages, exiting...");
                break;
            }
            for (ConsumerRecord<String, String> record : records) {
                System.out.printf("Consumed message: key=%s, value=%s, offset=%d%n",
                        record.key(), record.value(), record.offset());
                messageCount++;
            }
        }

        consumer.close();
    }

}

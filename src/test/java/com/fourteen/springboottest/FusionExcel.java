package com.fourteen.springboottest;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.read.listener.PageReadListener;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/9/15 15:33
 */
public class FusionExcel {

    public static void main(String[] args) {

        List<UserSource> userSourceList = new ArrayList<>();
        EasyExcel.read("D:\\update-sql\\bjs-source-temp.xlsx", UserSource.class,
                new PageReadListener<UserSource>(userSourceList::addAll)).sheet(0).doRead();

        List<RemarkSource> remarkSourceList = new ArrayList<>();
        EasyExcel.read("D:\\update-sql\\new-qyt.xlsx", RemarkSource.class,
                new PageReadListener<RemarkSource>(remarkSourceList::addAll)).sheet(0).doRead();

        Map<String, UserSource> map = userSourceList.stream().collect(Collectors.toMap(UserSource::getTableName, t -> t, (a, b) -> a));

        remarkSourceList.forEach(remarkSource -> {
            if (map.containsKey(remarkSource.getTableName())) {
                UserSource userSource = map.get(remarkSource.getTableName());
                remarkSource.setField(userSource.getField());
                remarkSource.setModule(userSource.getModule());
                remarkSource.setSource(userSource.getSource());
            }
        });

        EasyExcel.write("D:\\update-sql\\bjs-source.xlsx", RemarkSource.class)
                .sheet("模板")
                .doWrite(() -> remarkSourceList);
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class RemarkSource {
        @ExcelProperty(value = "表名")
        String tableName;

        @ExcelProperty(value = "schema")
        String tableSchema;

        @ExcelProperty(value = "备注")
        String remark;

        @ExcelProperty(value = "字段名")
        String field;

        @ExcelProperty(value = "涉及模块")
        String module;

        @ExcelProperty(value = "数据来源")
        String source;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class UserSource {
        @ExcelProperty(value = "表名")
        String tableName;

        @ExcelProperty(value = "schema")
        String tableSchema;

        @ExcelProperty(value = "字段名")
        String field;

        @ExcelProperty(value = "涉及模块")
        String module;

        @ExcelProperty(value = "数据来源")
        String source;
    }

}

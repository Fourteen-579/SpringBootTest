package com.fourteen.springboottest;

import cn.hutool.core.io.FileUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.read.listener.PageReadListener;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.junit.jupiter.api.Test;

import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/2/10 14:22
 */
public class CreateUpdateSql {

    public static final String REPLACE_UPDATE_SQL = "UPDATE `%s`.`%s` SET `%s` = REPLACE(%s, '%s', '%s') WHERE `%s` LIKE '%%%s%%';";
    public static final String UPDATE_SQL = "UPDATE `%s`.`%s` SET `%s` = '%s' WHERE `%s` = '%s';";
    public static final String QYT_TABLE_FILE = "D:\\update-sql\\qyt.xlsx";
    public static final String SC_TABLE_FILE = "D:\\update-sql\\sc.xlsx";
    public static final String DATA_FILE = "D:\\update-sql\\oldNewInfo.xlsx";
    public static final String SAVE_PATH = "D:\\update-sql\\";
    public static final String SUFFIX = ".BJ";
    public static final String SEARCH_SQL = "SELECT * FROM `%s`.`%s` WHERE `%s` in (%s);";

    //企业通&数仓-分文件
    @Test
    public void createSql() {
        List<OldNewInfo> oldNewInfoList = new ArrayList<>();
        EasyExcel.read(DATA_FILE, OldNewInfo.class, new PageReadListener<OldNewInfo>(oldNewInfoList::addAll)).sheet(0).doRead();

        createSqlAndFile("企业通", QYT_TABLE_FILE, oldNewInfoList);
        createSqlAndFile("数仓", SC_TABLE_FILE, oldNewInfoList);
    }

    //企业通&数仓-同一文件
    @Test
    public void createSql2() {
        List<OldNewInfo> oldNewInfoList = new ArrayList<>();
        EasyExcel.read(DATA_FILE, OldNewInfo.class, new PageReadListener<OldNewInfo>(oldNewInfoList::addAll)).sheet(0).doRead();

        createSql4(oldNewInfoList, QYT_TABLE_FILE, "D://update-sql//qyt-all.txt");
        createSql4(oldNewInfoList, SC_TABLE_FILE, "D://update-sql//sc-all.txt");
    }

    //企业通&数仓-查询语句
    @Test
    public void searchSql() {
        List<OldNewInfo> oldNewInfoList = new ArrayList<>();
        EasyExcel.read(DATA_FILE, OldNewInfo.class, new PageReadListener<OldNewInfo>(oldNewInfoList::addAll)).sheet(0).doRead();

        searchSql(oldNewInfoList, QYT_TABLE_FILE, "D://update-sql//qyt-search.txt");
        searchSql(oldNewInfoList, SC_TABLE_FILE, "D://update-sql//sc-search.txt");
    }

    private void searchSql(List<OldNewInfo> oldNewInfoList, String tableInfoPath, String savePath) {
        List<TableInfo> tableInfoList = new ArrayList<>();
        EasyExcel.read(tableInfoPath, TableInfo.class, new PageReadListener<TableInfo>(tableInfoList::addAll)).sheet(0).doRead();

        String whereCondition = oldNewInfoList.stream()
                .map(temp -> "'" + temp.getOldValue() + "'").collect(Collectors.joining(","));

        String whereCondition2 = oldNewInfoList.stream()
                .map(temp -> "'" + temp.getOldValue() + SUFFIX + "'").collect(Collectors.joining(","));

        List<String> sqlList = tableInfoList.stream()
                .map(temp -> String.format(SEARCH_SQL,
                        temp.getTableSchema(),
                        temp.getTableName(),
                        temp.getField(),
                        temp.getAddSuffix() == 1 ? whereCondition2 : whereCondition)).collect(Collectors.toList());

        FileUtil.writeLines(sqlList, savePath, StandardCharsets.UTF_8);
    }

    private void createSql4(List<OldNewInfo> oldNewInfoList, String tableInfoPath, String savePath) {
        List<TableInfo> tableInfoList = new ArrayList<>();
        EasyExcel.read(tableInfoPath, TableInfo.class, new PageReadListener<TableInfo>(tableInfoList::addAll)).sheet(0).doRead();
        List<String> sqlList = createSql(tableInfoList, oldNewInfoList);

        FileUtil.writeLines(sqlList, savePath, StandardCharsets.UTF_8);
    }

    private List<String> createSql(List<TableInfo> tableInfoList, List<OldNewInfo> oldNewInfoList) {

        List<String> result = new ArrayList<>();
        tableInfoList.forEach(tableInfo -> {
            List<String> sqlList = oldNewInfoList.stream().map(oldNewInfo -> {
                String sql;
                if (tableInfo.replace == 1) {
                    sql = String.format(REPLACE_UPDATE_SQL,
                            tableInfo.getTableSchema(),
                            tableInfo.getTableName(),
                            tableInfo.getField(),
                            tableInfo.getField(),
                            oldNewInfo.getOldValue(),
                            oldNewInfo.getNewValue(),
                            tableInfo.getField(),
                            oldNewInfo.getOldValue());
                } else {
                    if (tableInfo.addSuffix != null && tableInfo.addSuffix == 1) {
                        sql = String.format(UPDATE_SQL,
                                tableInfo.getTableSchema(),
                                tableInfo.getTableName(),
                                tableInfo.getField(),
                                oldNewInfo.getNewValue() + CreateUpdateSql.SUFFIX,
                                tableInfo.getField(),
                                oldNewInfo.getOldValue() + CreateUpdateSql.SUFFIX);
                    } else {
                        sql = String.format(UPDATE_SQL,
                                tableInfo.getTableSchema(),
                                tableInfo.getTableName(),
                                tableInfo.getField(),
                                oldNewInfo.getNewValue(),
                                tableInfo.getField(),
                                oldNewInfo.getOldValue());
                    }
                }
                return sql;
            }).collect(Collectors.toList());
            result.addAll(sqlList);
        });

        return result;
    }

    private void createSqlAndFile(String filePath, String tableInfoPath, List<OldNewInfo> oldNewInfoList) {
        List<TableInfo> tableInfoList = new ArrayList<>();
        EasyExcel.read(tableInfoPath, TableInfo.class, new PageReadListener<TableInfo>(tableInfoList::addAll)).sheet(0).doRead();

        tableInfoList.forEach(tableInfo -> {
            List<String> sqlList = oldNewInfoList.stream().map(oldNewInfo -> {
                String sql;
                if (tableInfo.replace == 1) {
                    sql = String.format(REPLACE_UPDATE_SQL,
                            tableInfo.getTableSchema(),
                            tableInfo.getTableName(),
                            tableInfo.getField(),
                            tableInfo.getField(),
                            oldNewInfo.getOldValue(),
                            oldNewInfo.getNewValue(),
                            tableInfo.getField(),
                            oldNewInfo.getOldValue());
                } else {
                    if (tableInfo.addSuffix == 1) {
                        sql = String.format(UPDATE_SQL,
                                tableInfo.getTableSchema(),
                                tableInfo.getTableName(),
                                tableInfo.getField(),
                                oldNewInfo.getNewValue() + CreateUpdateSql.SUFFIX,
                                tableInfo.getField(),
                                oldNewInfo.getOldValue() + CreateUpdateSql.SUFFIX);
                    } else {
                        sql = String.format(UPDATE_SQL,
                                tableInfo.getTableSchema(),
                                tableInfo.getTableName(),
                                tableInfo.getField(),
                                oldNewInfo.getNewValue(),
                                tableInfo.getField(),
                                oldNewInfo.getOldValue());
                    }
                }
                return sql;
            }).collect(Collectors.toList());

            FileUtil.writeLines(sqlList, SAVE_PATH + filePath + "\\" + tableInfo.getTableName() + "-" + tableInfo.getField() + ".txt", StandardCharsets.UTF_8);
        });
    }

    @Data
    public static class OldNewInfo {
        @ExcelProperty(value = "旧值")
        String oldValue;

        @ExcelProperty(value = "新值")
        String newValue;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class TableInfo {
        @ExcelProperty(value = "表名")
        String tableName;

        @ExcelProperty(value = "schema")
        String tableSchema;

        @ExcelProperty(value = "字段名")
        String field;

        @ExcelProperty(value = "是否添加后缀")
        Integer addSuffix;

        @ExcelProperty(value = "是否使用替换")
        Integer replace;
    }

}

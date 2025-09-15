package com.fourteen.springboottest;

import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.read.listener.PageReadListener;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.sql.*;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/8/21 10:48
 */
public class MySqlDataSource {

    private static final Map<String, Object> QYT = new HashMap<>();
    private static final Map<String, Object> SC = new HashMap<>();
    private static final String URL = "url";
    private static final String USERNAME = "username";
    private static final String PASSWORD = "password";
    private static final String SCHEMAS = "schemas";
    private static final String tableInfoPath = "tableInfoPath";
    private static final String newTableInfoPath = "newTableInfoPath";

    static {
        QYT.put("url", "jdbc:mysql://10.205.204.67:3306/webserver?serverTimezone=Asia/Shanghai&useUnicode=true&characterEncoding=utf-8&allowMultiQueries=true&useSSL=false&rewriteBatchedStatements=true");
        QYT.put("username", "root");
        QYT.put("password", "N8XmZ3OiMh");
        QYT.put("schemas", new String[]{"webserver", "ent_backend"});
        QYT.put("tableInfoPath", "D:\\update-sql\\qyt.xlsx");
        QYT.put("newTableInfoPath", "D:\\update-sql\\new-qyt.xlsx");

        SC.put("url", "jdbc:mysql://10.150.224.26:3306/nrs-enterprise?serverTimezone=Asia/Shanghai&characterEncoding=utf8&useUnicode=true&useSSL=false&allowMultiQueries=true");
        SC.put("username", "root");
        SC.put("password", "3xpYTaGL28");
        SC.put("schemas", new String[]{"nrs-enterprise", "qiyetong"});
        SC.put("tableInfoPath", "D:\\update-sql\\sc.xlsx");
        SC.put("newTableInfoPath", "D:\\update-sql\\new-sc.xlsx");
    }


    public static void main(String[] args) {
        getRemark();
    }

    static void getRemark() {
        Map<String, Object> map = SC;

        // 读取已有表信息
        List<TableInfo> tableInfoList = new ArrayList<>();
        EasyExcel.read((String) map.get(tableInfoPath), TableInfo.class,
                new PageReadListener<TableInfo>(tableInfoList::addAll)).sheet(0).doRead();

        List<TableInfo> list = tableInfoList.stream().filter(t -> StrUtil.isNotBlank(t.getField())).collect(Collectors.toList());

        // 已有表信息建立索引（schema+tableName）
        Map<String, List<TableInfo>> tableIndex = list.stream()
                .collect(Collectors.groupingBy(
                        t -> (t.getTableSchema() + "." + t.getTableName())));

        // 数据库连接信息
        String url = (String) map.get(URL);
        String username = (String) map.get(USERNAME);
        String password = (String) map.get(PASSWORD);
        String[] schemas = (String[]) map.get(SCHEMAS);

        List<TableRemarkInfo> newTableInfoList = new ArrayList<>();

        try (Connection conn = DriverManager.getConnection(url, username, password)) {
            // 获取数据库元数据
            DatabaseMetaData metaData = conn.getMetaData();

            for (String schema : schemas) {
                try (ResultSet tables = metaData.getTables(schema, null, "%", new String[]{"TABLE"})) {
                    while (tables.next()) {
                        String tableName = tables.getString("TABLE_NAME");
                        String remarks = tables.getString("REMARKS");

                        String key = (schema + "." + tableName).toLowerCase();
                        if (tableIndex.containsKey(key)) {
                            TableRemarkInfo newTable = new TableRemarkInfo(tableName, schema, remarks);
                            newTableInfoList.add(newTable);
                        }
                    }
                }
            }

            newTableInfoList.sort(
                    Comparator.comparing((TableRemarkInfo t) -> (t.tableSchema == null ? "" : t.tableSchema))
                            .thenComparing(t -> t.tableName == null ? "" : t.tableName, String.CASE_INSENSITIVE_ORDER)
            );

            EasyExcel.write((String) map.get(newTableInfoPath), TableRemarkInfo.class)
                    .sheet("模板")
                    .doWrite(() -> newTableInfoList);

        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    static void searchNotExistTable() {
        Map<String, Object> map = SC;

        // 读取已有表信息
        List<TableInfo> tableInfoList = new ArrayList<>();
        EasyExcel.read((String) map.get(tableInfoPath), TableInfo.class,
                new PageReadListener<TableInfo>(tableInfoList::addAll)).sheet(0).doRead();

        // 已有表信息建立索引（schema+tableName）
        Map<String, List<TableInfo>> tableIndex = tableInfoList.stream()
                .collect(Collectors.groupingBy(
                        t -> (t.getTableSchema() + "." + t.getTableName())));

        // 数据库连接信息
        String url = (String) map.get(URL);
        String username = (String) map.get(USERNAME);
        String password = (String) map.get(PASSWORD);
        String[] schemas = (String[]) map.get(SCHEMAS);

        List<TableInfo> newTableInfoList = new ArrayList<>();

        try (Connection conn = DriverManager.getConnection(url, username, password)) {
            // 获取数据库元数据
            DatabaseMetaData metaData = conn.getMetaData();

            for (String schema : schemas) {
                try (ResultSet tables = metaData.getTables(schema, null, "%", new String[]{"TABLE"})) {
                    while (tables.next()) {
                        String tableName = tables.getString("TABLE_NAME");

                        String key = (schema + "." + tableName).toLowerCase();
                        if (!tableIndex.containsKey(key)) {
                            // 新表
                            TableInfo newTable = new TableInfo(tableName, schema, 1, null, null, null);
                            newTableInfoList.add(newTable);
                            System.out.println("找到新表：Schema: " + schema + "，表名: " + tableName);
                        } else {
                            // 已有表，更新 isNewTable = 0
                            tableIndex.get(key).forEach(t -> t.setIsNewTable(0));
                        }
                    }
                }
            }

            tableInfoList.addAll(newTableInfoList);
            tableInfoList.sort(
                    Comparator.comparing((TableInfo t) -> (t.field == null || t.field.trim().isEmpty())) // true/false, 有字段的优先
                            .thenComparing(t -> t.tableSchema == null ? "" : t.tableSchema) // schema 排序
                            .thenComparing(t -> t.tableName == null ? "" : t.tableName, String.CASE_INSENSITIVE_ORDER) // 表名字母序
            );

            EasyExcel.write((String) map.get(newTableInfoPath), TableInfo.class)
                    .sheet("模板")
                    .doWrite(() -> tableInfoList);

        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class TableInfo {
        @ExcelProperty(value = "表名")
        String tableName;

        @ExcelProperty(value = "schema")
        String tableSchema;

        @ExcelProperty(value = "是否为新添加表")
        Integer isNewTable;

        @ExcelProperty(value = "字段名")
        String field;

        @ExcelProperty(value = "是否添加后缀")
        Integer addSuffix;

        @ExcelProperty(value = "是否使用替换")
        Integer replace;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class TableRemarkInfo {
        @ExcelProperty(value = "表名")
        String tableName;

        @ExcelProperty(value = "schema")
        String tableSchema;

        @ExcelProperty(value = "备注")
        String remark;
    }

}

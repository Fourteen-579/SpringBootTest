package com.fourteen.springboottest.util;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/12/30 14:34
 */
public class Md5Util {

    /**
     * 计算数据的 MD5 值（小写十六进制）
     */
    public static String md5HexLower(byte[] data) {
        try {
            MessageDigest md5 = MessageDigest.getInstance("MD5");
            byte[] digest = md5.digest(data);

            StringBuilder sb = new StringBuilder(digest.length * 2);
            for (byte b : digest) {
                sb.append(Character.forDigit((b >> 4) & 0xF, 16));
                sb.append(Character.forDigit(b & 0xF, 16));
            }
            return sb.toString(); // 默认就是小写
        } catch (Exception e) {
            throw new RuntimeException("MD5 calculation failed", e);
        }
    }

    public static String md5HexLower(String data) {
        return md5HexLower(data.getBytes(StandardCharsets.UTF_8));
    }
}

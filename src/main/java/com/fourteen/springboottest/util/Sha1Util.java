package com.fourteen.springboottest.util;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2025/12/30 13:43
 */
public class Sha1Util {

    public static String sha1(String data) {
        try {
            MessageDigest digest = MessageDigest.getInstance("SHA-1");
            byte[] bytes = digest.digest(data.getBytes(StandardCharsets.UTF_8));
            return toHex(bytes);
        } catch (Exception e) {
            throw new RuntimeException("SHA1 calculation failed", e);
        }
    }

    private static String toHex(byte[] bytes) {
        StringBuilder sb = new StringBuilder(bytes.length * 2);
        for (byte b : bytes) {
            sb.append(String.format("%02x", b & 0xff));
        }
        return sb.toString();
    }
}

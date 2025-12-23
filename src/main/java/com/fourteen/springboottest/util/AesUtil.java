package com.fourteen.springboottest.util;

import org.apache.commons.codec.binary.Base64;
import org.apache.commons.lang3.StringUtils;

import javax.crypto.Cipher;
import javax.crypto.KeyGenerator;
import javax.crypto.SecretKey;
import javax.crypto.SecretKeyFactory;
import javax.crypto.spec.DESKeySpec;
import javax.crypto.spec.IvParameterSpec;
import javax.crypto.spec.SecretKeySpec;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;

/**
 * @Description: Aes加密
 * @Author: Huang qingqing
 * @Date: 2020/11/17 14:29
 */
public class AesUtil {

    /**
     * 算法
     */
    private static final String ALGORITHMSTR = "AES/ECB/PKCS5Padding";

    private static final String KEY = "08cfbf6f49d676bf"; // 密钥
    private static final String XG_KEY = "QYT_XINGONG_AES_KEY_hdiudhwdqhic"; // 偏移向量

    private static final String KEY2 = "VCU4$K&S@2PQDQ%^";
    private static final String ALGORITHMSTR_CBC = "AES/CBC/PKCS5Padding";
    private static final String INIT_VECTOR = "VCU4$K&S@2PQDQ%^";
    private static final String Key3 = "B?aUG$$6@MXGE!PFz^wN36sX%QfV2sFT";

    public static void main(String[] args) throws IOException {
        //解密手机号
        System.out.println(aesDecrypt("G1MG2oO5fkTvFmNDMVPOSQ==", KEY));
//        System.out.println(aesEncrypt("kUYpOtaaNdQRC2Xnf9s0Ug==", KEY));
//        System.out.println(aesEncrypt("6873097591325414", XG_KEY));

        //加密手机号
//        System.out.println(aesEncrypt("MobileLoginReq(phone=13132029095, prefix=86, area=中国大陆, code=193632, type=1, uid=null, matterId=null)", KEY));

        //解密文件
//        String filePath = "C:\\Users\\Administrator\\Desktop\\三会电子签\\test.txt";
//        System.out.println(desFile(filePath));
    }

    public static byte[] aesDecrypt256(byte[] encryptStr) {
        byte[] decryptedStr = new byte[0];
        try {
            decryptedStr = aesDecryptBytes256(Base64.decodeBase64(encryptStr), Key3);
        } catch (Exception ex) {
            System.out.println("解密失败，失败原因：" + ex);
        }
        return decryptedStr;
    }

    private static byte[] aesDecryptBytes256(byte[] encryptBytes, String decryptKey) throws Exception {
        KeyGenerator kgen = KeyGenerator.getInstance("AES");
        IvParameterSpec ivParameterSpec = new IvParameterSpec(INIT_VECTOR.getBytes(StandardCharsets.UTF_8));
        kgen.init(256);
        Cipher cipher = Cipher.getInstance(ALGORITHMSTR_CBC);
        cipher.init(Cipher.DECRYPT_MODE, new SecretKeySpec(decryptKey.getBytes(), "AES"), ivParameterSpec);
        return cipher.doFinal(encryptBytes);
    }


    private static String desFile(String filePath) {
        try {
            byte[] bytes = Files.readAllBytes(Paths.get(filePath));
            String content = new String(bytes, StandardCharsets.UTF_8);

            String s = aesDecrypt(content, KEY);
            return cn.hutool.core.codec.Base64.decodeStr(s);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static void aes() throws UnsupportedEncodingException {
        //des密钥，固定值
        String keys = "TnKfEmmstNA=";
        //加密数据，约定使用的密钥，解密客户端通过appid从通行证组获取
        String dataSecret = "8c0914be";
        //加密后的密文
        String cryptStr = "sXb4c1POUhBfsfeJP1Au3A";
        //通过反向解码得到真正的密文"GT84R7+Q1Xyn5ETm/8NHhw=="
        String realCryptStr = cryptStr.replace("-", "+").replace("_", "/") + "====".substring(cryptStr.length() % 4 == 0 ? 4 : cryptStr.length() % 4);
        //System.out.println(realCryptStr);
        String umobPhone = DESDecrypt(Base64.decodeBase64(realCryptStr), Base64.decodeBase64(keys), dataSecret.getBytes("UTF-8"));
        System.out.println("密文：" + cryptStr + " 数据秘钥：" + dataSecret + " 解密后数据：" + umobPhone);
    }

    /**
     * 解密函数
     *
     * @param data 解密数据
     * @param key  密钥
     * @param iv   偏移向量
     * @return 返回解密后的数据
     */
    public static String DESDecrypt(byte[] data, byte[] key, byte[] iv) {
        try {
            // 从原始密匙数据创建一个DESKeySpec对象
            DESKeySpec dks = new DESKeySpec(key);

            // 创建一个密匙工厂，然后用它把DESKeySpec对象转换成
            // 一个SecretKey对象
            SecretKeyFactory keyFactory = SecretKeyFactory.getInstance("DES");
            SecretKey secretKey = keyFactory.generateSecret(dks);

            // using DES in CBC mode
            Cipher cipher = Cipher.getInstance("DES/CBC/PKCS5Padding");

            // 用密匙初始化Cipher对象
            IvParameterSpec param = new IvParameterSpec(iv);
            cipher.init(Cipher.DECRYPT_MODE, secretKey, param);

            // 正式执行解密操作
            byte decryptedData[] = cipher.doFinal(data);
            return new String(decryptedData);
        } catch (Exception e) {
            System.err.println("DES算法，解密出错。");
            e.printStackTrace();
        }

        return null;
    }


    /**
     * AES加密为base 64 code
     *
     * @param content    待加密的内容
     * @param encryptKey 加密密钥
     * @return 加密后的base 64 code
     */
    public static String aesEncrypt(String content, String encryptKey) {
        String encryptedStr = "";
        try {
            encryptedStr = base64Encode(aesEncryptToBytes(content, encryptKey));
        } catch (Exception ex) {
            System.out.println("AesUtil->aesEncrypt" + ex);
        }
        return encryptedStr;
    }

    /**
     * 将base 64 code AES解密
     *
     * @param encryptStr 待解密的base 64 code
     * @param decryptKey 解密密钥
     * @return 解密后的string
     */
    public static String aesDecrypt(String encryptStr, String decryptKey) {
        String decryptedStr = "";
        try {
            decryptedStr = StringUtils.isEmpty(encryptStr) ? null : aesDecryptByBytes(base64Decode(encryptStr), decryptKey);
        } catch (Exception ex) {
            System.out.println("AesUtil->aesDecrypt" + ex);
        }
        return decryptedStr;
    }


    /**
     * AES解密
     *
     * @param encryptBytes 待解密的byte[]
     * @param decryptKey   解密密钥
     * @return 解密后的String
     */
    private static String aesDecryptByBytes(byte[] encryptBytes, String decryptKey) throws Exception {
        KeyGenerator kgen = KeyGenerator.getInstance("AES");
        kgen.init(128);
        Cipher cipher = Cipher.getInstance(ALGORITHMSTR);
        cipher.init(Cipher.DECRYPT_MODE, new SecretKeySpec(decryptKey.getBytes(), "AES"));
        byte[] decryptBytes = cipher.doFinal(encryptBytes);
        return new String(decryptBytes);
    }


    /**
     * base 64 encode
     *
     * @param bytes 待编码的byte[]
     * @return 编码后的base 64 code
     */
    private static String base64Encode(byte[] bytes) {
        return Base64.encodeBase64String(bytes);
    }

    /**
     * base 64 decode
     *
     * @param base64Code 待解码的base 64 code
     * @return 解码后的byte[]
     */
    private static byte[] base64Decode(String base64Code) {
        return StringUtils.isEmpty(base64Code) ? null : Base64.decodeBase64(base64Code);
    }


    /**
     * AES加密
     *
     * @param content    待加密的内容
     * @param encryptKey 加密密钥
     * @return 加密后的byte[]
     */
    private static byte[] aesEncryptToBytes(String content, String encryptKey) throws Exception {
        KeyGenerator kgen = KeyGenerator.getInstance("AES");
        kgen.init(128);
        Cipher cipher = Cipher.getInstance(ALGORITHMSTR);
        cipher.init(Cipher.ENCRYPT_MODE, new SecretKeySpec(encryptKey.getBytes(), "AES"));

        return cipher.doFinal(content.getBytes(StandardCharsets.UTF_8));
    }

}

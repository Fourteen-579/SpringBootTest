package com.fourteen.springboottest;

import org.apache.commons.codec.binary.Base64;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import javax.crypto.Cipher;
import javax.crypto.KeyGenerator;
import javax.crypto.spec.IvParameterSpec;
import javax.crypto.spec.SecretKeySpec;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

/**
 * @author Fourteen_ksz
 * @version 1.0
 * @date 2024/5/11 16:11
 */
@SpringBootTest
public class TestZip {

    private static final String ALGORITHMSTR_CBC = "AES/CBC/PKCS5Padding";
    private static final String INIT_VECTOR = "VCU4$K&S@2PQDQ%^";


    public static byte[] readZipFileToBytes(File zipFile) {

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        try (FileInputStream fis = new FileInputStream(zipFile);
             ZipInputStream zipInputStream = new ZipInputStream(fis)) {

            byte[] buffer = new byte[1024];
            int len;
            while ((len = zipInputStream.read(buffer)) != -1) {
                outputStream.write(buffer, 0, len);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return outputStream.toByteArray();
    }

    public static void saveByteArrayOutputStreamAsZipFile(ByteArrayOutputStream byteArrayOutputStream, String outputFilePath) {
        try (FileOutputStream fos = new FileOutputStream(outputFilePath);
             ZipOutputStream zipOutputStream = new ZipOutputStream(fos)) {

            ZipEntry zipEntry = new ZipEntry("entryName"); // 替换为压缩包中的文件路径和文件名
            zipOutputStream.putNextEntry(zipEntry);
            zipOutputStream.write(byteArrayOutputStream.toByteArray());
            zipOutputStream.closeEntry();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void saveBytesToLocalFile(byte[] bytes, String filePath) {
        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            fos.write(bytes);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static byte[] aesDecrypt256(byte[] encryptStr, String decryptKey) {
        byte[] decryptedStr = new byte[0];
        try {
            decryptedStr = aesDecryptBytes256(Base64.decodeBase64(encryptStr), decryptKey);
        } catch (Exception ex) {
            //ESLogUtil.error("AesUtil->aesDecrypt",ex);
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

//    @Test
    void testZip() throws IOException {
        File file = new File("C:\\Users\\Administrator\\Desktop\\test.zip");
        byte[] bytes = readZipFileToBytes(file);


        ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(bytes);
        ZipInputStream zipInputStream = new ZipInputStream(byteArrayInputStream);

        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        ZipOutputStream zipOutputStream = new ZipOutputStream(byteArrayOutputStream);
        int i = 0;
        try {
            // 读取 ZIP 中的每个文件，并给每个文件重命名
            ZipEntry zipEntry = zipInputStream.getNextEntry();
            while (zipEntry != null) {
                String[] originName = zipEntry.getName().split("\\.");

                String newName = "test" + (i++) + ".pdf";
                zipOutputStream.putNextEntry(new ZipEntry(newName));

                // 将原文件内容写入到重命名后的 ZIP 中
                byte[] buffer = new byte[1024];
                int len;
                ByteArrayOutputStream arrayOutputStream = new ByteArrayOutputStream();
                while ((len = zipInputStream.read(buffer)) > 0) {
                    arrayOutputStream.write(buffer, 0, len);
                }
//                byte[] desBytes = AesUtil.aesDecrypt256(arrayOutputStream.toByteArray(), "B?aUG$$6@MXGE!PFz^wN36sX%QfV2sFT");
                zipOutputStream.write(arrayOutputStream.toByteArray());
                zipOutputStream.closeEntry();
                zipEntry = zipInputStream.getNextEntry();
            }

            saveByteArrayOutputStreamAsZipFile(byteArrayOutputStream, "C:\\Users\\Administrator\\Desktop\\test2.zip");
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            try {
                zipOutputStream.close();
                zipInputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


}

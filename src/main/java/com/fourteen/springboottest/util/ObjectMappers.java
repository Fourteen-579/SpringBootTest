package com.fourteen.springboottest.util;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.*;
import com.fasterxml.jackson.databind.type.ArrayType;
import com.fasterxml.jackson.databind.type.CollectionType;
import com.fasterxml.jackson.databind.type.MapType;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;
import com.fasterxml.jackson.datatype.jsr310.deser.LocalDateDeserializer;
import com.fasterxml.jackson.datatype.jsr310.deser.LocalDateTimeDeserializer;
import com.fasterxml.jackson.datatype.jsr310.ser.LocalDateSerializer;
import com.fasterxml.jackson.datatype.jsr310.ser.LocalDateTimeSerializer;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Map;
import java.util.Optional;

/**
 * jackson ObjectMapper util
 *
 * @author jiangjinhang
 * @since 2023/2/16
 */
@Slf4j
public class ObjectMappers {

    /**
     * global mapper
     */
    private static final ObjectMapper MAPPER = new ObjectMapper();

    static {
        // deserialization config
        // ignore unknown field in json
        MAPPER.configure(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES, false);
        MAPPER.configure(JsonParser.Feature.ALLOW_COMMENTS, true);
        MAPPER.configure(MapperFeature.ACCEPT_CASE_INSENSITIVE_PROPERTIES, true);


        // serialization config
        // close serialize date with timestamp
        MAPPER.configure(SerializationFeature.WRITE_DATES_AS_TIMESTAMPS, false);
        MAPPER.setDateFormat(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss"));

        JavaTimeModule javaTimeModule = new JavaTimeModule();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
        javaTimeModule.addSerializer(LocalDateTime.class, new LocalDateTimeSerializer(formatter));
        javaTimeModule.addDeserializer(LocalDateTime.class, new LocalDateTimeDeserializer(formatter));
        javaTimeModule.addSerializer(LocalDate.class, new LocalDateSerializer(formatter));
        javaTimeModule.addDeserializer(LocalDate.class, new LocalDateDeserializer(formatter));

        MAPPER.registerModule(javaTimeModule);
    }

    /**
     * serialize object to json string.
     *
     * @param o object of need to serialize
     * @return if serialize success return json str
     */
    public static Optional<String> writeAsJsonStr(Object o) {
        try {
            return Optional.of(MAPPER.writeValueAsString(o));
        } catch (Exception e) {
            log.error("serialize to json string failed.", e);
        }
        return Optional.empty();
    }

    public static String writeAsJsonStrThrow(Object o) {
        return writeAsJsonStr(o).orElseThrow(() -> new RuntimeException("序列化失败"));
    }

    /**
     * serialize object to json bytes.
     *
     * @param o object of need to serialize
     * @return if serialize success return json bytes
     */
    public static Optional<byte[]> writeAsJsonBytes(Object o) {
        try {
            return Optional.of(MAPPER.writeValueAsBytes(o));
        } catch (Exception e) {
            log.error("serialize to bytes failed.", e);
        }
        return Optional.empty();
    }

    /**
     * deserialize json string to object of type T
     *
     * @param jsonStr json string
     * @param <T>     target type, can not collection type or array type
     * @return if deserialize success return object of target type
     */
    public static <T> Optional<T> readAsObj(String jsonStr, Class<T> clazz) {
        try {
            return Optional.of(MAPPER.readValue(jsonStr, clazz));
        } catch (Exception e) {
            log.error("deserialize string failed.", e);
        }
        return Optional.empty();
    }

    public static <T> T readAsObjThrow(String jsonStr, Class<T> clazz) {
        return readAsObj(jsonStr, clazz).orElseThrow(() -> new RuntimeException("序列化失败"));
    }

    /**
     * deserialize json string to list of List< T >
     *
     * @param jsonStr      json string
     * @param elementClass target list element type
     * @param <T>          generic type
     * @return if deserialize success return List of target element type
     */
    public static <T> Optional<List<T>> readAsList(String jsonStr, Class<T> elementClass) {
        try {
            JavaType type = MAPPER.getTypeFactory().constructCollectionType(List.class, elementClass);
            return Optional.of(MAPPER.readValue(jsonStr, type));
        } catch (Exception e) {
            log.error("deserialize string failed.", e);
        }
        return Optional.empty();
    }

    public static <T> List<T> readAsListThrow(String jsonStr, Class<T> elementClass) {
        return readAsList(jsonStr, elementClass).orElseThrow(() -> new RuntimeException("序列化失败"));
    }

    /**
     * deserialize json string to array of type T[]
     *
     * @param jsonStr      json string
     * @param elementClass target array element type
     * @param <T>          generic type
     * @return if deserialize success return array of target element type
     */
    public static <T> Optional<T[]> readAsArray(String jsonStr, Class<T> elementClass) {
        try {
            JavaType type = MAPPER.getTypeFactory().constructArrayType(elementClass);
            return Optional.of(MAPPER.readValue(jsonStr, type));
        } catch (Exception e) {
            log.error("deserialize string failed.", e);
        }
        return Optional.empty();
    }

    /**
     * deserialize json string to map
     *
     * @param jsonStr json string
     * @return if deserialize success return instance of map
     */
    public static Optional<Map<String, ?>> readAsMap(String jsonStr) {
        try {
            MapType mapType = MAPPER.getTypeFactory().constructRawMapType(Map.class);
            return Optional.of(MAPPER.readValue(jsonStr, mapType));
        } catch (Exception e) {
            log.error("deserialize string failed.", e);
        }
        return Optional.empty();
    }

    public static Map<String, ?> readAsMapThrow(String jsonStr) {
        return readAsMap(jsonStr).orElseThrow(() -> new RuntimeException("序列化失败"));
    }

    /**
     * deserialize json bytes to object of type T
     *
     * @param bytes json bytes
     * @param <T>   target type, can not collection type or array type
     * @return if deserialize success return object of target type
     */
    public static <T> Optional<T> readAsObj(byte[] bytes, Class<T> clazz) {
        try {
            return Optional.of(MAPPER.readValue(bytes, clazz));
        } catch (Exception e) {
            log.error("deserialize bytes failed.", e);
        }
        return Optional.empty();
    }

    /**
     * deserialize json bytes to list of List< T >
     *
     * @param bytes        json bytes
     * @param elementClass target list element type
     * @param <T>          generic type
     * @return if deserialize success return List of target element type
     */
    public static <T> Optional<List<T>> readAsList(byte[] bytes, Class<T> elementClass) {
        try {
            CollectionType type = MAPPER.getTypeFactory().constructCollectionType(List.class, elementClass);
            return Optional.of(MAPPER.readValue(bytes, type));
        } catch (Exception e) {
            log.error("deserialize bytes failed.", e);
        }
        return Optional.empty();
    }

    /**
     * deserialize json bytes to array of type T[]
     *
     * @param bytes        json bytes
     * @param elementClass target array element type
     * @param <T>          generic type
     * @return if deserialize success return array of target element type
     */
    public static <T> Optional<T[]> readAsArray(byte[] bytes, Class<T> elementClass) {
        try {
            ArrayType type = MAPPER.getTypeFactory().constructArrayType(elementClass);
            return Optional.of(MAPPER.readValue(bytes, type));
        } catch (Exception e) {
            log.error("deserialize bytes failed.", e);
        }
        return Optional.empty();
    }

    /**
     * deserialize json bytes to map
     *
     * @param bytes json bytes
     * @return if deserialize success then return instance of map
     */
    public static Optional<Map<String, ?>> readAsMap(byte[] bytes) {
        try {
            MapType mapType = MAPPER.getTypeFactory().constructRawMapType(Map.class);
            return Optional.of(MAPPER.readValue(bytes, mapType));
        } catch (Exception e) {
            log.error("deserialize bytes failed.", e);
        }
        return Optional.empty();
    }

    public static <T> Optional<T> readAsObj(String result, TypeReference<T> typeReference) {
        try {
            return Optional.of(MAPPER.readValue(result, typeReference));
        } catch (Exception e) {
            log.error("deserialize bytes failed.", e);
        }
        return Optional.empty();
    }

    public static <T> T readAsObjThrow(String result, TypeReference<T> typeReference) {
        return readAsObj(result, typeReference).orElseThrow(() -> new RuntimeException("序列化失败"));
    }

    public static class PlainDoubleSerializer extends JsonSerializer<Double> {
        @Override
        public void serialize(Double value, JsonGenerator gen, SerializerProvider serializers) throws IOException {
            BigDecimal bigDecimal = new BigDecimal(value).setScale(10, RoundingMode.HALF_UP);
            gen.writeNumber(bigDecimal.toPlainString());
        }
    }
}

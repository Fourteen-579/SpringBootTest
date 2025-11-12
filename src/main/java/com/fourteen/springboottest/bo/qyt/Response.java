package com.fourteen.springboottest.bo.qyt;


import lombok.Data;

/**
 * @author: huangwei
 * @create: 2020/5/29 17:51
 */
@Data
public class Response<T> {
    private int code;
    private String msg;
    private T data;
    private String traceId;

    public Response() {
    }

    public Response(int code, String msg) {
        this.code = code;
        this.msg = msg;
    }

    public Response(int code, String msg, T data) {
        this.code = code;
        this.msg = msg;
        this.data = data;
    }

    public Response(Code code, T data) {
        this.code = code.code;
        this.msg = code.msg;
        this.data = data;
    }

    public Response(Code code) {
        this.code = code.code;
        this.msg = code.msg;
    }

    public static <T> Response<T> toResponse(Code code) {
        return new Response<>(code.code, code.msg);
    }

    public static <T> Response<T> toResponse(Code code, T t) {
        return new Response<>(code.code, code.msg, t);
    }

    public static <T> Response<T> toResponse(Code code, T data, String msg) {
        return new Response<>(code.code, msg, data);
    }

    public static <T> Response<T> success(T t) {
        return new Response<>(Code.SUCCESS, t);
    }

    public static <T> Response<T> fail(T t) {
        return new Response<>(Code.FAIL, t);
    }

    public int getCode() {
        return code;
    }

    public void setCode(int code) {
        this.code = code;
    }

    public String getMsg() {
        return msg;
    }

    public void setMsg(String msg) {
        this.msg = msg;
    }

    public T getData() {
        return data;
    }

    public void setData(T data) {
        this.data = data;
    }
}

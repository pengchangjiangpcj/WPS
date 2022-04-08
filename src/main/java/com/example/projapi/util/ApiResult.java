package com.example.projapi.util;

import lombok.Getter;
import lombok.Setter;

/**
 * @Author : pengchangjiang
 * @Date : 2021/5/10 18:09
 */
@Getter
@Setter
public class ApiResult {
    // 响应业务状态
    private Integer code;
    // 响应消息
    private String msg;
    // 响应中的数据
    private Object data;
    public static ApiResult build(Integer code, String msg, Object data) {
        return new ApiResult(code, msg, data);
    }

    public static ApiResult build(Integer code, String msg) {
        return new ApiResult(code, msg, null);
    }
    public static ApiResult build(String msg) {
        return new ApiResult(-1, msg, null);
    }

    public static ApiResult ok(Object data) {
        return new ApiResult(data);
    }
    public static ApiResult ok() {
        return new ApiResult(null);
    }
    public static ApiResult error(String msg) {
        return new ApiResult(-1, msg, null);
    }

    public static ApiResult error(String msg, Object data) {
        return new ApiResult(-1, msg, data);
    }

    public ApiResult() {}
    public ApiResult(Integer code, String msg, Object data) {
        this.code = code;
        this.msg = msg;
        this.data = data;
    }
    public ApiResult(Object data) {
        this.code = 0;
        this.msg = "SUCCESS";
        this.data = data;
    }

}

package cn.zhuangqf.tool.excel;

import java.io.Serializable;

/**
 * @author zhuangqf
 * @date 2018/2/6
 * @description excel导入行的详细信息
 */
public class ImportDetail<T> implements Serializable {

    /**
     * @description 导入状态
     * @value 0 导入成功
     */
    private int status = 0;
    /**
     * @description 导入信息
     * @value 成功 ""
     * 失败 异常信息
     */
    private String message = "";

    /**
     * @description 导入成功后读取的数据
     * @value 成功 T
     * 失败 null
     */
    private T data;

    public ImportDetail<T> success(T t) {
        this.status = 0;
        this.message = "";
        this.data = t;
        return this;
    }

    public ImportDetail<T> failure(String message) {
        this.status = 1;
        this.message = message;
        return this;
    }

    public int getStatus() {
        return status;
    }

    public String getMessage() {
        return message;
    }

    public T getData() {
        return data;
    }

}

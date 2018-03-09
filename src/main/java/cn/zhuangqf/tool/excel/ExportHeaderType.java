package cn.zhuangqf.tool.excel;

import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

/**
 * @author zhuangqf
 * @date 2018/2/8
 */
public enum ExportHeaderType {

    NUMBER, DATE, DATETIME, TIME, STRING, BOOLEAN;

    private static final Map<String, ExportHeaderType> mapper = new HashMap<>();

    static {
        Arrays.asList("int", "float", "double", "integer", "number", "numeric", "long").forEach(key -> mapper.put(key, NUMBER));
        mapper.put("date", DATE);
        mapper.put("datetime", DATETIME);
        mapper.put("time", TIME);
        Arrays.asList("bool", "boolean").forEach(key -> mapper.put(key, BOOLEAN));
    }

    public static ExportHeaderType fromString(String name) {
        name = name.toUpperCase();
        return mapper.getOrDefault(name, STRING);
    }
}

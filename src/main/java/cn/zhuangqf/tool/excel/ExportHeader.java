package cn.zhuangqf.tool.excel;

import org.apache.poi.ss.usermodel.CellType;

import java.util.function.Function;

/**
 * @author zhuangqf
 * @date 2018/2/8
 */
public class ExportHeader<T, R> {

    private String field;
    private String title;
    private ExportHeaderType type = ExportHeaderType.STRING;
    private boolean required = false;
    private Function<T, R> formatter;

    public ExportHeader(String field, String title) {
        this.field = field;
        this.title = title;
    }

    public ExportHeader(String field, String title, boolean required) {
        this.field = field;
        this.title = title;
        this.required = required;
    }

    public ExportHeader(String field, String title, String type, boolean required, Function<T, R> formatter) {
        this.field = field;
        this.title = title;
        this.type = ExportHeaderType.fromString(type);
        this.required = required;
        this.formatter = formatter;
    }

    public ExportHeader() {
    }

    public String getField() {
        return field;
    }

    public void setField(String field) {
        this.field = field;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public ExportHeaderType getType() {
        return type;
    }

    public void setType(String type) {
        this.type = ExportHeaderType.fromString(type);
    }

    public boolean isRequired() {
        return required;
    }

    public void setRequired(boolean required) {
        this.required = required;
    }

    public Function<T, R> getFormatter() {
        return formatter;
    }

    public void setFormatter(Function<T, R> formatter) {
        this.formatter = formatter;
    }

    public CellType getCellType() {
        switch (type) {
            case DATE:
                return CellType.NUMERIC;
            case TIME:
                return CellType.NUMERIC;
            case DATETIME:
                return CellType.NUMERIC;
            case NUMBER:
                return CellType.NUMERIC;
            case STRING:
                return CellType.STRING;
            case BOOLEAN:
                return CellType.BOOLEAN;
            default:
                return CellType.STRING;
        }
    }
}

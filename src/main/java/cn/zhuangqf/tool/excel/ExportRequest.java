package cn.zhuangqf.tool.excel;

import com.google.gson.Gson;
import com.google.gson.JsonArray;

import java.util.List;

/**
 * @author zhuangqf
 * @date 2018/2/9
 */
public class ExportRequest {

    private JsonArray data;
    private List<ExportHeader> headers;
    private String fileName = "导出文件.xls";
    private String sheetName;

    public JsonArray getData() {
        return data;
    }

    public void setData(Object data) {
        this.data = new Gson().toJsonTree(data).getAsJsonArray();
    }

    public List<ExportHeader> getHeaders() {
        return headers;
    }

    public void setHeaders(List<ExportHeader> headers) {
        this.headers = headers;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }
}

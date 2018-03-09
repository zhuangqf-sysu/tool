package cn.zhuangqf.tool.excel;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @author zhuangqf
 * @date 2018/2/6
 */
public class ImportResult<T> {

    private long total = 0;
    private long success = 0;
    private long failure = 0;
    private String sheetName;
    private List<ImportDetail<T>> data = new ArrayList<>();

    public void add(ImportDetail<T> detail) {
        this.total++;
        this.data.add(detail);
        if (detail.getStatus() == 0) {
            success++;
        } else {
            failure++;
        }
    }

    public long getTotal() {
        return total;
    }

    public long getSuccess() {
        return success;
    }

    public long getFailure() {
        return failure;
    }

    public List<ImportDetail<T>> getData() {
        return Collections.unmodifiableList(data);
    }

    public List<T> getSuccessData() {
        return data.stream().filter((detail) -> detail.getStatus() == 0).map(ImportDetail::getData).collect(Collectors.toList());
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }
}

package cn.zhuangqf.tool.excel;

import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author zhuangqf
 * @date 2018/2/9
 */
public class ExportWorkbook {

    private Workbook workbook;
    private Sheet sheet;
    private Row headerRow;
    private List<ExportHeader> headers;

    ExportWorkbook() {
        this.workbook = new HSSFWorkbook();
    }

    ExportWorkbook(String fileName) {
        if (null != fileName && fileName.endsWith(".xlsx")) {
            this.workbook = new XSSFWorkbook();
        } else {
            this.workbook = new HSSFWorkbook();
        }
    }

    public ExportWorkbook addSheet(String sheetName) {
        if (null == sheetName) {
            return addSheet();
        }
        this.sheet = workbook.createSheet(sheetName);
        this.headerRow = this.sheet.createRow(0);
        this.headers = new ArrayList<>();
        return this;
    }

    public ExportWorkbook addSheet() {
        this.sheet = workbook.createSheet();
        this.headerRow = this.sheet.createRow(0);
        this.headers = new ArrayList<>();
        return this;
    }

    public ExportWorkbook addHeader(ExportHeader header, int index) {
        headers.add(header);
        int length = Math.max((header.getTitle().getBytes().length + 4) * 256, 10 * 256);
        this.sheet.setColumnWidth(index, length);
        CreationHelper createHelper = workbook.getCreationHelper();
        Cell cell = headerRow.createCell(index);
        cell.setCellType(CellType.STRING);
        cell.setCellStyle(getHeaderCellStyle(header.isRequired()));
        cell.setCellValue(createHelper.createRichTextString(header.getTitle()));
        return this;
    }

    public ExportWorkbook addHeaders(List<ExportHeader> headers) {
        int begin = headerRow.getLastCellNum() + 1;
        for (int i = 0; i < headers.size(); i++) {
            addHeader(headers.get(i), i + begin);
        }
        return this;
    }

    public ExportWorkbook addData(JsonArray data) {
        int begin = sheet.getLastRowNum() + 1;
        for (int i = 0; i < data.size(); i++) {
            addData(data.get(i).getAsJsonObject(), i + begin);
        }
        return this;
    }

    public ExportWorkbook addData(JsonObject data, int index) {
        Row row = sheet.createRow(index);
        for (int i = 0; i < headers.size(); i++) {
            try {
                ExportHeader header = headers.get(i);
                Cell cell = row.createCell(i);
                JsonElement element = data.get(header.getField());
                String value = (null == element) ? "" : element.getAsString();
                cell.setCellStyle(getCellStyle(header.isRequired()));
                cell.setCellValue(value);
            } catch (Exception e) {

            }
        }
        return this;
    }

    public void save(OutputStream outputStream) throws IOException {
        workbook.write(outputStream);
    }

    /**
     * 列头单元格样式
     */
    private CellStyle getHeaderCellStyle(boolean required) {
        // 设置字体
        Font font = workbook.createFont();
        // 设置字体颜色
        if (required) {
            font.setColor(Font.COLOR_RED);
        }

        // 设置字体大小
        font.setFontHeightInPoints((short) 11);
        // 字体加粗
//            font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        font.setBold(true);
        // 设置字体名字
        font.setFontName("Courier New");
        // 设置样式;
        CellStyle style = workbook.createCellStyle();
        // 设置底边框;
//            style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderBottom(BorderStyle.THIN);
        // 设置底边框颜色;
        style.setBottomBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        HSSFColor.HSSFColorPredefined.BLACK.getIndex();
        // 设置左边框;
        style.setBorderLeft(BorderStyle.THIN);
        // 设置左边框颜色;
        style.setLeftBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        // 设置右边框;
        style.setBorderRight(BorderStyle.THIN);
        // 设置右边框颜色;
        style.setRightBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        // 设置顶边框;
        style.setBorderTop(BorderStyle.THIN);
        // 设置顶边框颜色;
        style.setTopBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        // 在样式用应用设置的字体;
        style.setFont(font);
        // 设置自动换行;
        style.setWrapText(false);
        // 设置水平对齐的样式为居中对齐;
        style.setAlignment(HorizontalAlignment.CENTER);
        // 设置垂直对齐的样式为居中对齐;
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        style.setWrapText(true);

        return style;

    }

    /**
     * 列数据信息单元格样式
     */
    private CellStyle getCellStyle(boolean isRequired) {
        // 设置字体
        Font font = workbook.createFont();
//         设置字体颜色
        if (isRequired) {
            font.setColor(Font.COLOR_RED);
        }
        // 设置字体名字
        font.setFontName("Courier New");
        // 设置样式;
        CellStyle style = workbook.createCellStyle();
        // 在样式用应用设置的字体;
        style.setFont(font);
        // 设置自动换行;
        style.setWrapText(false);

        return style;

    }

}

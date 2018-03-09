package cn.zhuangqf.tool.excel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.util.Collection;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

/**
 * @author zhuangqf
 * @date 2018/3/9
 */
public class ImportWorkbook implements Iterator<ImportWorkbook> {

    private final static SampleMap<String, String> SAMPLE_MAP = new SampleMap<>();
    private final static Function<Map<String, String>, Object> FUNCTION = map -> map;

    private Workbook workbook;
    private Sheet sheet;
    private Map<String, String> map;
    private Iterator<Sheet> sheetIterator;

    public ImportWorkbook(InputStream inputStream) throws IOException, InvalidFormatException {
        this.workbook = WorkbookFactory.create(inputStream);
        this.sheetIterator = this.workbook.sheetIterator();
    }

    private static <T> ImportResult<T> importSheet(Sheet sheet, Map<String, String> map, Function<Map<String, String>, T> function) {
        ImportResult<T> result = new ImportResult<>();
        result.setSheetName(sheet.getSheetName());
        Map<Integer, String> index =
                StreamSupport.stream(sheet.getRow(0).spliterator(), true)
                        .filter(cell -> !"".equals(map.getOrDefault(cell.getStringCellValue(), "")))
                        .collect(Collectors.toMap(
                                Cell::getColumnIndex,
                                cell -> map.get(cell.getStringCellValue())
                        ));
        StreamSupport.stream(sheet.spliterator(), true)
                .filter(row -> row.getRowNum() > 0)
                .forEach(row -> result.add(importRow(row, index, function)));
        return result;
    }

    private static <T> ImportDetail<T> importRow(Row row, Map<Integer, String> index, Function<Map<String, String>, T> function) {
        try {
            Map<String, String> rawData = StreamSupport.stream(row.spliterator(), true)
                    .collect(Collectors.toMap(
                            cell -> index.get(cell.getColumnIndex()),
                            cell -> {
                                cell.setCellType(CellType.STRING);
                                return cell.getStringCellValue();
                            },
                            (k, v) -> v
                    ));
            return new ImportDetail<T>().success(function.apply(rawData));
        } catch (Exception e) {
            e.printStackTrace();
            return new ImportDetail<T>().failure(e.getMessage());
        }
    }

    @Override
    public boolean hasNext() {
        return this.sheetIterator.hasNext();
    }

    @Override
    public ImportWorkbook next() {
        this.sheet = this.sheetIterator.next();
        this.map = SAMPLE_MAP;
        return this;
    }

    public ImportWorkbook setSheet(int i) {
        this.sheet = workbook.getSheetAt(i);
        this.map = SAMPLE_MAP;
        return this;
    }

    public ImportWorkbook setSheet(String sheetName) {
        this.sheet = workbook.getSheet(sheetName);
        this.map = SAMPLE_MAP;
        return this;
    }

    public String getSheetName() {
        return this.sheet.getSheetName();
    }

    public ImportWorkbook setMap(Map<String, String> map) {
        this.map = map;
        return this;
    }

    public ImportResult<Object> result() {
        return importSheet(this.sheet, this.map, FUNCTION);
    }

    public <T> ImportResult<T> result(Function<Map<String, String>, T> function) {
        return importSheet(this.sheet, this.map, function);
    }

    private static class SampleMap<K, V> implements Map<K, V> {

        @Override
        public int size() {
            return 0;
        }

        @Override
        public boolean isEmpty() {
            return false;
        }

        @Override
        public boolean containsKey(Object key) {
            return false;
        }

        @Override
        public boolean containsValue(Object value) {
            return false;
        }

        @Override
        public V get(Object key) {
            return (V) key;
        }

        @Override
        public V put(K key, V value) {
            return null;
        }

        @Override
        public V remove(Object key) {
            return null;
        }

        @Override
        public void putAll(Map<? extends K, ? extends V> m) {

        }

        @Override
        public void clear() {

        }

        @Override
        public Set<K> keySet() {
            return null;
        }

        @Override
        public Collection<V> values() {
            return null;
        }

        @Override
        public Set<Entry<K, V>> entrySet() {
            return null;
        }
    }

}

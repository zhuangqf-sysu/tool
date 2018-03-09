package cn.zhuangqf.tool.excel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.io.InputStream;
import java.util.Collection;
import java.util.Map;
import java.util.Set;
import java.util.function.Function;

/**
 * @author zhuangqf
 * @date 2018/2/6
 */
public class ExcelService {

    /**
     * @method importExcel
     * @param inputStream 数据源
     * @param map         可选 (header->field) 头部与ID的映射
     * @param function    可选 (map -> T) 对每行数据（map）执行操作，可以将一个数据转化为一个对象
     * @description 导入：只导入sheet[0]
     */
    public static <T> ImportResult<T> importExcel(InputStream inputStream, Map<String, String> map, Function<Map<String, String>, T> function) throws IOException, InvalidFormatException {
        return createImportWorkbook(inputStream).setSheet(0).setMap(map).result(function);
    }

    public static ImportResult importExcel(InputStream inputStream) throws IOException, InvalidFormatException {
        return createImportWorkbook(inputStream).setSheet(0).result();
    }

    public static ImportResult importExcel(InputStream inputStream, Map<String, String> map) throws IOException, InvalidFormatException {
        return createImportWorkbook(inputStream).setSheet(0).setMap(map).result();
    }

    public static ImportWorkbook createImportWorkbook(InputStream inputStream) throws IOException, InvalidFormatException {
        return new ImportWorkbook(inputStream);
    }

    public static ExportWorkbook createExportWorkbook(String fileName) {
        return new ExportWorkbook(fileName);
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

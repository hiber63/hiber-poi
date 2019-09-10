package cn.hiber.poi.util;

import cn.hiber.poi.core.RowParseAble;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.LinkedList;
import java.util.List;

/**
 * @author hiber
 */
public class ImportUtils {

    /**
     * @param importable
     * @param rowBeginIndex
     */
    public static <V> List<V> importExcelSimple(Sheet sheet, RowParseAble<V> importable, int rowBeginIndex) {
        List<V> list = new LinkedList<>();
        parseSheet(sheet, list, rowBeginIndex, importable);
        return list;
    }

    /**
     * @param sheet
     * @param dataList
     * @param rowBeginIndex
     * @param importable
     */
    private static <V> void parseSheet(Sheet sheet, List<V> dataList, int rowBeginIndex,
                                       RowParseAble<V> importable) {
        int lastRowNum = sheet.getLastRowNum();
        if(rowBeginIndex <= lastRowNum) {
            for(int i=rowBeginIndex; i<=lastRowNum; i++) {
                Row row = sheet.getRow(i);
                if(row==null) {
                    continue;//NullPointerException
                }
                V vo = importable.parseRow(row);
                if(vo != null) {
                    dataList.add(vo);
                }
            }
        }
    }

}

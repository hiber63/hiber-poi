package cn.hiber.poi.core;

import org.apache.poi.ss.usermodel.Row;
/**
 * @author hiber
 */
public interface RowParseAble<V> {

    V parseRow(Row row);

}

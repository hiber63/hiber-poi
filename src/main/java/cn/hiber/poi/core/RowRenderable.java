package cn.hiber.poi.core;

import org.apache.poi.ss.usermodel.Row;

/**
 * @author hiber
 */
public interface RowRenderable<V> {

    void renderRow(Row row, V vo, Object otherArgs);

}

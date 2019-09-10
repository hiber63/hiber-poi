package cn.hiber.poi.core;

import org.apache.poi.ss.usermodel.Sheet;


/**
 * @author hiber
 */
public interface SheetRenderable<V> {

	void renderSheet(Sheet sheet, V data);
}

package cn.hiber.poi.core;

import org.apache.poi.ss.usermodel.Sheet;

import java.util.Map;

/**
 * @author hiber
 */
public interface SheetParseAble {

	Map<String, Object> parseSheet(Sheet sheet);
	
}

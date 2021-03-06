package cn.hiber.poi.util;

import static cn.hiber.poi.constants.HiberPoiConstants.*;

import cn.hiber.poi.exception.ExcelFormatException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

/**
 * @author hiber
 */
public class FileUtils {

	public static String addPath(String... paths) {
		StringBuffer ret = new StringBuffer();
		if(paths != null && paths.length>=1) {
			for(int i=0,len=paths.length; i<len; i++) {
				ret.append(paths[i]).append(i==len-1?"":File.separator);
			}
		}
		return ret.toString();
 	}

	public static Workbook readExcel(String suffix,InputStream inputStream) {
		Workbook wb;
		try {
			if(SUFFIX_XLS.equals(suffix)) {
				wb = new HSSFWorkbook(inputStream);
			} else if(SUFFIX_XLSX.equals(suffix)) {
				wb = new XSSFWorkbook(inputStream);
			} else {
				throw new ExcelFormatException(WRONG_EXCEL_FORMAT);
			}
		} catch (IOException e) {
			throw new ExcelFormatException(WRONG_EXCEL_FORMAT);
		}
		return wb;
	}

	public static void close(Closeable... cs) {
		if(cs != null && cs.length > 0) {
			for(Closeable c:cs) {
				close(c);
			}
		}
	}

	public static void close(Closeable c) {
		if (c != null) {
			try {
				c.close();
			} catch (IOException e) {
				System.err.println(e);
			}
			c = null;
		}
	}

	/**
	 * copy row
	 *
	 * @param sourceRow
	 * @param targetRow
	 */
	public static void copyRow(Row sourceRow, Row targetRow) {
		targetRow.setHeight(sourceRow.getHeight());
		CellStyle rowStyle = targetRow.getRowStyle() ;
		if(rowStyle != null) {
			targetRow.setRowStyle(targetRow.getRowStyle());
		}
		int cols = sourceRow.getLastCellNum();
		for(int j=0; j<cols; j++) {
			Cell sourceCell = sourceRow.getCell(j) ;
			if (sourceCell != null) {
				Cell targetCell = targetRow.createCell(j) ;
				CellStyle cellStyle = sourceCell.getCellStyle() ;
				if(cellStyle != null) {
					targetCell.setCellStyle(cellStyle);
				}
			}
		}
	}

}

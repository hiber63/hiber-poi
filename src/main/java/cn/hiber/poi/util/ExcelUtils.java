package cn.hiber.poi.util;

import cn.hiber.poi.constants.HiberPoiConstants;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ExcelUtils {

    /**
     *
     * @param sheet 表
     * @param headIndex 头部索引，从1开始
     * @return
     */
    public static int countHeadCol(Sheet sheet, int headIndex) {
        Row row = sheet.getRow(headIndex-1);
        int count = row.getLastCellNum();
        while(true) {
            Cell cell = row.getCell(count - 1);
            if(cell != null) {
                cell.setCellType(CellType.STRING);
                if(StringUtils.isNotBlank(cell.getStringCellValue())) {
                    break;
                }
            }
            count--;
        }
        return count;
    }

    public static int countHeadCol(Sheet sheet) {
        return countHeadCol(sheet, HiberPoiConstants.ROW_HEAD_INDEX_DFT);
    }

}

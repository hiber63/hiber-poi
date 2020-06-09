package cn.hiber.poi.util;

import cn.hiber.poi.constants.HiberPoiConstants;
import cn.hiber.poi.core.RowRenderable;
import cn.hiber.poi.validation.HiberPoiParamAssert;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

/**
 * @author hiber
 */
public class ExportUtils {

    public static<V> void  downloadExcel(String templateUrl, RowRenderable<V> rowRenderable,
                                               List<V> dataList,
                                               OutputStream outputStream) throws IOException {
        Integer rowBeginIndex = HiberPoiConstants.ROW_BEGIN_INDEX_DFT;
        Integer colNumbers = HiberPoiConstants.COL_NUMBERS_DFT;
        downloadExcelCommon(templateUrl,rowRenderable,dataList,rowBeginIndex,colNumbers,outputStream);
    }

    public static<V> void  downloadExcelCommon(String templateUrl, RowRenderable<V> rowRenderable,
                                         List<V> dataList,Integer rowBeginIndex,Integer colNumbers,
                                         OutputStream outputStream) throws IOException {
        HiberPoiParamAssert.notBlank(templateUrl,"templateUrl");
        HiberPoiParamAssert.notNull(rowRenderable,"rowRenderable");
        HiberPoiParamAssert.notEmpty(dataList,"data");

        InputStream fis = null;
        Workbook wb = null;

        String url;
        try {
            File file = new File(templateUrl);
            if(file.exists()) {
                fis = new FileInputStream(templateUrl);
            } else {
                fis = new Object().getClass().getResourceAsStream(templateUrl);
            }
            String suffix = "."+templateUrl.substring(templateUrl.lastIndexOf(".") + 1);
            wb = FileUtils.readExcel(suffix, fis);
            Sheet sheet = wb.getSheetAt(0);
            colNumbers = colNumbers==HiberPoiConstants.COL_NUMBERS_DFT?ExcelUtils.countHeadCol(sheet, rowBeginIndex-1):colNumbers;
            renderSheet(sheet,dataList,rowBeginIndex,colNumbers,rowRenderable);
            wb.write(outputStream);
        } finally {
            FileUtils.close(wb,fis);
        }
        outputStream.flush();
    }

    public static <V> String exportExcelCommon(String templateUrl, String exportDirUrl, RowRenderable<V> rowRenderable,
                      List<V> dataList,Integer rowBeginIndex,Integer colNumbers, String fileName) throws IOException {
        HiberPoiParamAssert.notBlank(templateUrl,"templateUrl");
        HiberPoiParamAssert.notBlank(exportDirUrl,"exportDirUrl");
        HiberPoiParamAssert.notNull(rowRenderable,"rowRenderable");
        HiberPoiParamAssert.notEmpty(dataList,"data");

        FileOutputStream fos = null;
        FileInputStream fis = null;
        Workbook wb = null;

        String url;
        try {
            fis = new FileInputStream(templateUrl);
            String suffix = "."+templateUrl.substring(templateUrl.lastIndexOf(".") + 1);
            wb = FileUtils.readExcel(suffix, fis);
            Sheet sheet = wb.getSheetAt(0);
            renderSheet(sheet,dataList,rowBeginIndex,colNumbers,rowRenderable);
            File folder = new File(exportDirUrl);
            if (!folder.exists()) {
                folder.mkdirs();
            }
            url = exportDirUrl + fileName + suffix;
            fos = new FileOutputStream(url);
            wb.write(fos);
            fos.flush();
        } finally {
            FileUtils.close(fos,fis,wb);
        }
        return url;
    }

    private static <V> void renderSheet(Sheet sheet, List<V> dataList, Integer rowBeginIndex, Integer colNumbers, RowRenderable<V> rowRenderable) {
        HiberPoiParamAssert.notNull(rowRenderable,"rowRenderable");
        int index = rowBeginIndex-1;
        //colNumbers必须有值
        Row rowTemplate = sheet.getRow(index);
        if(rowTemplate == null) {
            rowTemplate = sheet.createRow(index);
            for(int i=0;i<colNumbers;i++) {
                rowTemplate.createCell(i);
            }
        }
        int rowIdx = index;
        for (V vo : dataList) {
            Row row;
            if(rowIdx != index) {
                row = sheet.createRow(rowIdx);
            } else {
                row = rowTemplate;
            }
            FileUtils.copyRow(rowTemplate, row) ;
            rowRenderable.renderRow(row, vo);
            rowIdx++;
        }
    }

    public static void renderRow(Row row, String... values) {
        renderRow(row, Arrays.asList(values));
    }

    public static void renderRow(Row row, Collection<String> values) {
        int colIndex = 0;
        for(String value : values) {
            row.getCell(colIndex++).setCellValue(value);
        }
    }

}

package cn.hiber.poi.util;

import cn.hiber.poi.core.RowRenderable;
import cn.hiber.poi.validation.HiberPoiParamAssert;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.List;

/**
 * @author hiber
 */
public class ExportUtils {

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
        Row rowTemplate;
        if(colNumbers!=null && colNumbers > 0) {
            rowTemplate = sheet.createRow(rowBeginIndex);
            for(int i=0;i<colNumbers;i++) {
                rowTemplate.createCell(i);
            }
        } else {
            rowTemplate = sheet.getRow(rowBeginIndex-1);
        }
        int rowIdx = rowBeginIndex;
        for (V vo : dataList) {
            Row row;
            if(rowIdx != rowBeginIndex) {
                row = sheet.createRow(rowIdx);
            } else {
                row = rowTemplate;
            }
            FileUtils.copyRow(rowTemplate, row) ;
            rowRenderable.renderRow(row, vo);
            rowIdx++;
        }
    }

}

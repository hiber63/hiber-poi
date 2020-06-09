package cn.hiber.poi.exp;

import cn.hiber.poi.constants.HiberPoiConstants;
import cn.hiber.poi.core.RowRenderable;
import cn.hiber.poi.util.ExcelUtils;
import cn.hiber.poi.util.ExportUtils;
import cn.hiber.poi.util.FileUtils;
import cn.hiber.poi.vo.Person;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Before;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;

public class ExportTest {

    List<Person> personList = new ArrayList<>();

    String templateUrl = "E://xx//export-template.xlsx";

    String exportDirUrl = "E://xx//";

    int rowBeginIndex = 2;

    RowRenderable<Person> rr;

    @Before
    public void preWork() {
        for(int i=0;i<100;i++) {
            Person p = new Person();
            p.setName("person"+i);
            p.setGender("man or woman");
            p.setAddr("addr"+i);
            personList.add(p);
        }
    }

    @Test
    public void exportSimplePage() throws Exception {
        rr = (row,vo) -> {
//            int colIndex = 0;
//            row.getCell(colIndex++).setCellValue(vo.getName());
//            row.getCell(colIndex++).setCellValue(vo.getGender());
//            row.getCell(colIndex++).setCellValue(vo.getAddr());
            ExportUtils.renderRow(row,vo.getName(),vo.getGender(),vo.getAddr());
        };
        ExportUtils.exportExcelCommon(templateUrl,exportDirUrl,rr,personList,rowBeginIndex,3,"test-export");
        System.out.println("done...");
    }

    @Test
    public void countCol() throws FileNotFoundException {
        FileInputStream fis = null;
        Workbook wb = null;
        String url;
        try {
            fis = new FileInputStream(templateUrl);
            String suffix = "."+templateUrl.substring(templateUrl.lastIndexOf(".") + 1);
            wb = FileUtils.readExcel(suffix, fis);
            Sheet sheet = wb.getSheetAt(0);
            int i = ExcelUtils.countHeadCol(sheet, 0);
            System.out.println(i);
        } finally {
            FileUtils.close(wb,fis);
        }

    }

}

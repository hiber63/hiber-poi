package cn.hiber.poi.imp;

import cn.hiber.poi.core.RowParseAble;
import cn.hiber.poi.util.FileUtils;
import cn.hiber.poi.util.ImportUtils;
import cn.hiber.poi.vo.Person;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.IOException;
import java.util.List;

/**
 * @author hiber
 */
public class ImportTest {

    @Test
    public void test1() throws IOException {
        Sheet sheet = new XSSFWorkbook(getClass().getClassLoader().getResourceAsStream(FileUtils.addPath("import","import-test.xlsx"))).getSheetAt(0);
        RowParseAble<Person> rpa = (row) -> {
            Person p = new Person();
            int i = 0;
            p.setName(row.getCell(i++).getStringCellValue());
            p.setGender(row.getCell(i++).getStringCellValue());
            p.setAddr(row.getCell(i).getStringCellValue());
            return p;
        };
        List<Person> list = ImportUtils.importExcelSimple(sheet, rpa, 2);
    }

}

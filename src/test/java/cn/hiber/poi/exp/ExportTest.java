package cn.hiber.poi.exp;

import cn.hiber.poi.core.RowRenderable;
import cn.hiber.poi.util.ExportUtils;
import cn.hiber.poi.vo.Person;

import org.junit.Before;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

public class ExportTest {

    List<Person> personList = new ArrayList<>();

    String templateUrl = "E://export-template.xlsx";

    String exportDirUrl = "E://";

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
            int colIndex = 0;
            row.getCell(colIndex++).setCellValue(vo.getName());
            row.getCell(colIndex++).setCellValue(vo.getGender());
            row.getCell(colIndex++).setCellValue(vo.getAddr());
        };
        ExportUtils.exportExcelCommon(templateUrl,exportDirUrl,rr,personList,rowBeginIndex,3,"test-export");
        System.out.println("done...");
    }


}

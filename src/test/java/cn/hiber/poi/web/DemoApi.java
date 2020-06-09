package cn.hiber.poi.web;

import cn.hiber.poi.core.RowRenderable;
import cn.hiber.poi.util.ExportUtils;
import cn.hiber.poi.vo.Person;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("/demo")
public class DemoApi {

    @GetMapping("download")
    public void download(HttpServletResponse response) throws IOException {
        responseExcel(response, "test_file_name.xlsx");
        List<Person> personList = new ArrayList<>();
        for(int i=0;i<100;i++) {
            Person p = new Person();
            p.setName("person"+i);
            p.setGender("man or woman");
            p.setAddr("addr"+i);
            personList.add(p);
        }
//        String templateUrl = "E:/xx/export-template.xlsx";
        String templateUrl = "/export/templates/export-template.xlsx";
        RowRenderable<Person> rr = (row,vo) -> {
            ExportUtils.renderRow(row,vo.getName(),vo.getGender(),vo.getAddr());
        };
        ExportUtils.downloadExcel(templateUrl,rr,personList,response.getOutputStream());
    }

    public HttpServletResponse responseExcel(HttpServletResponse response, String fileName) throws UnsupportedEncodingException {
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        fileName = URLEncoder.encode(fileName, "UTF-8");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName);
        return response;
    }



}

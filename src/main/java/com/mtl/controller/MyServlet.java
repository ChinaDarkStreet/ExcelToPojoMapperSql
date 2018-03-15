package com.mtl.controller;

import com.mtl.myMain.ExcelToPojo;
import com.mtl.myMain.ExcelToSql;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.IOUtils;
import org.apache.xmlbeans.impl.common.IOUtil;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;

@Controller
public class MyServlet {

    @RequestMapping("/doExcel")
    public void myMain(HttpServletResponse response, MultipartFile file) throws IOException, InvalidFormatException {
        String basePath = MyServlet.class.getResource("/").getPath();
        InputStream fileInputStream = file.getInputStream();
        File xlsx = new File(basePath, "catch.xlsx");
        if (!xlsx.exists()){
            xlsx.createNewFile();
        }
        FileOutputStream outputStream = new FileOutputStream(xlsx);
        IOUtils.copy(fileInputStream, outputStream);
        outputStream.close();
        fileInputStream.close();

        File sql = new File(basePath, "sql.txt");
        if (!sql.exists()){
            sql.createNewFile();
        }
        new ExcelToSql().myMain(xlsx, sql);

        response.setHeader("Content-Disposition", "attachment;filename=sql.txt");
        response.setContentType("application/octet-stream");
        FileInputStream inputStream = new FileInputStream(sql);
        IOUtils.copy(inputStream, response.getOutputStream());
        inputStream.close();
    }

    @RequestMapping("excelToPojo")
    public void excelToPojo(HttpServletResponse response, MultipartFile file) throws IOException, InvalidFormatException {
        String basePath = MyServlet.class.getResource("/").getPath();
        InputStream fileInputStream = file.getInputStream();
        InputStream inputStream = new ExcelToPojo().parseExcel(fileInputStream);

        response.setHeader("Content-Disposition", "attachment;filename=result.zip");
        response.setContentType("application/octet-stream");
        IOUtils.copy(inputStream, response.getOutputStream());
        inputStream.close();
    }
}

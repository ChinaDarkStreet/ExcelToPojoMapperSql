package com.mtl.myMain;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.*;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class ExcelToPojo {
    private static Map<String, String> typeMap;
    private static Map<String, String> sqlMap;

    static {
        typeMap = new HashMap<>();
        typeMap.put("字符", "String");
        typeMap.put("小数", "double");
        typeMap.put("整数", "int");

        sqlMap = new HashMap<>();
        sqlMap.put("字符", "varchar2");
        sqlMap.put("小数", "number");
        sqlMap.put("整数", "number");
    }

    private String basePath = "D:\\MTL\\IdeaProject\\ExcelToSql\\src\\main\\java\\com\\mtl\\";
    private StringBuilder sqlStr = new StringBuilder();

    @Test
    public void myTest() throws IOException, InvalidFormatException {
//        new ExcelToPojo().parseExcel(new File("D:\\MTL\\Desktop\\test.xlsx"));
//        GenerateMapper.generateSql(new File("D:\\MTL\\IdeaProject\\ExcelToSql\\src\\main\\java\\com\\mtl\\pojo\\Student.java"), new File("D:\\MTL\\IdeaProject\\ExcelToSql\\src\\main\\resources\\1.sql"));
    }

    public InputStream parseExcel(InputStream excel) throws IOException, InvalidFormatException {
        XSSFWorkbook sheets = new XSSFWorkbook(excel);
        Iterator<Sheet> iterator = sheets.sheetIterator();

        File zip = new File(basePath, "result.zip");
        if (!zip.exists()) {
            zip.createNewFile();
        }
        ZipOutputStream zipOutputStream = new ZipOutputStream(new FileOutputStream(zip));

        while (iterator.hasNext()){
            generatePojo(iterator.next(), zipOutputStream);
        }

        zipOutputStream.putNextEntry(new ZipEntry("createTable.sql"));
        zipOutputStream.write(sqlStr.toString().getBytes());
        zipOutputStream.flush();
        zipOutputStream.close();

        return new FileInputStream(zip);
//        File sql = new File(basePath, "createTable.sql");
//        if (!sql.exists())
//            sql.createNewFile();
//        FileWriter sqlWriter = new FileWriter(sql);
//        sqlWriter.write(sqlStr.toString());
//        sqlWriter.flush();
//        sqlWriter.close();

    }

    /**
     * 生成实体类的字符串的方法
     * @param next 就是当前的sheet
     * @param zipOS     输出需要zip流
     * @throws IOException      文件io流异常
     */
    private void generatePojo(Sheet next, ZipOutputStream zipOS) throws IOException {
        String pojoName = next.getSheetName();

        StringBuilder sb = new StringBuilder();
        StringBuilder end = new StringBuilder();
        StringBuilder mapperStr = new StringBuilder();
        sb.append("package com.mtl.pojo;\n\npublic class ").append(pojoName).append(" {\n\n");
        mapperStr.append("package com.mtl.mapper;\n\nimport org.apache.ibatis.annotations.Insert;\nimport com.mtl.pojo.").append(pojoName)
                .append(";\n\npublic interface ").append(pojoName).append("Mapper {\n    @Insert(\"insert into ").append(pojoName)
                .append(" values (");
        sqlStr.append("\ncreate table ").append(pojoName).append("\n(\n");

        for (int i = 2; i <= next.getLastRowNum(); i++) {
            Row row = next.getRow(i);
            String fieldName = row.getCell(1).getStringCellValue();
            String fieldType = row.getCell(3).getStringCellValue();
            Cell cell4 = row.getCell(4);
            cell4.setCellType(CellType.STRING);
            String fieldLength = cell4.getStringCellValue();
            String firstToLower = firstToLower(fieldName);
            String type = typeMap.get(fieldType);
            sb.append("    private ").append(type).append(" ").append(firstToLower).append(";\n");
            end.append("    public ").append(type).append(" get").append(fieldName).append("() {\n").append("        return ")
                    .append(firstToLower).append(";\n    }\n\n    public void set").append(fieldName)
                    .append("(").append(type).append(" ").append(firstToLower).append(") {\n        this.").append(firstToLower)
                    .append(" = ").append(firstToLower).append(";\n    }\n\n");
            if (i == 2){
                mapperStr.append("#{").append(firstToLower).append("}");
                sqlStr.append("  ").append(firstToLower).append(" ").append(sqlMap.get(fieldType)).append("(").append(fieldLength).append(")");
            } else {
                mapperStr.append(", #{").append(firstToLower).append("}");
                sqlStr.append(",\n  ").append(firstToLower).append(" ").append(sqlMap.get(fieldType)).append("(").append(fieldLength).append(")");
            }
        }

        mapperStr.append(")\")\n    int insert(Class object);\n}");
        sqlStr.append("\n)\n;\n");
        sb.append("\n");
        end.append("}\n");

        zipOS.putNextEntry(new ZipEntry("mapper/"+pojoName+"Mapper.java"));
        zipOS.write(mapperStr.toString().getBytes());
        zipOS.putNextEntry(new ZipEntry("pojo/" + pojoName + ".java"));
        zipOS.write(sb.append(end.toString()).toString().getBytes());
        zipOS.flush();
//        File pojo = new File(basePath, "pojo/" + pojoName + ".java");
//        File mapper = new File(basePath, "mapper/" + pojoName + "Mapper.java");
//        if (!mapper.exists())
//            mapper.createNewFile();
////        System.out.println("path = " + pojo.getPath());
//        if (!pojo.exists())
//            pojo.createNewFile();
//
//        FileWriter fileWriter = new FileWriter(mapper);
//        fileWriter.write(mapperStr.toString());
//        fileWriter.flush();
//        fileWriter.close();
//
//        FileWriter writer = new FileWriter(pojo);
//        writer.write(sb.toString() + end.toString());
//        writer.flush();
//        writer.close();

    }

    private String firstToLower(String fieldName) {
        return fieldName.substring(0, 1).toLowerCase() + fieldName.substring(1);
    }
}

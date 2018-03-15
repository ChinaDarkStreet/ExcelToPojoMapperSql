package com.mtl.myMain;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.*;

public class ExcelToSql {
    private static Map<String, String> map = new TreeMap<String, String>();
    Map<String, Integer> repeatList = new TreeMap<>();
    String sheetName;

    static {
        map.put("0", "零");
        map.put("1", "一");
        map.put("2", "二");
        map.put("3", "三");
        map.put("4", "四");
        map.put("5", "五");
        map.put("6", "六");
        map.put("7", "七");
        map.put("8", "八");
        map.put("9", "九");
    }

    public void myMain(File file, File targerFile) throws IOException, InvalidFormatException {
        XSSFWorkbook sheets = new XSSFWorkbook(file);
        Iterator<Sheet> sheetIterator = sheets.sheetIterator();
        parseSheets(sheetIterator, targerFile);
        sheets.close();
    }

    private void parseSheets(Iterator<Sheet> sheetIterator, File targerFile) throws IOException {
        FileWriter writer = new FileWriter(targerFile);
        writer.write("");
        writer.flush();
        writer.close();

        Sheet next = sheetIterator.next();
        this.sheetName = next.getSheetName();

        int lastRowNum = next.getLastRowNum();
        Row row = null;
        ArrayList<String> enFieldList = new ArrayList<>();
        ArrayList<String> cnFieldList = new ArrayList<>();
        String tableName = "";
        String cnTableName = "";
        String enName;
        String cnName;
        String enField;
        String cnField;
        for (int i = 3; i < lastRowNum + 1; i++) {
            System.out.println("row count = " + i);
            row = next.getRow(i);
            enName =  doStr(row.getCell(3).getStringCellValue());
            cnName =  doStr(row.getCell(4).getStringCellValue());
            enField = doStr(row.getCell(6).getStringCellValue());
            cnField = doStr(row.getCell(7).getStringCellValue());

            if (i == 24){
                System.out.println("hah ");
            }
            if (!tableName.equals(enName)){
                generateSQL(enFieldList, cnFieldList, tableName, cnTableName, targerFile);
                tableName = enName;
                cnTableName = cnName;
                enFieldList.clear();
                cnFieldList.clear();
                repeatList.clear();
            }

            //如果中文字段或者英文字段为空, 直接跳过添加
            if (cnField == null) {
                continue;
            }
            enFieldList.add(enField);

            //判断是不是有同名中文字段
            int strIndex = getStrIndex(cnFieldList, cnField);
            Integer repeatIndex = repeatList.get(cnField);
            if (strIndex >-1){
                String value = cnFieldList.get(strIndex);
                repeatList.put(value, 2);
                cnFieldList.set(strIndex, value + 0);
                cnFieldList.add(cnField + 1);
            }else if (repeatIndex != null){
                Integer size = repeatList.get(cnField);
                cnFieldList.add(cnField + size);                    //添加字段
                repeatList.put(cnField, size + 1);                  //重复字段个数  +1
            }else {
                cnFieldList.add(cnField);
            }
        }
    }

    private void generateSQL(ArrayList<String> enFieldList, ArrayList<String> cnFieldList, String tableName, String cnName, File targerFile) {
        if (enFieldList.size() == 0){
            return;
        }
        try {
            FileWriter writer = new FileWriter(targerFile, true);
            StringBuilder sb = new StringBuilder();
            sb.append(cnName).append(":\n").append("select\n");
            for (int i = 0; i < cnFieldList.size(); i++) {
                String enField = enFieldList.get(i);
                String cnField = cnFieldList.get(i);
                if (enField.equals(cnField)){
                    if (i == cnFieldList.size() - 1){
                        sb.append("    ").append(enField).append("\n");
                    }else {
                        sb.append("    ").append(enField).append(",\n");
                    }
                }else {
                    if (i == cnFieldList.size() - 1){
                        sb.append("    ").append(enField).append(" as ").append(cnField).append("\n");
                    }else {
                        sb.append("    ").append(enField).append(" as ").append(cnField).append(",\n");
                    }
                }
            }
            sb.append("from ").append(tableName).append(";\n").append("store * from ").append(cnName).append(" into ../").append(sheetName)
                    .append("/").append(tableName).append("_").append(cnName).append(".qvd(qvd);\n").append("drop table ").append(cnName).append(";\n\n");

            writer.write(sb.toString());
            writer.flush();
            writer.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private String doStr(String old){
        if (old.equals("")){
            return null;
        }
        String newStr = old.replaceAll("\\(", "_").replaceAll("\\)", "").replaceAll("<", "_").replaceAll(">", "").replaceAll("/", "_").replaceAll("（", "_").replaceAll("）", "").replaceAll("\\[", "_").replaceAll("\\]", "");
        newStr = doFirstChar(newStr);
        return newStr;
    }

    private String doFirstChar(String str) {
        String first = str.substring(0, 1);
        first = map.get(first);
        if (first != null) {
            return first + str.substring(1);
        }
        return str;
    }

    private int getStrIndex(List list, String value){
        for (int i = 0; i < list.size(); i++) {
            if (value.equals(list.get(i))){
                return i;
            }
        }
        return -1;
    }
}

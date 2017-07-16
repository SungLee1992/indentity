package com.cie.indetity;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

    public static void main(String[] args) throws FileNotFoundException, InvalidFormatException, IOException {
        // TODO Auto-generated method stub
        ExcelUtils excel = new ExcelUtils();
        File trueFile = new File("D:/2.xlsx");
        //File falseFile = new File("D:/false.xlsx");
        File srcFile = new File("D:/source.xlsx");
        List<ArrayList<String>> strLists = new ArrayList<ArrayList<String>>();
        List<ArrayList<String>> trueLists = new ArrayList<ArrayList<String>>();
        // List<ArrayList<String>> falseLists = new ArrayList<ArrayList<String>>();
        XSSFWorkbook tbook = new XSSFWorkbook();// 创建工作文档对象
        Sheet tsheet = tbook.createSheet("sheet1");// 创建工作簿
        // XSSFWorkbook fbook = new XSSFWorkbook();// 创建工作文档对象
        // Sheet fsheet = fbook.createSheet("sheet1");// 创建工作簿

        Row tRow, fRow;
        Cell cell;

        System.out.println("开始读源文件===========================");
        strLists = excel.poiReadXExcel(srcFile);

        System.out.println("开始验证身份证号码=======================");
        int i = 1;
        // int j = 1;
        for (ArrayList<String> strList : strLists) {
            boolean flag = IdcardUtils.validateCard(strList.get(3));
            if (!flag) {
                //正确
                //trueLists.add(strList);
                // 数据
                tRow = tsheet.createRow(i); //从第二行开始
                for (int c = 0; c < strList.size(); c++) {
                    //写入数据
                    cell = tRow.createCell(c); // 创建数据列
                    cell.setCellValue(strList.get(c)); // 赋值
                }
                i++;
            }
            /*else {
                //错误
                //falseLists.add(strList);
                fRow = fsheet.createRow(j); //从第二行开始
                for (int c = 0; c < strList.size(); c++) {
                    //写入数据
                    cell = fRow.createCell(c); // 创建数据列
                    cell.setCellValue(strList.get(c)); // 赋值
                }
                j++;
            }*/
        }

        System.out.println("开始写入文件==========================");
        //excel.exportExcel(trueLists, trueFile);
        // excel.exportExcel(falseLists, falseFile);
        System.out.println("准备导出Excel.............");
        //导出Excel
        // 第一种方式：写入文件
        try {
            // 创建文件流
            OutputStream tout = new FileOutputStream(trueFile);
            // 写入数据
            tbook.write(tout);
            // 创建文件流
            //OutputStream fout = new FileOutputStream(trueFile);
            // 写入数据
            //fbook.write(fout);
            // 关闭文件流
            tout.flush();
            tout.close();
            // 关闭文件流
            //fout.flush();
            //fout.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

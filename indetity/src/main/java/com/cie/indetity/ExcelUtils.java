package com.cie.indetity;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
    private String           val      = null;
    private SimpleDateFormat fmt      = new SimpleDateFormat("yyyy-MM-dd"); //日期格式yyyy-mm-dd
    private DecimalFormat    df       = new DecimalFormat("0");            //数字格式，防止长数字成为科学计数法形式，或者int变为double形式
    List<ArrayList<String>>  strLists = new ArrayList<ArrayList<String>>(); //存放Excel中的数据

    /**
     * @author: 李嵩
     * 2016-8-27 下午7:04:26
     * @return List<ArrayList<String>>
     * @throws FileNotFoundException,IOException,InvalidFormatException
     * TODO POI方式解析EXCEL表格(2007以上版本，xlsx)
     */
    public List<ArrayList<String>> poiReadXExcel(File srcfile) throws FileNotFoundException, IOException, InvalidFormatException {
        FileInputStream input = new FileInputStream(srcfile); //读取的文件路径   
        XSSFWorkbook wb = new XSSFWorkbook(new BufferedInputStream(input));
        XSSFSheet sheet = wb.getSheetAt(0); //获取第一张表

        int rowNum = sheet.getPhysicalNumberOfRows();//得到数据的行数
        System.out.println("行数：" + rowNum);
        strLists.clear();

        //遍历行
        for (int i = 20000; i < rowNum; i++) //Excel从第2行开始有数据
        {
            List<String> strList = new ArrayList<String>();
            XSSFRow row = sheet.getRow(i);
            int colNum = row.getPhysicalNumberOfCells();//得到当前行中存在数据的列数
            //遍历列
            for (int j = 0; j < colNum; j++) {
                XSSFCell cell = row.getCell(j);
                if(cell==null){
                    System.out.println("第"+(i+1)+"行,"+"第"+(j+1)+"列");
                    continue;
                }
                strList.add(getXCellVal(cell));
            }
            strLists.add((ArrayList<String>) strList); //存储该行
        }

        //打印
        /*for (ArrayList<String> stringList : strLists) {
            for (String str : stringList) {
                System.out.print(str + "  ");
            }
            System.out.println();
        }*/
        return strLists;
    }

    /**
     * @author: 李嵩
     * 2016-8-25 下午7:40:05
     * @param cell
     * @return String
     * TODO 获取单元格中的值
     */
    private String getXCellVal(XSSFCell cell) {
        switch (cell.getCellType()) {
            case XSSFCell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    val = fmt.format(cell.getDateCellValue()); //日期型
                } else {
                    val = df.format(cell.getNumericCellValue()); //数字型
                }
                break;
            case XSSFCell.CELL_TYPE_STRING: //文本类型
                val = cell.getStringCellValue();
                break;
            case XSSFCell.CELL_TYPE_BOOLEAN: //布尔型
                val = String.valueOf(cell.getBooleanCellValue());
                break;
            case XSSFCell.CELL_TYPE_BLANK: //空白
                val = cell.getStringCellValue();
                break;
            case XSSFCell.CELL_TYPE_ERROR: //错误
                val = "错误";
                break;
            case XSSFCell.CELL_TYPE_FORMULA: //公式
                try {
                    val = String.valueOf(cell.getStringCellValue());
                } catch (IllegalStateException e) {
                    val = String.valueOf(cell.getNumericCellValue());
                }
                break;
            default:
                val = cell.getRichStringCellValue() == null ?  "null" : cell.getRichStringCellValue().toString();
        }
        return val;

    }

    public void exportExcel(List<ArrayList<String>> strLists,File file) {
        XSSFWorkbook book = new XSSFWorkbook();// 创建工作文档对象
        Sheet sheet = book.createSheet("sheet1");// 创建工作簿

        Row row;
        Cell cell;

        row = sheet.createRow(0);// 创建表头行 
        // 数据
        for (int i = 0; i < strLists.size(); i++) {
            row = sheet.createRow(i+1);     //从第二行开始
            for (int c = 0; c < strLists.get(i).size(); c++) {
              //写入数据
                cell = row.createCell(c);   // 创建数据列
                cell.setCellValue(strLists.get(i).get(c));      // 赋值
            }
        }
        System.out.println("准备导出Excel.............");
        //导出Excel
        // 第一种方式：写入文件
        try {
            // 创建文件流
            OutputStream out = new FileOutputStream(file);
            // 写入数据
            book.write(out);
            // 关闭文件流
            out.flush();
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}

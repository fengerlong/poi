package com.poi.utils.poidemo.demo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 创建工作簿，使用工作簿（workbook）对excel建模
 * 工作簿 —— 也就是整个excel
 */
public class WorkBookDemo {

    public static void main(String[] args) throws IOException {
        //后缀为.xls
        Workbook excel1997 = new HSSFWorkbook();
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        excel1997.write(fileOut);
        fileOut.close();

        //后缀为.xlsx
        Workbook excel2007 = new XSSFWorkbook(); // excel 2007
        fileOut = new FileOutputStream("workbook.xlsx");
        excel2007.write(fileOut);
        fileOut.close();
    }
}

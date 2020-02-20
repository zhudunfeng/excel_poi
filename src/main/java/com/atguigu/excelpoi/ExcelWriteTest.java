package com.atguigu.excelpoi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

/**
 * @author ADun
 * @date 2020/2/20
 */
public class ExcelWriteTest {

    @Test
    public void textWrite03() throws IOException {
        //创建工作簿
        Workbook wb = new HSSFWorkbook();

        //创建工作单
        Sheet sheet1 = wb.createSheet("new sheet");
        Sheet sheet2 = wb.createSheet("second sheet");

        //创建行对象（第0行）
        Row row = sheet1.createRow(0);

        //创建单元格
        Cell cell = row.createCell(0);
        Cell cel2 = row.createCell(1);
        Cell cel3 = row.createCell(2);


        cell.setCellValue("hello excel");
        cel2.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));
        cel3.setCellValue("138888888");


        //创建输出流
        FileOutputStream out = new FileOutputStream("E:/project/190805/excel_poi/test-write03.xls");

        wb.write(out);

        out.close();

        System.out.println("文件生成成功");
    }
}

package com.qingfeng.test;


import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

public class PoiExcelTest {
    public static final String SAMPLE_XLSX_FILE_PATH="form.xlsx";
    /**
     * 通过POI库解析Excel表格
     *
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {
        //通过工厂创建工作簿
        Workbook workbook = WorkbookFactory.create(new FileInputStream(new File(SAMPLE_XLSX_FILE_PATH)));
        /*
         * getNumberOfSheets()方法：用于获取表格数量
         * getSheetAt(i)方法：用于获取指定的表格(下标从0开始)
         * */
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            //打开对应表格
            Sheet sheet = workbook.getSheetAt(i);
            //通过迭代器遍历表格
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                //获取表格的行
                Row row = rowIterator.next();
                //通过迭代器遍历行
                Iterator<Cell> cellIterator = row.iterator();
                while (cellIterator.hasNext()) {
                    //获取单元格内容
                    Cell cell = cellIterator.next();
                    //格式化单元格内容
                    DataFormatter dataFormatter = new DataFormatter();
                    String cellValue = dataFormatter.formatCellValue(cell);
                    System.out.print(cellValue + "\t");


                }


            }
        }
        workbook.close();
    }
}

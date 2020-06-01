package com.qingfeng.test;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
public class PoiExcelTest2 {
    public static final String SAMPLE_XLSX_FILE_PATH = "form.xlsx";
    /**
     * 通过POI库解析Excel表格
     *
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {
        //通过工厂创建工作簿
        Workbook workbook = WorkbookFactory.create(new FileInputStream(new File(SAMPLE_XLSX_FILE_PATH)));
        //获取第一个表格
        Sheet sheet = workbook.getSheetAt(0);
        DataFormatter dataFormatter = new DataFormatter();
        //打印单元格内容
        sheet.forEach(row -> {
            row.forEach(cell -> {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            });
            System.out.println();
        });
        //关掉工作簿
        workbook.close();
    }
}

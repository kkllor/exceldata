package com.kkllor.exceldata;

import com.kkllor.exceldata.entity.ResultBean;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ReadWriteExcelFile {

    public static void readXLSFile() throws IOException {
        InputStream ExcelFileToRead = new FileInputStream("C:/Test.xls");
        HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);

        HSSFSheet sheet = wb.getSheetAt(0);
        HSSFRow row;
        HSSFCell cell;

        Iterator rows = sheet.rowIterator();

        while (rows.hasNext()) {
            row = (HSSFRow) rows.next();
            Iterator cells = row.cellIterator();

            while (cells.hasNext()) {
                cell = (HSSFCell) cells.next();

                if (cell.getCellType() == CellType.STRING) {
                    System.out.print(cell.getStringCellValue() + " ");
                } else if (cell.getCellType() == CellType.NUMERIC) {
                    System.out.print(cell.getNumericCellValue() + " ");
                } else {
                    // U Can Handel Boolean, Formula, Errors
                }
            }
            System.out.println();
        }

    }

    public static void writeXLSFile() throws IOException {

        String excelFileName = "C:/Test.xls";// name of excel file

        String sheetName = "Sheet1";// name of sheet

        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet(sheetName);

        // iterating r number of rows
        for (int r = 0; r < 5; r++) {
            HSSFRow row = sheet.createRow(r);

            // iterating c number of columns
            for (int c = 0; c < 5; c++) {
                HSSFCell cell = row.createCell(c);

                cell.setCellValue("Cell " + r + " " + c);
            }
        }

        FileOutputStream fileOut = new FileOutputStream(excelFileName);

        // write this workbook to an Outputstream.
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
    }

    public static List<ResultBean> readXLSXFile(String fileName) throws IOException {
        List<ResultBean> resultBeanList = new ArrayList<>();


        InputStream ExcelFileToRead = new FileInputStream(fileName);
        XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);

        XSSFSheet sheet = wb.getSheetAt(0);
        XSSFRow row;
        XSSFCell cell;

        Iterator rows = sheet.rowIterator();
        long index = 0;
        while (rows.hasNext()) {
            row = (XSSFRow) rows.next();
            Iterator cells = row.cellIterator();
            while (cells.hasNext()) {
                ResultBean resultBean = new ResultBean();
                cell = (XSSFCell) cells.next();

                if (cell.getCellType() == CellType.STRING) {
                    resultBean.setOriginData(cell.getStringCellValue());
                    resultBean.setIndex(index);
                    resultBeanList.add(resultBean);
                }

                index++;
            }
        }
        return resultBeanList;
    }

    public static void writeXLSXFile(String fileName, List<ResultBean> resultBeanList) throws IOException {

        String sheetName = "result";// name of sheet

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);

        // iterating r number of rows
        for (int r = 0; r < resultBeanList.size(); r++) {
            ResultBean resultBean = resultBeanList.get(r);
            XSSFRow row = sheet.createRow(r);
            // iterating c number of columns
            for (int c = 0; c < 4; c++) {
                XSSFCell cell = row.createCell(c);
                if (c == 0) {
                    cell.setCellValue(resultBean.getOriginData());
                } else if (c == 1) {
                    cell.setCellValue(resultBean.getName());
                } else if (c == 2) {
                    cell.setCellValue(resultBean.getRange());
                } else if (c == 3) {
                    cell.setCellValue(resultBean.getType());
                }
            }
        }

        FileOutputStream fileOut = new FileOutputStream(fileName);

        // write this workbook to an Outputstream.
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
    }
}

package com.company;

import org.apache.poi.xssf.eventusermodel.XSSFBReader;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;

public class Main {


// Creates a new workbook based on file workbook.xls
    SXSSFWorkbook newSheet(){

        try {
// File name and output declaration.
            //TO DO: Change output path.
            File file = new File("/Users/zachsharpe/IdeaProjects/Metric/ExcelOutput");
            FileOutputStream out = new FileOutputStream(file);

// Workbook, sheet, row, and cell creation
            SXSSFWorkbook workbook = new SXSSFWorkbook();
            SXSSFSheet sheet = workbook.createSheet();
            SXSSFRow row = sheet.createRow(0);
            workbook.write(out);
            return workbook;

        }catch (Exception exception){
            System.out.println(exception);
            exception.printStackTrace();
            return null;
        }

    }
    public static void main(String[] args) {

        Main workbook = new Main();
        workbook.newSheet();

    }
}

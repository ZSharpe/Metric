package com.company;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileView;
import javax.swing.plaf.FileChooserUI;
import java.io.File;
import java.io.FileOutputStream;

public class Main {

    public Main() {
    }

    // Read Excel File and places information into an array.
    private void readSheet(){


    }

    // Creates a new workbook based on file workbook.xls
    HSSFWorkbook createSheet(){

        try {
// File name and output declaration.
            //TO DO: Change output path.
            File file = new File("ExcelFile5.xls");
            FileOutputStream out = new FileOutputStream(file);

// Workbook, sheet, row, and cell creation
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet();
            HSSFRow row = sheet.createRow(0);
            HSSFCell cell = row.createCell(0);
            cell.setCellValue("Zach");
            workbook.write(out);
            return workbook;

        }catch (Exception exception){
            System.out.println("Exception: " + exception);
            exception.printStackTrace();
            return null;
        }


    }
    public static void main(String[] args) {

        Main workbook = new Main();
        workbook.createSheet();
        workbook.readSheet();

    }
}

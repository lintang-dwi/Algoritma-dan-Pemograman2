/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Pertemuan_7;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author Lintang Dwi
 */
public class Read_xls {

    public static void main(String[] args) {

    }

    public static void readFromExcel(String urlexcel) throws FileNotFoundException, IOException {

        HSSFWorkbook myexcel = new HSSFWorkbook(new FileInputStream(urlexcel));
        HSSFSheet myexcelSheet = myexcel.getSheet("asas");
        FormulaEvaluator formulaEv = myexcel.getCreationHelper().createFormulaEvaluator();

        for (Row row : myexcelSheet) {
            for (org.apache.poi.ss.usermodel.Cell cell : row) {
                switch (formulaEv.evaluateInCell(cell).getCellType()) {
                      case Cell.CELL_TYPE_NUMERIC:
                        System.out.println(cell.getNumericCellValue() + "\t\t");
                        break;
                    case Cell.CELL_TYPE_STRING:
                        System.out.println(cell.getStringCellValue() + "\t\t");
                        break;
                        
                }
            }
        }
    }
}

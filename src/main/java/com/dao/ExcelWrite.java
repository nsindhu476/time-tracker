
package com.dao;

import com.models.EmployeeConnectionDetails;
import com.models.EmployeeDetails;



import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import static com.util.Constants.FILE_NAME;
import static com.util.Constants.FILE_NAME1;
import static com.util.Constants.LID_DETAILS_XLSX;
import static com.util.Constants.columnHeaders;
import static com.util.Constants.columnHeaders1;

@Component
public class ExcelWrite {
  

    public void write(EmployeeConnectionDetails empDetails) throws IOException {
        // Create a Workbook
        Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

        // Create a Sheet
        Sheet sheet = workbook.createSheet("ConnectionDetails");

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.BLUE.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        // Create header cells
        for (int i = 0; i < columnHeaders1.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columnHeaders1[i]);
            cell.setCellStyle(headerCellStyle);
        }

    

        // Create first entry after header in the sheet
        Row row = sheet.createRow(1);
        row.createCell(0).setCellValue(empDetails.getLid());
        Cell Cell = row.createCell(1);
        Cell.setCellValue(empDetails.getEmail());
        Cell.setCellType(CellType.STRING);

        Cell Cell2 = row.createCell(2);
        Cell2.setCellType(CellType.STRING);
        Cell2.setCellValue(empDetails.getContype());

        // Resize all columns to fit the content size
        for (int i = 0; i < columnHeaders1.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = null;

        fileOut = new FileOutputStream(FILE_NAME1);
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();

    }

    public void update(EmployeeConnectionDetails empDetails, String lid,String email,String contype) throws IOException, InvalidFormatException {
        // Obtain a workbook from the excel file
        FileInputStream file = new FileInputStream(FILE_NAME1);
        Workbook workbook = WorkbookFactory.create(file);

        if (Objects.nonNull(workbook)) {
            // Get Sheet at index 0
            boolean lidExists = false;
            Sheet sheet = workbook.getSheetAt(0);
            if (Objects.nonNull(sheet)) {
                for (Row myrow : sheet) {
                    Cell lidCell = myrow.getCell(0);
                    lidCell.setCellType(CellType.STRING);
                    if (empDetails.getLid().equalsIgnoreCase(String.valueOf(lidCell))) {
                        lidExists = true;
                        
                            Cell cell = myrow.getCell(1);
                            if (cell == null) {
                                cell = myrow.createCell(1);
                                cell.setCellType(CellType.STRING);
                                cell.setCellValue(empDetails.getEmail());
                            }
                            
                  
                            Cell cell2 = myrow.getCell(2);
                            if (cell2 == null)
                                cell2 = myrow.createCell(2);
                            cell2.setCellType(CellType.STRING);
                            cell2.setCellValue(empDetails.getContype());
                       
                    }
                }
            } else {
                // Create a Sheet
                sheet = workbook.createSheet();

                // Create a Font for styling header cells
                Font headerFont = workbook.createFont();
                headerFont.setBold(true);
                headerFont.setFontHeightInPoints((short) 14);
                headerFont.setColor(IndexedColors.RED.getIndex());

                // Create a CellStyle with the font
                CellStyle headerCellStyle = workbook.createCellStyle();
                headerCellStyle.setFont(headerFont);

                // Create a Row
                Row headerRow = sheet.createRow(0);

                // Create header cells
                for (int i = 0; i < columnHeaders1.length; i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(columnHeaders1[i]);
                    cell.setCellStyle(headerCellStyle);
                }
            }

            if (!lidExists) {
                int rowNum = sheet.getPhysicalNumberOfRows();
                Row row = sheet.createRow(rowNum);

                row.createCell(0).setCellValue(empDetails.getLid());

                Cell Cell = row.createCell(1);
               Cell.setCellValue(empDetails.getEmail());
                Cell.setCellType(CellType.STRING);

                Cell Cell2 = row.createCell(2);
               Cell2.setCellType(CellType.STRING);
               Cell2.setCellValue(empDetails.getContype());

                // Resize all columns to fit the content size
                for (int i = 0; i < columnHeaders1.length; i++) {
                    sheet.autoSizeColumn(i);
                }

            }
            // Write the output to the file
            FileOutputStream fileOut = null;
            file.close();
            fileOut = new FileOutputStream(FILE_NAME1);
            workbook.write(fileOut);
            fileOut.close();

            // Closing the workbook
            workbook.close();
        }
    }

   
}
package com.hdfc.reportgeneration.helper;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

public class ExcelHelper {
    public static String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    public static String TYPE1 ="application/vnd.ms-excel";
    static String[] HEADERs = { "Id", "Title", "Description", "Published" };
    static String SHEET = "Tutorials";

    public static boolean hasExcelFormat(MultipartFile file) {

        if (!TYPE.equals(file.getContentType())) {
            return false;
        }

        return true;
    }

//    public static ByteArrayInputStream tutorialsToExcel(List<Map<String,Object>> tutorials) {
//
//        try (Workbook workbook = new HSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream();) {
//            Sheet sheet = workbook.createSheet(SHEET);
//
//            // Header
//            Row headerRow = sheet.createRow(0);
//
//            for (int col = 0; col < HEADERs.length; col++) {
//                Cell cell = headerRow.createCell(col);
//                cell.setCellValue(HEADERs[col]);
//            }
//
//            int rowIdx = 1;
//            for (Tutorial tutorial : tutorials) {
//                Row row = sheet.createRow(rowIdx++);
//
//                row.createCell(0).setCellValue(tutorial.getId());
//                row.createCell(1).setCellValue(tutorial.getTitle());
//                row.createCell(2).setCellValue(tutorial.getDescription());
//                row.createCell(3).setCellValue(tutorial.isPublished());
//            }
//
//            workbook.write(out);
//            return new ByteArrayInputStream(out.toByteArray());
//        } catch (IOException e) {
//            throw new RuntimeException("fail to import data to Excel file: " + e.getMessage());
//        }
//    }
    public XSSFWorkbook excelUpload(InputStream is,String type) {
    try {
        Workbook workbook=null;
        if(type.equalsIgnoreCase("xlsx"))
             workbook = new XSSFWorkbook(is);
        else
             workbook = new HSSFWorkbook(is);
        int sheetsInWorkbook = workbook.getNumberOfSheets();
        System.out.println(sheetsInWorkbook);
        return  verifyDataInExcelBookAllSheets(workbook,sheetsInWorkbook);

    } catch (IOException e) {
        throw new RuntimeException("fail to parse Excel file: " + e.getMessage());
    }
    catch (Exception e){
        e.printStackTrace();
    }
    return null;
    }

    private XSSFWorkbook verifyDataInExcelBookAllSheets(Workbook workbook,int sheetCounts) throws IOException {
        XSSFWorkbook wb=new XSSFWorkbook();
//        CellStyle style=wb.createCellStyle();
//        style.setFillForegroundColor(IndexedColors.RED.getIndex());

        XSSFCellStyle style=wb.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.RED.getIndex());
        style.setFillPattern(FillPatternType.DIAMONDS);
        Sheet s1 = workbook.getSheetAt(0);
        Sheet s2 = workbook.getSheetAt(1);
        Sheet sheet = wb.createSheet("report");
        //System.out.println("*********** Sheet Name : " + s1.getSheetName() + "*************" + s2.getSheetName());
        int rowCounts = s2.getPhysicalNumberOfRows();
        for (int j = 0; j < rowCounts; j++) {
            Row row = sheet.createRow(j);
            // Iterating through each cell
            int cellCounts = s2.getRow(j).getPhysicalNumberOfCells();
            for (int k = 0; k < cellCounts; k++) {
                Cell c1 = s1.getRow(j).getCell(k);
                Cell c2 = s2.getRow(j).getCell(k);
                //Cell cell = row.createCell(k);
                XSSFCell cell= (XSSFCell) row.createCell(k);


                if (c2.getCellType().equals(c1.getCellType())) {
                    getCell(c1,c2,cell,style);

                } else
                {
                    // If cell types are not same, exit comparison
                    System.out.println("Non matching cell type.");
                }

            }
        }


//        FileOutputStream file = new FileOutputStream("d:\\report\\style.xlsx");
//        wb.write(file);
//        file.close();
        return wb;
    }
    private void getCell(Cell c1,Cell c2,XSSFCell cell,XSSFCellStyle style){

        if (c2.getCellType() == CellType.STRING) {
            String v1 = c1.getStringCellValue();
            String v2 = c2.getStringCellValue();
            if(!v1.equals(v2)) {

                cell.setCellStyle(style);
            }

                cell.setCellValue(c2.getStringCellValue());

            //System.out.println("Its matched : "+ v1 + " === "+ v2);
        }
        if (c2.getCellType() == CellType.NUMERIC) {
            // If cell type is numeric, we need to check if data is of Date type
            if (DateUtil.isCellDateFormatted(c1) | DateUtil.isCellDateFormatted(c2)) {
                // Need to use DataFormatter to get data in given style otherwise it will come as time stamp
                DataFormatter df = new DataFormatter();
                String v1 = df.formatCellValue(c1);
                String v2 = df.formatCellValue(c2);
                if(!v1.equals(v2)) {
                    cell.setCellStyle(style);
                }
                cell.setCellValue(c2.getNumericCellValue());
               // System.out.println("Its matched : "+ v1 + " === "+ v2);
            } else {
                double v1 = c1.getNumericCellValue();
                double v2 = c2.getNumericCellValue();
                if(v2 !=v1) {
                    cell.setCellStyle(style);
                }
                cell.setCellValue(c2.getNumericCellValue());
               // System.out.println("Its matched : "+ v1 + " === "+ v2);
            }
        }
        if (c2.getCellType() == CellType.BOOLEAN) {
            boolean v1 = c1.getBooleanCellValue();
            boolean v2 = c2.getBooleanCellValue();
            cell.setCellValue(c2.getBooleanCellValue());
            //System.out.println("Its matched : "+ v1 + " === "+ v2);
        }
    }
    private void downloadWorkbook(){

    }
}

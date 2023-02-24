package IBM2;

//import org.apache.poi.hssf.record.formula.functions.Cell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class WorkBookBean {

    public Map<Integer, List<String>> readExcel(String fileLocation) throws Exception {

        FileInputStream fis = new FileInputStream(new File(fileLocation));
//creating workbook instance that refers to .xls file
        HSSFWorkbook wb = new HSSFWorkbook(fis);
//creating a Sheet object to retrieve the object
        HSSFSheet sheet = wb.getSheetAt(0);
//evaluating cell type
        Integer i = 0;
        Map<Integer, List<String>> data = new HashMap<Integer, List<String>>();
        FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
        for (Row row : sheet)     //iteration over row using for each loop
        {
            List<String> newList = new ArrayList<String>();
            data.put(i, newList);
            List<String> retrievedList = data.get(i);
            for (Cell cell : row)    //iteration over cell using for each loop
            {
                switch (formulaEvaluator.evaluateInCell(cell).getCellType()) {
                    case Cell.CELL_TYPE_NUMERIC:   //field that represents numeric cell type
//getting the value of the cell as a number
//                        System.out.print(cell.getNumericCellValue() + "\t\t");
                        double celN = cell.getNumericCellValue();
                        retrievedList.add( celN + "");
                        break;
                    case Cell.CELL_TYPE_STRING:    //field that represents string cell type
//getting the value of the cell as a string
//                        System.out.print(cell.getStringCellValue() + "\t\t");
                        String celS = cell.getStringCellValue();
                        retrievedList.add(celS + "");
                        break;

                    default:
                        retrievedList.add("");
                }
            }

        }
        return data;
    }


        public String writeExcel(String name, int age, int rowNo, String fileLocation) throws Exception {

    Workbook wb = new HSSFWorkbook();
    Sheet sheet = wb.createSheet("Persons");
        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 4000);

    Row header = sheet.createRow(0);

    CellStyle headerStyle = wb.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
    short x =1;
        headerStyle.setFillPattern(x);
    //         FillPatternType x = FillPatternType.SOLID_FOREGROUND;
//        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
    Font font = sheet.getWorkbook().createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        headerStyle.setFont(font);

    Cell headerCell = header.createCell(0);
        headerCell.setCellValue("Name");
        headerCell.setCellStyle(headerStyle);

    headerCell = header.createCell(1);
        headerCell.setCellValue("Age");
        headerCell.setCellStyle(headerStyle);

    CellStyle style = wb.createCellStyle();
        style.setWrapText(true);

    Row row = sheet.createRow(rowNo);
    Cell cell = row.createCell(0);
        cell.setCellValue(name);
        cell.setCellStyle(style);

    cell = row.createCell(1);
        cell.setCellValue(age);
        cell.setCellStyle(style);


//    String fileLocation = path.substring(0, path.length() - 1) + "temp.xlsx";
//    String fileLocation = "/Users/omaribrahimelrifai/Documents/bpm_java/javaExcel/src/main/java/com/example/JavaBeans/temp.xlsx";


    FileOutputStream outputStream = new FileOutputStream(fileLocation);
        wb.write(outputStream);
        outputStream.close();
        return "Success";
    }


}


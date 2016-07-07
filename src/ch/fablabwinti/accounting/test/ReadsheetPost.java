package ch.fablabwinti.accounting.test;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

public class ReadsheetPost {

    static XSSFRow row;

    public static void main(String[] args) throws Exception {
        FileInputStream fis = new FileInputStream(new File("20141231_post_b.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet spreadsheet = workbook.getSheetAt(0);
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        CellValue cellValue;
        Iterator <Row> rowIterator = spreadsheet.iterator();
        while (rowIterator.hasNext()) {
            row = (XSSFRow) rowIterator.next();
            Iterator <Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                System.out.print(row.getRowNum() + "/" + cell.getColumnIndex() + " ");
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            System.out.print(cell.getDateCellValue() + " (DATE)\t\t ");
                        } else {
                            System.out.print(cell.getNumericCellValue() + " \t\t ");
                        }
                        break;
                    case Cell.CELL_TYPE_STRING:
                        System.out.print(cell.getStringCellValue() + " \t\t " );
                        break;
                    case Cell.CELL_TYPE_FORMULA:
                        cellValue = evaluator.evaluate(cell);
                        System.out.print(cell.getCellFormula() + " (FORMULA) => " + cellValue.getNumberValue() + "\t\t");
                        break;
                }
            }
            System.out.println();
        }
        fis.close();
    }
}

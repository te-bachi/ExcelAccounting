package ch.fablabwinti.accounting.main;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.util.Date;

/**
 *
 */
public class CellCreator {

    private Cell cell;

    public CellCreator(XSSFRow row, int columnIndex, CellStyle cellStyle) {
        cell = row.createCell(columnIndex);
        cell.setCellStyle(cellStyle);
    }

    public Cell createCell(String value) {
        cell.setCellValue(value);
        return cell;
    }

    public Cell createCell(int value) {
        cell.setCellValue(value);
        return cell;
    }

    public Cell createCell(double value) {
        cell.setCellValue(value);
        return cell;
    }

    public Cell createCell(Date value) {
        cell.setCellValue(value);
        return cell;
    }
}

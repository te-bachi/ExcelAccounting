package ch.fablabwinti.accounting.cell;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.math.BigDecimal;
import java.util.Date;

/**
 *
 */
public class CustomCell {
    protected XSSFRow row;
    protected Cell    cell;

    public CustomCell(XSSFRow row, int columnIndex) {
        this.row    = row;
        this.cell   = row.getCell(columnIndex);
    }

    protected String cellTypeToString(int cellType) {
        if (cellType == Cell.CELL_TYPE_NUMERIC) return "NUMERIC";
        if (cellType == Cell.CELL_TYPE_STRING)  return "STRING";
        if (cellType == Cell.CELL_TYPE_FORMULA) return "FORMULA";
        if (cellType == Cell.CELL_TYPE_BLANK)   return "BLANK";
        if (cellType == Cell.CELL_TYPE_BOOLEAN) return "BOOLEAN";
        if (cellType == Cell.CELL_TYPE_ERROR)   return "ERROR";
        return "UNKNOW";
    }

    public int getInt() {
        return Double.valueOf(cell.getNumericCellValue()).intValue();
    }

    public Date getDate() {
        return cell.getDateCellValue();
    }

    public String getString() {
        if (cell == null) {
            return "";
        }
        return cell.getStringCellValue();
    }

    public BigDecimal getBigDecimal() {
        return new BigDecimal(cell.getNumericCellValue());
    }
}

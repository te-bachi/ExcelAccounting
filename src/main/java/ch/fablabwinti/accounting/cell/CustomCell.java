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

    protected String cellTypeToString(CellType cellType) {
        if (cellType == CellType.NUMERIC) return "NUMERIC";
        if (cellType == CellType.STRING)  return "STRING";
        if (cellType == CellType.FORMULA) return "FORMULA";
        if (cellType == CellType.BLANK)   return "BLANK";
        if (cellType == CellType.BOOLEAN) return "BOOLEAN";
        if (cellType == CellType.ERROR)   return "ERROR";
        return "UNKNOW";
    }

    public int getInt() {
        return Double.valueOf(cell.getNumericCellValue()).intValue();
    }

    public String getDoubleString() {
        return String.valueOf(cell.getNumericCellValue());
    }
    public String getBigDecimalString() {
        return new BigDecimal(cell.getNumericCellValue()).setScale(2, BigDecimal.ROUND_HALF_EVEN).toString();
    }

    public Cell getCell() {
        return cell;
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
        return new BigDecimal(cell.getNumericCellValue()).setScale(2, BigDecimal.ROUND_HALF_EVEN);
    }
}

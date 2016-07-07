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

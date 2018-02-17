package ch.fablabwinti.accounting.cell;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;


/**
 *
 */
public class CustomStringCell extends CustomCell {
    public CustomStringCell(XSSFRow row, int columnIndex) throws CustomCellException {
        super(row, columnIndex);
        if (cell == null) {

        } else if (cell.getCellType() != Cell.CELL_TYPE_STRING && cell.getCellType() != Cell.CELL_TYPE_BLANK) {
            throw new CustomCellException(row.getRowNum() + "/" + cell.getColumnIndex() + " is not of type string (or blank) but " + cellTypeToString(cell.getCellType()) + "!");
        }
    }
}

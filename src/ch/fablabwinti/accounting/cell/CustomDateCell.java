package ch.fablabwinti.accounting.cell;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;

/**
 *
 */
public class CustomDateCell extends CustomCell {
    public CustomDateCell(XSSFRow row, int columnIndex) throws CustomCellException {
        super(row, columnIndex);
        if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK || cell.getCellType() != Cell.CELL_TYPE_NUMERIC || !DateUtil.isCellDateFormatted(cell)) {
            throw new CustomCellException(row.getRowNum() + "/" + cell.getColumnIndex() + " is not of type date but " + cellTypeToString(cell.getCellType()) + "!");
        }
    }
}

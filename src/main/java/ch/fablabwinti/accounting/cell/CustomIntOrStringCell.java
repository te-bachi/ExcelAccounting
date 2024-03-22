package ch.fablabwinti.accounting.cell;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class CustomIntOrStringCell extends CustomCell {
    public CustomIntOrStringCell(XSSFRow row, int columnIndex) throws CustomCellException {
        super(row, columnIndex);
        if (cell == null) {

        } else if (cell.getCellType() != CellType.NUMERIC && cell.getCellType() != CellType.STRING && cell.getCellType() != CellType.BLANK) {
            throw new CustomCellException(row.getRowNum() + "/" + cell.getColumnIndex() + " is not of type string (or blank) but " + cellTypeToString(cell.getCellType()) + "!");
        }
    }
}

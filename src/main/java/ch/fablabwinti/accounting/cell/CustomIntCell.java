package ch.fablabwinti.accounting.cell;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;

/**
 *
 */
public class CustomIntCell extends CustomCell {

    public CustomIntCell(XSSFRow row, int columnIndex) throws CustomCellException {
        this(row, columnIndex, null);
    }

    public CustomIntCell(XSSFRow row, int columnIndex, FormulaEvaluator evaluator) throws CustomCellException {
        super(row, columnIndex);
        if (cell == null) {
            throw new CustomCellException(row.getRowNum() + "/" + columnIndex + " is not defined!");
        } else if (cell.getCellType() == CellType.FORMULA && evaluator != null) {
            try {
                evaluator.evaluateFormulaCell(cell);
            } catch (RuntimeException e) {
                throw e;
            }
            //System.out.println(row.getRowNum() + "/" + cell.getColumnIndex() + ": formula value = " + cell.getNumericCellValue());
        } else if (cell.getCellType() == CellType.BLANK || cell.getCellType() != CellType.NUMERIC) {
            throw new CustomCellException(row.getRowNum() + "/" + cell.getColumnIndex() + " is not of type number but " + cellTypeToString(cell.getCellType()) + "!");
        }
    }
}

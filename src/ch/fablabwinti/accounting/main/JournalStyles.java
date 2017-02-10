package ch.fablabwinti.accounting.main;

import org.apache.poi.ss.usermodel.*;

/**
 *
 */
public class JournalStyles {
    public Font             normalFont;
    public Font             boldFont;

    public CreationHelper   createHelper;
    public CellStyle        normalStyle;
    public CellStyle        boldStyle;
    public CellStyle        dateStyle;
    public CellStyle        numberStyle;
    public CellStyle        boldNumberStyle;
    public CellStyle        totalStyle;

    public JournalStyles(Workbook workbook) {


        /* Fonts */
        normalFont          = workbook.createFont();
        normalFont.setFontHeightInPoints((short)11);
        normalFont.setFontName("Arial");

        boldFont            = workbook.createFont();
        boldFont.setFontHeightInPoints((short)11);
        boldFont.setFontName("Arial");
        boldFont.setBoldweight(Font.BOLDWEIGHT_BOLD);

        /* Normal Style */
        normalStyle         = workbook.createCellStyle();
        normalStyle.setFont(normalFont);

        /* Bold Style */
        boldStyle           = workbook.createCellStyle();
        boldStyle.setFont(boldFont);

        /* Total Style */
        totalStyle          = workbook.createCellStyle();
        totalStyle.setFont(boldFont);
        totalStyle.setBorderBottom(CellStyle.BORDER_THIN);
        totalStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        totalStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("#,##0.00"));

        /* Date Style */
        dateStyle           = workbook.createCellStyle();
        createHelper        = workbook.getCreationHelper();
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("DD.MM.YYYY"));

        /* Number Style */
        numberStyle         = workbook.createCellStyle();
        numberStyle.setFont(normalFont);
        numberStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("#,##0.00"));

        /* Bold Number Style */
        boldNumberStyle     = workbook.createCellStyle();
        boldNumberStyle.setFont(boldFont);
        boldNumberStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("#,##0.00"));
    }
}

package ch.fablabwinti.accounting.main;

import org.apache.poi.ss.usermodel.*;

/**
 *
 */
public class JournalStyles {
    public Font             normalFont;
    public Font             boldFont;
    public Font             smallFont;

    public CreationHelper   createHelper;
    public CellStyle        normalStyle;
    public CellStyle        boldStyle;
    public CellStyle        dateStyle;
    public CellStyle        numberStyle;
    public CellStyle        boldNumberStyle;
    public CellStyle        totalStyle;
    public CellStyle        accountTitleStyle;
    public CellStyle        accountHeaderStyle;
    public CellStyle        accountTotalStyle;

    public CellStyle        balanceStyle;

    public JournalStyles(Workbook workbook) {


        /* Fonts */
        normalFont          = workbook.createFont();
        normalFont.setFontHeightInPoints((short)11);
        normalFont.setFontName("Arial");

        boldFont            = workbook.createFont();
        boldFont.setFontHeightInPoints((short)11);
        boldFont.setFontName("Arial");
        boldFont.setBoldweight(Font.BOLDWEIGHT_BOLD);

        smallFont            = workbook.createFont();
        smallFont.setFontHeightInPoints((short)8);
        smallFont.setFontName("Arial");

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
        dateStyle.setFont(normalFont);
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("DD.MM.YYYY"));

        /* Number Style */
        numberStyle         = workbook.createCellStyle();
        numberStyle.setFont(normalFont);
        numberStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("#,##0.00"));

        /* Bold Number Style */
        boldNumberStyle     = workbook.createCellStyle();
        boldNumberStyle.setFont(boldFont);
        boldNumberStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("#,##0.00"));

        /* Account Title Style */
        accountTitleStyle   = workbook.createCellStyle();
        accountTitleStyle.setFont(boldFont);
        accountTitleStyle.setBorderBottom(CellStyle.BORDER_THIN);
        accountTitleStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());

        /* Account Header Style */
        accountHeaderStyle   = workbook.createCellStyle();
        accountHeaderStyle.setFont(smallFont);
        accountHeaderStyle.setBorderBottom(CellStyle.BORDER_THIN);
        accountHeaderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());

        /* Account Total Style */
        accountTotalStyle   = workbook.createCellStyle();
        accountTotalStyle.setFont(boldFont);
        accountTotalStyle.setBorderBottom(CellStyle.BORDER_DOUBLE);
        accountTotalStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        accountTotalStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("#,##0.00"));


        /* Account Total Style */
        balanceStyle        = workbook.createCellStyle();
        balanceStyle.setFont(boldFont);
        balanceStyle.setBorderTop(CellStyle.BORDER_THIN);
        balanceStyle.setBorderBottom(CellStyle.BORDER_DOUBLE);
        balanceStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        balanceStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        balanceStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("#,##0.00"));
    }
}

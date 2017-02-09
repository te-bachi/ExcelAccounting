package ch.fablabwinti.accounting.main;

import ch.fablabwinti.accounting.Account;
import ch.fablabwinti.accounting.AccountList;
import ch.fablabwinti.accounting.TitleAccount;
import ch.fablabwinti.accounting.Transaction;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 *
 */
public class JournalAccumulator {

    public static int ACCOUNT_NUM_COLUMNS                       = 2;

    public static int ACCOUNT_ASSET_NR                          = 1;
    public static int ACCOUNT_LIABILITY_NR                      = 2;
    public static int[] ACCOUNT_INCOME_NR                       = { 3, 7 };
    public static int[] ACCOUNT_EXPENSE_NR                      = { 4, 5, 6, 9 };

    public static int COLUMN_WIDTH_RATIO                        = 260;
                                                                /* Assets,  Liabilities */
    public static int[] OUTPUT_COLUMN_BALANCE                   = { 0,      5 }; /* offset */

    public static int OUTPUT_COLUMN_BALANCE_TITLE_NR            = 0;
    public static int OUTPUT_COLUMN_BALANCE_ACCOUNT_NR          = 1;
    public static int OUTPUT_COLUMN_BALANCE_ACCOUNT             = 2;
    public static int OUTPUT_COLUMN_BALANCE_AMOUNT              = 3;
    public static int OUTPUT_COLUMN_BALANCE_TOTAL               = 4;

    /* 7 px per point */
    public static double OUTPUT_COLUMN_BALANCE_TITLE_NR_WIDTH      = 4.7142;    /* 33 px */
    public static double OUTPUT_COLUMN_BALANCE_ACCOUNT_NR_WIDTH    = 9.2857;    /* 65 px */
    public static double OUTPUT_COLUMN_BALANCE_ACCOUNT_WIDTH       = 40.7142;   /* 285 px */
    public static double OUTPUT_COLUMN_BALANCE_AMOUNT_WIDTH        = 16.7142;   /* 16 px */
    public static double OUTPUT_COLUMN_BALANCE_TOTAL_WIDTH         = 16.7142;   /* 16 px */

    private AccountPlanExport   accountPlan;
    private JournalExport       journalExport;
    private File                outputFile;

    public JournalAccumulator(File accountPlanFile, File journalExportFile) throws Exception {
        String path = journalExportFile.getPath();

        accountPlan = new AccountPlanExport();
        accountPlan.parseInput(accountPlanFile, 6);

        journalExport = new JournalExport();
        journalExport.parseInput(journalExportFile, null, new AccountList(accountPlan.getAccountList()));

        for (int i = 0; i < 1024; i++) {
            outputFile = new File(path.substring(0, path.lastIndexOf('.')) + "_output_" + i + path.substring(path.lastIndexOf('.'), path.length()));
            if (!outputFile.exists()) {
                break;
            }
        }
    }

    public void collectChildren(List<Account> childrenList, Account account) {
        for (Account child : account.getChildList()) {
            childrenList.add(child);
            collectChildren(childrenList, child);
        }
    }

    public void collectChildren(List<Account> childrenList, List<Account> accountList) {
        for (Account account : accountList) {
            childrenList.add(account);
            collectChildren(childrenList, account);
        }
    }

    public void exportOutput() throws Exception {
        FileOutputStream    out;
        XSSFWorkbook        workbook;
        XSSFSheet           spreadsheetBilanz;
        XSSFSheet           spreadsheetErfolg;
        XSSFSheet           spreadsheetKonten;
        int                 columnOffset;
        int                 balanceMaxIdx;
        AccountList         rootList;
        AccountList         accountList;
        Account             rootAccount;
        List<Account>       incomeList;
        List<Account>       expenseList;
        List<Account>       childrenList;
        Transaction         transaction;

        rootList            = new AccountList(accountPlan.getRootList());
        accountList         = new AccountList(accountPlan.getAccountList());
        incomeList          = new ArrayList<>();
        expenseList         = new ArrayList<>();
        childrenList        = new ArrayList<>();

        workbook            = new XSSFWorkbook();

        /*** Bilanz **********************************************************/
        spreadsheetBilanz   = workbook.createSheet("Bilanz");

        for (columnOffset = 0; columnOffset < OUTPUT_COLUMN_BALANCE.length; columnOffset++) {
            setColumnWidthForBalanceProfitAndLoss(spreadsheetBilanz, columnOffset);
            /*
            spreadsheetBilanz.getColumnHelper().setColWidth(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_TITLE_NR,   OUTPUT_COLUMN_BALANCE_TITLE_NR_WIDTH);
            spreadsheetBilanz.getColumnHelper().setColWidth(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_ACCOUNT_NR, OUTPUT_COLUMN_BALANCE_ACCOUNT_NR_WIDTH);
            spreadsheetBilanz.getColumnHelper().setColWidth(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_ACCOUNT,    OUTPUT_COLUMN_BALANCE_ACCOUNT_WIDTH);
            spreadsheetBilanz.getColumnHelper().setColWidth(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_AMOUNT,     OUTPUT_COLUMN_BALANCE_AMOUNT_WIDTH);
            spreadsheetBilanz.getColumnHelper().setColWidth(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_TOTAL,      OUTPUT_COLUMN_BALANCE_TOTAL_WIDTH);

            spreadsheetBilanz.getColumnHelper().setCustomWidth(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_TITLE_NR,   true);
            spreadsheetBilanz.getColumnHelper().setCustomWidth(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_ACCOUNT_NR, true);
            spreadsheetBilanz.getColumnHelper().setCustomWidth(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_ACCOUNT,    true);
            spreadsheetBilanz.getColumnHelper().setCustomWidth(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_AMOUNT,     true);
            spreadsheetBilanz.getColumnHelper().setCustomWidth(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_TOTAL,      true);
            */
        }

        rootAccount     = null;
        balanceMaxIdx   = 0;

        for (columnOffset = 0; columnOffset < ACCOUNT_NUM_COLUMNS; columnOffset++) {
            if (columnOffset == 0) {
                rootAccount = rootList.remove(ACCOUNT_ASSET_NR);
            } else if (columnOffset == 1) {
                rootAccount = rootList.remove(ACCOUNT_LIABILITY_NR);
            }
            childrenList.clear();
            collectChildren(childrenList, rootAccount);

            writeCells(workbook, spreadsheetBilanz, childrenList, columnOffset, balanceMaxIdx);

            /* increase max. row number for balance if necessary */
            if (balanceMaxIdx < childrenList.size()) {
                balanceMaxIdx = childrenList.size();
            }
        }


        /*** Erfolgsrechnung *************************************************/
        spreadsheetErfolg           = workbook.createSheet("Erfolgsrechnung");

        for (columnOffset = 0; columnOffset < OUTPUT_COLUMN_BALANCE.length; columnOffset++) {
            setColumnWidthForBalanceProfitAndLoss(spreadsheetErfolg, columnOffset);
        }

        while (!rootList.isEmpty()) {
            rootAccount = rootList.removeNext();
            if (checkArray(ACCOUNT_INCOME_NR, rootAccount.getNumber())) {
                incomeList.add(rootAccount);
            } else if (checkArray(ACCOUNT_EXPENSE_NR, rootAccount.getNumber())) {
                expenseList.add(rootAccount);
            } else {
                System.out.println(rootAccount.getNumber() + " \"" + rootAccount.getName() + "\" is neither an income nor an expense account");
                continue;
            }
        }

        rootAccount     = null;
        balanceMaxIdx   = 0;

        for (columnOffset = 0; columnOffset < ACCOUNT_NUM_COLUMNS; columnOffset++) {
            if (columnOffset == 0) {
                childrenList.clear();
                collectChildren(childrenList, expenseList);
            } else if (columnOffset == 1) {
                childrenList.clear();
                collectChildren(childrenList, incomeList);
            }

            writeCells(workbook, spreadsheetErfolg, childrenList, columnOffset, balanceMaxIdx);

            /* increase max. row number for balance if necessary */
            if (balanceMaxIdx < childrenList.size()) {
                balanceMaxIdx = childrenList.size();
            }
        }

        /*** Konten **********************************************************/
        spreadsheetKonten   = workbook.createSheet("Konten");

        out = new FileOutputStream(outputFile);
        workbook.write(out);
        out.close();
    }

    private void writeCells(XSSFWorkbook workbook, XSSFSheet sheet, List<Account> childrenList, int columnOffset, int balanceMaxIdx) {
        Cell                cell;
        XSSFRow             row;
        DataFormat          format;
        CreationHelper      createHelper;
        CellStyle           normalStyle;
        CellStyle           boldStyle;
        CellStyle           dateStyle;
        CellStyle           numberStyle;
        CellStyle           boldNumberStyle;
        CellStyle           totalStyle;
        Font                normalFont;
        Font                boldFont;

        Account             account;

        double              total;
        boolean             writeTotal;
        int                 rowIdx;

        total           = 0.0;
        writeTotal      = false;

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

        rowIdx          = 0;
        for (rowIdx = 0; rowIdx < childrenList.size(); rowIdx++) {
            account = childrenList.get(rowIdx);
            if (account instanceof TitleAccount) {

                if (rowIdx >= balanceMaxIdx) {
                    row = sheet.createRow(rowIdx);
                } else {
                    row = sheet.getRow(rowIdx);
                }

                cell = row.createCell(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_TITLE_NR);
                cell.setCellValue(account.getNumber());
                cell.setCellStyle(boldStyle);

                cell = row.createCell(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_ACCOUNT_NR);
                cell.setCellValue(account.getName());
                cell.setCellStyle(boldStyle);

                total = account.getTotal().doubleValue();
            } else {

                if (rowIdx >= balanceMaxIdx) {
                    row = sheet.createRow(rowIdx);
                } else {
                    row = sheet.getRow(rowIdx);
                }

                cell = row.createCell(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_ACCOUNT_NR);
                cell.setCellValue(account.getNumber());
                cell.setCellStyle(normalStyle);

                cell = row.createCell(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_ACCOUNT);
                cell.setCellValue(account.getName());
                cell.setCellStyle(normalStyle);

                cell = row.createCell(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_AMOUNT);
                cell.setCellValue(account.getTotal().doubleValue());
                cell.setCellStyle(numberStyle);

                    /* peek next account and check if it's a TitleAccount... */
                if ((rowIdx + 1) < childrenList.size()) {
                    account = childrenList.get(rowIdx + 1);
                    if (account instanceof TitleAccount) {
                        writeTotal = true;
                    }

                    /*  or is it the last one in the list? */
                } else {
                    writeTotal = true;
                }

                if (writeTotal) {
                    writeTotal = false;
                    cell = row.createCell(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_TOTAL);
                    cell.setCellValue(total);
                    cell.setCellStyle(totalStyle);
                }
            }

        }
    }

    public static void setColumnWidthForBalanceProfitAndLoss(XSSFSheet sheet, int idx) {
        sheet.getColumnHelper().setColWidth(OUTPUT_COLUMN_BALANCE[idx] + OUTPUT_COLUMN_BALANCE_TITLE_NR,   OUTPUT_COLUMN_BALANCE_TITLE_NR_WIDTH);
        sheet.getColumnHelper().setColWidth(OUTPUT_COLUMN_BALANCE[idx] + OUTPUT_COLUMN_BALANCE_ACCOUNT_NR, OUTPUT_COLUMN_BALANCE_ACCOUNT_NR_WIDTH);
        sheet.getColumnHelper().setColWidth(OUTPUT_COLUMN_BALANCE[idx] + OUTPUT_COLUMN_BALANCE_ACCOUNT,    OUTPUT_COLUMN_BALANCE_ACCOUNT_WIDTH);
        sheet.getColumnHelper().setColWidth(OUTPUT_COLUMN_BALANCE[idx] + OUTPUT_COLUMN_BALANCE_AMOUNT,     OUTPUT_COLUMN_BALANCE_AMOUNT_WIDTH);
        sheet.getColumnHelper().setColWidth(OUTPUT_COLUMN_BALANCE[idx] + OUTPUT_COLUMN_BALANCE_TOTAL,      OUTPUT_COLUMN_BALANCE_TOTAL_WIDTH);

        sheet.getColumnHelper().setCustomWidth(OUTPUT_COLUMN_BALANCE[idx] + OUTPUT_COLUMN_BALANCE_TITLE_NR,   true);
        sheet.getColumnHelper().setCustomWidth(OUTPUT_COLUMN_BALANCE[idx] + OUTPUT_COLUMN_BALANCE_ACCOUNT_NR, true);
        sheet.getColumnHelper().setCustomWidth(OUTPUT_COLUMN_BALANCE[idx] + OUTPUT_COLUMN_BALANCE_ACCOUNT,    true);
        sheet.getColumnHelper().setCustomWidth(OUTPUT_COLUMN_BALANCE[idx] + OUTPUT_COLUMN_BALANCE_AMOUNT,     true);
        sheet.getColumnHelper().setCustomWidth(OUTPUT_COLUMN_BALANCE[idx] + OUTPUT_COLUMN_BALANCE_TOTAL,      true);
    }

    public static boolean checkArray(int[] array, int value) {
        for (int i : array) {
            if (i == value) {
                return true;
            }
        }
        return false;
    }

    public static void main(String[] args) throws Exception {
        JournalAccumulator journalAccumulator;

        if (args.length < 2) {
            System.out.println("<account plan> <journal>");
        }

        journalAccumulator = new JournalAccumulator(new File(args[0]), new File(args[1]));
        journalAccumulator.exportOutput();

        System.out.println("done");
    }
}

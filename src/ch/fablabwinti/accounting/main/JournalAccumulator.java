package ch.fablabwinti.accounting.main;

import ch.fablabwinti.accounting.Account;
import ch.fablabwinti.accounting.AccountList;
import ch.fablabwinti.accounting.TitleAccount;
import ch.fablabwinti.accounting.Transaction;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 *
 */
public class JournalAccumulator {

    public static int ACCOUNT_NUM_COLUMNS                       = 2;

    public static int ACCOUNT_ASSET_NR                          = 1;
    public static int ACCOUNT_LIABILITY_NR                      = 2;
    public static int[] ACCOUNT_INCOME_NR                       = { 3, 7 };
    public static int[] ACCOUNT_EXPENSE_NR                      = { 4, 5, 6, 8, 9 };

                                                                /* Assets,  Liabilities */
    public static int[] OUTPUT_COLUMN_BALANCE                   = { 0,      5 }; /* offset */

    public static int OUTPUT_COLUMN_BALANCE_TITLE_NR            = 0;
    public static int OUTPUT_COLUMN_BALANCE_ACCOUNT_NR          = 1;
    public static int OUTPUT_COLUMN_BALANCE_ACCOUNT             = 2;
    public static int OUTPUT_COLUMN_BALANCE_AMOUNT              = 3;
    public static int OUTPUT_COLUMN_BALANCE_TOTAL               = 4;

    public static int OUTPUT_COLUMN_TRANSACTION_NR              = 0;
    public static int OUTPUT_COLUMN_TRANSACTION_DATE            = 1;
    public static int OUTPUT_COLUMN_TRANSACTION_DEBIT_ACCOUNT   = 2;
    public static int OUTPUT_COLUMN_TRANSACTION_CREDIT_ACCOUNT  = 3;
    public static int OUTPUT_COLUMN_TRANSACTION_TEXT            = 4;
    public static int OUTPUT_COLUMN_TRANSACTION_DEBIT_AMOUNT    = 5;
    public static int OUTPUT_COLUMN_TRANSACTION_CREDIT_AMOUNT   = 6;
    public static int OUTPUT_COLUMN_TRANSACTION_TOTAL           = 7;

    /* 7 px per point */
    public static double OUTPUT_COLUMN_BALANCE_TITLE_NR_WIDTH      = 4.7142;    /* 33 px */
    public static double OUTPUT_COLUMN_BALANCE_ACCOUNT_NR_WIDTH    = 9.2857;    /* 65 px */
    public static double OUTPUT_COLUMN_BALANCE_ACCOUNT_WIDTH       = 40.7142;   /* 285 px */
    public static double OUTPUT_COLUMN_BALANCE_AMOUNT_WIDTH        = 16.7142;   /* 117 px */
    public static double OUTPUT_COLUMN_BALANCE_TOTAL_WIDTH         = 16.7142;   /* 117 px */

    public static double OUTPUT_COLUMN_TRANSACTION_NR_WIDTH             = 9.2857;    /* 65 px */
    public static double OUTPUT_COLUMN_TRANSACTION_DATE_WIDTH           = 11.4285;   /* 80 px */
    public static double OUTPUT_COLUMN_TRANSACTION_DEBIT_ACCOUNT_WIDTH  = 9.2857;    /* 65 px */
    public static double OUTPUT_COLUMN_TRANSACTION_CREDIT_ACCOUNT_WIDTH = 9.2857;    /* 65 px */
    public static double OUTPUT_COLUMN_TRANSACTION_TEXT_WIDTH           = 40.7142;   /* 285 px */
    public static double OUTPUT_COLUMN_TRANSACTION_DEBIT_AMOUNT_WIDTH   = 16.7142;   /* 117 px */
    public static double OUTPUT_COLUMN_TRANSACTION_CREDIT_AMOUNT_WIDTH  = 16.7142;   /* 117 px */
    public static double OUTPUT_COLUMN_TRANSACTION_TOTAL_WIDTH          = 16.7142;   /* 117 px */

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
        JournalStyles       styles;

        rootList            = new AccountList(accountPlan.getRootList());
        accountList         = new AccountList(accountPlan.getAccountList());
        incomeList          = new ArrayList<>();
        expenseList         = new ArrayList<>();
        childrenList        = new ArrayList<>();

        workbook            = new XSSFWorkbook();
        styles              = new JournalStyles(workbook);

        /*** Bilanz **********************************************************/
        spreadsheetBilanz   = workbook.createSheet("Bilanz");

        for (columnOffset = 0; columnOffset < OUTPUT_COLUMN_BALANCE.length; columnOffset++) {
            setColumnWidthForBalanceProfitAndLoss(spreadsheetBilanz, columnOffset);
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

            writeCellsForBalanceProfitAndLoss(workbook, spreadsheetBilanz, styles, childrenList, columnOffset, balanceMaxIdx);

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

            writeCellsForBalanceProfitAndLoss(workbook, spreadsheetErfolg, styles, childrenList, columnOffset, balanceMaxIdx);

            /* increase max. row number for balance if necessary */
            if (balanceMaxIdx < childrenList.size()) {
                balanceMaxIdx = childrenList.size();
            }
        }

        /*** Konten **********************************************************/
        spreadsheetKonten   = workbook.createSheet("Konten");
        setColumnWidthForAccountTransactions(spreadsheetKonten);
        writeCellsForAccountTransactions(workbook, spreadsheetKonten, styles, accountList);

        out = new FileOutputStream(outputFile);
        workbook.write(out);
        out.close();
    }

    private void writeCellsForBalanceProfitAndLoss(XSSFWorkbook workbook, XSSFSheet sheet, JournalStyles styles, List<Account> childrenList, int columnOffset, int balanceMaxIdx) {
        Cell                cell;
        XSSFRow             row;

        Account             account;

        double              total;
        boolean             writeTotal;
        int                 rowIdx;

        total           = 0.0;
        writeTotal      = false;

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
                cell.setCellStyle(styles.boldStyle);

                cell = row.createCell(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_ACCOUNT_NR);
                cell.setCellValue(account.getName());
                cell.setCellStyle(styles.boldStyle);

                total = account.getTotal().doubleValue();
            } else {

                if (rowIdx >= balanceMaxIdx) {
                    row = sheet.createRow(rowIdx);
                } else {
                    row = sheet.getRow(rowIdx);
                }

                cell = row.createCell(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_ACCOUNT_NR);
                cell.setCellValue(account.getNumber());
                cell.setCellStyle(styles.normalStyle);

                cell = row.createCell(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_ACCOUNT);
                cell.setCellValue(account.getName());
                cell.setCellStyle(styles.normalStyle);

                cell = row.createCell(OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_AMOUNT);
                cell.setCellValue(account.getTotal().doubleValue());
                cell.setCellStyle(styles.numberStyle);

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
                    cell.setCellStyle(styles.totalStyle);
                }
            }

        }
    }

    private void writeCellsForAccountTransactions(XSSFWorkbook workbook, XSSFSheet sheet, JournalStyles styles, AccountList accountList) {
        Cell                cell;
        XSSFRow             row;

        Account             account;
        List<Transaction>   transactionList;
        Transaction         transaction;

        double              total;
        int                 accountIdx;
        int                 transactionIdx;
        int                 rowIdx;

        rowIdx = 0;
        for (accountIdx = 0; accountIdx < accountList.size(); accountIdx++) {
            account         = accountList.get(accountIdx);
            if (!(account instanceof TitleAccount)) {
                transactionList = account.getTransactionList();

                /* Account Title */
                row = sheet.createRow(rowIdx++);

                /* Account Nr */
                createCell(row, OUTPUT_COLUMN_TRANSACTION_NR, account.getNumber(), styles.accountTitleStyle);

                /* Account Name */
                createCell(row, OUTPUT_COLUMN_TRANSACTION_DATE, account.getName(), styles.accountTitleStyle);

                /* Blank */
                createCell(row, OUTPUT_COLUMN_TRANSACTION_DEBIT_ACCOUNT,    "", styles.accountTitleStyle);
                createCell(row, OUTPUT_COLUMN_TRANSACTION_CREDIT_ACCOUNT,   "", styles.accountTitleStyle);
                createCell(row, OUTPUT_COLUMN_TRANSACTION_TEXT,             "", styles.accountTitleStyle);
                createCell(row, OUTPUT_COLUMN_TRANSACTION_DEBIT_AMOUNT,     "", styles.accountTitleStyle);
                createCell(row, OUTPUT_COLUMN_TRANSACTION_CREDIT_AMOUNT,    "", styles.accountTitleStyle);
                createCell(row, OUTPUT_COLUMN_TRANSACTION_TOTAL,            "", styles.accountTitleStyle);

                sheet.addMergedRegion(
                        new CellRangeAddress(
                                row.getRowNum(),
                                row.getRowNum(),
                                OUTPUT_COLUMN_TRANSACTION_DATE,
                                OUTPUT_COLUMN_TRANSACTION_TOTAL
                        )
                );

                /* Account Header */
                row = sheet.createRow(rowIdx++);

                createCell(row, OUTPUT_COLUMN_TRANSACTION_NR,               "Nr.",       styles.accountHeaderStyle);
                createCell(row, OUTPUT_COLUMN_TRANSACTION_DATE,             "Datum",     styles.accountHeaderStyle);
                createCell(row, OUTPUT_COLUMN_TRANSACTION_DEBIT_ACCOUNT,    "Soll Nr.",  styles.accountHeaderStyle);
                createCell(row, OUTPUT_COLUMN_TRANSACTION_CREDIT_ACCOUNT,   "Haben Nr.", styles.accountHeaderStyle);
                createCell(row, OUTPUT_COLUMN_TRANSACTION_TEXT,             "Text",      styles.accountHeaderStyle);
                createCell(row, OUTPUT_COLUMN_TRANSACTION_DEBIT_AMOUNT,     "Soll",      styles.accountHeaderStyle);
                createCell(row, OUTPUT_COLUMN_TRANSACTION_CREDIT_AMOUNT,    "Haben",     styles.accountHeaderStyle);
                createCell(row, OUTPUT_COLUMN_TRANSACTION_TOTAL,            "Total",     styles.accountHeaderStyle);

                total = 0.0;
                for (transactionIdx = 0; transactionIdx < transactionList.size(); transactionIdx++) {
                    transaction = transactionList.get(transactionIdx);
                    row = sheet.createRow(rowIdx++);

                    createCell(row, OUTPUT_COLUMN_TRANSACTION_NR,               transaction.getNr(),                    styles.normalStyle);
                    createCell(row, OUTPUT_COLUMN_TRANSACTION_DATE,             transaction.getDate(),                  styles.dateStyle);
                    createCell(row, OUTPUT_COLUMN_TRANSACTION_DEBIT_ACCOUNT,    transaction.getDebit().getNumber(),     styles.normalStyle);
                    createCell(row, OUTPUT_COLUMN_TRANSACTION_CREDIT_ACCOUNT,   transaction.getCredit().getNumber(),    styles.normalStyle);
                    createCell(row, OUTPUT_COLUMN_TRANSACTION_TEXT,             transaction.getText(),                  styles.normalStyle);

                    if (account.getNumber() == transaction.getDebit().getNumber()) {
                        createCell(row, OUTPUT_COLUMN_TRANSACTION_DEBIT_AMOUNT,  transaction.getAmount().doubleValue(), styles.numberStyle);
                        total += transaction.getAmount().doubleValue();
                    } else {
                        createCell(row, OUTPUT_COLUMN_TRANSACTION_CREDIT_AMOUNT, transaction.getAmount().doubleValue(), styles.numberStyle);
                        total -= transaction.getAmount().doubleValue();
                    }

                    cell = row.createCell(OUTPUT_COLUMN_TRANSACTION_TOTAL);
                    cell.setCellValue(total);
                    if ((transactionIdx + 1) < transactionList.size()) {
                        cell.setCellStyle(styles.numberStyle);
                    } else {
                        cell.setCellStyle(styles.accountTotalStyle);
                    }
                }
                rowIdx++;
                rowIdx++;
            }
        }
    }

    private Cell createCell(XSSFRow row, int columnIndex, String value, CellStyle cellStyle) {
        Cell cell;
        cell = row.createCell(columnIndex);
        cell.setCellValue(value);
        cell.setCellStyle(cellStyle);
        return cell;
    }

    private Cell createCell(XSSFRow row, int columnIndex, int value, CellStyle cellStyle) {
        Cell cell;
        cell = row.createCell(columnIndex);
        cell.setCellValue(value);
        cell.setCellStyle(cellStyle);
        return cell;
    }

    private Cell createCell(XSSFRow row, int columnIndex, double value, CellStyle cellStyle) {
        Cell cell;
        cell = row.createCell(columnIndex);
        cell.setCellValue(value);
        cell.setCellStyle(cellStyle);
        return cell;
    }


    private Cell createCell(XSSFRow row, int columnIndex, Date value, CellStyle cellStyle) {
        Cell cell;
        cell = row.createCell(columnIndex);
        cell.setCellValue(value);
        cell.setCellStyle(cellStyle);
        return cell;
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

    public static void setColumnWidthForAccountTransactions(XSSFSheet sheet) {
        sheet.getColumnHelper().setColWidth(OUTPUT_COLUMN_TRANSACTION_NR,               OUTPUT_COLUMN_TRANSACTION_NR_WIDTH);
        sheet.getColumnHelper().setColWidth(OUTPUT_COLUMN_TRANSACTION_DATE,             OUTPUT_COLUMN_TRANSACTION_DATE_WIDTH);
        sheet.getColumnHelper().setColWidth(OUTPUT_COLUMN_TRANSACTION_DEBIT_ACCOUNT,    OUTPUT_COLUMN_TRANSACTION_DEBIT_ACCOUNT_WIDTH);
        sheet.getColumnHelper().setColWidth(OUTPUT_COLUMN_TRANSACTION_CREDIT_ACCOUNT,   OUTPUT_COLUMN_TRANSACTION_CREDIT_ACCOUNT_WIDTH);
        sheet.getColumnHelper().setColWidth(OUTPUT_COLUMN_TRANSACTION_TEXT,             OUTPUT_COLUMN_TRANSACTION_TEXT_WIDTH);
        sheet.getColumnHelper().setColWidth(OUTPUT_COLUMN_TRANSACTION_DEBIT_AMOUNT,     OUTPUT_COLUMN_TRANSACTION_DEBIT_AMOUNT_WIDTH);
        sheet.getColumnHelper().setColWidth(OUTPUT_COLUMN_TRANSACTION_CREDIT_AMOUNT,    OUTPUT_COLUMN_TRANSACTION_CREDIT_AMOUNT_WIDTH);
        sheet.getColumnHelper().setColWidth(OUTPUT_COLUMN_TRANSACTION_TOTAL,            OUTPUT_COLUMN_TRANSACTION_TOTAL_WIDTH);

        sheet.getColumnHelper().setCustomWidth(OUTPUT_COLUMN_TRANSACTION_NR,            true);
        sheet.getColumnHelper().setCustomWidth(OUTPUT_COLUMN_TRANSACTION_DATE,          true);
        sheet.getColumnHelper().setCustomWidth(OUTPUT_COLUMN_TRANSACTION_DEBIT_ACCOUNT, true);
        sheet.getColumnHelper().setCustomWidth(OUTPUT_COLUMN_TRANSACTION_CREDIT_ACCOUNT,true);
        sheet.getColumnHelper().setCustomWidth(OUTPUT_COLUMN_TRANSACTION_TEXT,          true);
        sheet.getColumnHelper().setCustomWidth(OUTPUT_COLUMN_TRANSACTION_DEBIT_AMOUNT,  true);
        sheet.getColumnHelper().setCustomWidth(OUTPUT_COLUMN_TRANSACTION_CREDIT_AMOUNT, true);
        sheet.getColumnHelper().setCustomWidth(OUTPUT_COLUMN_TRANSACTION_TOTAL,         true);
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

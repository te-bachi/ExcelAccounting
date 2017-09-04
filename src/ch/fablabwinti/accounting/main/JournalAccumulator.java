package ch.fablabwinti.accounting.main;

import ch.fablabwinti.accounting.Account;
import ch.fablabwinti.accounting.AccountList;
import ch.fablabwinti.accounting.TitleAccount;
import ch.fablabwinti.accounting.Transaction;
import org.apache.poi.ss.util.CellRangeAddress;
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

    private static String OUTPUT_SHEET_BALANCE_STR                       = "Bilanz";
    private static String OUTPUT_SHEET_PROFIT_LOSS_STR                   = "Erfolgsrechnung";
    private static String OUTPUT_SHEET_ACCOUNTS                          = "Konten";

    private static int ACCOUNT_NUM_COLUMNS                               = 2;

    /* Account Numbers */
    private static int ACCOUNT_ASSET_NR                                  = 1;
    private static int ACCOUNT_LIABILITY_NR                              = 2;
    private static int[] ACCOUNT_EXPENSE_NR                              = { 4, 5, 6, 8, 9 };
    private static int[] ACCOUNT_INCOME_NR                               = { 3, 7 };
    private static int ACCOUNT_BALANCE_PROFIT_TAXABLE                    = 2978;
    private static int ACCOUNT_BALANCE_PROFIT_TAXLESS                    = 2979;
    private static int ACCOUNT_PROFIT_LOSS_PROFIT_TAXABLE                = 9200;
    private static int ACCOUNT_PROFIT_LOSS_PROFIT_TAXLESS                = 9201;
    private static int ACCOUNT_INITIAL                                   = 9999;

    /* */                                                                /* Assets,  Liabilities */
    private static int[] OUTPUT_COLUMN_BALANCE                           = { 0,      5 }; /* offset */

    private static int OUTPUT_COLUMN_BALANCE_TITLE_NR                    = 0;
    private static int OUTPUT_COLUMN_BALANCE_ACCOUNT_NR                  = 1;
    private static int OUTPUT_COLUMN_BALANCE_ACCOUNT                     = 2;
    private static int OUTPUT_COLUMN_BALANCE_AMOUNT                      = 3;
    private static int OUTPUT_COLUMN_BALANCE_TOTAL                       = 4;

    private static int OUTPUT_COLUMN_TRANSACTION_NR                      = 0;
    private static int OUTPUT_COLUMN_TRANSACTION_DATE                    = 1;
    private static int OUTPUT_COLUMN_TRANSACTION_DEBIT_ACCOUNT           = 2;
    private static int OUTPUT_COLUMN_TRANSACTION_CREDIT_ACCOUNT          = 3;
    private static int OUTPUT_COLUMN_TRANSACTION_TEXT                    = 4;
    private static int OUTPUT_COLUMN_TRANSACTION_DEBIT_AMOUNT            = 5;
    private static int OUTPUT_COLUMN_TRANSACTION_CREDIT_AMOUNT           = 6;
    private static int OUTPUT_COLUMN_TRANSACTION_TOTAL                   = 7;

    private static String OUTPUT_COLUMN_TRANSACTION_NR_STR               = "Nr.";
    private static String OUTPUT_COLUMN_TRANSACTION_DATE_STR             = "Datum";
    private static String OUTPUT_COLUMN_TRANSACTION_DEBIT_ACCOUNT_STR    = "Soll Nr.";
    private static String OUTPUT_COLUMN_TRANSACTION_CREDIT_ACCOUNT_STR   = "Haben Nr.";
    private static String OUTPUT_COLUMN_TRANSACTION_TEXT_STR             = "Text";
    private static String OUTPUT_COLUMN_TRANSACTION_DEBIT_AMOUNT_STR     = "Soll";
    private static String OUTPUT_COLUMN_TRANSACTION_CREDIT_AMOUNT_STR    = "Haben";
    private static String OUTPUT_COLUMN_TRANSACTION_TOTAL_STR            = "Total";

    /* 7 px per point */
    private static double OUTPUT_COLUMN_BALANCE_TITLE_NR_WIDTH           = 4.7142;    /* 33 px */
    private static double OUTPUT_COLUMN_BALANCE_ACCOUNT_NR_WIDTH         = 9.2857;    /* 65 px */
    private static double OUTPUT_COLUMN_BALANCE_ACCOUNT_WIDTH            = 40.7142;   /* 285 px */
    private static double OUTPUT_COLUMN_BALANCE_AMOUNT_WIDTH             = 16.7142;   /* 117 px */
    private static double OUTPUT_COLUMN_BALANCE_TOTAL_WIDTH              = 16.7142;   /* 117 px */

    private static double OUTPUT_COLUMN_TRANSACTION_NR_WIDTH             = 9.2857;    /* 65 px */
    private static double OUTPUT_COLUMN_TRANSACTION_DATE_WIDTH           = 11.4285;   /* 80 px */
    private static double OUTPUT_COLUMN_TRANSACTION_DEBIT_ACCOUNT_WIDTH  = 9.2857;    /* 65 px */
    private static double OUTPUT_COLUMN_TRANSACTION_CREDIT_ACCOUNT_WIDTH = 9.2857;    /* 65 px */
    private static double OUTPUT_COLUMN_TRANSACTION_TEXT_WIDTH           = 40.7142;   /* 285 px */
    private static double OUTPUT_COLUMN_TRANSACTION_DEBIT_AMOUNT_WIDTH   = 16.7142;   /* 117 px */
    private static double OUTPUT_COLUMN_TRANSACTION_CREDIT_AMOUNT_WIDTH  = 16.7142;   /* 117 px */
    private static double OUTPUT_COLUMN_TRANSACTION_TOTAL_WIDTH          = 16.7142;   /* 117 px */

    private AccountPlanExport   accountPlan;
    private JournalExport       journalExport;
    private File                outputFile;

    public JournalAccumulator(File accountPlanFile, File journalExportFile) throws Exception {
        String path = journalExportFile.getPath();

        /* Account plan */
        accountPlan = new AccountPlanExport();
        accountPlan.parseInput(accountPlanFile, 6);

        /* Journal */
        journalExport = new JournalExport();
        journalExport.parseInput(journalExportFile, null, new AccountList(accountPlan.getAccountList()));

        /* Create output file */
        for (int i = 0; i < 1024; i++) {
            outputFile = new File(path.substring(0, path.lastIndexOf('.')) + "_output_" + i + path.substring(path.lastIndexOf('.'), path.length()));
            if (!outputFile.exists()) {
                break;
            }
        }
    }

    private void collectChildren(List<Account> childrenList, Account account) {
        for (Account child : account.getChildList()) {
            childrenList.add(child);
            collectChildren(childrenList, child);
        }
    }

    private void collectChildren(List<Account> childrenList, List<Account> accountList) {
        for (Account account : accountList) {
            childrenList.add(account);
            collectChildren(childrenList, account);
        }
    }

    private void calculateAccounts() {

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
        Account             assetAccount;
        Account             liabilityAccount;
        List<Account>       incomeList;
        List<Account>       expenseList;
        List<Account>       childrenList;
        JournalStyles       styles;
        double              neg;

        assetAccount        = null;
        liabilityAccount    = null;

        /* All accounts in a tree structure */
        rootList            = new AccountList(accountPlan.getRootList());

        /* All accounts in a flat structure */
        accountList         = new AccountList(accountPlan.getAccountList());

        /* Empty lists */
        incomeList          = new ArrayList<>();
        expenseList         = new ArrayList<>();
        childrenList        = new ArrayList<>();

        workbook            = new XSSFWorkbook();
        styles              = new JournalStyles(workbook);


        /*** Create Sheets ***************************************************/
        spreadsheetBilanz   = workbook.createSheet(OUTPUT_SHEET_BALANCE_STR);
        spreadsheetErfolg   = workbook.createSheet(OUTPUT_SHEET_PROFIT_LOSS_STR);
        spreadsheetKonten   = workbook.createSheet(OUTPUT_SHEET_ACCOUNTS);

        /*** Erfolgsrechnung *************************************************/

        /* Set sheet width => splitted in expense and income section */
        for (columnOffset = 0; columnOffset < OUTPUT_COLUMN_BALANCE.length; columnOffset++) {
            setColumnWidthForBalanceProfitAndLoss(spreadsheetErfolg, columnOffset);
        }

        /* Split root list into expense- and income list */
        while (!rootList.isEmpty()) {
            /* Remove an account from the list */
            rootAccount = rootList.removeNext();

            /* check if is an expense account ... */
            if (checkArray(ACCOUNT_EXPENSE_NR, rootAccount.getNumber())) {
                /* add to expense list */
                expenseList.add(rootAccount);

            /* ... or an income account */
            } else if (checkArray(ACCOUNT_INCOME_NR, rootAccount.getNumber())) {
                /* add to income list */
                incomeList.add(rootAccount);

            /* warning if it's neither an expense nor an income account */
            } else {
                if (rootAccount.getNumber() == ACCOUNT_ASSET_NR) {
                    assetAccount = rootAccount;
                } else if (rootAccount.getNumber() == ACCOUNT_LIABILITY_NR) {
                    liabilityAccount = rootAccount;
                } else {
                    System.out.println(rootAccount.getNumber() + " \"" + rootAccount.getName() + "\" is neither an income nor an expense account");
                }
            }
        }

        balanceMaxIdx   = 0;
        neg             = 1.0;



        /* Es wird zuerst über die ExpenseList, danach über die IncomeList iteriert.
         * Dabei werden alle Kinder- und Kindes-Kinder-Konten in eine Liste getan: die ChildrenList
         * Diese ChildrenList ist gleich wie in der Ausgabedatei im Excel
         *
         * Collect all children from all expense or income accounts and write it to the sheet
         */
        for (columnOffset = 0; columnOffset < ACCOUNT_NUM_COLUMNS; columnOffset++) {

            /* expense list =
             *   4 material expenses
             *   5 staff expenses
             *   6 energy expenses
             **/
            if (columnOffset == 0) {
                childrenList.clear();
                collectChildren(childrenList, expenseList);
                neg = 1.0;

            /* income list =
             *   3 operating revenue
             *   7 membership fee
             **/
            } else if (columnOffset == 1) {
                childrenList.clear();
                collectChildren(childrenList, incomeList);
                neg = -1.0;
            }

            /* Write child-accounts to the sheet.
             * rowIdx >= balanceMaxIdx --> create new row, otherwise re-use old row */
            writeCellsForBalanceProfitAndLoss(spreadsheetErfolg, styles, childrenList, columnOffset, balanceMaxIdx, neg);

            /*
             * Create new row or re-use old row?
             * What row is the total row?
             * increase max. row number for profit and loss calculation if necessary.
             *
             *  Expense  | Income
             * __________|__________
             * A1        | L1
             * A2        |
             * __________|__________
             * Total     | Total
             * __________|__________
             **/
            if (balanceMaxIdx < childrenList.size()) {
                balanceMaxIdx = childrenList.size();
            }
        }

        /*** Bilanz **********************************************************/

        /* Set sheet width => splitted in asset and liability section */
        for (columnOffset = 0; columnOffset < OUTPUT_COLUMN_BALANCE.length; columnOffset++) {
            setColumnWidthForBalanceProfitAndLoss(spreadsheetBilanz, columnOffset);
        }

        rootAccount     = null;
        balanceMaxIdx   = 0;

        /* Remove the asset and the liability account from the root list,
         * collect all children and write it to the sheet */
        for (columnOffset = 0; columnOffset < ACCOUNT_NUM_COLUMNS; columnOffset++) {
            /* asset */
            if (columnOffset == 0) {
                rootAccount = assetAccount;
                neg         = 1.0;

            /* liability */
            } else if (columnOffset == 1) {
                rootAccount = liabilityAccount;
                neg         = -1.0;
            }

            /* Clear children list */
            childrenList.clear();

            /* collect all children from account */
            collectChildren(childrenList, rootAccount);

            /* Write child-accounts to the sheet
             * rowIdx >= balanceMaxIdx --> create new row, otherwise re-use old row */
            writeCellsForBalanceProfitAndLoss(spreadsheetBilanz, styles, childrenList, columnOffset, balanceMaxIdx, neg);

            /**
             * Create new row or re-use old row?
             * What row is the total row?
             * increase max. row number for balance if necessary.
             *
             *  Asset    | Liability
             * __________|__________
             * A1        | L1
             * A2        |
             * __________|__________
             * Total     | Total
             * __________|__________
             **/
            if (balanceMaxIdx < childrenList.size()) {
                balanceMaxIdx = childrenList.size();
            }
        }

        /*** Konten **********************************************************/
        setColumnWidthForAccountTransactions(spreadsheetKonten);
        writeCellsForAccountTransactions(spreadsheetKonten, styles, accountList);

        out = new FileOutputStream(outputFile);
        workbook.write(out);
        out.close();
    }

    /**
     * Balance, Profit and Loss
     * Write cells to sheet.
     *
     * @param sheet
     * @param styles
     * @param childrenList
     * @param columnOffset
     * @param balanceMaxIdx
     */
    private void writeCellsForBalanceProfitAndLoss(XSSFSheet sheet, JournalStyles styles, List<Account> childrenList, int columnOffset, int balanceMaxIdx, double neg) {
        XSSFRow             row;

        Account             account;

        double              total;
        boolean             writeTotal;
        int                 rowIdx;

        total           = 0.0;
        writeTotal      = false;

        /* Iterate over the whole children list (list of accounts) */
        for (rowIdx = 0; rowIdx < childrenList.size(); rowIdx++) {
            account = childrenList.get(rowIdx);

            /* bypass initial account */
            if (account.getNumber() == ACCOUNT_INITIAL) {
                continue;
            }

            /* create or re-use row */
            if (rowIdx >= balanceMaxIdx) {
                row = sheet.createRow(rowIdx);
            } else {
                row = sheet.getRow(rowIdx);
            }
            /* Check if the account is a TitleAccount ... */
            if (account instanceof TitleAccount) {

                /*** Write to sheet */
                /* Account Nr. */
                new CellCreator(row, OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_TITLE_NR,      styles.boldStyle).createCell(account.getNumber());
                /* Account Name */
                new CellCreator(row, OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_ACCOUNT_NR,    styles.boldStyle).createCell(account.getName());

                total = account.getTotal().doubleValue();

            /* ... or NOT a TitleAccount account */
            } else {

                /*** Write to sheet */
                /* Account Nr. */
                new CellCreator(row, OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_ACCOUNT_NR,    styles.normalStyle).createCell(account.getNumber());

                /* Account Name */
                new CellCreator(row, OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_ACCOUNT,       styles.normalStyle).createCell(account.getName());

                /* Account Total (Calculated in TODO)*/
                new CellCreator(row, OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_AMOUNT,        styles.numberStyle).createCell(neg * account.getTotal().doubleValue());

                /* Current account is NOT the last in the list ... */
                if ((rowIdx + 1) < childrenList.size()) {
                    /* peek next account and check if it's a TitleAccount... */
                    account = childrenList.get(rowIdx + 1);
                    if (account instanceof TitleAccount) {
                        writeTotal = true;
                    }

                /*  ... or is it the last one in the list? */
                } else {
                    writeTotal = true;
                }

                /* Write total to sheet */
                if (writeTotal) {
                    writeTotal = false;
                    new CellCreator(row, OUTPUT_COLUMN_BALANCE[columnOffset] + OUTPUT_COLUMN_BALANCE_TOTAL,     styles.totalStyle).createCell(neg * total);
                }
            }

        }
    }

    /**
     * Account Transaction
     * Write cells to sheet.
     *
     * @param sheet
     * @param styles
     * @param accountList
     */
    private void writeCellsForAccountTransactions(XSSFSheet sheet, JournalStyles styles, AccountList accountList) {
        XSSFRow             row;

        Account             account;
        List<Transaction>   transactionList;
        Transaction         transaction;

        double              total;
        int                 accountIdx;
        int                 transactionIdx;
        int                 rowIdx;

        rowIdx = 0;

        /* Iterate over the flat unsorted account list */
        for (accountIdx = 0; accountIdx < accountList.size(); accountIdx++) {
            account         = accountList.get(accountIdx);

            /* Only write accounts that are NOT TitleAccounts */
            if (!(account instanceof TitleAccount)) {

                /*** Write to sheet:
                 *
                 * 9999 Account Name
                 * ------------------------------------------------------------------------------------
                 * | nr | date | debit nr | credit nr | text | debit amount | credit amount | total |
                 **/

                /* Account Title */
                row = sheet.createRow(rowIdx++);

                /* Account Nr */
                new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_NR,              styles.accountTitleStyle).createCell(account.getNumber());

                /* Account Name */
                new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_DATE,            styles.accountTitleStyle).createCell(account.getName());

                /* Blank */
                new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_DEBIT_ACCOUNT,   styles.accountTitleStyle).createCell("");
                new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_CREDIT_ACCOUNT,  styles.accountTitleStyle).createCell("");
                new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_TEXT,            styles.accountTitleStyle).createCell("");
                new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_DEBIT_AMOUNT,    styles.accountTitleStyle).createCell("");
                new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_CREDIT_AMOUNT,   styles.accountTitleStyle).createCell("");
                new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_TOTAL,           styles.accountTitleStyle).createCell("");

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

                new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_NR,              styles.accountHeaderStyle).createCell(OUTPUT_COLUMN_TRANSACTION_NR_STR);
                new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_DATE,            styles.accountHeaderStyle).createCell(OUTPUT_COLUMN_TRANSACTION_DATE_STR);
                new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_DEBIT_ACCOUNT,   styles.accountHeaderStyle).createCell(OUTPUT_COLUMN_TRANSACTION_DEBIT_ACCOUNT_STR);
                new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_CREDIT_ACCOUNT,  styles.accountHeaderStyle).createCell(OUTPUT_COLUMN_TRANSACTION_CREDIT_ACCOUNT_STR);
                new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_TEXT,            styles.accountHeaderStyle).createCell(OUTPUT_COLUMN_TRANSACTION_TEXT_STR);
                new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_DEBIT_AMOUNT,    styles.accountHeaderStyle).createCell(OUTPUT_COLUMN_TRANSACTION_DEBIT_AMOUNT_STR);
                new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_CREDIT_AMOUNT,   styles.accountHeaderStyle).createCell(OUTPUT_COLUMN_TRANSACTION_CREDIT_AMOUNT_STR);
                new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_TOTAL,           styles.accountHeaderStyle).createCell(OUTPUT_COLUMN_TRANSACTION_TOTAL_STR);

                total = 0.0;

                /* Get transaction list of an account */
                transactionList = account.getTransactionList();

                /* Iterate over transaction list */
                for (transactionIdx = 0; transactionIdx < transactionList.size(); transactionIdx++) {
                    transaction = transactionList.get(transactionIdx);
                    row = sheet.createRow(rowIdx++);

                    /*** Write to sheet */
                    new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_NR,              styles.normalStyle).createCell(transaction.getNr());
                    new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_DATE,            styles.dateStyle).createCell(transaction.getDate());
                    new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_DEBIT_ACCOUNT,   styles.normalStyle).createCell(transaction.getDebit().getNumber());
                    new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_CREDIT_ACCOUNT,  styles.normalStyle).createCell(transaction.getCredit().getNumber());
                    new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_TEXT,            styles.normalStyle).createCell(transaction.getText());

                    /* This transaction is a debit ...*/
                    if (account.getNumber() == transaction.getDebit().getNumber()) {
                        new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_DEBIT_AMOUNT,    styles.numberStyle).createCell(transaction.getAmount().doubleValue());
                        total += transaction.getAmount().doubleValue();

                    /* ... or a credit from the account perspective */
                    } else {
                        new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_CREDIT_AMOUNT,   styles.numberStyle).createCell(transaction.getAmount().doubleValue());
                        total -= transaction.getAmount().doubleValue();
                    }

                    /* Current transaction is NOT the last in the list ... */
                    if ((transactionIdx + 1) < transactionList.size()) {
                        new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_TOTAL,            styles.numberStyle).createCell(total);

                    /*  ... or is it the last one in the list? */
                    } else {
                        new CellCreator(row, OUTPUT_COLUMN_TRANSACTION_TOTAL,            styles.accountTotalStyle).createCell(total);
                    }
                }
                rowIdx++;
                rowIdx++;
            }
        }
    }

    private static void setColumnWidthForBalanceProfitAndLoss(XSSFSheet sheet, int idx) {
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

    private static void setColumnWidthForAccountTransactions(XSSFSheet sheet) {
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

    private static boolean checkArray(int[] array, int value) {
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

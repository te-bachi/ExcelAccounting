package ch.fablabwinti.accounting.main;

import ch.fablabwinti.accounting.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 *
 */
public class AccountPlanExport {

    public static int COLUMN_WIDTH_RATIO                = 260;

    public static int INPUT_COLUMN_ACCOUNT_TYPE         = 1;
    public static int INPUT_COLUMN_ACCOUNT_NR           = 2;
    public static int INPUT_COLUMN_ACCOUNT_TEXT         = 3;

    public static int[] OUTPUT_COLUMN_ACCOUNT_TEXT      = {  0, 3 };
    public static int[] OUTPUT_COLUMN_ACCOUNT_NR        = {  1, 2 };
    public static int[] OUTPUT_COLUMN_ACCOUNT_TYPE      = {  2, 0 };
    public static int[] OUTPUT_COLUMN_ACCOUNT_TITLE     = { -1, 1 };

    public static int OUTPUT_COLUMN_ACCOUNT_TEXT_WIDTH  = 50;
    public static int OUTPUT_COLUMN_ACCOUNT_NR_WIDTH    = 9;
    public static int OUTPUT_COLUMN_ACCOUNT_TYPE_WIDTH  = 20;

    public static String ACCOUNT_TYPE_ASSET             = "Aktivkonto";
    public static String ACCOUNT_TYPE_LIABILITY         = "Passivkonto";
    public static String ACCOUNT_TYPE_INCOME            = "Ertragskonto";
    public static String ACCOUNT_TYPE_EXPENSE           = "Aufwandskonto";

    private List<Account>   rootList;
    private List<Account>   accountList;

    public AccountPlanExport() {
        /* Root list with only root TitleAccounts
         * Every account could have children. */
        rootList    = new ArrayList<>();

        /* A list of all accounts, not sorted. */
        accountList = new ArrayList<>();
    }

    public void parseInput(File inputFile, int level) throws Exception {
        FileInputStream     in;
        XSSFWorkbook        workbook;
        XSSFSheet           spreadsheet;
        FormulaEvaluator    evaluator;
        Iterator<Row>       rowIterator;
        XSSFRow             row;
        Iterator<Cell>      cellIterator;
        Cell                cell;
        CellValue           cellValue;
        String              accountText;
        int                 accountNumber;
        Account             account;
        Account             parent;
        int                 depth;
        int                 parentDepth;
        int                 depthDiff;

        parent      = null;

        in          = new FileInputStream(inputFile);
        workbook    = new XSSFWorkbook(in);
        spreadsheet = workbook.getSheetAt(0);
        evaluator   = workbook.getCreationHelper().createFormulaEvaluator();
        rowIterator = spreadsheet.iterator();
        while (rowIterator.hasNext()) {
            row  = (XSSFRow) rowIterator.next();

            /* First column: Account Text */
            cell = row.getCell(INPUT_COLUMN_ACCOUNT_TEXT);

            /* Only go further if the cell is NOT blank => otherwise skip row */
            if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) {
                if (cell.getCellType() != Cell.CELL_TYPE_STRING) {
                    System.out.println(row.getRowNum() + "/" + cell.getColumnIndex() + " is not of type string but " + cell.getCellType() + "!");
                    continue;
                }
                accountText = cell.getStringCellValue();

                /* Second column: Account Nr */
                cell = row.getCell(INPUT_COLUMN_ACCOUNT_NR);
                if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK || cell.getCellType() != Cell.CELL_TYPE_NUMERIC) {
                    System.out.println(row.getRowNum() + "/" + cell.getColumnIndex() + " is not of type number but " + cell.getCellType() + "!");
                    continue;
                }
                accountNumber = Double.valueOf(cell.getNumericCellValue()).intValue();

                /* Third column: Account Type (String of empty [TitleAccount]) */
                cell = row.getCell(INPUT_COLUMN_ACCOUNT_TYPE);

                /* Check if Account Type is empty/blank => It's a TitleAccount and there are children */
                if ((cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK)) {
                    /**
                     * Check if the current accountNumber is in range, ex. 34 < 99
                     * Math.pow(10, 2) - 1 = 100 - 1 = 99
                     * 'level' is a static parameter of this function
                     * The user can decide, how many TitleAccounts he want can
                     * can reduce sub-levels.
                     **/
                    if (accountNumber < (new Double(Math.pow(10, level)).intValue() - 1)) {
                        try {
                            /* The same but using the length of the accountNumber, not the accountNumber itself ?!*/
                            depth = Integer.valueOf(accountNumber).toString().length();

                            /* Is there a parent? */
                            if (parent != null) {
                                /* Get depth of parent */
                                parentDepth = Integer.valueOf(parent.getNumber()).toString().length();

                                /**
                                 * The actual depth is lower than the parent depth.
                                 * ex.
                                 * parent:    1234  AccountText1 => parentDepth = 4
                                 * current:   32    AccountText1 => depth       = 2
                                 * depthDiff: 4 - 2 + 1 = 3
                                 *
                                 * We have to go up 3 levels to get a depth-1 TitleAccount
                                 **/
                                if (depth <= parentDepth) {
                                    depthDiff = parentDepth - depth + 1;
                                    do {
                                        if (parent != null) {
                                            parent = parent.getParent();
                                        }
                                        depthDiff--;
                                    } while (depthDiff > 0);
                                }
                                /* If there is a parent, use parent to create a new TitleAccount... */
                                if (parent != null) {
                                    /* Replace parent with a new TitleAccount.
                                     * Internally the new TitleAccount is added to the parent */
                                    parent = new TitleAccount(parent, accountNumber, accountText);
                                    accountList.add(parent);

                                /* ... otherwise use a root TitleAccount (no parent) */
                                } else {
                                    parent = new TitleAccount(accountNumber, accountText);
                                    rootList.add(parent);
                                    accountList.add(parent);
                                }
                            /* ... otherwise use a root TitleAccount (no parent) */
                            } else {
                                parent = new TitleAccount(accountNumber, accountText);
                                rootList.add(parent);
                                accountList.add(parent);
                            }
                        } catch (NumberFormatException e) {
                            /* bad account number format: skip journal entry */
                        }
                    }

                /* cell is NOT null and NOT blank => if it's NOT string, skip row */
                } else if (cell.getCellType() != Cell.CELL_TYPE_STRING) {
                    System.out.println(row.getRowNum() + "/" + cell.getColumnIndex() + " is not of type string but " + cell.getCellType() + "!");
                    continue;
                /* cell is of type string */
                } else {
                    if (cell.getStringCellValue().equals(ACCOUNT_TYPE_ASSET)) {
                        account = new AssetAccount(parent, accountNumber, accountText);
                    } else if (cell.getStringCellValue().equals(ACCOUNT_TYPE_LIABILITY)) {
                        account = new LiabilityAccount(parent, accountNumber, accountText);
                    } else if (cell.getStringCellValue().equals(ACCOUNT_TYPE_INCOME)) {
                        account = new IncomeAccount(parent, accountNumber, accountText);
                    } else if (cell.getStringCellValue().equals(ACCOUNT_TYPE_EXPENSE)) {
                        account = new ExpenseAccount(parent, accountNumber, accountText);
                    } else {
                        System.out.println(row.getRowNum() + "/" + cell.getColumnIndex() + " is not a defined type: \"" + cell.getStringCellValue() + "\"!");
                        continue;
                    }
                    accountList.add(account);
                }
            }
        }
        in.close();
    }

    public void exportOutput(File outputFile) throws Exception {
        FileOutputStream    out;
        XSSFWorkbook        workbook;
        XSSFSheet           spreadsheet[];
        XSSFRow             row;
        Cell                cell;
        int                 i;
        int                 k;
        int                 m;
        Account             account;

        workbook        = new XSSFWorkbook();
        spreadsheet     = new XSSFSheet[2];
        spreadsheet[0]  = workbook.createSheet("journal");
        spreadsheet[1]  = workbook.createSheet("buha");

        /* Iterate over the two newly created sheets */
        for (i = 0; i < spreadsheet.length; i++) {
            spreadsheet[i].setColumnWidth(OUTPUT_COLUMN_ACCOUNT_TEXT[i], COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_ACCOUNT_TEXT_WIDTH);
            spreadsheet[i].setColumnWidth(OUTPUT_COLUMN_ACCOUNT_NR[i], COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_ACCOUNT_NR_WIDTH);
            spreadsheet[i].setColumnWidth(OUTPUT_COLUMN_ACCOUNT_TYPE[i], COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_ACCOUNT_TYPE_WIDTH);

            /* Iterate over the whole account list (unsorted) */
            for (k = 0, m = 0; k < accountList.size(); k++) {
                account = accountList.get(k);

                if (account instanceof TitleAccount) {
                    if (i > 0) {
                        row = spreadsheet[i].createRow(m++);

                        cell = row.createCell(OUTPUT_COLUMN_ACCOUNT_TITLE[i]);
                        cell.setCellValue(account.getNumber() + " " + account.getName());
                    }
                } else {
                    row = spreadsheet[i].createRow(m++);

                    cell = row.createCell(OUTPUT_COLUMN_ACCOUNT_TEXT[i]);
                    cell.setCellValue(account.getName());

                    cell = row.createCell(OUTPUT_COLUMN_ACCOUNT_NR[i]);
                    cell.setCellValue(Integer.valueOf(account.getNumber()));

                    cell = row.createCell(OUTPUT_COLUMN_ACCOUNT_TYPE[i]);
                    if (account instanceof AssetAccount) {
                        cell.setCellValue(ACCOUNT_TYPE_ASSET);
                    } else if (account instanceof LiabilityAccount) {
                        cell.setCellValue(ACCOUNT_TYPE_LIABILITY);
                    } else if (account instanceof IncomeAccount) {
                        cell.setCellValue(ACCOUNT_TYPE_INCOME);
                    } else if (account instanceof ExpenseAccount) {
                        cell.setCellValue(ACCOUNT_TYPE_EXPENSE);
                    } else {
                        System.out.print(row.getRowNum() + "/" + cell.getColumnIndex() + " object is a strange thing!");
                        return;
                    }
                }
            }
        }

        out = new FileOutputStream(outputFile);
        workbook.write(out);
        out.close();
    }

    public List<Account> getRootList() {
        return rootList;
    }

    public List<Account> getAccountList() {
        return accountList;
    }

    public Account findAccount(int number) {
        for (Account account : accountList) {
            if (account.getNumber() == number) {
                return account;
            }
        }
        return null;
    }

    public static void main(String[] args) throws Exception {
        File                inputFile;
        File                outputFile;
        int                 level;
        AccountPlanExport   main;

        /* 1 = input  [0]
           2 = level  [1]
           3 = output [2]
         */
        if (args.length < 1 || args.length > 3) {
            System.out.println("<input> [<level> [<output>]]");
            return;
        }
        inputFile = new File(args[0]);

        if (args.length > 1) {
            level = Integer.valueOf(args[1]).intValue();
        } else {
            level = 2;
        }

        if (args.length < 3) {
            outputFile = new File(args[0].substring(0, args[0].lastIndexOf('.'))  + "_output" + args[0].substring(args[0].lastIndexOf('.'), args[0].length()));
        } else {
            outputFile = new File(args[2]);
        }

        if (!inputFile.exists()) {
            System.out.println("Input file \"" + inputFile.getAbsolutePath() + "\" doesn't exist!");
            return;
        }

        if (outputFile.exists()) {
            System.out.println("Output file \"" + outputFile.getAbsolutePath() + "\" exist! Abort");
            return;
        }

        main = new AccountPlanExport();

        main.parseInput(inputFile, level);
        if ((main.accountList.size() > 0)) {
            main.exportOutput(outputFile);
            System.out.println("Done!");
        } else {
            System.out.println("No parsed data found!");
        }

    }
}

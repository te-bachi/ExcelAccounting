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
    public static int INPUT_COLUMN_ACCOUNT_SUBNR        = 3;
    public static int INPUT_COLUMN_ACCOUNT_TEXT         = 4;
    public static int INPUT_COLUMN_ACCOUNT_SUBTEXT      = 5;

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

    public static int ACCOUNT_NR_DEPTH[]                = {1, 2, 4, 6, 0};

    private List<Account>   rootList;
    private List<Account>   accountList;

    /* not used */
    enum Type {
        COMPACT,
        FULL
    };

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
        String              accountNumberStr;
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

            /* Account Text */
            cell = row.getCell(INPUT_COLUMN_ACCOUNT_TEXT);

            /* Only go further if the cell is NOT blank => otherwise skip row */
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                if (cell.getCellType() != CellType.STRING) {
                    System.out.println(row.getRowNum() + "/" + cell.getColumnIndex() + " is not of type string but " + cell.getCellType() + "!");
                    continue;
                }
                accountText = cell.getStringCellValue();

                /* Account Sub-Text */
                cell = row.getCell(INPUT_COLUMN_ACCOUNT_SUBTEXT);
                if (cell != null && cell.getCellType() != CellType.BLANK) {
                    if (cell.getCellType() != CellType.STRING) {
                        System.out.println(row.getRowNum() + "/" + cell.getColumnIndex() + " is not of type string but " + cell.getCellType() + "!");
                        continue;
                    }
                    accountText += " " + cell.getStringCellValue();
                }

                /* Account Nr */
                cell = row.getCell(INPUT_COLUMN_ACCOUNT_NR);
                if (cell == null || cell.getCellType() == CellType.BLANK || cell.getCellType() != CellType.STRING) {
                    System.out.println(row.getRowNum() + "/" + cell.getColumnIndex() + " is not of type string but " + cell.getCellType() + "!");
                    continue;
                }
                accountNumberStr = cell.getStringCellValue();

                /* Account Sub-Nr */
                cell = row.getCell(INPUT_COLUMN_ACCOUNT_SUBNR);
                if (cell != null && cell.getCellType() != CellType.BLANK) {
                    if (cell.getCellType() != CellType.STRING) {
                        System.out.println(row.getRowNum() + "/" + cell.getColumnIndex() + " is not of type string but " + cell.getCellType() + "!");
                        continue;
                    }
                    accountNumberStr += cell.getStringCellValue();
                }

                accountNumber = Double.valueOf(accountNumberStr).intValue();

                /* Third column: Account Type (String of empty [TitleAccount]) */
                cell = row.getCell(INPUT_COLUMN_ACCOUNT_TYPE);

                try {
                    /**
                     * Check if the current accountNumber is in range.
                     * 'level' is a static parameter of this function.
                     * The user can decide, how many Accounts he want
                     * to reduce sub-levels.
                     */
                    depth = Integer.valueOf(accountNumber).toString().length();

                    if (depth <= level) {


                        if (parent != null) {
                            parentDepth = Integer.valueOf(parent.getNumber()).toString().length();

                            if (depth <= parentDepth) {
                                /* find account index */
                                int idx;
                                for (idx = 0; idx < ACCOUNT_NR_DEPTH.length; idx++) {
                                    if (ACCOUNT_NR_DEPTH[idx] == depth) {
                                        break;
                                    }
                                }
                                if (idx == ACCOUNT_NR_DEPTH.length) {
                                    System.out.println("name=\"" + accountText + "\": can't find new parent!");
                                }

                                int idxParent;
                                for (idxParent = 0; idxParent < ACCOUNT_NR_DEPTH.length; idxParent++) {
                                    if (ACCOUNT_NR_DEPTH[idxParent] == parentDepth) {
                                        break;
                                    }
                                }
                                if (idxParent == ACCOUNT_NR_DEPTH.length) {
                                    System.out.println("parent=\"" + parent.getName() + "\": can't find new parent!");
                                }

                                int idxDiff = idxParent - idx + 1;
                                do {
                                    if (parent != null) {
                                        parent = parent.getParent();
                                    }
                                    idxDiff--;
                                } while (idxDiff > 0);

                            }
                        }

                        /* Check if Account Type is empty/blank => It's a TitleAccount and there are children */
                        /*                                                ============                        */
                        if ((cell == null || cell.getCellType() == CellType.BLANK)) {

                            /* Is there a parent? */
                            if (parent != null) {

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

                        /* cell is NOT null and NOT blank => if it's NOT string, skip row */
                        } else if (cell.getCellType() != CellType.STRING) {
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
                } catch (NumberFormatException e) {
                    /* bad account number format: skip journal entry */
                    System.out.println("Account plan " + row.getRowNum() + "/" + cell.getColumnIndex() + ": bad account number format");
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

    public static void main(String[] args) throws Exception {
        File                inputFile;
        File                outputFile;
        Type                type;           /* not used */
        int                 level;
        AccountPlanExport   main;

        /* 1 = "compact" | "full"   [0]
           2 = input                [1]
           3 = level                [2]
           4 = output               [3]
         */
        String              argType;
        String              argInput;
        String              argLevel;
        String              argOutput;
        int                 argTypeIdx      = 0;
        int                 argInputIdx     = 1;
        int                 argLevelIdx     = 2;
        int                 argOutputIdx    = 3;

        if (args.length < 2 || args.length > 4) {
            System.out.println("(\"compact\" | \"full\") <input> [<level> [<output>]]");
            return;
        }

        /* argType */
        argType = args[argTypeIdx];
        if (argType.equals("compact")) {
            type = Type.COMPACT;
        } else if (argType.equals("full")) {
            type = Type.FULL;
        }

        /* argInput */
        argInput = args[argInputIdx];
        inputFile = new File(argInput);

        if (args.length > argLevelIdx) {
            argLevel = args[argLevelIdx];
            level  = Integer.valueOf(argLevel).intValue();
        } else {
            level = 2;
        }

        if (args.length < argOutputIdx) {

            outputFile = new File(argInput.substring(0, argInput.lastIndexOf('.'))  + "_output" + argInput.substring(argInput.lastIndexOf('.'), argInput.length()));
        } else {
            argOutput = args[argOutputIdx];
            outputFile = new File(argOutput);
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

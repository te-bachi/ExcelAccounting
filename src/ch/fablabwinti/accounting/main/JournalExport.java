package ch.fablabwinti.accounting.main;

import ch.fablabwinti.accounting.*;
import ch.fablabwinti.accounting.cell.CustomCell;
import ch.fablabwinti.accounting.cell.CustomCellException;
import ch.fablabwinti.accounting.cell.CustomIntCell;
import ch.fablabwinti.accounting.cell.CustomStringCell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.*;

/**
 *
 */
public class JournalExport {

    public static int COLUMN_WIDTH_RATIO                    = 260;

    /* [0]=Post [1]=Kasse [2]=SumUp [3]=Restkonto [4]=Guthaben */
                                                            /* [0] [1] [2] [3] [4] */
    public static int[] INPUT_COLUMN_JOURNAL_NR             = {  0,  0,  0,  0,  0};
    public static int[] INPUT_COLUMN_JOURNAL_DATE           = {  1,  3,  1,  1,  3};
    public static int[] INPUT_COLUMN_JOURNAL_DEBIT_NR       = {  5,  5,  5,  5,  4};
    public static int[] INPUT_COLUMN_JOURNAL_CREDIT_NR      = {  6,  6,  6,  6,  5};
    public static int[] INPUT_COLUMN_JOURNAL_VALUE          = { 17, 24, 12, 13, 23};
    public static int[] INPUT_COLUMN_JOURNAL_TEXT           = {  9, 11,  9,  9, 10};

    public static int OUTPUT_COLUMN_JOURNAL_NR              = 0;
    public static int OUTPUT_COLUMN_JOURNAL_DATE            = 1;
    public static int OUTPUT_COLUMN_JOURNAL_DEBIT_NR        = 2;
    public static int OUTPUT_COLUMN_JOURNAL_CREDIT_NR       = 3;
    public static int OUTPUT_COLUMN_JOURNAL_VALUE           = 5;
    public static int OUTPUT_COLUMN_JOURNAL_TEXT            = 6;

    public static int OUTPUT_COLUMN_JOURNAL_NR_WIDTH        = 6;
    public static int OUTPUT_COLUMN_JOURNAL_DATE_WIDTH      = 14;
    public static int OUTPUT_COLUMN_JOURNAL_DEBIT_NR_WIDTH  = 6;
    public static int OUTPUT_COLUMN_JOURNAL_CREDIT_NR_WIDTH = 6;
    public static int OUTPUT_COLUMN_JOURNAL_VALUE_WIDTH     = 10;
    public static int OUTPUT_COLUMN_JOURNAL_TEXT_WIDTH      = 50;

    private List<Transaction> transactionList;

    public JournalExport() {
        transactionList = new ArrayList<>();
    }

    /**
     *
     * @param inputFile
     * @param filter
     * @param accountList flatten account list from "Kontenplan"
     * @throws Exception
     */
    public void parseInput(File inputFile, JournalFilter filter, AccountList accountList) throws Exception {
        FileInputStream     in;
        XSSFWorkbook        workbook;
        XSSFSheet           spreadsheet;
        FormulaEvaluator    evaluator;
        Iterator<Row>       rowIterator;
        XSSFRow             row;
        Iterator<Cell>      cellIterator;
        CustomCell          cell;
        CellValue           cellValue;
        int                 i;
        int                 k;
        Transaction         transaction;
        int                 journalNr;
        Date                journalDate;
        int                 journalDebitNr;
        int                 journalCreditNr;
        BigDecimal          journalValue;
        String              journalText;
        Account             debit;
        Account             credit;

        in          = new FileInputStream(inputFile);
        workbook    = new XSSFWorkbook(in);
        evaluator   = workbook.getCreationHelper().createFormulaEvaluator();
        for (i = 0; i < 5; i++) {
            spreadsheet = workbook.getSheetAt(i);
            System.out.println("Parsing sheet " + i);
            rowIterator = spreadsheet.iterator();
            while (rowIterator.hasNext()) {
                row = (XSSFRow) rowIterator.next();
                if (row.getRowNum() > 0) {
                    try {
                        /* try to fetch cell as INTEGER or die with exception */
                        cell = new CustomIntCell(row, INPUT_COLUMN_JOURNAL_NR[i]);
                        journalNr = cell.getInt();

                        /* try to fetch cell as INTEGER or die with exception */
                        cell = new CustomIntCell(row, INPUT_COLUMN_JOURNAL_DATE[i]);
                        journalDate = cell.getDate();

                        /* try to fetch cell as INTEGER or die with exception */
                        cell = new CustomIntCell(row, INPUT_COLUMN_JOURNAL_DEBIT_NR[i], evaluator);
                        journalDebitNr = cell.getInt();

                        /* try to fetch cell as INTEGER or die with exception */
                        cell = new CustomIntCell(row, INPUT_COLUMN_JOURNAL_CREDIT_NR[i], evaluator);
                        journalCreditNr = cell.getInt();

                        /* try to fetch cell as INTEGER or die with exception */
                        cell = new CustomIntCell(row, INPUT_COLUMN_JOURNAL_VALUE[i], evaluator);
                        journalValue = cell.getBigDecimal();

                        /* try to fetch cell as STRING or die with exception */
                        cell = new CustomStringCell(row, INPUT_COLUMN_JOURNAL_TEXT[i]);
                        journalText = cell.getString();

                        /* if there is no filter (normal) or a filter is attached and returns true */
                        if (filter == null || filter.filter(journalNr, journalDate, journalDebitNr, journalCreditNr, journalValue, journalText)) {
                            /* flatten account list from "Kontenplan", if available */
                            if (accountList != null) {
                                /* debit  = Account object -> find with account number from journal
                                 * credit = Account object -> find with account number from journal */
                                debit  = accountList.find(journalDebitNr);
                                credit = accountList.find(journalCreditNr);
                                if (debit == null || credit == null) {
                                    if (debit == null) {
                                        System.out.println("Can't find account nr " + journalDebitNr + " in account list");
                                    }
                                    if (credit == null) {
                                        System.out.println("Can't find account nr " + journalCreditNr + " in account list");
                                    }
                                    continue;
                                }
                                /* with two account objects */
                                transaction = new Transaction(journalNr, journalDate, debit, credit, journalValue, journalText);

                                /* === book transaction (add/remove money from an account) === */
                                debit.addTransaction(transaction);
                                credit.addTransaction(transaction);
                            } else {
                                /* only with account numbers */
                                transaction = new Transaction(journalNr, journalDate, journalDebitNr, journalCreditNr, journalValue, journalText);
                            }
                            transactionList.add(transaction);
                        }
                    } catch (CustomCellException e) {
                        System.out.println(e.getMessage());
                    } catch (IllegalStateException e) {
                        System.out.println("row " + row.getRowNum() + " has illegal cells");
                        throw e;
                    }
                }
            }
        }
        in.close();
    }

    public void exportOutput(File outputFile) throws Exception {
        FileOutputStream    out;
        XSSFWorkbook        workbook;
        XSSFSheet           spreadsheet;
        XSSFRow             row;
        Cell                cell;
        CellStyle           dateStyle;
        CreationHelper      createHelper;
        int                 i;
        int                 k;
        Transaction         transaction;

        workbook        = new XSSFWorkbook();
        spreadsheet     = workbook.createSheet("journal");
        dateStyle       = workbook.createCellStyle();
        createHelper    = workbook.getCreationHelper();
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("DD.MM.YYYY"));

        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_NR,        COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_NR_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_DATE,      COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_DATE_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_DEBIT_NR,  COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_DEBIT_NR_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_CREDIT_NR, COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_CREDIT_NR_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_VALUE,     COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_VALUE_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_TEXT,      COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_TEXT_WIDTH);

        for (k = 0; k < transactionList.size(); k++) {
            transaction = transactionList.get(k);
            row = spreadsheet.createRow(k);

            cell = row.createCell(OUTPUT_COLUMN_JOURNAL_NR);
            cell.setCellValue(transaction.getNr());

            cell = row.createCell(OUTPUT_COLUMN_JOURNAL_DATE);
            cell.setCellValue(transaction.getDate());
            cell.setCellStyle(dateStyle);

            cell = row.createCell(OUTPUT_COLUMN_JOURNAL_DEBIT_NR);
            cell.setCellValue(transaction.getDebit().getNumber());

            cell = row.createCell(OUTPUT_COLUMN_JOURNAL_CREDIT_NR);
            cell.setCellValue(transaction.getCredit().getNumber());

            cell = row.createCell(OUTPUT_COLUMN_JOURNAL_VALUE);
            cell.setCellValue(transaction.getAmount().doubleValue());

            cell = row.createCell(OUTPUT_COLUMN_JOURNAL_TEXT);
            cell.setCellValue(transaction.getText());
        }
        out = new FileOutputStream(outputFile);
        workbook.write(out);
        out.close();
    }

    public static void main(String[] args) throws Exception {
        File                inputFile;
        File                outputFile;
        JournalExport       main;
        JournalFilter       filter;

        if (args.length < 1 || args.length > 2) {
            System.out.println("<input> [<account filter>]");
            return;
        }
        inputFile   = new File(args[0]);
        outputFile  = new File(args[0].substring(0, args[0].lastIndexOf('.'))  + "_output" + args[0].substring(args[0].lastIndexOf('.'), args[0].length()));
        filter      = null;

        if (args.length == 2) {
            try {
                filter = new AccountJournalFilter(Integer.valueOf(args[1]).intValue());
            } catch (NumberFormatException e) {
                System.out.println(e.getMessage());
                return;
            }
        }

        if (!inputFile.exists()) {
            System.out.println("Input file \"" + inputFile.getAbsolutePath() + "\" doesn't exist!");
            return;
        }

        if (outputFile.exists()) {
            System.out.println("Output file \"" + outputFile.getAbsolutePath() + "\" exist! Abort");
            return;
        }

        main = new JournalExport();

        main.parseInput(inputFile, filter, null);
        if ((main.transactionList.size() > 0)) {
            Collections.sort(main.transactionList, new TransactionComparator());
            main.exportOutput(outputFile);
            System.out.println("Done!");
        } else {
            System.out.println("No parsed data found!");
        }

    }
}

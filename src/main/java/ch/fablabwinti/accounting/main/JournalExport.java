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

    private static int COLUMN_WIDTH_RATIO                    = 260;

    /* [0]=Post [1]=Kasse [2]=SumUp [3]=Restkonto [4]=Guthaben */
                                                            /* [0] [1] [2] [3] [4] */
    private static int[] INPUT_COLUMN_JOURNAL_NR                = {  0,  0,  0,  0,  0};
    private static int[] INPUT_COLUMN_JOURNAL_DATE              = {  1,  3,  1,  1,  3};
    private static int[] INPUT_COLUMN_JOURNAL_DEBIT_NR          = {  5,  5,  5,  5,  4};
    private static int[] INPUT_COLUMN_JOURNAL_CREDIT_NR         = {  6,  6,  6,  6,  5};
    private static int[] INPUT_COLUMN_JOURNAL_VALUE             = { 17, 24, 12, 13, 23};
    private static int[] INPUT_COLUMN_JOURNAL_TEXT              = {  9, 11,  9,  9, 10};
    private static int[] INPUT_COLUMN_JOURNAL_LASTNAME          = { 11, 13, 10, 11, 12};
    private static int[] INPUT_COLUMN_JOURNAL_FIRSTNAME         = { 12, 12, 11, 12, 11};

    private static int OUTPUT_COLUMN_JOURNAL_NR                 = 0;
    private static int OUTPUT_COLUMN_JOURNAL_DATE               = 1;
    private static int OUTPUT_COLUMN_JOURNAL_DEBIT_NR           = 2;
    private static int OUTPUT_COLUMN_JOURNAL_CREDIT_NR          = 3;
    private static int OUTPUT_COLUMN_JOURNAL_VALUE              = 4;
    private static int OUTPUT_COLUMN_JOURNAL_TEXT               = 5;
    private static int OUTPUT_COLUMN_JOURNAL_LASTNAME           = 6;
    private static int OUTPUT_COLUMN_JOURNAL_FIRSTNAME          = 7;

    private static int OUTPUT_COLUMN_JOURNAL_NR_WIDTH           = 6;
    private static int OUTPUT_COLUMN_JOURNAL_DATE_WIDTH         = 14;
    private static int OUTPUT_COLUMN_JOURNAL_DEBIT_NR_WIDTH     = 8;
    private static int OUTPUT_COLUMN_JOURNAL_CREDIT_NR_WIDTH    = 8;
    private static int OUTPUT_COLUMN_JOURNAL_VALUE_WIDTH        = 10;
    private static int OUTPUT_COLUMN_JOURNAL_TEXT_WIDTH         = 50;
    private static int OUTPUT_COLUMN_JOURNAL_LASTNAME_WIDTH     = 15;
    private static int OUTPUT_COLUMN_JOURNAL_FIRSTNAME_WIDTH    = 15;

    private static String OUTPUT_COLUMN_JOURNAL_NR_STR          = "Nr.";
    private static String OUTPUT_COLUMN_JOURNAL_DATE_STR        = "Datum";
    private static String OUTPUT_COLUMN_JOURNAL_DEBIT_NR_STR    = "Soll Nr.";
    private static String OUTPUT_COLUMN_JOURNAL_CREDIT_NR_STR   = "Haben Nr.";
    private static String OUTPUT_COLUMN_JOURNAL_VALUE_STR       = "Betrag";
    private static String OUTPUT_COLUMN_JOURNAL_TEXT_STR        = "Text";
    private static String OUTPUT_COLUMN_JOURNAL_LASTNAME_STR    = "Nachname";
    private static String OUTPUT_COLUMN_JOURNAL_FIRSTNAME_STR   = "Vorname";

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
        String              journalLastname;
        String              journalFirstname;
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

                        /* try to fetch cell as STRING or die with exception */
                        cell = new CustomStringCell(row, INPUT_COLUMN_JOURNAL_LASTNAME[i]);
                        journalLastname = cell.getString();

                        /* try to fetch cell as STRING or die with exception */
                        cell = new CustomStringCell(row, INPUT_COLUMN_JOURNAL_FIRSTNAME[i]);
                        journalFirstname = cell.getString();

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

                                /* debit == credit */
                                if (debit.getNumber() == credit.getNumber()) {
                                    System.out.println(row.getRowNum() + ": debit and credit are the same, skipping");
                                    continue;
                                }

                                /* with two account objects */
                                transaction = new Transaction(journalNr, journalDate, debit, credit, journalValue, journalText, journalLastname, journalFirstname);

                                /* === book transaction (add/remove money from an account) === */
                                debit.addTransaction(transaction);
                                credit.addTransaction(transaction);
                            } else {
                                /* only with account numbers */
                                transaction = new Transaction(journalNr, journalDate, journalDebitNr, journalCreditNr, journalValue, journalText, journalLastname, journalFirstname);
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
        JournalStyles       styles;
        Transaction         transaction;

        workbook        = new XSSFWorkbook();
        styles          = new JournalStyles(workbook);
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
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_LASTNAME,  COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_LASTNAME_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_FIRSTNAME, COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_FIRSTNAME_WIDTH);

        /* header */
        row = spreadsheet.createRow(0);
        new CellCreator(row, OUTPUT_COLUMN_JOURNAL_NR,              styles.accountHeaderStyle).createCell(OUTPUT_COLUMN_JOURNAL_NR_STR);
        new CellCreator(row, OUTPUT_COLUMN_JOURNAL_DATE,            styles.accountHeaderStyle).createCell(OUTPUT_COLUMN_JOURNAL_DATE_STR);
        new CellCreator(row, OUTPUT_COLUMN_JOURNAL_DEBIT_NR,        styles.accountHeaderStyle).createCell(OUTPUT_COLUMN_JOURNAL_DEBIT_NR_STR);
        new CellCreator(row, OUTPUT_COLUMN_JOURNAL_CREDIT_NR,       styles.accountHeaderStyle).createCell(OUTPUT_COLUMN_JOURNAL_CREDIT_NR_STR);
        new CellCreator(row, OUTPUT_COLUMN_JOURNAL_VALUE,           styles.accountHeaderStyle).createCell(OUTPUT_COLUMN_JOURNAL_VALUE_STR);
        new CellCreator(row, OUTPUT_COLUMN_JOURNAL_TEXT,            styles.accountHeaderStyle).createCell(OUTPUT_COLUMN_JOURNAL_TEXT_STR);
        new CellCreator(row, OUTPUT_COLUMN_JOURNAL_LASTNAME,        styles.accountHeaderStyle).createCell(OUTPUT_COLUMN_JOURNAL_LASTNAME_STR);
        new CellCreator(row, OUTPUT_COLUMN_JOURNAL_FIRSTNAME,       styles.accountHeaderStyle).createCell(OUTPUT_COLUMN_JOURNAL_FIRSTNAME_STR);

        for (k = 0; k < transactionList.size(); k++) {
            transaction = transactionList.get(k);
            row = spreadsheet.createRow(k + 1); // + 1 for header

            new CellCreator(row, OUTPUT_COLUMN_JOURNAL_NR,              styles.normalStyle).createCell(transaction.getNr());
            new CellCreator(row, OUTPUT_COLUMN_JOURNAL_DATE,            styles.dateStyle  ).createCell(transaction.getDate());
            new CellCreator(row, OUTPUT_COLUMN_JOURNAL_DEBIT_NR,        styles.normalStyle).createCell(transaction.getDebit().getNumber());
            new CellCreator(row, OUTPUT_COLUMN_JOURNAL_CREDIT_NR,       styles.normalStyle).createCell(transaction.getCredit().getNumber());
            new CellCreator(row, OUTPUT_COLUMN_JOURNAL_VALUE,           styles.numberStyle).createCell(transaction.getAmount().doubleValue());
            new CellCreator(row, OUTPUT_COLUMN_JOURNAL_TEXT,            styles.normalStyle).createCell(transaction.getText());
            new CellCreator(row, OUTPUT_COLUMN_JOURNAL_LASTNAME,        styles.normalStyle).createCell(transaction.getLastname());
            new CellCreator(row, OUTPUT_COLUMN_JOURNAL_FIRSTNAME,       styles.normalStyle).createCell(transaction.getFirstname());

            /*
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
            */
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

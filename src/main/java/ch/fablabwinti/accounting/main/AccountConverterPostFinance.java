package ch.fablabwinti.accounting.main;

import ch.fablabwinti.accounting.*;
import ch.fablabwinti.accounting.cell.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

public class AccountConverterPostFinance {

    private static int COLUMN_WIDTH_RATIO               = 260;

    private static int OUTPUT_COLUMN_DATE               = 2;
    private static int OUTPUT_COLUMN_TEXT               = 3;
    private static int OUTPUT_COLUMN_DEBIT              = 4;
    private static int OUTPUT_COLUMN_CREDIT             = 5;
    private static int OUTPUT_COLUMN_VALUTA             = 6;
    private static int OUTPUT_COLUMN_TOTAL              = 7;
    private static int OUTPUT_COLUMN_DESC               = 8;
    private static int OUTPUT_COLUMN_LASTNAME           = 9;
    private static int OUTPUT_COLUMN_FIRSTNAME          = 10;

    private static int OUTPUT_COLUMN_DATE_WIDTH         = 14;
    private static int OUTPUT_COLUMN_TEXT_WIDTH         = 45;
    private static int OUTPUT_COLUMN_DEBIT_WIDTH        = 10;
    private static int OUTPUT_COLUMN_CREDIT_WIDTH       = 10;
    private static int OUTPUT_COLUMN_VALUTA_WIDTH       = 14;
    private static int OUTPUT_COLUMN_TOTAL_WIDTH        = 10;
    private static int OUTPUT_COLUMN_DESC_WIDTH         = 38;
    private static int OUTPUT_COLUMN_LASTNAME_WIDTH     = 24;
    private static int OUTPUT_COLUMN_FIRSTNAME_WIDTH    = 24;

    private SimpleDateFormat                    dateFormat;
    private ArrayList<PostFinanceTransaction>   transactions;

    private class PostFinanceTransaction {
        public  int         startIndex;
        public  int         stopIndex;
        public  Date        date;
        public  BigDecimal  amount;
        public  boolean     debit;
        public  boolean     credit;
        public  BigDecimal  total;
        public  int         textCount;
        public  String      text[];

        public PostFinanceTransaction() {
            text = new String[20];
        }

        public String toString() {
            return dateFormat.format(date);
        }

        public int indexDiff() {
            return stopIndex - startIndex - 1;
        }
    }

    public AccountConverterPostFinance() {
        transactions = new ArrayList<>();
        dateFormat = new SimpleDateFormat("dd.MM.yyyy");
    }

    private void parseAmount(XSSFSheet spreadsheet, int startRowIndex, int i, int startColumnIndex, PostFinanceTransaction transaction) throws CustomCellException {
        CustomCell              customCell;
        String                  text;
        /* amount positive? */
        customCell = new CustomStringCell(spreadsheet.getRow(startRowIndex + i), startColumnIndex);
        text = customCell.getString();
        if (text != "" && text.contains("+")) {
            text = text.substring(0, text.indexOf("+") + 1);

            /* amount negative */
        } else {
            customCell = new CustomStringCell(spreadsheet.getRow(startRowIndex + i), startColumnIndex + 1);
            text = customCell.getString();
            if (text != "" && text.contains("-")) {
                text = text.substring(0, text.indexOf("-") + 1);
            }
        }

        if (text != "") {

            transaction.debit = text.endsWith("+");
            transaction.credit = text.endsWith("-");

            text = text.replaceAll("[+-]+", "");
            text = text.replaceAll("\'", "");

            try {
                transaction.amount = new BigDecimal(text);
            } catch (NumberFormatException e) {
                System.out.println("Amount can't be parsed on cell " + (customCell.getCell().getRowIndex() + 1) + "/" + (customCell.getCell().getColumnIndex() + 1));
            }
        } else {
            System.out.println("Amount in cell can't be parsed " + (customCell.getCell().getRowIndex() + 1) + "/" + (customCell.getCell().getColumnIndex() + 1));
        }
    }


    public void parseInput(File inputFile) throws Exception {
        FileInputStream         in;
        XSSFWorkbook            workbook;
        XSSFSheet               spreadsheet;
        FormulaEvaluator        evaluator;
        Iterator<Row>           rowIterator;
        XSSFRow                 row;
        Cell                    cell;
        CustomCell              customCell;
        PostFinanceTransaction  transaction;
        int                     startIndex;
        int                     i;
        String                  text;
        SimpleDateFormat        formatter;

        in          = new FileInputStream(inputFile);
        workbook    = new XSSFWorkbook(in);
        spreadsheet = workbook.getSheetAt(0);
        evaluator   = workbook.getCreationHelper().createFormulaEvaluator();
        rowIterator = spreadsheet.iterator();
        startIndex  = 0;
        formatter   = new SimpleDateFormat("dd.MM.yyyy");

        /* Prefetch next "Details" */
        while (rowIterator.hasNext()) {

            row  = (XSSFRow) rowIterator.next();
            /* Account Text */
            cell = row.getCell(0);

            /* Debug */
            /*
            if (startIndex >= 24 && startIndex <= 26) {
                System.out.print("debug ("+ (cell.getRowIndex() + 1) + "/" + (cell.getColumnIndex() + 1) + "): ");
                switch (cell.getCellType()) {
                    case STRING:
                        System.out.println(cell.getStringCellValue());
                        break;

                    case NUMERIC:
                        System.out.println(cell.getNumericCellValue());
                        break;

                    default:
                        System.out.println("<unknow>");
                        break;
                }

                System.out.print("");
            }
            */

            /* Only go further if the cell is "Details" */
            if (cell != null && cell.getCellType() == CellType.STRING && cell.getStringCellValue().equalsIgnoreCase("Details")) {
                transaction             = new PostFinanceTransaction();
                transaction.startIndex  = startIndex;
                transaction.stopIndex   = cell.getRowIndex();

                /* Date */
                try {
                    try {
                        customCell = new CustomStringCell(spreadsheet.getRow(startIndex), 0);
                        transaction.date = formatter.parse(customCell.getString());
                    } catch (ParseException e3) {
                        //
                    }

                    /* Text */
                    customCell              = new CustomStringCell(spreadsheet.getRow(startIndex), 1);
                    transaction.text[0]     = customCell.getString();
                } catch (CustomCellException e2) {
                    try {
                        customCell = new CustomStringCell(spreadsheet.getRow(startIndex), 1);
                        transaction.date = formatter.parse(customCell.getString());
                    } catch (ParseException e3) {
                        //
                    }

                    /* Text */
                    customCell              = new CustomStringCell(spreadsheet.getRow(startIndex), 2);
                    transaction.text[0]     = customCell.getString();
                }

                /* multi-liner */
                if (transaction.indexDiff() > 0) {
                    for (i = 0; i < transaction.indexDiff(); i++) {
                        try {
                            customCell = new CustomStringCell(spreadsheet.getRow(startIndex + 1 + i), 0);
                            text = customCell.getString();
                            /* Last row */
                            if (i == (transaction.indexDiff() - 1)) {

                                /* end text */
                                transaction.textCount = transaction.indexDiff() + 1;
                                transaction.text[i + 1] = text;

                                parseAmount(spreadsheet, startIndex + 1, i, 1, transaction);

                            /* NOT last row */
                            } else {
                                transaction.text[i + 1] = text;
                            }

                        } catch (CustomCellException e) {
                            customCell = new CustomIntCell(spreadsheet.getRow(startIndex + 1 + i), 0);
                            transaction.text[i + 1] = customCell.getBigDecimalString();
                        }
                    }
                /* one-liner */
                } else {
                    parseAmount(spreadsheet, startIndex, 0, 2, transaction);
                    transaction.textCount = 1;
                }

                transactions.add(transaction);

                startIndex = cell.getRowIndex() + 1;
            }

            /*
            cell = row.getCell(1);

            if (cell.getCellType() != CellType.STRING) {
                System.out.println(row.getRowNum() + "/" + cell.getColumnIndex() + " is not of type string but " + cell.getCellType() + "!");
                continue;
            }

            if (cell.getCellType() == CellType.NUMERIC) {
                accountNumberStr = Integer.toString((int) cell.getNumericCellValue());
            }
            */
        }
        in.close();
    }

    public void exportOutput(File outputFile) throws Exception {
        FileOutputStream        out;
        XSSFWorkbook            workbook;
        XSSFSheet               spreadsheet;
        XSSFRow                 row;
        Cell                    cell;
        JournalStyles           styles;
        int                     i;
        int                     k;
        int                     m;
        PostFinanceTransaction  transaction;

        workbook        = new XSSFWorkbook();
        styles          = new JournalStyles(workbook);
        spreadsheet     = workbook.createSheet("output");

        spreadsheet.setColumnWidth(OUTPUT_COLUMN_DATE,      COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_DATE_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_TEXT,      COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_TEXT_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_DEBIT,     COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_DEBIT_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_CREDIT,    COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_CREDIT_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_VALUTA,    COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_VALUTA_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_TOTAL,     COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_TOTAL_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_DESC,      COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_DESC_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_LASTNAME,  COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_LASTNAME_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_FIRSTNAME, COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_FIRSTNAME_WIDTH);


        for (i = 0, k = 0; i < transactions.size(); i++) {
            transaction = transactions.get(i);

            row = spreadsheet.createRow(k + 1); // + 1 for header

            new CellCreator(row, OUTPUT_COLUMN_DATE, styles.dateStyle)  .createCell(transaction.date);
            if (transaction.debit) {
                new CellCreator(row, OUTPUT_COLUMN_DEBIT, styles.numberStyle).createCell(transaction.amount.doubleValue());
            }
            if (transaction.credit) {
                new CellCreator(row, OUTPUT_COLUMN_CREDIT, styles.numberStyle).createCell(transaction.amount.doubleValue());
            }
            //new CellCreator(row, OUTPUT_COLUMN_VALUTA,      styles.dateStyle)  .createCell(transaction.???);
            //new CellCreator(row, OUTPUT_COLUMN_TOTAL,       styles.numberStyle).createCell(transaction.???);
            //new CellCreator(row, OUTPUT_COLUMN_DESC,        styles.normalStyle).createCell(transaction.???);
            //new CellCreator(row, OUTPUT_COLUMN_LASTNAME,    styles.normalStyle).createCell(transaction.???);
            //new CellCreator(row, OUTPUT_COLUMN_FIRSTNAME,   styles.normalStyle).createCell(transaction.???);

            for (m = 0; m < transaction.textCount; m++) {
                new CellCreator(row, OUTPUT_COLUMN_TEXT, styles.normalStyle).createCell(transaction.text[m]);
                if (m != (transaction.textCount - 1)) {
                    k++;
                    row = spreadsheet.createRow(k + 1); // + 1 for header
                }
            }
            k++;
        }

        out = new FileOutputStream(outputFile);
        workbook.write(out);
        out.close();
    }

    public static void main(String[] args) throws Exception {
        AccountConverterPostFinance main;
        String                      input;
        File                        inputFile;
        File                        outputFile;

        if (args.length < 1) {
            System.out.println("<input>");
            return;
        }

        /* argInput */
        input = args[0];
        inputFile = new File(input);

        outputFile = new File(input.substring(0, input.lastIndexOf('.'))  + "_output" + input.substring(input.lastIndexOf('.'), input.length()));


        if (!inputFile.exists()) {
            System.out.println("Input file \"" + inputFile.getAbsolutePath() + "\" doesn't exist! Abort");
            return;
        }

        if (outputFile.exists()) {
            System.out.println("Output file \"" + outputFile.getAbsolutePath() + "\" exist! Abort");
            return;
        }

        main = new AccountConverterPostFinance();

        main.parseInput(inputFile);
        if ((main.transactions.size() > 0)) {
            main.exportOutput(outputFile);
            System.out.println("Done!");
        } else {
            System.out.println("No parsed data found!");
        }
    }


}

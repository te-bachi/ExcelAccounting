package ch.fablabwinti.accounting.main;

import ch.fablabwinti.accounting.Account;
import ch.fablabwinti.accounting.AccountList;
import ch.fablabwinti.accounting.JournalFilter;
import ch.fablabwinti.accounting.Transaction;
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
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;

public class AccountConvertTwint {

    private static int INPUT_COLUMN_TWINT_DATE          = 1;
    private static int INPUT_COLUMN_TWINT_STATUS        = 3;
    private static int INPUT_COLUMN_TWINT_AMOUNT_TOTAL  = 4;
    private static int INPUT_COLUMN_TWINT_AMOUNT_FEE    = 21;
    private static int INPUT_COLUMN_TWINT_FIRSTNAME     = 30;
    private static int INPUT_COLUMN_TWINT_LASTNAME      = 31;
    private static int INPUT_COLUMN_TWINT_COMMENT       = 43;

    private static int COLUMN_WIDTH_RATIO               = 260;

    private static int OUTPUT_COLUMN_DATE               = 0;
    private static int OUTPUT_COLUMN_DEBIT              = 1;
    private static int OUTPUT_COLUMN_CREDIT             = 2;
    private static int OUTPUT_COLUMN_AMOUNT             = 3;
    private static int OUTPUT_COLUMN_LASTNAME           = 4;
    private static int OUTPUT_COLUMN_FIRSTNAME          = 5;
    private static int OUTPUT_COLUMN_COMMENT            = 6;

    private static int OUTPUT_COLUMN_DATE_WIDTH         = 14;
    private static int OUTPUT_COLUMN_DEBIT_WIDTH        = 10;
    private static int OUTPUT_COLUMN_CREDIT_WIDTH       = 10;
    private static int OUTPUT_COLUMN_AMOUNT_WIDTH       = 14;

    private SimpleDateFormat dateFormat;
    private ArrayList<TwintTransaction> transactions;

    private class TwintTransaction {
        public Date date;
        public BigDecimal amountTotal;
        public BigDecimal amountFee;

        public String firstname;

        public String lastname;

        public String comment;

        public TwintTransaction() {
            //
        }

        public String toString() {
            return dateFormat.format(date);
        }

        public Date getDate() {
            return date;
        }

        public void setDate(Date date) {
            this.date = date;
        }

        public BigDecimal getAmountTotal() {
            return amountTotal;
        }

        public void setAmountTotal(BigDecimal amountTotal) {
            this.amountTotal = amountTotal;
        }

        public BigDecimal getAmountFee() {
            return amountFee;
        }

        public void setAmountFee(BigDecimal amountFee) {
            this.amountFee = amountFee;
        }

        public String getFirstname() {
            return firstname;
        }

        public void setFirstname(String firstname) {
            this.firstname = firstname;
        }

        public String getLastname() {
            return lastname;
        }

        public void setLastname(String lastname) {
            this.lastname = lastname;
        }

        public String getComment() {
            return comment;
        }

        public void setComment(String comment) {
            this.comment = comment;
        }
    }

    public AccountConvertTwint() {
        transactions = new ArrayList<>();
        dateFormat = new SimpleDateFormat("dd.MM.yyyy");
    }

    public void parseInput(File inputFile) throws Exception {
        FileInputStream     in;
        XSSFWorkbook        workbook;
        XSSFSheet           spreadsheet;
        Iterator<Row>       rowIterator;
        XSSFRow             row;
        CustomCell          cell;
        TwintTransaction    transaction;

        in          = new FileInputStream(inputFile);
        workbook    = new XSSFWorkbook(in);
        spreadsheet = workbook.getSheetAt(0);
        System.out.println("Parsing sheet: " + spreadsheet.getSheetName());
        rowIterator = spreadsheet.iterator();
        while (rowIterator.hasNext()) {
            row = (XSSFRow) rowIterator.next();
            if (row.getRowNum() > 0) {
                try {

                    cell = new CustomStringCell(row, INPUT_COLUMN_TWINT_STATUS);

                    if (cell.getString().equals("succeeded")) {
                        transaction = new TwintTransaction();

                        /* try to fetch cell as INTEGER or die with exception */
                        cell = new CustomIntCell(row, INPUT_COLUMN_TWINT_DATE);
                        transaction.setDate(cell.getDate());

                        /* try to fetch cell as INTEGER or die with exception */
                        cell = new CustomIntCell(row, INPUT_COLUMN_TWINT_AMOUNT_TOTAL);
                        transaction.setAmountTotal(cell.getBigDecimal());

                        /* try to fetch cell as INTEGER or die with exception */
                        cell = new CustomIntCell(row, INPUT_COLUMN_TWINT_AMOUNT_FEE);
                        transaction.setAmountFee(cell.getBigDecimal());

                        /* try to fetch cell as STRING or die with exception */
                        try {
                            cell = new CustomStringCell(row, INPUT_COLUMN_TWINT_FIRSTNAME);
                            transaction.setFirstname(cell.getString());
                        } catch (CustomCellException e) {
                            System.out.println("<IGNORED> " + e.getMessage());
                            transaction.setFirstname("-");
                        }

                        /* try to fetch cell as STRING or die with exception */
                        try {
                            cell = new CustomStringCell(row, INPUT_COLUMN_TWINT_LASTNAME);
                            transaction.setLastname(cell.getString());
                        } catch (CustomCellException e) {
                            System.out.println("<IGNORED> " + e.getMessage());
                            transaction.setLastname("-");
                        }

                        /* try to fetch cell as STRING or die with exception */
                        try {
                            cell = new CustomStringCell(row, INPUT_COLUMN_TWINT_COMMENT);
                            transaction.setComment(cell.getString());
                        } catch (CustomCellException e) {
                            System.out.println("<IGNORED> " + e.getMessage());
                            transaction.setComment("");
                        }

                        transactions.add(transaction);
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

    public void exportOutput(File outputFile) throws Exception {
        FileOutputStream out;
        XSSFWorkbook            workbook;
        XSSFSheet               spreadsheet;
        XSSFRow                 row;
        Cell                    cell;
        JournalStyles           styles;
        int                     i;
        int                     k;
        int                     m;
        TwintTransaction        transaction;

        workbook        = new XSSFWorkbook();
        styles          = new JournalStyles(workbook);
        spreadsheet     = workbook.createSheet("output");

        spreadsheet.setColumnWidth(OUTPUT_COLUMN_DATE,      COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_DATE_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_DEBIT,     COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_DEBIT_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_CREDIT,    COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_CREDIT_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_AMOUNT,    COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_AMOUNT_WIDTH);


        for (i = 0, k = 0; i < transactions.size(); i++) {
            transaction = transactions.get(i);

            row = spreadsheet.createRow(k + 1); // + 1 for header
            new CellCreator(row, OUTPUT_COLUMN_DATE,        styles.dateStyle)  .createCell(dateFormat.format(transaction.date));
            new CellCreator(row, OUTPUT_COLUMN_DEBIT,       styles.dateStyle)  .createCell("TWINT");
            new CellCreator(row, OUTPUT_COLUMN_CREDIT,      styles.dateStyle)  .createCell("Lasercutter");
            new CellCreator(row, OUTPUT_COLUMN_AMOUNT,      styles.numberStyle).createCell(transaction.amountTotal.doubleValue());
            new CellCreator(row, OUTPUT_COLUMN_LASTNAME,    styles.normalStyle).createCell(transaction.getLastname());
            new CellCreator(row, OUTPUT_COLUMN_FIRSTNAME,   styles.normalStyle).createCell(transaction.getFirstname());
            new CellCreator(row, OUTPUT_COLUMN_COMMENT,     styles.normalStyle).createCell(transaction.getComment());

            k++;

            row = spreadsheet.createRow(k + 1);
            new CellCreator(row, OUTPUT_COLUMN_DATE,        styles.dateStyle)  .createCell(dateFormat.format(transaction.date));
            new CellCreator(row, OUTPUT_COLUMN_DEBIT,       styles.dateStyle)  .createCell("TWINT Gebühren\n");
            new CellCreator(row, OUTPUT_COLUMN_CREDIT,      styles.dateStyle)  .createCell("TWINT");
            new CellCreator(row, OUTPUT_COLUMN_AMOUNT,      styles.numberStyle).createCell(transaction.amountFee.doubleValue());
            new CellCreator(row, OUTPUT_COLUMN_LASTNAME,    styles.normalStyle).createCell(transaction.getLastname());
            new CellCreator(row, OUTPUT_COLUMN_FIRSTNAME,   styles.normalStyle).createCell(transaction.getFirstname());

            k++;

        }

        out = new FileOutputStream(outputFile);
        workbook.write(out);
        out.close();
    }

    public static void main(String[] args) throws Exception {
        AccountConvertTwint         main;
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

        main = new AccountConvertTwint();

        main.parseInput(inputFile);
        if ((main.transactions.size() > 0)) {
            main.exportOutput(outputFile);
            System.out.println("Done!");
        } else {
            System.out.println("No parsed data found!");
        }
    }

}


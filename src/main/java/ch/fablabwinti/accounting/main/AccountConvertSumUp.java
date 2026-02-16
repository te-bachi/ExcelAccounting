package ch.fablabwinti.accounting.main;

import ch.fablabwinti.accounting.cell.CustomCell;
import ch.fablabwinti.accounting.cell.CustomCellException;
import ch.fablabwinti.accounting.cell.CustomIntCell;
import ch.fablabwinti.accounting.cell.CustomStringCell;
import ch.fablabwinti.checkout.EmptyFilter;
import ch.fablabwinti.checkout.Item;
import com.opencsv.bean.CsvToBeanBuilder;
import com.opencsv.bean.CsvToBeanFilter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;

public class AccountConvertSumUp {

    private static int INPUT_COLUMN_SUMUP_DATE          = 1;
    private static int INPUT_COLUMN_SUMUP_STATUS        = 4;
    private static int INPUT_COLUMN_SUMUP_AMOUNT_TOTAL  = 12;
    private static int INPUT_COLUMN_SUMUP_AMOUNT_FEE    = 16;
    private static int INPUT_COLUMN_SUMUP_COMMENT       = 11;

    private static int COLUMN_WIDTH_RATIO               = 260;

    private static int OUTPUT_COLUMN_DATE               = 0;
    private static int OUTPUT_COLUMN_DEBIT              = 1;
    private static int OUTPUT_COLUMN_CREDIT             = 2;
    private static int OUTPUT_COLUMN_AMOUNT             = 3;
    private static int OUTPUT_COLUMN_LASTNAME           = 4;
    private static int OUTPUT_COLUMN_FIRSTNAME          = 5;
    private static int OUTPUT_COLUMN_COMMENT            = 6;
    private static int OUTPUT_COLUMN_TYPE               = 7;

    private static int OUTPUT_COLUMN_DATE_WIDTH         = 14;
    private static int OUTPUT_COLUMN_DEBIT_WIDTH        = 10;
    private static int OUTPUT_COLUMN_CREDIT_WIDTH       = 10;
    private static int OUTPUT_COLUMN_AMOUNT_WIDTH       = 14;
    private static int OUTPUT_COLUMN_TYPE_WIDTH         = 10;

    private SimpleDateFormat dateFormat;
    private List<SumUpTransaction> sumUpTransactions;
    private List<Item> checkoutTransactions;

    private class SumUpTransaction {
        public Date date;
        public BigDecimal amountTotal;
        public BigDecimal amountFee;

        public String comment;

        public SumUpTransaction() {
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

        public String getComment() {
            return comment;
        }

        public void setComment(String comment) {
            this.comment = comment;
        }
    }

    public AccountConvertSumUp() {
        sumUpTransactions = new ArrayList<>();
        checkoutTransactions = new ArrayList<>();
        dateFormat = new SimpleDateFormat("dd.MM.yyyy");
    }

    public void parseSumUpInput(File inputFile) throws Exception {
        FileInputStream in;
        XSSFWorkbook workbook;
        XSSFSheet spreadsheet;
        Iterator<Row> rowIterator;
        XSSFRow row;
        CustomCell cell;
        SumUpTransaction transaction;
        SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");

        in          = new FileInputStream(inputFile);
        workbook    = new XSSFWorkbook(in);
        spreadsheet = workbook.getSheetAt(0);
        System.out.println("Parsing sheet: " + spreadsheet.getSheetName());
        rowIterator = spreadsheet.iterator();
        while (rowIterator.hasNext()) {
            row = (XSSFRow) rowIterator.next();
            if (row.getRowNum() > 0) {
                try {

                    cell = new CustomStringCell(row, INPUT_COLUMN_SUMUP_STATUS);

                    if (cell.getString().equals("Erfolgreich")) {
                        transaction = new SumUpTransaction();

                        /* try to fetch cell as INTEGER or die with exception */
                        cell = new CustomStringCell(row, INPUT_COLUMN_SUMUP_DATE);
                        transaction.setDate(format.parse(cell.getString()));

                        /* try to fetch cell as INTEGER or die with exception */
                        cell = new CustomIntCell(row, INPUT_COLUMN_SUMUP_AMOUNT_TOTAL);
                        transaction.setAmountTotal(cell.getBigDecimal());

                        /* try to fetch cell as INTEGER or die with exception */
                        cell = new CustomIntCell(row, INPUT_COLUMN_SUMUP_AMOUNT_FEE);
                        transaction.setAmountFee(cell.getBigDecimal());

                        /* try to fetch cell as STRING or die with exception */
                        try {
                            cell = new CustomStringCell(row, INPUT_COLUMN_SUMUP_COMMENT);
                            transaction.setComment(cell.getString());
                        } catch (CustomCellException e) {
                            System.out.println("<IGNORED> " + e.getMessage());
                            transaction.setComment("");
                        }

                        sumUpTransactions.add(transaction);
                    }
                } catch (CustomCellException e) {
                    System.out.println(e.getMessage());
                } catch (IllegalStateException e) {
                    System.out.println("row " + row.getRowNum() + " has illegal cells");
                    throw e;
                }
            }
        }
        Collections.sort(sumUpTransactions, new Comparator<SumUpTransaction>() {
            @Override
            public int compare(SumUpTransaction a, SumUpTransaction b)
            {
                return a.getDate().compareTo(b.getDate());
            }
        });
    }



    private void parseCheckoutInput(File folder) throws IOException {
        char            cvsSplitBy  = ';';
        Class           klass       = Item.class;
        CsvToBeanFilter filter      = new EmptyFilter();
        List<Item> sublist;

        /* iterate over the folder */
        for (File file : folder.listFiles()) {

            /* if there are subfolders, ignore it */
            if (file.isDirectory()) {
                continue;
            }

            /* if the filename doesn't start with "Position", ignore it */
            if (!file.getName().startsWith("Positionen")) {
                continue;
            }

            System.out.println("Processing " + file.getName());

            /* parse CSV file */
            try {
                sublist = new CsvToBeanBuilder(new InputStreamReader(new FileInputStream(file), "Cp1252"))
                        .withSeparator(cvsSplitBy)
                        .withIgnoreQuotations(true)
                        .withType(klass)
                        .withFilter(filter)
                        .build()
                        .parse();

                checkoutTransactions.addAll(sublist);
            } catch (Exception e) {
                System.err.println("=== Exception in file " + file.getName());
                throw e;
            }
        }
        Collections.sort(checkoutTransactions, new Comparator<Item>() {
            @Override
            public int compare(Item a, Item b)
            {
                return a.getDate().compareTo(b.getDate());
            }
        });
    }
    public String getKeyByValue(Properties props, String valueToCheck) {
        return props.entrySet().stream()
                .filter(entry -> valueToCheck.equals(entry.getValue()))
                .map(Map.Entry::getKey)
                .map(Object::toString) // Sicherstellen, dass es ein String ist
                .findFirst()
                .orElse(null); // Oder eine Exception werfen / Optional zurückgeben
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
        SumUpTransaction        transaction;

        Properties kassenblattProperties = new Properties();
        kassenblattProperties.load(new FileInputStream("kassenblatt.properties"));

        workbook        = new XSSFWorkbook();
        styles          = new JournalStyles(workbook);
        spreadsheet     = workbook.createSheet("output");

        spreadsheet.setColumnWidth(OUTPUT_COLUMN_DATE,      COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_DATE_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_DEBIT,     COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_DEBIT_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_CREDIT,    COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_CREDIT_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_AMOUNT,    COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_AMOUNT_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_TYPE,      COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_TYPE_WIDTH);


        if (checkoutTransactions.size() == 0) {
            for (i = 0, k = 0; i < sumUpTransactions.size(); i++) {
                transaction = sumUpTransactions.get(i);

                row = spreadsheet.createRow(k + 1); // + 1 for header
                new CellCreator(row, OUTPUT_COLUMN_DATE, styles.dateStyle).createCell(dateFormat.format(transaction.date));
                //new CellCreator(row, OUTPUT_COLUMN_DEBIT, styles.dateStyle).createCell("SumUp");
                //new CellCreator(row, OUTPUT_COLUMN_CREDIT, styles.dateStyle).createCell("");
                new CellCreator(row, OUTPUT_COLUMN_DEBIT, styles.normalStyle).createCell(1030);
                new CellCreator(row, OUTPUT_COLUMN_CREDIT, styles.normalStyle).createCell("");
                new CellCreator(row, OUTPUT_COLUMN_AMOUNT, styles.numberStyle).createCell(transaction.amountTotal.doubleValue());
                new CellCreator(row, OUTPUT_COLUMN_LASTNAME, styles.normalStyle).createCell("");
                new CellCreator(row, OUTPUT_COLUMN_FIRSTNAME, styles.normalStyle).createCell("");
                new CellCreator(row, OUTPUT_COLUMN_COMMENT, styles.normalStyle).createCell(transaction.getComment());

                k++;

                row = spreadsheet.createRow(k + 1);
                new CellCreator(row, OUTPUT_COLUMN_DATE, styles.dateStyle).createCell(dateFormat.format(transaction.date));
                //new CellCreator(row, OUTPUT_COLUMN_DEBIT, styles.dateStyle).createCell("SumUp Gebühren");
                //new CellCreator(row, OUTPUT_COLUMN_CREDIT, styles.dateStyle).createCell("SumUp");
                new CellCreator(row, OUTPUT_COLUMN_DEBIT, styles.normalStyle).createCell(6841);
                new CellCreator(row, OUTPUT_COLUMN_CREDIT, styles.normalStyle).createCell(1030);
                new CellCreator(row, OUTPUT_COLUMN_AMOUNT, styles.numberStyle).createCell(transaction.amountFee.doubleValue());
                new CellCreator(row, OUTPUT_COLUMN_LASTNAME, styles.normalStyle).createCell("");
                new CellCreator(row, OUTPUT_COLUMN_FIRSTNAME, styles.normalStyle).createCell("");

                k++;
            }
        } else {
            for (i = 0, m = 0, k = 0; i < sumUpTransactions.size() || m < checkoutTransactions.size();) {
                SumUpTransaction sumUpTransaction = null;
                Item checkoutTransaction = null;
                try {
                    sumUpTransaction = sumUpTransactions.get(i);
                } catch (IndexOutOfBoundsException e) {
                    //
                }

                try {
                    checkoutTransaction = checkoutTransactions.get(m);
                } catch (IndexOutOfBoundsException e) {
                    //
                }

                /* Checkout */
                if (m < checkoutTransactions.size() && (sumUpTransaction == null ||  checkoutTransaction.getDate().compareTo(sumUpTransaction.getDate()) <= 0)) {

                    if (checkoutTransaction.getPaymentMethod().equalsIgnoreCase("SumUp")) {
                        String firstname = "";
                        String lastname = "";
                        if (checkoutTransaction.getMember() != null) {
                            String[] member = checkoutTransaction.getMember().split(" ");

                            if (member.length >= 1) {
                                firstname = member[0];
                            }

                            if (member.length >= 2) {
                                lastname = member[1];
                            }

                            if (member.length >= 3) {
                                System.out.println("Member has long name: \"" + checkoutTransaction.getMember() + "\"");
                            }
                        }
                        row = spreadsheet.createRow(k + 1); // + 1 for header
                        new CellCreator(row, OUTPUT_COLUMN_DATE, styles.dateStyle).createCell(dateFormat.format(checkoutTransaction.getDate()));
                        //new CellCreator(row, OUTPUT_COLUMN_DEBIT, styles.dateStyle).createCell("SumUp");
                        //new CellCreator(row, OUTPUT_COLUMN_CREDIT, styles.dateStyle).createCell(checkoutTransaction.getPosition());
                        new CellCreator(row, OUTPUT_COLUMN_DEBIT, styles.normalStyle).createCell(1030);
                        //new CellCreator(row, OUTPUT_COLUMN_CREDIT, styles.dateStyle).createCell(checkoutTransaction.getPosition());
                        try {
                            new CellCreator(row, OUTPUT_COLUMN_CREDIT, styles.normalStyle).createCell(Integer.valueOf(kassenblattProperties.getProperty(checkoutTransaction.getPosition())).intValue());
                        } catch (NumberFormatException e) {
                            new CellCreator(row, OUTPUT_COLUMN_CREDIT, styles.normalStyle).createCell(checkoutTransaction.getPosition());
                        }
                        new CellCreator(row, OUTPUT_COLUMN_AMOUNT, styles.numberStyle).createCell(Optional.ofNullable(checkoutTransaction.getAmount()).map(c -> c.doubleValue()).orElse(Double.valueOf(0.00)));
                        new CellCreator(row, OUTPUT_COLUMN_LASTNAME, styles.normalStyle).createCell(lastname);
                        new CellCreator(row, OUTPUT_COLUMN_FIRSTNAME, styles.normalStyle).createCell(firstname);
                        new CellCreator(row, OUTPUT_COLUMN_COMMENT, styles.normalStyle).createCell(checkoutTransaction.getText());
                        new CellCreator(row, OUTPUT_COLUMN_TYPE, styles.normalStyle).createCell("Kassenblatt");

                        k++;
                    }

                    m++;

                    /* TWINT */
                } else {

                    row = spreadsheet.createRow(k + 1); // + 1 for header
                    new CellCreator(row, OUTPUT_COLUMN_DATE, styles.dateStyle).createCell(dateFormat.format(sumUpTransaction.date));
                    //new CellCreator(row, OUTPUT_COLUMN_DEBIT, styles.dateStyle).createCell("SumUp");
                    //new CellCreator(row, OUTPUT_COLUMN_CREDIT, styles.dateStyle).createCell("");
                    new CellCreator(row, OUTPUT_COLUMN_DEBIT, styles.normalStyle).createCell(1030);
                    new CellCreator(row, OUTPUT_COLUMN_CREDIT, styles.normalStyle).createCell("");
                    new CellCreator(row, OUTPUT_COLUMN_AMOUNT, styles.numberStyle).createCell(sumUpTransaction.amountTotal.doubleValue());
                    new CellCreator(row, OUTPUT_COLUMN_LASTNAME, styles.normalStyle).createCell("");
                    new CellCreator(row, OUTPUT_COLUMN_FIRSTNAME, styles.normalStyle).createCell("");
                    new CellCreator(row, OUTPUT_COLUMN_COMMENT, styles.normalStyle).createCell(sumUpTransaction.getComment());
                    new CellCreator(row, OUTPUT_COLUMN_TYPE, styles.normalStyle).createCell("SumUp");

                    k++;

                    row = spreadsheet.createRow(k + 1);
                    new CellCreator(row, OUTPUT_COLUMN_DATE, styles.dateStyle).createCell(dateFormat.format(sumUpTransaction.date));
                    //new CellCreator(row, OUTPUT_COLUMN_DEBIT, styles.dateStyle).createCell("SumUp Gebühren");
                    //new CellCreator(row, OUTPUT_COLUMN_CREDIT, styles.dateStyle).createCell("SumUp");
                    new CellCreator(row, OUTPUT_COLUMN_DEBIT, styles.normalStyle).createCell(6841);
                    new CellCreator(row, OUTPUT_COLUMN_CREDIT, styles.normalStyle).createCell(1030);
                    new CellCreator(row, OUTPUT_COLUMN_AMOUNT, styles.numberStyle).createCell(sumUpTransaction.amountFee.doubleValue());
                    new CellCreator(row, OUTPUT_COLUMN_LASTNAME, styles.normalStyle).createCell("");
                    new CellCreator(row, OUTPUT_COLUMN_FIRSTNAME, styles.normalStyle).createCell("");
                    // OUTPUT_COLUMN_COMMENT
                    new CellCreator(row, OUTPUT_COLUMN_TYPE, styles.normalStyle).createCell("SumUp");

                    k++;

                    i++;
                }
            }
        }

        out = new FileOutputStream(outputFile);
        workbook.write(out);
        out.close();
    }

    public static void main(String[] args) throws Exception {
        AccountConvertSumUp         main;
        String                      input;
        File                        inputRaiseNowFile;
        File                        inputCheckoutFolder;
        File                        outputFile;

        if (args.length < 2) {
            System.out.println("<input raisenow xlsx> (<input checkout folder>)");
            return;
        }

        /* argInput */
        input = args[0];
        inputRaiseNowFile = new File(input);

        outputFile = new File(input.substring(0, input.lastIndexOf('.'))  + "_output" + input.substring(input.lastIndexOf('.'), input.length()));


        if (!inputRaiseNowFile.exists()) {
            System.out.println("Input file \"" + inputRaiseNowFile.getAbsolutePath() + "\" doesn't exist! Abort");
            return;
        }

        if (outputFile.exists()) {
            System.out.println("Output file \"" + outputFile.getAbsolutePath() + "\" exist! Abort");
            return;
        }

        inputCheckoutFolder = null;
        if (args.length == 2) {
            inputCheckoutFolder = new File(args[1]);
        }

        main = new AccountConvertSumUp();

        main.parseSumUpInput(inputRaiseNowFile);
        if (inputCheckoutFolder != null) {
            main.parseCheckoutInput(inputCheckoutFolder);
        }

        if ((main.sumUpTransactions.size() > 0)) {
            main.exportOutput(outputFile);
            System.out.println("Done!");
        } else {
            System.out.println("No parsed data found!");
        }
    }
}

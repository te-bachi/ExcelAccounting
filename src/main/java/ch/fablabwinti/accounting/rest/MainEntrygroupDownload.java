package ch.fablabwinti.accounting.rest;

import ch.fablabwinti.accounting.main.CellCreator;
import ch.fablabwinti.accounting.main.JournalStyles;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.net.URI;
import java.net.URISyntaxException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

public class MainEntrygroupDownload {


    private static int COLUMN_WIDTH_RATIO               = 260;

    private static int OUTPUT_COLUMN_DATE               = 1;
    private static int OUTPUT_COLUMN_DEBIT              = 2;
    private static int OUTPUT_COLUMN_CREDIT             = 3;
    private static int OUTPUT_COLUMN_TITLE              = 4;
    private static int OUTPUT_COLUMN_AMOUNT             = 5;

    private static int OUTPUT_COLUMN_DATE_WIDTH         = 14;
    private static int OUTPUT_COLUMN_DEBIT_WIDTH        = 10;
    private static int OUTPUT_COLUMN_CREDIT_WIDTH       = 10;
    private static int OUTPUT_COLUMN_TITLE_WIDTH        = 45;
    private static int OUTPUT_COLUMN_AMOUNT_WIDTH       = 14;

    private Properties idProperties;
    private RestClient restClient;

    private List<Entrygroup> entrygroupList;

    private class Entrygroup {

        public String date;
        public String title;
        public String amount;
        public int debit;
        public int credit;

        public Entrygroup() {
            //
        }

        public Entrygroup(String date, String title, String amount, int debit, int credit) {
            this.date = date;
            this.title = title;
            this.amount = amount;
            this.debit = debit;
            this.credit = credit;
        }

    }


    MainEntrygroupDownload() throws IOException {

        idProperties = new Properties();
        idProperties.load(new FileInputStream("id.properties"));
        restClient = new RestClient();
        entrygroupList = new ArrayList<>();
    }

    public void process() throws URISyntaxException, Exception {
        RestObjectList entrygroupList = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/entrygroup"), RestObjectList.class);
        try {
            for (int entrygroupId : entrygroupList.getObjects()) {
                RestObject entrygroup = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/entrygroup/" + entrygroupId), RestObject.class);
                String dateString = entrygroup.getProperties().get("date");
                String title = entrygroup.getProperties().get("title");
                int debitId = entrygroup.getLinks().get("account")[0];
                int creditId = entrygroup.getLinks().get("account")[1];
                int entryId = entrygroup.getChildren().get("entry")[0];

                entrygroup = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/account/" + debitId), RestObject.class);
                int debit = Integer.valueOf(entrygroup.getProperties().get("title").split(" ")[0]);

                entrygroup = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/account/" + creditId), RestObject.class);
                int credit = Integer.valueOf(entrygroup.getProperties().get("title").split(" ")[0]);

                entrygroup = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/entry/" + entryId), RestObject.class);
                String amountString = entrygroup.getProperties().get("amount");

                System.out.println(" > " + dateString + " " + title + " " + debit + " " + credit + " " + amountString);

                this.entrygroupList.add(new Entrygroup(dateString, title, amountString, debit, credit));
            }
        } catch (NullPointerException e) {
            System.out.println("error: " + entrygroupList.getError());
        }
    }

    public void parseOutput(File outputFile) throws IOException, ParseException {
        FileOutputStream out;
        XSSFWorkbook workbook;
        XSSFSheet spreadsheet;
        XSSFRow row;
        Cell cell;
        JournalStyles styles;
        int                     i;
        int                     k;
        int                     m;
        Entrygroup entrygroup;

        workbook        = new XSSFWorkbook();
        styles          = new JournalStyles(workbook);
        spreadsheet     = workbook.createSheet("output");

        SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");

        spreadsheet.setColumnWidth(OUTPUT_COLUMN_DATE,      COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_DATE_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_DEBIT,     COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_DEBIT_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_CREDIT,    COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_CREDIT_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_TITLE,    COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_TITLE_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_AMOUNT,     COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_AMOUNT_WIDTH);


        for (i = 0, k = 0; i < entrygroupList.size(); i++) {
            entrygroup = entrygroupList.get(i);

            row = spreadsheet.createRow(i); // + 1 for header

            new CellCreator(row, OUTPUT_COLUMN_DATE, styles.dateStyle).createCell(formatter.parse(entrygroup.date));
            new CellCreator(row, OUTPUT_COLUMN_DEBIT, styles.normalStyle).createCell(entrygroup.debit);
            new CellCreator(row, OUTPUT_COLUMN_CREDIT, styles.normalStyle).createCell(entrygroup.credit);
            new CellCreator(row, OUTPUT_COLUMN_TITLE, styles.normalStyle).createCell(entrygroup.title);
            new CellCreator(row, OUTPUT_COLUMN_AMOUNT, styles.numberStyle).createCell(Double.valueOf(entrygroup.amount).doubleValue());
        }

        out = new FileOutputStream(outputFile);
        workbook.write(out);
        out.close();
    }

    public static void main(String[] args) throws IOException, URISyntaxException, Exception {
        File outputFile;

        if (args.length < 1 || args.length > 1) {
            System.out.println("<input>");
            return;
        }

        outputFile = new File(args[0]);

        if (outputFile.exists()) {
            System.out.println("Output file \"" + outputFile.getAbsolutePath() + "\" does exist!");
            return;
        }
        MainEntrygroupDownload main = new MainEntrygroupDownload();
        main.process();
        main.parseOutput(outputFile);
        System.out.println("Done!");
    }
}

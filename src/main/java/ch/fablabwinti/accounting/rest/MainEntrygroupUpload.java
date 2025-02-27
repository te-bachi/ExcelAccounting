package ch.fablabwinti.accounting.rest;

import ch.fablabwinti.accounting.cell.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;


public class MainEntrygroupUpload {

    public static int INPUT_ROW_SKIP = 1;
    public static int INPUT_COLUMN_DATE = 0;
    public static int INPUT_COLUMN_AMOUNT_DEBIT = 1;
    public static int INPUT_COLUMN_AMOUNT_CREDIT = 2;
    public static int INPUT_COLUMN_AMOUNT = 3;
    public static int INPUT_COLUMN_DEBIT_ID = 4;
    public static int INPUT_COLUMN_CREDIT_ID = 5;
    public static int INPUT_COLUMN_DESCRIPTION = 6;
    public static int INPUT_COLUMN_SURNAME = 7;
    public static int INPUT_COLUMN_FIRSTNAME = 8;

    private Properties idProperties;
    private RestClient restClient;
    private List<Entrygroup> entrygroupList;


    public MainEntrygroupUpload() throws IOException {
        idProperties = new Properties();
        idProperties.load(new FileInputStream("id.properties"));
        restClient = new RestClient();
        entrygroupList = new ArrayList<>();
    }

    private class Entrygroup {

        public String date;
        public String title;
        public String amount;
        public int debitId;
        public int creditId;

        public Entrygroup() {
            //
        }

        public Entrygroup(String date, String title, String amount, int debitId, int creditId) {
            this.date = date;
            this.title = title;
            this.amount = amount;
            this.debitId = debitId;
            this.creditId = creditId;
        }

    }

    public void parseInput(File inputFile) throws Exception {
        FileInputStream in;
        XSSFWorkbook workbook;
        XSSFSheet spreadsheet;
        FormulaEvaluator evaluator;
        Iterator<Row> rowIterator;
        XSSFRow row;
        CustomCell cell;
        Entrygroup entrygroup;
        SimpleDateFormat formatter;

        formatter = new SimpleDateFormat("yyyy-MM-dd");
        in = new FileInputStream(inputFile);
        workbook = new XSSFWorkbook(in);
        spreadsheet = workbook.getSheetAt(0);
        evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        rowIterator = spreadsheet.iterator();

        boolean isSkipped = true;

        for (int skipCount = 0; skipCount < MainEntrygroupUpload.INPUT_ROW_SKIP; skipCount++) {
            isSkipped = rowIterator.hasNext();
            if (!isSkipped) {
                break;
            }
            rowIterator.next();
        }
        if (isSkipped) {
            while (rowIterator.hasNext()) {
                row = (XSSFRow) rowIterator.next();
                entrygroup = new Entrygroup();

                /* Date */
                cell = new CustomDateCell(row, MainEntrygroupUpload.INPUT_COLUMN_DATE);
                entrygroup.date = formatter.format(cell.getDate());

                /* Amount */
                cell = new CustomIntCell(row, MainEntrygroupUpload.INPUT_COLUMN_AMOUNT, evaluator);
                entrygroup.amount = cell.getBigDecimalString();

                /* Debit Id */
                cell = new CustomIntCell(row, MainEntrygroupUpload.INPUT_COLUMN_DEBIT_ID);
                entrygroup.debitId = Integer.valueOf(idProperties.getProperty(String.valueOf(cell.getInt()))).intValue();

                /* Credit Id */
                cell = new CustomIntCell(row, MainEntrygroupUpload.INPUT_COLUMN_CREDIT_ID);
                entrygroup.creditId = Integer.valueOf(idProperties.getProperty(String.valueOf(cell.getInt()))).intValue();

                /* Description */
                cell = new CustomStringCell(row, MainEntrygroupUpload.INPUT_COLUMN_DESCRIPTION);
                entrygroup.title = cell.getString();
                entrygroup.title += ", ";

                /* Surname */
                cell = new CustomStringCell(row, MainEntrygroupUpload.INPUT_COLUMN_SURNAME);
                entrygroup.title += cell.getString();

                /* Firstname */
                try {
                    cell = new CustomStringCell(row, MainEntrygroupUpload.INPUT_COLUMN_FIRSTNAME);
                    entrygroup.title += " ";
                    entrygroup.title += cell.getString();
                } catch (CustomCellException e) {
                    System.out.println("e: " + e.getMessage());
                }

                entrygroupList.add(entrygroup);
            }
        }
    }

    public void process() throws URISyntaxException, Exception {

        //Entrygroup entrygroup = new Entrygroup("2024-01-24", "Test Andreas XYZ", "50.20", 361, 2298);
        for (Entrygroup entrygroup : entrygroupList) {
            postEntrygroup(297, entrygroup);
        }

    }

    public void postEntrygroup(int periodId, Entrygroup entrygroup) throws URISyntaxException, Exception {
        String jsonString = "{" +
                "  \"properties\": {" +
                "    \"date\": \"" + entrygroup.date + "\"," +
                "    \"title\": \"" + entrygroup.title + "\"" +
                "  }," +
                "  \"children\": {" +
                "    \"entry\": [" +
                "      {\n" +
                "        \"properties\": {" +
                "          \"amount\": " + entrygroup.amount +
                "        }," +
                "        \"links\": {" +
                "          \"debit\": [" +
                "            " + entrygroup.debitId +
                "          ]," +
                "          \"credit\": [" +
                "            " + entrygroup.creditId +
                "          ]" +
                "        }" +
                "      }" +
                "    ]" +
                "  }," +
                "  \"parents\": [" +
                "    " + periodId +
                "  ]" +
                "}";

        this.restClient.postRequest(new URI("https://fablabwinti.webling.ch/api/1/entrygroup"), jsonString);
    }

    public static void main(String[] args) throws IOException, URISyntaxException, Exception {
        File inputFile;

        if (args.length < 1 || args.length > 1) {
            System.out.println("<input>");
            return;
        }

        inputFile = new File(args[0]);

        if (!inputFile.exists()) {
            System.out.println("Input file \"" + inputFile.getAbsolutePath() + "\" doesn't exist!");
            return;
        }
        MainEntrygroupUpload main = new MainEntrygroupUpload();
        main.parseInput(inputFile);
        main.process();
        System.out.println("Done!");
    }
}

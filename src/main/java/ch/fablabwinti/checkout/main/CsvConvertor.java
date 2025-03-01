package ch.fablabwinti.checkout.main;

import ch.fablabwinti.checkout.EmptyFilter;
import ch.fablabwinti.checkout.Item;
import com.opencsv.bean.CsvToBeanBuilder;
import com.opencsv.bean.CsvToBeanFilter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.ParseException;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;
import java.util.Date;
import java.text.SimpleDateFormat;

/**
 *
 */
public class CsvConvertor {
    public static int OUTPUT_ROW_OFFSET                             = 1;

    public static int COLUMN_WIDTH_RATIO                            = 260;

    /* Only for "Kasse" */
    public static int OUTPUT_COLUMN_JOURNAL_NR                      = 0;
    public static int OUTPUT_COLUMN_JOURNAL_DATE                    = 1;
    public static int OUTPUT_COLUMN_JOURNAL_SOLL                    = 2;
    public static int OUTPUT_COLUMN_JOURNAL_HABEN                   = 3;
    public static int OUTPUT_COLUMN_JOURNAL_TEXT                    = 4;
    public static int OUTPUT_COLUMN_JOURNAL_FIRSTNAME               = 5;
    public static int OUTPUT_COLUMN_JOURNAL_LASTNAME                = 6;
    public static int OUTPUT_COLUMN_JOURNAL_MACHINE_3D              = 7;
    public static int OUTPUT_COLUMN_JOURNAL_MACHINE_LATHE           = 8;
    public static int OUTPUT_COLUMN_JOURNAL_MACHINE_LASER           = 9;
    public static int OUTPUT_COLUMN_JOURNAL_DRINKS                  = 10;
    public static int OUTPUT_COLUMN_JOURNAL_OTHER                   = 11;
    public static int OUTPUT_COLUMN_JOURNAL_WORKSHOP                = 12;
    public static int OUTPUT_COLUMN_JOURNAL_MEMBERSHIP_FEE          = 13;
    public static int OUTPUT_COLUMN_JOURNAL_PURCHASE                = 14;

    public static int OUTPUT_COLUMN_JOURNAL_NR_WIDTH                = 6;
    public static int OUTPUT_COLUMN_JOURNAL_DATE_WIDTH              = 14;
    public static int OUTPUT_COLUMN_JOURNAL_SOLL_WIDTH              = 10;
    public static int OUTPUT_COLUMN_JOURNAL_HABEN_WIDTH             = 10;
    public static int OUTPUT_COLUMN_JOURNAL_TEXT_WIDTH              = 50;
    public static int OUTPUT_COLUMN_JOURNAL_FIRSTNAME_WIDTH         = 20;
    public static int OUTPUT_COLUMN_JOURNAL_LASTNAME_WIDTH          = 20;
    public static int OUTPUT_COLUMN_JOURNAL_MACHINE_3D_WIDTH        = 14;
    public static int OUTPUT_COLUMN_JOURNAL_MACHINE_LATHE_WIDTH     = 14;
    public static int OUTPUT_COLUMN_JOURNAL_MACHINE_LASER_WIDTH     = 14;
    public static int OUTPUT_COLUMN_JOURNAL_DRINKS_WIDTH            = 14;
    public static int OUTPUT_COLUMN_JOURNAL_OTHER_WIDTH             = 14;
    public static int OUTPUT_COLUMN_JOURNAL_WORKSHOP_WIDTH          = 14;
    public static int OUTPUT_COLUMN_JOURNAL_MEMBERSHIP_FEE_WIDTH    = 14;
    public static int OUTPUT_COLUMN_JOURNAL_PURCHASE_WIDTH          = 14;

    private List<Item>      list;

    // debug
    private Date            debugDate;

    public CsvConvertor() {
        list = new ArrayList<Item>();

        // debug
        try {
            debugDate = new SimpleDateFormat("dd.MM.yyyy").parse("08.10.2020");
        } catch (ParseException e) {
            System.out.println("ParseException on SimpleDateFormat");
            System.exit(-1);
        }
    }

    private void parseInput(File folder) throws IOException {
        char            cvsSplitBy  = ';';
        Class           klass       = Item.class;
        CsvToBeanFilter filter      = new EmptyFilter();
        List<Item>      sublist;

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

                for (Item item : sublist) {
                    if (item.getDate().compareTo(debugDate) == 0) {
                        System.out.println("break");
                    }
                }

                list.addAll(sublist);
            } catch (Exception e) {
                System.err.println("=== Exception in file " + file.getName());
                throw e;
            }
        }
        Collections.sort(list, new Comparator<Item>() {
            @Override
            public int compare(Item a, Item b)
            {
                return a.getDate().compareTo(b.getDate());
            }
        });
    }

    public void exportOutput(File outputFile) throws IOException {
        FileOutputStream    out;
        XSSFWorkbook        workbook;
        XSSFSheet           spreadsheetCash;
        XSSFSheet           spreadsheetCredit;
        XSSFSheet           spreadsheetSumUp;
        XSSFSheet           spreadsheetTwint;
        XSSFRow             row;
        Cell                cell;
        CellStyle           dateStyle;
        CreationHelper      createHelper;
        Item                item;
        int                 k;
        int                 kCash;
        int                 kCredit;
        int                 kSumUp;
        int                 kTwint;
        int                 positionIdx;
        String              sollStr = null;
        String              habenStr = null;

        workbook            = new XSSFWorkbook();
        spreadsheetCash     = workbook.createSheet("cash");
        spreadsheetCredit   = workbook.createSheet("credit");
        spreadsheetSumUp    = workbook.createSheet("sumup");
        spreadsheetTwint    = workbook.createSheet("twint");
        dateStyle           = workbook.createCellStyle();
        createHelper        = workbook.getCreationHelper();
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("DD.MM.YYYY"));

        spreadsheetCash.setColumnWidth(OUTPUT_COLUMN_JOURNAL_NR,            COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_NR_WIDTH);
        spreadsheetCash.setColumnWidth(OUTPUT_COLUMN_JOURNAL_DATE,          COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_DATE_WIDTH);
        spreadsheetCash.setColumnWidth(OUTPUT_COLUMN_JOURNAL_SOLL,          COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_SOLL_WIDTH);
        spreadsheetCash.setColumnWidth(OUTPUT_COLUMN_JOURNAL_HABEN,         COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_HABEN_WIDTH);
        spreadsheetCash.setColumnWidth(OUTPUT_COLUMN_JOURNAL_TEXT,          COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_TEXT_WIDTH);
        spreadsheetCash.setColumnWidth(OUTPUT_COLUMN_JOURNAL_FIRSTNAME,     COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_FIRSTNAME_WIDTH);
        spreadsheetCash.setColumnWidth(OUTPUT_COLUMN_JOURNAL_LASTNAME,      COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_LASTNAME_WIDTH);
        spreadsheetCash.setColumnWidth(OUTPUT_COLUMN_JOURNAL_MACHINE_3D,    COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_MACHINE_3D_WIDTH);
        spreadsheetCash.setColumnWidth(OUTPUT_COLUMN_JOURNAL_MACHINE_LATHE, COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_MACHINE_LATHE_WIDTH);
        spreadsheetCash.setColumnWidth(OUTPUT_COLUMN_JOURNAL_MACHINE_LASER, COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_MACHINE_LASER_WIDTH);
        spreadsheetCash.setColumnWidth(OUTPUT_COLUMN_JOURNAL_DRINKS,        COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_DRINKS_WIDTH);
        spreadsheetCash.setColumnWidth(OUTPUT_COLUMN_JOURNAL_OTHER,         COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_OTHER_WIDTH);
        spreadsheetCash.setColumnWidth(OUTPUT_COLUMN_JOURNAL_WORKSHOP,      COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_WORKSHOP_WIDTH);
        spreadsheetCash.setColumnWidth(OUTPUT_COLUMN_JOURNAL_MEMBERSHIP_FEE,COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_MEMBERSHIP_FEE_WIDTH);
        spreadsheetCash.setColumnWidth(OUTPUT_COLUMN_JOURNAL_PURCHASE,      COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_PURCHASE_WIDTH);

        /* header */
        row = spreadsheetCash.createRow(0);
        row.createCell(OUTPUT_COLUMN_JOURNAL_NR).setCellValue("nr");
        row.createCell(OUTPUT_COLUMN_JOURNAL_DATE).setCellValue("date");
        row.createCell(OUTPUT_COLUMN_JOURNAL_SOLL).setCellValue("soll");
        row.createCell(OUTPUT_COLUMN_JOURNAL_HABEN).setCellValue("haben");
        row.createCell(OUTPUT_COLUMN_JOURNAL_TEXT).setCellValue("text");
        row.createCell(OUTPUT_COLUMN_JOURNAL_FIRSTNAME).setCellValue("firstname");
        row.createCell(OUTPUT_COLUMN_JOURNAL_LASTNAME).setCellValue("lastname");
        row.createCell(OUTPUT_COLUMN_JOURNAL_MACHINE_3D).setCellValue("3d drucker");
        row.createCell(OUTPUT_COLUMN_JOURNAL_MACHINE_LATHE).setCellValue("drehbank");
        row.createCell(OUTPUT_COLUMN_JOURNAL_MACHINE_LASER).setCellValue("laser-cutter");
        row.createCell(OUTPUT_COLUMN_JOURNAL_DRINKS).setCellValue("getränke");
        row.createCell(OUTPUT_COLUMN_JOURNAL_OTHER).setCellValue("sonstige");
        row.createCell(OUTPUT_COLUMN_JOURNAL_WORKSHOP).setCellValue("workshops");
        row.createCell(OUTPUT_COLUMN_JOURNAL_MEMBERSHIP_FEE).setCellValue("mitglied");
        row.createCell(OUTPUT_COLUMN_JOURNAL_PURCHASE).setCellValue("kauf");

        kCash   = 0;
        kCredit = 0;
        kSumUp  = 0;
        kTwint  = 0;
        for (k = 0; k < list.size(); k++) {
            item = list.get(k);
            if (item.getPaymentMethod().equalsIgnoreCase("Bar")) {
                row = spreadsheetCash.createRow(kCash + OUTPUT_ROW_OFFSET);
                kCash++;
            } else if (item.getPaymentMethod().equalsIgnoreCase("Guthaben")) {
                row = spreadsheetCredit.createRow(kCredit + OUTPUT_ROW_OFFSET);
                kCredit++;
            } else if (item.getPaymentMethod().equalsIgnoreCase("SumUp")) {
                row = spreadsheetSumUp.createRow(kSumUp + OUTPUT_ROW_OFFSET);
                kSumUp += 2;
            } else if (item.getPaymentMethod().equalsIgnoreCase("Twint")) {
                row = spreadsheetTwint.createRow(kTwint + OUTPUT_ROW_OFFSET);
                kTwint += 2;
            } else {
                System.out.println("Unknow payment method: \"" + item.getPaymentMethod() + "\"");
            }

            cell = row.createCell(OUTPUT_COLUMN_JOURNAL_NR);
            cell.setCellValue(0);

            cell = row.createCell(OUTPUT_COLUMN_JOURNAL_DATE);
            cell.setCellValue(item.getDate());
            cell.setCellStyle(dateStyle);

            if (item.getMember() != null) {
                String[] member = item.getMember().split(" ");

                if (member.length >= 1) {
                    cell = row.createCell(OUTPUT_COLUMN_JOURNAL_FIRSTNAME);
                    cell.setCellValue(member[0]);
                }

                if (member.length >= 2) {
                    cell = row.createCell(OUTPUT_COLUMN_JOURNAL_LASTNAME);
                    cell.setCellValue(member[1]);
                }

                if (member.length >= 3) {
                    System.out.println("Member has long name: \"" + item.getMember() + "\"");
                }
            }

            sollStr = "Kasse";

            if (item.getPaymentMethod().equals("SumUp")) {
                sollStr = "SumUp";
            }

            if (item.getPosition() == null) {
                positionIdx = OUTPUT_COLUMN_JOURNAL_OTHER;
            } else if (item.getPosition().equals("Lasercutter") || item.getPosition().equals("LaserSaur")) {
                habenStr = "Lasercutter LaserSaur";

                positionIdx = OUTPUT_COLUMN_JOURNAL_MACHINE_LASER;

            } else if (item.getPosition().equals("Trotec")) {
                habenStr = "Lasercutter Trotec";
                positionIdx = OUTPUT_COLUMN_JOURNAL_MACHINE_LASER;

            } else if (item.getPosition().equals("3D Drucker")) {
                habenStr = "3D Drucker";
                positionIdx = OUTPUT_COLUMN_JOURNAL_MACHINE_3D;

            } else if (item.getPosition().equals("Fräsmaschine")) {
                habenStr = "Fräsmaschine";
                positionIdx = OUTPUT_COLUMN_JOURNAL_OTHER;

            } else if (item.getPosition().equals("Drehbank")) {
                habenStr = "Drehbank";
                positionIdx = OUTPUT_COLUMN_JOURNAL_MACHINE_LATHE;

            } else if (item.getPosition().equals("Getränke/Food")) {
                habenStr = "Getränke Verkauf";
                positionIdx = OUTPUT_COLUMN_JOURNAL_DRINKS;

            } else if (item.getPosition().equals("Mitgliederbeitrag")) {
                habenStr = "Mitgliederbeitrag";
                positionIdx = OUTPUT_COLUMN_JOURNAL_MEMBERSHIP_FEE;

            } else if (item.getPosition().equals("Elektronikbauteile")) {
                habenStr = "Elektronikbauteil Verkauf";
                positionIdx = OUTPUT_COLUMN_JOURNAL_OTHER;

            } else if (item.getPosition().equals("FabLab Kit")) {
                habenStr = "FabLab-Kits Verkauf";
                positionIdx = OUTPUT_COLUMN_JOURNAL_OTHER;

            } else if (item.getPosition().equals("Spende")) {
                habenStr = "Spenden";
                positionIdx = OUTPUT_COLUMN_JOURNAL_OTHER;

            } else if (item.getPosition().equals("Boxmiete")) {
                habenStr = "Mieteinnahmen Stoffboxen";
                if (item.getText() == null || item.getText().isEmpty()) {
                    item.setText("Boxmiete");
                }
                positionIdx = OUTPUT_COLUMN_JOURNAL_OTHER;

            } else if (item.getPosition().equals("Kursgebühren")) {
                habenStr = "Workshop EggBot Einnahmen";
                if (item.getText() == null || item.getText().isEmpty()) {
                    item.setText("Workshop");
                }
                positionIdx = OUTPUT_COLUMN_JOURNAL_OTHER;

            } else if (item.getPosition().equals("Plattenmaterial")) {
                habenStr = "Plattenmaterial Verkauf";
                if (item.getText() == null || item.getText().isEmpty()) {
                    item.setText("Plattenmaterial");
                }
                positionIdx = OUTPUT_COLUMN_JOURNAL_OTHER;

            } else if (item.getPosition().equals("Spesen")) {
                sollStr = "Forderungen Postüberweisung";
                habenStr = "Kasse";
                positionIdx = OUTPUT_COLUMN_JOURNAL_PURCHASE;

            } else if (item.getPosition().equals("Diverses")) {
                habenStr = "-";
                positionIdx = OUTPUT_COLUMN_JOURNAL_OTHER;
            } else {
                System.out.println("Konto \"" + item.getPosition() + "\" nicht gefunden");
                habenStr = "-";
                positionIdx = OUTPUT_COLUMN_JOURNAL_OTHER;
            }

            cell = row.createCell(OUTPUT_COLUMN_JOURNAL_TEXT);
            cell.setCellValue(item.getText());

            /* Only for "Kasse" */

            //if ((positionIdx < OUTPUT_COLUMN_JOURNAL_OTHER || positionIdx > OUTPUT_COLUMN_JOURNAL_OTHER) && positionIdx < OUTPUT_COLUMN_JOURNAL_PURCHASE) {
                row.createCell(OUTPUT_COLUMN_JOURNAL_SOLL).setCellValue(sollStr);
                row.createCell(OUTPUT_COLUMN_JOURNAL_HABEN).setCellValue(habenStr);
            //}

            cell = row.createCell(positionIdx);
            if (item.getAmount() != null) {
                cell.setCellValue(item.getAmount().doubleValue());
            } else {
                cell.setCellValue(0.0f);
            }
        }
        out = new FileOutputStream(outputFile);
        workbook.write(out);
        out.close();
    }

    public static void main(String[] args) {
        File            folder;
        CsvConvertor    convertor;


        if (args.length < 1) {
            System.out.println("path to CSV directory file?");
            System.exit(-1);
        }


        try {
            /* creates a folder class from the argument */
            folder = new File(args[0]);
            if (!folder.exists()) {
                System.out.println("file does not exist!");
                System.exit(-2);
            }

            /* if the argument is a file, use the folder of the file */
            if (folder.isFile()) {
                folder = folder.getParentFile();
            } else {
                folder = folder;
            }

            convertor = new CsvConvertor();
            System.out.println("=== Parse input ===");
            convertor.parseInput(folder);
            System.out.println("=== Export output ===");
            convertor.exportOutput(new File(folder.getPath() + "/output.xlsx"));

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}

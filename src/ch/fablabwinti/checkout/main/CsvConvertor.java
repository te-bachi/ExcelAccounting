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
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

/**
 *
 */
public class CsvConvertor {
    public static int COLUMN_WIDTH_RATIO                            = 260;

    /* Only for "Kasse" */
    public static int OUTPUT_COLUMN_JOURNAL_NR                      = 0;
    public static int OUTPUT_COLUMN_JOURNAL_DATE                    = 1;
    public static int OUTPUT_COLUMN_JOURNAL_TEXT                    = 2;
    public static int OUTPUT_COLUMN_JOURNAL_FIRSTNAME               = 3;
    public static int OUTPUT_COLUMN_JOURNAL_LASTNAME                = 4;
    public static int OUTPUT_COLUMN_JOURNAL_MACHINE_3D              = 5;
    public static int OUTPUT_COLUMN_JOURNAL_MACHINE_LATHE           = 6;
    public static int OUTPUT_COLUMN_JOURNAL_MACHINE_LASER           = 7;
    public static int OUTPUT_COLUMN_JOURNAL_DRINKS                  = 8;
    public static int OUTPUT_COLUMN_JOURNAL_OTHER                   = 9;
    public static int OUTPUT_COLUMN_JOURNAL_WORKSHOP                = 10;
    public static int OUTPUT_COLUMN_JOURNAL_MEMBERSHIP_FEE          = 11;
    public static int OUTPUT_COLUMN_JOURNAL_PURCHASE                = 12;

    public static int OUTPUT_COLUMN_JOURNAL_NR_WIDTH                = 6;
    public static int OUTPUT_COLUMN_JOURNAL_DATE_WIDTH              = 14;
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

    public CsvConvertor() {
        list = new ArrayList<Item>();
    }

    private void parseInput(File folder) throws IOException {
        char            cvsSplitBy  = ';';
        Class           klass       = Item.class;
        CsvToBeanFilter filter      = new EmptyFilter();

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
            list.addAll(new CsvToBeanBuilder(new InputStreamReader(new FileInputStream(file), "Cp1252"))
                    .withSeparator(cvsSplitBy)
                    .withIgnoreQuotations(true)
                    .withType(klass)
                    .withFilter(filter)
                    .build()
                    .parse());
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
        XSSFSheet           spreadsheet;
        XSSFRow             row;
        Cell                cell;
        CellStyle           dateStyle;
        CreationHelper      createHelper;
        Item                item;
        int                 k;
        int                 positionIdx;

        workbook        = new XSSFWorkbook();
        spreadsheet     = workbook.createSheet("cash");
        dateStyle       = workbook.createCellStyle();
        createHelper    = workbook.getCreationHelper();
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("DD.MM.YYYY"));

        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_NR,            COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_NR_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_DATE,          COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_DATE_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_TEXT,          COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_TEXT_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_FIRSTNAME,     COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_FIRSTNAME_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_LASTNAME,      COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_LASTNAME_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_MACHINE_3D,    COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_MACHINE_3D_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_MACHINE_LATHE, COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_MACHINE_LATHE_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_MACHINE_LASER, COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_MACHINE_LASER_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_DRINKS,        COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_DRINKS_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_OTHER,         COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_OTHER_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_WORKSHOP,      COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_WORKSHOP_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_MEMBERSHIP_FEE,COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_MEMBERSHIP_FEE_WIDTH);
        spreadsheet.setColumnWidth(OUTPUT_COLUMN_JOURNAL_PURCHASE,      COLUMN_WIDTH_RATIO * OUTPUT_COLUMN_JOURNAL_PURCHASE_WIDTH);

        for (k = 0; k < list.size(); k++) {
            item = list.get(k);
            row = spreadsheet.createRow(k);

            cell = row.createCell(OUTPUT_COLUMN_JOURNAL_NR);
            cell.setCellValue(0);

            cell = row.createCell(OUTPUT_COLUMN_JOURNAL_DATE);
            cell.setCellValue(item.getDate());
            cell.setCellStyle(dateStyle);

            cell = row.createCell(OUTPUT_COLUMN_JOURNAL_TEXT);
            cell.setCellValue(item.getText());

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

            if (item.getPosition() == null) {
                positionIdx = OUTPUT_COLUMN_JOURNAL_OTHER;
            } else if (item.getPosition().equals("Lasercutter")) {
                positionIdx = OUTPUT_COLUMN_JOURNAL_MACHINE_LASER;
            } else if (item.getPosition().equals("3D Drucker")) {
                positionIdx = OUTPUT_COLUMN_JOURNAL_MACHINE_3D;
            } else if (item.getPosition().equals("Fräsmaschine")) {
                positionIdx = OUTPUT_COLUMN_JOURNAL_OTHER;
            } else if (item.getPosition().equals("Drehbank")) {
                positionIdx = OUTPUT_COLUMN_JOURNAL_MACHINE_LATHE;
            } else if (item.getPosition().equals("Getränke/Food")) {
                positionIdx = OUTPUT_COLUMN_JOURNAL_DRINKS;
            } else if (item.getPosition().equals("Mitgliederbeitrag")) {
                positionIdx = OUTPUT_COLUMN_JOURNAL_MACHINE_LASER;
            } else {
                positionIdx = OUTPUT_COLUMN_JOURNAL_OTHER;
            }

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
            convertor.parseInput(folder);
            convertor.exportOutput(new File(folder.getPath() + "/output.xlsx"));

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}

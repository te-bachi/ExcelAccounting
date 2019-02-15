package ch.fablabwinti.accounting.main;

import ch.fablabwinti.accounting.*;
import ch.fablabwinti.accounting.cell.CustomCell;
import ch.fablabwinti.accounting.cell.CustomCellException;
import ch.fablabwinti.accounting.cell.CustomIntCell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.*;

public class AccountSummaryExport {

    public static int   COLUMN_WIDTH_RATIO                  = 260;

    private static int[] INPUT_TRANSACTION_DATE             = { 1, 4, 7, 10 };
    private static int[] INPUT_TRANSACTION_VALUE            = { 2, 5, 8, 11 };


    private static int   OUTPUT_ROW_OFFSET                  = 2;
    private static int   OUTPUT_TRANSACTION_DATE            = 1;
    private static int[] OUTPUT_TRANSACTION_VALUE           = { 2, 3, 4, 5 };

    private static int   OUTPUT_TRANSACTION_DATE_WIDTH     = 14;
    private static int   OUTPUT_TRANSACTION_VALUE_WIDTH    = 10;


    private AccountSummaryList accountSummaryList;

    public AccountSummaryExport() {
        accountSummaryList = new AccountSummaryList();
    }

    public void parseInput(File inputFile) throws Exception {
        FileInputStream     in;
        XSSFWorkbook        workbook;
        XSSFSheet           spreadsheet;
        FormulaEvaluator    evaluator;
        Iterator<Row>       rowIterator;
        XSSFRow             row;
        CustomCell          cell;
        int                 i;
        Date                date;
        BigDecimal          value;

        in          = new FileInputStream(inputFile);
        workbook    = new XSSFWorkbook(in);
        evaluator   = workbook.getCreationHelper().createFormulaEvaluator();
        spreadsheet = workbook.getSheetAt(0);
        System.out.println("Parsing sheet " + 0);
        for (i = 0; i < 4; i++) {
            rowIterator = spreadsheet.iterator();

            while (rowIterator.hasNext()) {
                row = (XSSFRow) rowIterator.next();
                if (row.getRowNum() > 0) {
                    try {

                        //System.out.println("column " + INPUT_TRANSACTION_DATE[i] + " row " + row.getRowNum());
                        /* try to fetch cell as INTEGER or die with exception */
                        cell = new CustomIntCell(row, INPUT_TRANSACTION_DATE[i]);
                        date = cell.getDate();

                        /* try to fetch cell as INTEGER or die with exception */
                        cell = new CustomIntCell(row, INPUT_TRANSACTION_VALUE[i], evaluator);
                        value = cell.getBigDecimal();

                        accountSummaryList.addEntry(date, i, value);
                    } catch (CustomCellException e) {
                        System.out.println(e.getMessage());
                        break;
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
        CellStyle dateStyle;
        CreationHelper      createHelper;
        int                 i;
        int                 k;
        Date                date;
        int                 accountNr;
        double[]            value = {0.0, 0.0, 0.0, 0.0};
        AccountSummary      accountSummary;

        workbook        = new XSSFWorkbook();
        spreadsheet     = workbook.createSheet("journal");
        dateStyle       = workbook.createCellStyle();
        createHelper    = workbook.getCreationHelper();
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("DD.MM.YYYY"));

        spreadsheet.setColumnWidth(OUTPUT_TRANSACTION_DATE,        COLUMN_WIDTH_RATIO * OUTPUT_TRANSACTION_DATE_WIDTH);
        for (i = 0; i < 4; i++) {
            spreadsheet.setColumnWidth(OUTPUT_TRANSACTION_VALUE[i], COLUMN_WIDTH_RATIO * OUTPUT_TRANSACTION_VALUE_WIDTH);
        }

        row = spreadsheet.createRow(0 + OUTPUT_ROW_OFFSET - 1);

        for (i = 0; i < 4; i++) {
            cell = row.createCell(OUTPUT_TRANSACTION_VALUE[i]);
            cell.setCellValue(i);
        }

        Set<Date> keys = accountSummaryList.getKeys();
        List list = new ArrayList(keys);
        Collections.sort(list);
        Iterator<Date> dateIterator = list.iterator();
        for (k = 0; k < keys.size(); k++) {
            date = dateIterator.next();
            accountSummary = accountSummaryList.getDate(date);
            row = spreadsheet.createRow(k + OUTPUT_ROW_OFFSET);

            cell = row.createCell(OUTPUT_TRANSACTION_DATE);
            cell.setCellValue(date);
            cell.setCellStyle(dateStyle);

            for (accountNr = 0; accountNr < value.length; accountNr++) {
                try {
                    //accountNr = accountSummary.getAccountNr(i);
                    value[accountNr] = accountSummary.getValue(accountNr).doubleValue();
                } catch (NullPointerException e) {
                    //
                }
                cell = row.createCell(OUTPUT_TRANSACTION_VALUE[accountNr]);
                cell.setCellValue(value[accountNr]);
            }
        }
        out = new FileOutputStream(outputFile);
        workbook.write(out);
        out.close();
    }

    public static void main(String[] args) throws Exception {
        File                    inputFile;
        File                    outputFile;
        AccountSummaryExport    main;

        if (args.length < 1 || args.length > 1) {
            System.out.println("<input>");
            return;
        }
        inputFile   = new File(args[0]);
        outputFile  = new File(args[0].substring(0, args[0].lastIndexOf('.'))  + "_output" + args[0].substring(args[0].lastIndexOf('.'), args[0].length()));

        if (!inputFile.exists()) {
            System.out.println("Input file \"" + inputFile.getAbsolutePath() + "\" doesn't exist!");
            return;
        }

        if (outputFile.exists()) {
            System.out.println("Output file \"" + outputFile.getAbsolutePath() + "\" exist! Abort");
            return;
        }

        main = new AccountSummaryExport();

        main.parseInput(inputFile);
        main.exportOutput(outputFile);

        AccountSummary accountSummary = main.accountSummaryList.getDate("31.12.2018");
        System.out.println("Done");
    }
}

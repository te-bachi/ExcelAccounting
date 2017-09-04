package ch.fablabwinti.checkout.main;

import ch.fablabwinti.accounting.cell.CustomCell;
import ch.fablabwinti.accounting.cell.CustomCellException;
import ch.fablabwinti.accounting.cell.CustomStringCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Iterator;

/**
 *
 */
public class CashSheetMerger {
    public static void main(String[] args) {
        if (args.length < 1) {
            System.out.println("no path specified!");
            System.exit(-1);
        }

        File                folder;
        BufferedReader      br          = null;
        String              line;
        String              cvsSplitBy  = ";";
        StringBuilder       builder;
        String[]            lineArray;

        FileInputStream     in;
        XSSFWorkbook        workbook;
        XSSFSheet           spreadsheet;
        FormulaEvaluator    evaluator;
        Iterator<Row>       rowIterator;
        XSSFRow             row;
        CustomCell          cell;
        int                 idx = 0;
        SimpleDateFormat    dateFormat = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss");
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

            /* iterate over the folder */
            for (File file : folder.listFiles()) {
                /* if there are subfolders, ignore it */
                if (file.isDirectory()) {
                    continue;
                }


                in = new FileInputStream(file);
                workbook = new XSSFWorkbook(in);
                spreadsheet = workbook.getSheetAt(0);
                evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                rowIterator = spreadsheet.iterator();
                while (rowIterator.hasNext()) {
                    row = (XSSFRow) rowIterator.next();
                    if (row.getRowNum() > 26 && row.getRowNum() < 48) {
                        cell = new CustomStringCell(row, 8);
                        String value = cell.getString();
                        if (value.length() > 0) {
                            idx++;
                            System.out.println(String.valueOf(idx) + ": " + dateFormat.format(file.lastModified()) + ", " + value);
                        }
                    }
                }
            }
        } catch (CustomCellException e) {
            System.out.println("error: " + e.getMessage());
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (br != null) {
                try {
                    br.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}

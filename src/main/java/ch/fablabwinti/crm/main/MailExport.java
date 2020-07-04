package ch.fablabwinti.crm.main;

import ch.fablabwinti.accounting.cell.CustomCell;
import ch.fablabwinti.accounting.cell.CustomCellException;
import ch.fablabwinti.accounting.cell.CustomStringCell;
import ch.fablabwinti.crm.Address;

import ch.fablabwinti.crm.AddressComparator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;

/**
 *
 */
public class MailExport {

    public static int INPUT_COLUMN_FIRSTNAME    = 2;
    public static int INPUT_COLUMN_LASTNAME     = 3;
    public static int INPUT_COLUMN_EMAIL        = 7;

    private List<Address> addressList;

    public MailExport() {
        addressList = new ArrayList<>();
    }

    public void parseInput(File inputFile) throws Exception {
        FileInputStream     in;
        XSSFWorkbook        workbook;
        XSSFSheet           spreadsheet;
        Iterator<Row>       rowIterator;
        XSSFRow             row;
        CustomCell          cell;
        int                 i;
        Address             address;

        in          = new FileInputStream(inputFile);
        workbook    = new XSSFWorkbook(in);

        spreadsheet = workbook.getSheetAt(1);
        System.out.println("Parsing sheet " + spreadsheet.getSheetName());
        rowIterator = spreadsheet.iterator();
        while (rowIterator.hasNext()) {
            row = (XSSFRow) rowIterator.next();
            if (row.getRowNum() > 0) {
                try {
                    address = new Address();

                    cell = new CustomStringCell(row, INPUT_COLUMN_FIRSTNAME);
                    address.setFirstname(cell.getString());

                    cell = new CustomStringCell(row, INPUT_COLUMN_LASTNAME);
                    address.setLastname(cell.getString());

                    cell = new CustomStringCell(row, INPUT_COLUMN_EMAIL);
                    address.setEmail(cell.getString());

                    addressList.add(address);
                } catch (CustomCellException e) {
                    System.out.println(e.getMessage());
                }
            }
        }
        in.close();
    }

    public static void main(String[] args) throws Exception {
        File                inputFile;
        File                outputFile;
        MailExport          main;

        if (args.length < 1 || args.length > 2) {
            System.out.println("<input> [<account filter>]");
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

        main = new MailExport();

        main.parseInput(inputFile);
        if ((main.addressList.size() > 0)) {
            Collections.sort(main.addressList, new AddressComparator());
            System.out.println("Done!");
            for (Address a : main.addressList) {
                //System.out.println("\"" + a.getFirstname() + " " + a.getLastname() + "\" <" + a.getEmail() + ">");
                System.out.print(" " + a.getFirstname() + " " + a.getLastname() + " <" + a.getEmail() + ">,");
            }
        } else {
            System.out.println("No parsed data found!");
        }

    }
}

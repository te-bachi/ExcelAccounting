package ch.fablabwinti.checkout.main;

import java.io.*;

/**
 *
 */
public class CsvConvertor {
    public static void main(String[] args) {
        if (args.length < 1) {
            System.out.println("path to CSV file?");
            System.exit(-1);
        }
        String          csvFile     = null;
        BufferedReader  br          = null;
        String          line;
        String          cvsSplitBy  = ";";
        StringBuilder   builder;
        String[]        lineArray;
        File            folder;
        try {
            folder = new File(args[0]);
            if (!folder.exists()) {
                System.out.println("file does not exist!");
                System.exit(-2);
            }
            if (folder.isFile()) {
                folder = folder.getParentFile();
            } else {
                folder = folder;
            }
            for (File file : folder.listFiles()) {
                br = new BufferedReader(new InputStreamReader(new FileInputStream(file), "ISO-8859-15"));
                while ((line = br.readLine()) != null) {
                    /* Delete double quotes */
                    builder = new StringBuilder(line);
                    builder.delete(0, 1);
                    builder.delete(builder.length() - 1, builder.length());
                    line = builder.toString();

                    /* Check if string is not empty */
                    if (line.length() > 0) {
                        lineArray = line.split(cvsSplitBy);

                        System.out.println(lineArray[0]);
                    }

                }
            }

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

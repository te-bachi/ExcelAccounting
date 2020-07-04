package ch.fablabwinti.checkout.main;

import java.io.FileReader;
import java.io.StringReader;

import com.opencsv.CSVParser;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import com.opencsv.bean.CsvToBeanBuilder;

public class BlaTest {
    public static void main(String[] args) throws Exception {

        String filename = "C:\\Users\\bachman0\\Documents\\FabLab\\Buchhaltung\\2017\\KassenblattAblage-20170903T071453Z-001\\KassenblattAblage\\CSV\\Positionen_2017-09-01.csv";
        final CSVParser parser =
                new CSVParserBuilder()
                        .withSeparator(';')
                        .withIgnoreQuotations(true)
                        .build();
        final CSVReader reader =
                new CSVReaderBuilder(new FileReader(filename))
                        .withCSVParser(parser)
                        .build();
        String[] line;

        while ((line = reader.readNext()) != null && line.length == 6) {
            System.out.println("bla");
        }
    }
}

package ch.fablabwinti.checkout.main;

import ch.fablabwinti.checkout.EmptyFilter;
import ch.fablabwinti.checkout.Item;
import com.opencsv.bean.CsvToBean;
import com.opencsv.bean.CsvToBeanBuilder;

import java.io.FileReader;
import java.util.List;

public class Bla2Test {
    public static void main(String[] args) throws Exception {

        //String filename = "C:\\Users\\bachman0\\Documents\\FabLab\\Buchhaltung\\2017\\KassenblattAblage-20170903T071453Z-001\\KassenblattAblage\\CSV\\Positionen_2017-09-01.csv";
        String filename = "C:\\Users\\bachman0\\Documents\\FabLab\\Buchhaltung\\2017\\KassenblattAblage-20170903T071453Z-001\\KassenblattAblage\\CSV\\Positionen_2017-06-17.csv";

        List<Item> list = new CsvToBeanBuilder(new FileReader(filename))
                .withSeparator(';')
                .withIgnoreQuotations(true)
                .withType(Item.class)
                .withFilter(new EmptyFilter())
                .build()
                .parse();
        list.size();

    }
}

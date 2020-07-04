package ch.fablabwinti.checkout;

import com.opencsv.bean.CsvToBeanFilter;
import com.opencsv.bean.MappingStrategy;

public class EmptyFilter implements CsvToBeanFilter {

    public boolean allowLine(String[] line) {
        String value = line[0];
        return !value.isEmpty();
    }
}
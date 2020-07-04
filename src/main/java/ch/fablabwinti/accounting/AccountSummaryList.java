package ch.fablabwinti.accounting;

import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Set;

public class AccountSummaryList {

    private HashMap<Date, AccountSummary> map;

    public AccountSummaryList() {
        map = new HashMap<>();
    }

    public void addEntry(Date date, int accountNr, BigDecimal value) {
        try {
            DateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");
            Date testDate = dateFormat.parse("31.12.2018");

            AccountSummary accountSummary = map.get(date);
            if (accountSummary == null) {
                accountSummary = new AccountSummary(date);
                map.put(date, accountSummary);
                if (date.equals(testDate)) {
                    System.out.println("create entry with value " + value.toString());
                }
            } else {
                if (date.equals(testDate)) {
                    System.out.println("add entry with value " + value.toString());
                }
            }
            accountSummary.addEntry(accountNr, value);
        } catch (ParseException e) {

        }
    }

    public AccountSummary getDate(String dateStr) throws ParseException {
        return getDate(new SimpleDateFormat("dd.MM.yyyy").parse(dateStr));
    }

    public AccountSummary getDate(Date date) {
        return map.get(date);
    }

    public Set<Date> getKeys() {
        return map.keySet();
    }
}

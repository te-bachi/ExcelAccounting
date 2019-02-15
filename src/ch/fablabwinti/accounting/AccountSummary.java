package ch.fablabwinti.accounting;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class AccountSummary {

    /*
    public class Entry {
        public int accountNr;
        public BigDecimal value;

        public Entry() {

        }

        public Entry(int accountNr, BigDecimal value) {
            this.accountNr = accountNr;
            this.value = value;
        }
    }
    */

    private Date date;
    //private List<Entry> entryList;
    //private int[] accountNr;
    private BigDecimal[] value;

    public AccountSummary(Date date) {
        this.date = date;
        //this.entryList = new ArrayList<>();
        //this.accountNr = new int[4];
        this.value = new BigDecimal[4];
    }

    public void addEntry(int accountNr, BigDecimal value) {
        //this.entryList.add(new Entry(accountNr, value));
        this.value[accountNr] = value;
    }

    /*
    public int getAccountNr(int idx) {
        return entryList.get(idx).accountNr;
    }

    public BigDecimal getValue(int idx) {
        return entryList.get(idx).value;
    }
    */

    public BigDecimal getValue(int accountNr) {
        return this.value[accountNr];
    }
}

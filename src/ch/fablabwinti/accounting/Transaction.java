package ch.fablabwinti.accounting;

import java.math.BigDecimal;
import java.util.Date;

/**
 *
 */
public class Transaction {
    private int         nr;
    private Date        date;
    private Account     debit;
    private Account     credit;
    private BigDecimal  amount;
    private String      text;

    public Transaction() {

    }

    public Transaction(int nr, Date date, int debit, int credit, BigDecimal amount, String text) {
        this.nr     = nr;
        this.date   = date;
        this.debit  = new TitleAccount(debit, "");
        this.credit = new TitleAccount(credit, "");
        this.amount = amount;
        this.text   = text;
    }

    public Transaction(int nr, Date date, Account debit, Account credit, BigDecimal amount, String text) {
        this.nr     = nr;
        this.date   = date;
        this.debit  = debit;
        this.credit = credit;
        this.amount = amount;
        this.text   = text;
    }

    public int getNr() {
        return nr;
    }

    public void setNr(int nr) {
        this.nr = nr;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public Account getDebit() {
        return debit;
    }

    public void setDebit(Account debit) {
        this.debit = debit;
    }

    public Account getCredit() {
        return credit;
    }

    public void setCredit(Account credit) {
        this.credit = credit;
    }

    public BigDecimal getAmount() {
        return amount;
    }

    public void setAmount(BigDecimal amount) {
        this.amount = amount;
    }

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }
}

package ch.fablabwinti.accounting;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

/**
 *
 */
public abstract class Account {
    private Account             parent;
    private List<Account>       childList;
    private int                 number;
    private String              name;
    private BigDecimal          total;
    private List<Transaction>   transactionList;

    public Account(Account parent, int number, String name) {
        this.parent             = parent;
        this.childList          = new ArrayList<>();
        if (parent != null) {
            parent.addChild(this);
        }
        this.number             = number;
        this.name               = name;
        this.total              = new BigDecimal(0.0);
        this.transactionList    = new ArrayList<>();
    }

    public Account (int number, String name) {
        this(null, number, name);
    }

    public Account getParent() {
        return parent;
    }

    public void setParent(Account parent) {
        this.parent = parent;
    }

    public List<Account> getChildList() {
        return childList;
    }

    public void setChildList(List<Account> childList) {
        this.childList = childList;
    }

    public void addChild(Account child) {
        this.childList.add(child);
    }

    public int getNumber() {
        return number;
    }

    public void setNumber(int number) {
        this.number = number;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public BigDecimal getTotal() {
        return total;
    }

    public void setTotal(BigDecimal total) {
        this.total = total;
    }

    public List<Transaction> getTransactionList() {
        return transactionList;
    }

    public void setTransactionList(List<Transaction> transactionList) {
        this.transactionList = transactionList;
    }
}
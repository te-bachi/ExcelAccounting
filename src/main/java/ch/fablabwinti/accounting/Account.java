package ch.fablabwinti.accounting;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Comparator;
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
    private List<String>        keywordList;
    private List<Transaction>   transactionList;

    public Account(Account parent, int number, String name, List<String> keywordList) {
        this.parent             = parent;
        this.keywordList        = keywordList;
        this.childList          = new ArrayList<>();
        if (parent != null) {
            parent.addChild(this);
        }
        this.number             = number;
        this.name               = name;
        this.total              = new BigDecimal(0.0);
        this.transactionList    = new ArrayList<>();
    }

    public Account (Account parent, int number, String name) {
        this(parent, number, name, null);
    }

    public Account (int number, String name) {
        this(null, number, name, null);
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

    public List<String> getKeywordList() {
        return keywordList;
    }

    public void setKeywordList(List<String> keywordList) {
        this.keywordList = keywordList;
    }

    public List<Transaction> getTransactionList() {
        return transactionList;
    }

    public void setTransactionList(List<Transaction> transactionList) {
        this.transactionList = transactionList;
    }

    /**
     * Add total to Account and all parent TitleAccounts
     * @param amount
     */
    public void addTotal(BigDecimal amount) {
        total = total.add(amount);
        if (parent != null) {
            parent.addTotal(amount);
        }
    }

    /**
     * Subtract total from Account and all parent TitleAccounts
     * @param amount
     */
    public void subtractTotal(BigDecimal amount) {
        total = total.subtract(amount);
        if (parent != null) {
            parent.subtractTotal(amount);
        }
    }

    public void addTransaction(Transaction transaction) {
        this.transactionList.add(transaction);
        this.transactionList.sort(new Comparator<Transaction>() {
            @Override
            public int compare(Transaction o1, Transaction o2) {
                int a = o1.getDate().compareTo(o2.getDate());
                if (a == 0) {
                    if (o1.getNr() > o2.getNr()) {
                        return 1;
                    } else {
                        return -1;
                    }
                }
                return a;
            }
        });

        /* Debit */
        if (transaction.getDebit().getNumber() == number) {
            addTotal(transaction.getAmount());
        }

        /* Credit */
        if (transaction.getCredit().getNumber() == number) {
            subtractTotal(transaction.getAmount());
        }
    }

    @Override
    public String toString() {
        return name;
    }
}

package ch.fablabwinti.accounting;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 *
 */
public class AccountList implements Iterable<Account> {
    private List<Account> accountList;
    private boolean       matchExact;

    public AccountList() {
        this(new ArrayList(), true);
    }

    public AccountList(List<Account> accountList) {
        this(accountList, true);
    }

    public AccountList(List<Account> accountList, boolean matchExact) {
        this.accountList = accountList;
        this.matchExact  = matchExact;
    }

    public void add(Account account) {
        accountList.add(account);
    }

    public Account get(int index) {
        return accountList.get(index);
    }

    public int size() {
        return accountList.size();
    }

    public boolean isEmpty() {
        return accountList.isEmpty();
    }

    public Account removeNext() {
        return accountList.remove(0);
    }

    public void clear() {
        for (Account account : accountList) {
            account.setTotal(new BigDecimal(0.0));
        }
    }

    /**
     * Find account number only in the first layer, not in the deeper layers
     *
     * @param accountNr
     * @return
     * @throws AccountNotFoundException
     */
    public Account find(int accountNr) throws AccountNotFoundException {
        Account match = null;

        /* Find *ONLY* exact account */
        if (matchExact) {
            for (Account account : accountList) {
                if (account.getNumber() == accountNr) {
                    match = account;
                    break;
                }
            }

        /* Find exact or parent account */
        } else {
            /* Calculate depth */
            try {
                int     depth;
                int     firstDepth;
                int     reducedAccountNr;
                boolean finish = false;

                depth       = Integer.valueOf(accountNr).toString().length();
                firstDepth  = depth;
                do {
                    reducedAccountNr = accountNr /  Double.valueOf(Math.pow(10, firstDepth - depth)).intValue();
                    for (Account account : accountList) {
                        if (account.getNumber() == reducedAccountNr) {
                            match = account;
                            finish = true;
                            break;
                        }
                    }
                    depth--;
                } while (depth > 0 && !finish);
            } catch (NumberFormatException e) {
                //
            }
        }

        if (match == null) {
            throw new AccountNotFoundException(accountNr);
        }

        return match;
    }

    public Account remove(int accountNr) {
        Account match = null;
        Account account;
        int     i;

        for (i = 0; i < accountList.size(); i++) {
            account = accountList.get(i);
            if (account.getNumber() == accountNr) {
                match = accountList.remove(i);
                break;
            }
        }
        return match;
    }

    @Override
    public Iterator<Account> iterator() {
        return accountList.iterator();
    }
}

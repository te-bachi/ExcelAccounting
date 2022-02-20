package ch.fablabwinti.checkout;

import com.opencsv.bean.CsvBindByPosition;
import com.opencsv.bean.CsvDate;

import java.math.BigDecimal;
import java.util.Date;

/**
 *
 */
public class Item {
    @CsvBindByPosition(position = 0)
    @CsvDate("dd.MM.yyyy")
    private Date        date;

    @CsvBindByPosition(position = 1)
    private String position;

    @CsvBindByPosition(position = 2)
    private String      member;

    @CsvBindByPosition(position = 3)
    private String      text;

    @CsvBindByPosition(position = 4)
    private int         duration;

    @CsvBindByPosition(position = 5)
    private BigDecimal  amount;

    @CsvBindByPosition(position = 6)
    private String  paymentMethod;

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public String getPosition() {
        return position;
    }

    public void setPosition(String position) {
        this.position = position;
    }

    public String getMember() {
        return member;
    }

    public void setMember(String member) {
        this.member = member;
    }

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    public int getDuration() {
        return duration;
    }

    public void setDuration(int duration) {
        this.duration = duration;
    }

    public BigDecimal getAmount() {
        return amount;
    }

    public void setAmount(BigDecimal amount) {
        this.amount = amount;
    }

    public String getPaymentMethod() {
        return paymentMethod;
    }

    public void setPaymentMethod(String paymentMethod) {
        this.paymentMethod = paymentMethod;
    }
}

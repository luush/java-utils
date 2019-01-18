package test;

import java.util.Date;

public class ExcelBean {
    private String name;

    private Integer age;

    private double balance;

    private long idCard;

    private boolean flag;

    private Date updateTime;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public double getBalance() {
        return balance;
    }

    public void setBalance(double balance) {
        this.balance = balance;
    }

    public long getIdCard() {
        return idCard;
    }

    public void setIdCard(long idCard) {
        this.idCard = idCard;
    }

    public boolean getFlag() {
        return flag;
    }

    public void setFlag(boolean flag) {
        this.flag = flag;
    }

    public Date getUpdateTime() {
        return updateTime;
    }

    public void setUpdateTime(Date updateTime) {
        this.updateTime = updateTime;
    }

    @Override
    public String toString() {
        return "ExcelBean{" +
                "name='" + name + '\'' +
                ", age=" + age +
                ", balance=" + balance +
                ", idCard=" + idCard +
                ", flag=" + flag +
                ", updateTime=" + updateTime +
                '}';
    }
}

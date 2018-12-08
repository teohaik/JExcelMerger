package gr.teohaik.jexcelmerger;

public class Employee {
  
    String aem;
    String surname;
    String name;
    String fatherOrHusbandName;
    String birthdate;
    String afm;
    String amka;
    String address;
    String bankAccount;
    String bank;
    
    String hireDate;
    String leaveDate;
    int monthsInsured;
    
    double average3Years;
    double average5Years;
   
    double efapax;
    double efapaxTaken;
    
    int servedDays, servedMonths, servedYears;
    int strikeDays, strikeMonths, strikeYears;
    
    int parentalLeaveDays, parentalLeaveMonths, parentalLeaveYears;
    int unpaidLeaveDays, unpaidLeaveMonths, unpaidLeaveYears;
    int unjustifiedAbsenceDays, unjustifiedAbsenceMonths, unjustifiedAbsenceYears;
    int pauseDays, pauseMonths, pauseYears;
    
    String praxiAponomis;
    
    String telephone1;
    String telephone2;
    String telephone3;

    public void setAem(String aem) {
        this.aem = aem;
    }

    public void setSurname(String surname) {
        this.surname = surname;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setFatherOrHusbandName(String fatherOrHusbandName) {
        this.fatherOrHusbandName = fatherOrHusbandName;
    }

    public void setBirthdate(String birthdate) {
        this.birthdate = birthdate;
    }

    public void setAfm(String afm) {
        this.afm = afm;
    }

    public void setAmka(String amka) {
        this.amka = amka;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public void setTelephone1(String telephone1) {
        this.telephone1 = telephone1;
    }

    public void setTelephone2(String telephone2) {
        this.telephone2 = telephone2;
    }

    public void setTelephone3(String telephone3) {
        this.telephone3 = telephone3;
    }

    public void setBankAccount(String bankAccount) {
        this.bankAccount = bankAccount;
    }

    public void setBank(String bank) {
        this.bank = bank;
    }

    public void setHireDate(String hireDate) {
        this.hireDate = hireDate;
    }

    public void setLeaveDate(String leaveDate) {
        this.leaveDate = leaveDate;
    }

    public void setMonthsInsured(int monthsInsured) {
        this.monthsInsured = monthsInsured;
    }

    public void setAverage3Years(double average3Years) {
        this.average3Years = average3Years;
    }

    public void setAverage5Years(double average5Years) {
        this.average5Years = average5Years;
    }

    public void setEfapax(double efapax) {
        this.efapax = efapax;
    }

    public void setEfapaxTaken(double efapaxTaken) {
        this.efapaxTaken = efapaxTaken;
    }

    public void setStrikeDays(int strikeDays) {
        this.strikeDays = strikeDays;
    }

    public void setStrikeMonths(int strikeMonths) {
        this.strikeMonths = strikeMonths;
    }

    public void setStrikeYears(int strikeYears) {
        this.strikeYears = strikeYears;
    }

    


    @Override
    public String toString() {
        return "Employee{" + "aem=" + aem + ", surname=" + surname + ", name=" + name + ", fatherOrHusbandName=" + fatherOrHusbandName + ", birthdate=" + birthdate + ", afm=" + afm + ", amka=" + amka + ", address=" + address + ", telephone1=" + telephone1 + ", telephone2=" + telephone2 + ", telephone3=" + telephone3 + ", bankAccount=" + bankAccount + ", bank=" + bank + '}';
    }
    
    
}

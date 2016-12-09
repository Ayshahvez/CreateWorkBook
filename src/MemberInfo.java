import java.util.ArrayList;
import java.util.List;

/**
 * Created by Ayshahvez on 11/10/2016.
 */
public class MemberInfo {
    public MemberInfo(String empID, String fname, String lname, String dob) {
        super();
        this.empID = empID;
        this.fname = fname;
        this.lname = lname;
        this.DOB=dob;

    }

    public MemberInfo(String empID, String fname, String lname, String sex, String reconStatus, String DOB, String DOH, String DOE) {
        this.empID = empID;
        this.fname = fname;
        this.lname = lname;
        this.sex = sex;
        ReconStatus = reconStatus;
        this.DOB = DOB;
        this.DOH = DOH;
        this.DOE = DOE;
    }



    public MemberInfo(){
        super();
    }

    public String getEmpID() {
        return empID;
    }

    public void setEmpID(String empID) {
        this.empID = empID;
    }


    public String getFname() {
        return fname;
    }

    public void setFname(String fname) {
        this.fname = fname;
    }

    public String getLname() {
        return lname;
    }

    public void setLname(String lname) {
        this.lname = lname;
    }



    public String getDOB() {
        return DOB;
    }

    public void setDOB(String DOB) {
        this.DOB = DOB;
    }

    private String empID;
    private String fname;
    private String lname;
 //   private String DOB;

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    public String getReconStatus() {
        return ReconStatus;
    }

    public void setReconStatus(String reconStatus) {
        ReconStatus = reconStatus;
    }

    public String getDOH() {
        return DOH;
    }

    public void setDOH(String DOH) {
        this.DOH = DOH;
    }

    public String getDOE() {
        return DOE;
    }

    public void setDOE(String DOE) {
        this.DOE = DOE;
    }

    private String sex;
    private String ReconStatus;
    private String DOB;
    private String DOH;
    private String DOE;

    public List<String> getPensionableSalary() {
        return PensionableSalary;
    }

    public void setPensionableSalary(List<String> pensionableSalary) {
        PensionableSalary = pensionableSalary;
    }

    private List<String> PensionableSalary = new ArrayList<String>();


}

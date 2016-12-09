import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;

import static java.util.Calendar.DATE;
import static java.util.Calendar.MONTH;
import static java.util.Calendar.YEAR;

/**
 * Created by Ayshahvez on 12/4/2016.
 */
public class TemplateSheets {

    public static Calendar getCalendar(Date date) {
        Calendar cal = Calendar.getInstance(Locale.US);
        cal.setTime(date);
        return cal;
    }

    public static int getDiffYears(Date first, Date last) {
        Calendar a = getCalendar(first);
        Calendar b = getCalendar(last);
        int diff = b.get(YEAR) - a.get(YEAR);
        if (a.get(MONTH) > b.get(MONTH) ||
                (a.get(MONTH) == b.get(MONTH) && a.get(DATE) > b.get(DATE))) {
            diff--;
        }
        return diff;
    }

    public static String getDate(int year, int month, int day){

        return LocalDate.of(year, month, day).plusYears(1).minusDays(1).format(DateTimeFormatter.ofPattern("yyyy.MM.dd"));
    }

    public static void  Create_Template_Active_Sheet(Date StartDate, Date EndDate, String PensionPlanName) throws IOException {
        try {
            LocalDate localStartDate= StartDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
            LocalDate localEndDate = EndDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();



          /*  int EndYear = 2015;
            int EndMonth = 12;
            int EndDay=31;

            int StartYear=2004;
            int StartMonth=01;
            int StartDay= 01;*/

           int EndYear = localEndDate.getYear();
            int EndMonth = localEndDate.getMonthValue();
            int EndDay=localEndDate.getDayOfMonth();

            int StartYear=localStartDate.getYear();
            int StartMonth=localStartDate.getMonthValue();
            int StartDay= localStartDate.getDayOfMonth();

            XSSFWorkbook workbook = new XSSFWorkbook();
            MemberModel model = new MemberModel();


            // TitleModel model = new TitleModel();
            XSSFSheet sheet2 = workbook.createSheet("Actives");
            XSSFRow rowHeading1 = sheet2.createRow(0);
            rowHeading1.createCell(0).setCellValue(PensionPlanName);
            XSSFRow row2 = sheet2.createRow(1);
            row2.createCell(0).setCellValue("ACTUARIAL FUNDING VALUATION AS AT "+EndYear+"."+EndMonth+"."+EndDay);
            XSSFRow row3 = sheet2.createRow(2);
            row3.createCell(0).setCellValue("ACCUMULATION OF ACTIVE MEMBERS' ACCOUNT BALANCES");

            //TESTING
            XSSFRow rowHeading2 = sheet2.createRow(6);
            rowHeading2.createCell(0).setCellValue("Employee's Number");
            rowHeading2.createCell(1).setCellValue("Last Name");
            rowHeading2.createCell(2).setCellValue("First Name");
            rowHeading2.createCell(3).setCellValue("Sex");
            rowHeading2.createCell(4).setCellValue("Recon Status");
            rowHeading2.createCell(5).setCellValue("Date of Birth (DOB)");
            rowHeading2.createCell(6).setCellValue("Date of Hire (DOH)");
            rowHeading2.createCell(7).setCellValue("Date of Enrolment (DOE)");


            List<Integer> PensionableSalary = new ArrayList<Integer>();
            //int years = EndYear-StartYear;

            DateFormat df = new SimpleDateFormat("yyyy.MM.dd");
            // Date startDate = df.parse("2012.01.01");
            Date startDate = df.parse(StartYear+"."+StartMonth+"."+StartDay);
            Date endDate = df.parse(EndYear+"."+EndMonth+"."+EndDay);

            int years= getDiffYears(startDate,endDate);

            for(int h=0;h<=years+1;h++){
                PensionableSalary.add(h,StartYear+h);
            }

            int Contindex=0;
            int leap=0;

            GregorianCalendar cal = new GregorianCalendar();

            for(int t=0;t<=years+1;t++){

                rowHeading2.createCell(8+t).setCellValue("Pensionable Salary Earned " + (PensionableSalary.get(t))+"."+StartMonth+"."+StartDay +" to "
                        + getDate(PensionableSalary.get(t), StartMonth, StartDay));

                // LocalDate.of(StartYear, StartMonth, StartDay).plus(365, ChronoUnit.DAYS);
                Contindex = 8+t;
            }

            //Members Age Col
            rowHeading2.createCell(Contindex++).setCellValue("Members Age as at "  + EndYear + "." + EndMonth + "." + EndDay);
            rowHeading2.createCell(Contindex++).setCellValue("Pensionable Service as at "  + EndYear + "." + EndMonth + "." + EndDay);

            XSSFRow rowHeading3 = sheet2.createRow(5);
            int newIndex = Contindex;

            rowHeading3.createCell(newIndex).setCellValue("Acc'd Cont'ns. Plus Credited Interest up to " + StartYear + "." + StartMonth + "." + StartDay);

            int StartCol=newIndex;
            int LastCol=newIndex+3;
            sheet2.addMergedRegion(new CellRangeAddress(5, 5,StartCol,LastCol));
            newIndex+=4; //jump over 4 cell
            int tmp = newIndex;
            rowHeading3.createCell(newIndex).setCellValue("Contributions During Plan Year " + PensionableSalary.get(0)+"."+StartMonth+"." +StartDay+ " to " + getDate(PensionableSalary.get(0), StartMonth, StartDay));

            for(int h=0;h<PensionableSalary.size()-1;h++){
                tmp+=4;

                rowHeading3.createCell(tmp).setCellValue("Acc'd Cont'ns. Plus Credited Interest up to " + getDate(PensionableSalary.get(h), StartMonth, StartDay));
                StartCol = tmp;
                LastCol = tmp + 3;
                sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol));
                newIndex += 4; //jump over 4 cell
                tmp+=4;

                if(h==PensionableSalary.size()-2)
                {
                    rowHeading3.createCell(tmp).setCellValue("Account Balance as at " + EndYear + "." + EndMonth + "." + EndDay);

                    break;
                }
                rowHeading3.createCell(tmp).setCellValue("Contributions During Plan Year " + PensionableSalary.get(h+1) + "." + StartMonth + "." + StartDay + " to " + getDate(PensionableSalary.get(h+1), StartMonth, StartDay));

                StartCol = tmp;
                LastCol = tmp + 3;
                sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol));
                newIndex += 4; //jump over 4 cell
            }


            for(int k=0;k<(PensionableSalary.size())*2;k++) {
                //Members Age Col
                if(k==(PensionableSalary.size()*2)-1) break;
                rowHeading2.createCell(Contindex++).setCellValue("Employees Basic");
                rowHeading2.createCell(Contindex++).setCellValue("Employees' Optional");
                rowHeading2.createCell(Contindex++).setCellValue("Employers' Required");
                rowHeading2.createCell(Contindex++).setCellValue("Employers' Optional");
                //    rowHeading2.createCell(Contindex++).setCellValue("Fees");
            }

            int r=7;
            for(MemberInfo M : model.findAll()){
                Row row = sheet2.createRow(r);

                //id col
                Cell cellId = row.createCell(0);
                cellId.setCellValue(M.getEmpID());
                //fname col
                Cell cellFName = row.createCell(1);
                cellFName.setCellValue(M.getFname());

                //Lname col
                Cell cellLName = row.createCell(2);
                cellLName.setCellValue(M.getLname());

                //Lname col
                Cell cellSex = row.createCell(3);
                cellSex.setCellValue(M.getSex());

                //Lname col
                Cell cellreconStatus = row.createCell(4);
                cellreconStatus.setCellValue(M.getReconStatus());

                //Lname col
                Cell cellDOB = row.createCell(5);
                cellDOB.setCellValue(M.getDOB());

                //Lname col
                Cell cellDOH = row.createCell(6);
                cellDOH.setCellValue(M.getDOH());

                //Lname col
                Cell cellDOE = row.createCell(7);
                cellDOE.setCellValue(M.getDOE());

                r++;
            }

            //autofit
            for(int x=0;x<Contindex;x++){
                //   sheet.autoSizeColumn(x);
                sheet2.autoSizeColumn(x);

            }


            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(
                    new File("C:\\Users\\akonowalchuk\\GFRAM\\Template_Active_Sheet.xlsx"));
            workbook.write(out);
            out.close();
            workbook.close();
            System.out.println("Template_Active_Sheet.xlsx written successfully" );
        } catch (Exception e) {
            e.printStackTrace();
            //  System.out.print(e.getMessage());
        }
    }

    public static void  Create_Template_Terminee_Sheet(Date StartDate, Date EndDate, String PensionPlanname) throws IOException {
        try {

            LocalDate localStartDate= StartDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
            LocalDate localEndDate = EndDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();


            int EndYear = localEndDate.getYear();
            int EndMonth = localEndDate.getMonthValue();
            int EndDay=localEndDate.getDayOfMonth();

            int StartYear=localStartDate.getYear();
            int StartMonth=localStartDate.getMonthValue();
            int StartDay= localStartDate.getDayOfMonth();

        /*    int EndYear = 2015;
            int EndMonth = 12;
            int EndDay=31;

            int StartYear=2004;
            int StartMonth=01;
            int StartDay= 01;
*/
            XSSFWorkbook workbook = new XSSFWorkbook();
            MemberModel model = new MemberModel();

            // TitleModel model = new TitleModel();
            XSSFSheet sheet2 = workbook.createSheet("Terminees");


            XSSFRow rowHeading1 = sheet2.createRow(0);
            rowHeading1.createCell(0).setCellValue(PensionPlanname);
            XSSFRow row2 = sheet2.createRow(1);
            row2.createCell(0).setCellValue("ACTUARIAL FUNDING VALUATION AS AT "+EndYear+"."+EndMonth+"."+EndDay);
            XSSFRow row3 = sheet2.createRow(2);
            row3.createCell(0).setCellValue("ACCUMULATION OF ACTIVE MEMBERS' ACCOUNT BALANCES");

            //TESTING
            XSSFRow rowHeading2 = sheet2.createRow(6);
            rowHeading2.createCell(0).setCellValue("Employee's Number");
            rowHeading2.createCell(1).setCellValue("Last Name");
            rowHeading2.createCell(2).setCellValue("First Name");
            rowHeading2.createCell(3).setCellValue("Sex");
            rowHeading2.createCell(4).setCellValue("Type of Termination");
            rowHeading2.createCell(5).setCellValue("Date of Birth (DOB)");
            rowHeading2.createCell(6).setCellValue("Date of Hire (DOH)");
            rowHeading2.createCell(7).setCellValue("Date of Enrolment (DOE)");
            rowHeading2.createCell(8).setCellValue("Date of Termination (DOT)");
            rowHeading2.createCell(9).setCellValue("Start of Plan Year of Termination");
            rowHeading2.createCell(10).setCellValue("End of Plan Year of Termination");
            rowHeading2.createCell(11).setCellValue("End of Plan Year of Enrolment");
            rowHeading2.createCell(12).setCellValue("Date of Refund");
            rowHeading2.createCell(13).setCellValue("Period from DOE to DOR");
            rowHeading2.createCell(14).setCellValue("Period from Start of Plan Year of Termination to DOR");
            rowHeading2.createCell(15).setCellValue("Period from DOE to End of Plan Year of Enrolment/DOR");
            rowHeading2.createCell(16).setCellValue("Period from DOR to End of Plan Year of Termination");
            rowHeading2.createCell(17).setCellValue("Pensonable Service up to DOT");
            rowHeading2.createCell(18).setCellValue("Type of Termination(Vested or Non-Vested");
            List<Integer> PensionableSalary = new ArrayList<Integer>();
            //int years = EndYear-StartYear;

            DateFormat df = new SimpleDateFormat("yyyy.MM.dd");
            // Date startDate = df.parse("2012.01.01");
            Date startDate = df.parse(StartYear+"."+StartMonth+"."+StartDay);
            Date endDate = df.parse(EndYear+"."+EndMonth+"."+EndDay);

            int years= getDiffYears(startDate,endDate);

            for(int h=0;h<=years+1;h++){
                PensionableSalary.add(h,StartYear+h);
            }

            int Contindex=rowHeading2.getLastCellNum();


            XSSFRow rowHeading3 = sheet2.createRow(5);
            int newIndex = Contindex;


            rowHeading3.createCell(newIndex).setCellValue("Acc'd Cont'ns. Plus Credited Interest up to " + StartYear + "." + StartMonth + "." + StartDay);

            int StartCol=newIndex;
            int LastCol=newIndex+3;
            sheet2.addMergedRegion(new CellRangeAddress(5, 5,StartCol,LastCol));
            newIndex+=4; //jump over 4 cell
            int tmp = newIndex;
            rowHeading3.createCell(newIndex).setCellValue("Contributions During Plan Year " + PensionableSalary.get(0)+"."+StartMonth+"." +StartDay+ " to " + getDate(PensionableSalary.get(0), StartMonth, StartDay));
            int j=0;
            for(int h=0;h<PensionableSalary.size()-1;h++){
                tmp+=4;


                rowHeading3.createCell(tmp).setCellValue("Acc'd Cont'ns. Plus Credited Interest up to " + getDate(PensionableSalary.get(h), StartMonth, StartDay));
                StartCol = tmp;
                LastCol = tmp + 3;
                sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol));
                newIndex += 4; //jump over 4 cell
                tmp+=4;

                if(h==PensionableSalary.size()-2) //when it gets to end of loop
                {
                    for(int k=0;k<(PensionableSalary.size())*2;k++) {
                        //Members Age Col
                        if(k==(PensionableSalary.size()*2)-1) break;
                        rowHeading2.createCell(Contindex++).setCellValue("Employees Basic");
                        rowHeading2.createCell(Contindex++).setCellValue("Employees' Optional");
                        rowHeading2.createCell(Contindex++).setCellValue("Employers' Required");
                        rowHeading2.createCell(Contindex++).setCellValue("Employers' Optional");
                        //    rowHeading2.createCell(Contindex++).setCellValue("Fees");
                    }
                    rowHeading3.createCell(tmp++).setCellValue("Vested Balances at Date of Refund as Computed by GFRAM");
                    rowHeading3.createCell(tmp+=2).setCellValue("Amounts Refunded to Member Upon Termination, based n Data Submitted");
                    rowHeading3.createCell(tmp+=2).setCellValue("Under/Over Payment at Date of Refund");
                    rowHeading2.createCell(Contindex++).setCellValue("Employees' Basic Contributions plus Credited Interest");
                    rowHeading2.createCell(Contindex++).setCellValue("Employees Optional Plus Credited Interest");
                    rowHeading2.createCell(Contindex++).setCellValue("Employer's Contributions Plus Credited Interest");


                    rowHeading2.createCell(Contindex++).setCellValue("Employees' Basic Contributions plus Credited Interest");
                    rowHeading2.createCell(Contindex++).setCellValue("Employees Optional Plus Credited Interest");

                    rowHeading2.createCell(Contindex++).setCellValue("Employees' Basic Contributions plus Credited Interest");
                    rowHeading2.createCell(Contindex++).setCellValue("Employees Optional Plus Credited Interest");
                    break;
                }
                rowHeading3.createCell(tmp).setCellValue("Contributions During Plan Year " + PensionableSalary.get(h+1) + "." + StartMonth + "." + StartDay + " to " + getDate(PensionableSalary.get(h+1), StartMonth, StartDay));

                StartCol = tmp;
                LastCol = tmp + 3;
                sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol));
                newIndex += 4; //jump over 4 cell

            }
            int r=7;
            for(MemberInfo M : model.findAll()){
                Row row = sheet2.createRow(r);

                //id col
                Cell cellId = row.createCell(0);
                cellId.setCellValue(M.getEmpID());
                //fname col
                Cell cellFName = row.createCell(1);
                cellFName.setCellValue(M.getFname());

                //Lname col
                Cell cellLName = row.createCell(2);
                cellLName.setCellValue(M.getLname());

                //Lname col
                Cell cellSex = row.createCell(3);
                cellSex.setCellValue(M.getSex());

                //Lname col
                Cell cellreconStatus = row.createCell(4);
                cellreconStatus.setCellValue(M.getReconStatus());

                //Lname col
                Cell cellDOB = row.createCell(5);
                cellDOB.setCellValue(M.getDOB());

                //Lname col
                Cell cellDOH = row.createCell(6);
                cellDOH.setCellValue(M.getDOH());

                //Lname col
                Cell cellDOE = row.createCell(7);
                cellDOE.setCellValue(M.getDOE());

                r++;
            }

            //autofit
            for(int x=0;x<Contindex;x++){
                //   sheet.autoSizeColumn(x);
                sheet2.autoSizeColumn(x);

            }


            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(
                    new File("C:\\Users\\akonowalchuk\\GFRAM\\Template_Terminee_Sheet.xlsx"));
            workbook.write(out);
            out.close();
            workbook.close();
            System.out.println("Template_Terminee_Sheet.xlsx written successfully" );
        } catch (Exception e) {
            e.printStackTrace();
            //  System.out.print(e.getMessage());
        }
    }

}

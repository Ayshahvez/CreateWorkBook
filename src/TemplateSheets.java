import com.sun.prism.paint.Color;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import javax.swing.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.NoSuchFileException;
import java.text.DateFormat;
import java.text.ParseException;
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
        if (a.get(MONTH) > b.get(MONTH) || (a.get(MONTH) == b.get(MONTH) && a.get(DATE) > b.get(DATE))) {
            diff--;
        }
        return diff;
    }

    public static String getDate(int year, int month, int day) {

        return LocalDate.of(year, month, day).plusYears(1).minusDays(1).format(DateTimeFormatter.ofPattern("yyyy.MM.dd"));
    }

    public static void Create_Template_Fees_Active_Sheet(String StartDate, String EndDate, String PensionPlanName, String workingDir) throws IOException {
        try {
            String str[] = StartDate.split("/");
            int StartMonth = Integer.parseInt(str[0]);
            int StartDay = Integer.parseInt(str[1]);
            int StartYear = Integer.parseInt(str[2]);

            String str2[] = EndDate.split("/");
            int EndMonth = Integer.parseInt(str2[0]);
            int EndDay = Integer.parseInt(str2[1]);
            int EndYear = Integer.parseInt(str2[2]);

            XSSFWorkbook workbook = new XSSFWorkbook();
            MemberModel model = new MemberModel();


            // TitleModel model = new TitleModel();
            XSSFSheet sheet2 = workbook.createSheet("Actives");
            XSSFRow rowHeading1 = sheet2.createRow(0);
            rowHeading1.createCell(0).setCellValue(PensionPlanName);
            XSSFRow row2 = sheet2.createRow(1);
            row2.createCell(0).setCellValue("ACTUARIAL FUNDING VALUATION AS AT " + EndYear + "." + EndMonth + "." + EndDay);
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
            Date startDate = df.parse(StartYear + "." + StartMonth + "." + StartDay);
            Date endDate = df.parse(EndYear + "." + EndMonth + "." + EndDay);

            int years = getDiffYears(startDate, endDate) + 2;
            System.out.println("years:" + years);
            for (int h = 0; h < years; h++) {
                PensionableSalary.add(h, StartYear + h);
            }

            int Contindex = 0;
            int leap = 0;

            GregorianCalendar cal = new GregorianCalendar();

            for (int t = 0; t < years; t++) {

                rowHeading2.createCell(8 + t).setCellValue("Pensionable Salary Earned " + (PensionableSalary.get(t)) + "." + StartMonth + "." + StartDay + " to " + getDate(PensionableSalary.get(t), StartMonth, StartDay));

                // LocalDate.of(StartYear, StartMonth, StartDay).plus(365, ChronoUnit.DAYS);
                Contindex = 8 + t;
            }

            //Members Age Col
            rowHeading2.createCell(Contindex++).setCellValue("Members Age as at " + EndYear + "." + EndMonth + "." + EndDay);
            rowHeading2.createCell(Contindex++).setCellValue("Pensionable Service as at " + EndYear + "." + EndMonth + "." + EndDay);


            XSSFRow rowHeading3 = sheet2.createRow(5);
            int newIndex = Contindex;

            rowHeading3.createCell(newIndex).setCellValue("Acc'd Cont'ns. Plus Credited Interest up to " + StartYear + "." + StartMonth + "." + StartDay);
//to merge cells
            int StartCol = newIndex;  //start merging from this column
            int LastCol = newIndex + 3;//end merging at this column
            sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol)); //do the merge

            newIndex += 4; //jump over 4 cell to the next employee basic column
            int tmp = newIndex;

            rowHeading3.createCell(newIndex).setCellValue("Contributions During Plan Year " + PensionableSalary.get(0) + "." + StartMonth + "." + StartDay + " to " + getDate(PensionableSalary.get(0), StartMonth, StartDay));
            //  rowHeading3.createCell(tmp++).setCellValue(PensionableSalary.get(0));
            tmp = newIndex;
            StartCol = newIndex;  //start merging from this column
            LastCol = newIndex + 3;//end merging at this column
            sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol)); //do the merge

            for (int h = 0; h < PensionableSalary.size(); h++) {
                tmp += 5;

                rowHeading3.createCell(tmp).setCellValue("Acc'd Cont'ns. Plus Credited Interest up to " + getDate(PensionableSalary.get(h), StartMonth, StartDay));
                StartCol = tmp;
                LastCol = tmp + 3;
                sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol));
                newIndex += 4; //jump over 4 cell
                tmp += 4;

                rowHeading3.createCell(tmp).setCellValue("Contributions During Plan Year " + PensionableSalary.get(h + 1) + "." + StartMonth + "." + StartDay + " to " + getDate(PensionableSalary.get(h + 1), StartMonth, StartDay));
                rowHeading3.createCell(newIndex++).setCellValue(getDate(PensionableSalary.get(h), StartMonth, StartDay));

                if (h == PensionableSalary.size() - 2) {
                    rowHeading3.createCell(tmp).setCellValue("Account Balance as at " + EndYear + "." + EndMonth + "." + EndDay);
                    break;
                }


                StartCol = tmp;
                LastCol = tmp + 3;
                sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol));
                newIndex += 4; //jump over 4 cell
            }


     /*       for (int k = 0; k < (PensionableSalary.size()) * 2; k++) {
                //Members Age Col
                if (k == (PensionableSalary.size() * 2) - 1) break;
                rowHeading2.createCell(Contindex++).setCellValue("Employees Basic");
                rowHeading2.createCell(Contindex++).setCellValue("Employees' Optional");
                rowHeading2.createCell(Contindex++).setCellValue("Employers' Required");
                rowHeading2.createCell(Contindex++).setCellValue("Employers' Optional");
                if(k>0) {

//int tempInd = Contindex +
                    rowHeading2.createCell(Contindex++).setCellValue("Fees");
                }
            }*/

            for (int k = 0; k < PensionableSalary.size(); k++) {
                //Members Age Col
                //      if (k == PensionableSalary.size() - 1) break;
                rowHeading2.createCell(Contindex++).setCellValue("Employees Basic");
                rowHeading2.createCell(Contindex++).setCellValue("Employees' Optional");
                rowHeading2.createCell(Contindex++).setCellValue("Employers' Required");
                rowHeading2.createCell(Contindex++).setCellValue("Employers' Optional");

                if (k == PensionableSalary.size() - 1) break;
                rowHeading2.createCell(Contindex++).setCellValue("Employees Basic");
                rowHeading2.createCell(Contindex++).setCellValue("Employees' Optional");
                rowHeading2.createCell(Contindex++).setCellValue("Employers' Required");
                rowHeading2.createCell(Contindex++).setCellValue("Employers' Optional");
                //    if(k>0) {

//int tempInd = Contindex +
                rowHeading2.createCell(Contindex++).setCellValue("FEES");

            }


            int r = 7;
            for (MemberInfo M : model.findAll()) {
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
            for (int x = 0; x < Contindex; x++) {
                //   sheet.autoSizeColumn(x);
                sheet2.autoSizeColumn(x);

            }

            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File(workingDir + "\\Template_Active_Sheet.xlsx"));
            workbook.write(out);
            out.close();
            workbook.close();
            System.out.println("Template_Active_Sheet.xlsx written successfully");
        } catch (NoSuchFileException e1) {
            JOptionPane.showMessageDialog(null, "Please ensure the Plan Requirements are set, then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void Create_Template_Active_Sheet(String StartDate, String EndDate, String PensionPlanName, String workingDir) throws IOException {
        try {

            String str[] = StartDate.split("/");
            int StartMonth = Integer.parseInt(str[0]);
            int StartDay = Integer.parseInt(str[1]);
            int StartYear = Integer.parseInt(str[2]);

            String str2[] = EndDate.split("/");
            int EndMonth = Integer.parseInt(str2[0]);
            int EndDay = Integer.parseInt(str2[1]);
            int EndYear = Integer.parseInt(str2[2]);

   /*         LocalDate localStartDate= StartDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
            LocalDate localEndDate = EndDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
*/

          /*  int EndYear = 2015;
            int EndMonth = 12;
            int EndDay=31;

            int StartYear=2004;
            int StartMonth=01;
            int StartDay= 01;*/


           /*int EndYear = localEndDate.getYear();
            int EndMonth = localEndDate.getMonthValue();
            int EndDay=localEndDate.getDayOfMonth();

            int StartYear=localStartDate.getYear();
            int StartMonth=localStartDate.getMonthValue();
            int StartDay= localStartDate.getDayOfMonth();*/

            //   System.out.print(StartDay +"/"+ StartMonth +"/" + StartYear);
            //    System.out.print(EndDay +"/"+ EndMonth +"/" + EndYear);
            XSSFWorkbook workbook = new XSSFWorkbook();
            MemberModel model = new MemberModel();


            // TitleModel model = new TitleModel();
            XSSFSheet sheet2 = workbook.createSheet("Actives");
            XSSFRow rowHeading1 = sheet2.createRow(0);
            rowHeading1.createCell(0).setCellValue(PensionPlanName);
            XSSFRow row2 = sheet2.createRow(1);
            row2.createCell(0).setCellValue("ACTUARIAL FUNDING VALUATION AS AT " + EndYear + "." + EndMonth + "." + EndDay);
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
            Date startDate = df.parse(StartYear + "." + StartMonth + "." + StartDay);
            Date endDate = df.parse(EndYear + "." + EndMonth + "." + EndDay);

            int years = getDiffYears(startDate, endDate) + 2;
            System.out.println("years:" + years);
            for (int h = 0; h < years; h++) {
                PensionableSalary.add(h, StartYear + h);
            }

            int Contindex = 0;
            int leap = 0;

            GregorianCalendar cal = new GregorianCalendar();

            for (int t = 0; t < years; t++) {

                rowHeading2.createCell(8 + t).setCellValue("Pensionable Salary Earned " + (PensionableSalary.get(t)) + "." + StartMonth + "." + StartDay + " to " + getDate(PensionableSalary.get(t), StartMonth, StartDay));

                // LocalDate.of(StartYear, StartMonth, StartDay).plus(365, ChronoUnit.DAYS);
                Contindex = 8 + t;
            }

            //Members Age Col
            rowHeading2.createCell(Contindex++).setCellValue("Members Age as at " + EndYear + "." + EndMonth + "." + EndDay);
            rowHeading2.createCell(Contindex++).setCellValue("Pensionable Service as at " + EndYear + "." + EndMonth + "." + EndDay);

            XSSFRow rowHeading3 = sheet2.createRow(5);
            int newIndex = Contindex;

            rowHeading3.createCell(newIndex).setCellValue("Acc'd Cont'ns. Plus Credited Interest up to " + StartYear + "." + StartMonth + "." + StartDay);

            int StartCol = newIndex;
            int LastCol = newIndex + 3;
            sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol));
            newIndex += 4; //jump over 4 cell
            int tmp = newIndex;

            rowHeading3.createCell(newIndex).setCellValue("Contributions During Plan Year " + PensionableSalary.get(0) + "." + StartMonth + "." + StartDay + " to " + getDate(PensionableSalary.get(0), StartMonth, StartDay));
            StartCol = newIndex;
            LastCol = newIndex + 3;
            sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol));
            for (int h = 0; h < PensionableSalary.size() - 1; h++) {
                tmp += 4;

                rowHeading3.createCell(tmp).setCellValue("Acc'd Cont'ns. Plus Credited Interest up to " + getDate(PensionableSalary.get(h), StartMonth, StartDay));
                StartCol = tmp;
                LastCol = tmp + 3;
                sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol));
                newIndex += 4; //jump over 4 cell
                tmp += 4;

                if (h == PensionableSalary.size() - 2) {
                    rowHeading3.createCell(tmp).setCellValue("Account Balance as at " + EndYear + "." + EndMonth + "." + EndDay);

                    break;
                }
                rowHeading3.createCell(tmp).setCellValue("Contributions During Plan Year " + PensionableSalary.get(h + 1) + "." + StartMonth + "." + StartDay + " to " + getDate(PensionableSalary.get(h + 1), StartMonth, StartDay));

                StartCol = tmp;
                LastCol = tmp + 3;
                sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol));
                newIndex += 4; //jump over 4 cell
            }


            for (int k = 0; k < (PensionableSalary.size()) * 2; k++) {
                //Members Age Col
                if (k == (PensionableSalary.size() * 2) - 1) break;
                rowHeading2.createCell(Contindex++).setCellValue("Employees Basic");
                rowHeading2.createCell(Contindex++).setCellValue("Employees' Optional");
                rowHeading2.createCell(Contindex++).setCellValue("Employers' Required");
                rowHeading2.createCell(Contindex++).setCellValue("Employers' Optional");
                //    rowHeading2.createCell(Contindex++).setCellValue("Fees");
            }

            int r = 7;
            for (MemberInfo M : model.findAll()) {
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
            for (int x = 0; x < Contindex; x++) {
                //   sheet.autoSizeColumn(x);
                sheet2.autoSizeColumn(x);

            }

            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File(workingDir + "\\Template_Active_Sheet.xlsx"));
            workbook.write(out);
            out.close();
            workbook.close();
            System.out.println("Template_Active_Sheet.xlsx written successfully");
        } catch (NoSuchFileException e1) {
            JOptionPane.showMessageDialog(null, "Please ensure the Plan Requirements are set, then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void Create_Template_Terminee_Sheet(String StartDate, String EndDate, String PensionPlanname, String workingDir) throws IOException {
        try {
     /*       LocalDate localStartDate= StartDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
            LocalDate localEndDate = EndDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();*/
            String str[] = StartDate.split("/");
            int StartMonth = Integer.parseInt(str[0]);
            int StartDay = Integer.parseInt(str[1]);
            int StartYear = Integer.parseInt(str[2]);

            String str2[] = EndDate.split("/");
            int EndMonth = Integer.parseInt(str2[0]);
            int EndDay = Integer.parseInt(str2[1]);
            int EndYear = Integer.parseInt(str2[2]);

        /*    int EndYear = localEndDate.getYear();
            int EndMonth = localEndDate.getMonthValue();
            int EndDay=localEndDate.getDayOfMonth();

            int StartYear=localStartDate.getYear();
            int StartMonth=localStartDate.getMonthValue();
            int StartDay= localStartDate.getDayOfMonth();*/

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
            row2.createCell(0).setCellValue("ACTUARIAL FUNDING VALUATION AS AT " + EndYear + "." + EndMonth + "." + EndDay);
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
            Date startDate = df.parse(StartYear + "." + StartMonth + "." + StartDay);
            Date endDate = df.parse(EndYear + "." + EndMonth + "." + EndDay);

            int years = getDiffYears(startDate, endDate) + 2;

            for (int h = 0; h < years; h++) {
                PensionableSalary.add(h, StartYear + h);
            }

            int Contindex = rowHeading2.getLastCellNum();


            XSSFRow rowHeading3 = sheet2.createRow(5);
            int newIndex = Contindex;


            rowHeading3.createCell(newIndex).setCellValue("Acc'd Cont'ns. Plus Credited Interest up to " + StartYear + "." + StartMonth + "." + StartDay);

            int StartCol = newIndex;
            int LastCol = newIndex + 3;
            sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol));
            newIndex += 4; //jump over 4 cell
            int tmp = newIndex;

            rowHeading3.createCell(newIndex).setCellValue("Contributions During Plan Year " + PensionableSalary.get(0) + "." + StartMonth + "." + StartDay + " to " + getDate(PensionableSalary.get(0), StartMonth, StartDay));
            StartCol = newIndex;
            LastCol = newIndex + 3;
            sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol));

            for (int h = 0; h < PensionableSalary.size() - 1; h++) {
                tmp += 4;
                rowHeading3.createCell(tmp).setCellValue("Acc'd Cont'ns. Plus Credited Interest up to " + getDate(PensionableSalary.get(h), StartMonth, StartDay));
                StartCol = tmp;
                LastCol = tmp + 3;
                sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol));
                newIndex += 4; //jump over 4 cell
                tmp += 4;

                if (h == PensionableSalary.size() - 2) //when it gets to end of loop
                {
                    for (int k = 0; k < (PensionableSalary.size()) * 2; k++) {
                        //Members Age Col
                        if (k == (PensionableSalary.size() * 2) - 1) break;
                        rowHeading2.createCell(Contindex++).setCellValue("Employees Basic");
                        rowHeading2.createCell(Contindex++).setCellValue("Employees' Optional");
                        rowHeading2.createCell(Contindex++).setCellValue("Employers' Required");
                        rowHeading2.createCell(Contindex++).setCellValue("Employers' Optional");
                        //    rowHeading2.createCell(Contindex++).setCellValue("Fees");
                    }
                    rowHeading3.createCell(tmp++).setCellValue("Vested Balances at Date of Refund as Computed by GFRAM");
                    rowHeading3.createCell(tmp += 2).setCellValue("Amounts Refunded to Member Upon Termination, based n Data Submitted");
                    rowHeading3.createCell(tmp += 2).setCellValue("Under/Over Payment at Date of Refund");
                    rowHeading2.createCell(Contindex++).setCellValue("Employees' Basic Contributions plus Credited Interest");
                    rowHeading2.createCell(Contindex++).setCellValue("Employees Optional Plus Credited Interest");
                    rowHeading2.createCell(Contindex++).setCellValue("Employer's Contributions Plus Credited Interest");


                    rowHeading2.createCell(Contindex++).setCellValue("Employees' Basic Contributions plus Credited Interest");
                    rowHeading2.createCell(Contindex++).setCellValue("Employees Optional Plus Credited Interest");

                    rowHeading2.createCell(Contindex++).setCellValue("Employees' Basic Contributions plus Credited Interest");
                    rowHeading2.createCell(Contindex++).setCellValue("Employees Optional Plus Credited Interest");
                    rowHeading2.createCell(Contindex++).setCellValue("Vesting %");
                    rowHeading2.createCell(Contindex++).setCellValue("Er Non-Vested Bal as at "+EndYear+"."+EndMonth+"."+EndYear);
                    break;
                }
                rowHeading3.createCell(tmp).setCellValue("Contributions During Plan Year " + PensionableSalary.get(h + 1) + "." + StartMonth + "." + StartDay + " to " + getDate(PensionableSalary.get(h + 1), StartMonth, StartDay));

                StartCol = tmp;
                LastCol = tmp + 3;
                sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol));
                newIndex += 4; //jump over 4 cell

            }
            int r = 7;
            for (MemberInfo M : model.findAll()) {
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
            for (int x = 0; x < Contindex; x++) {
                //   sheet.autoSizeColumn(x);
                sheet2.autoSizeColumn(x);

            }


            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(
                    //  new File("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Template_Terminee_Sheet.xlsx"));
                    new File(workingDir + "\\Template_Terminee_Sheet.xlsx"));
            workbook.write(out);
            out.close();
            workbook.close();
            System.out.println("Template_Terminee_Sheet.xlsx written successfully");
        } catch (NoSuchFileException e1) {
            JOptionPane.showMessageDialog(null, "Please ensure the Plan Requirements are set, then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void Create_Template_Fees_Terminee_Sheet(String StartDate, String EndDate, String PensionPlanname, String workingDir) throws IOException {
        try {
     /*       LocalDate localStartDate= StartDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
            LocalDate localEndDate = EndDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();*/
            String str[] = StartDate.split("/");
            int StartMonth = Integer.parseInt(str[0]);
            int StartDay = Integer.parseInt(str[1]);
            int StartYear = Integer.parseInt(str[2]);

            String str2[] = EndDate.split("/");
            int EndMonth = Integer.parseInt(str2[0]);
            int EndDay = Integer.parseInt(str2[1]);
            int EndYear = Integer.parseInt(str2[2]);

        /*    int EndYear = localEndDate.getYear();
            int EndMonth = localEndDate.getMonthValue();
            int EndDay=localEndDate.getDayOfMonth();

            int StartYear=localStartDate.getYear();
            int StartMonth=localStartDate.getMonthValue();
            int StartDay= localStartDate.getDayOfMonth();*/

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
            row2.createCell(0).setCellValue("ACTUARIAL FUNDING VALUATION AS AT " + EndYear + "." + EndMonth + "." + EndDay);
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
            Date startDate = df.parse(StartYear + "." + StartMonth + "." + StartDay);
            Date endDate = df.parse(EndYear + "." + EndMonth + "." + EndDay);

            int years = getDiffYears(startDate, endDate) + 2;

            for (int h = 0; h < years; h++) {
                PensionableSalary.add(h, StartYear + h);
            }

            int Contindex = rowHeading2.getLastCellNum();


            XSSFRow rowHeading3 = sheet2.createRow(5);
            int newIndex = Contindex;


            rowHeading3.createCell(newIndex).setCellValue("Acc'd Cont'ns. Plus Credited Interest up to " + StartYear + "." + StartMonth + "." + StartDay);

            int StartCol = newIndex;
            int LastCol = newIndex + 3;
            sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol));
            newIndex += 4; //jump over 4 cell
            int tmp = newIndex;


            rowHeading3.createCell(newIndex).setCellValue("Contributions During Plan Year " + PensionableSalary.get(0) + "." + StartMonth + "." + StartDay + " to " + getDate(PensionableSalary.get(0), StartMonth, StartDay));
            StartCol = newIndex;
            LastCol = newIndex + 3;
            sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol));


            for (int h = 0; h < PensionableSalary.size(); h++) {
                tmp += 5;
                rowHeading3.createCell(tmp).setCellValue("Acc'd Cont'ns. Plus Credited Interest up to " + getDate(PensionableSalary.get(h), StartMonth, StartDay));
                StartCol = tmp;
                LastCol = tmp + 3;
                sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol));
                newIndex += 4; //jump over 4 cell
                tmp += 4;

                rowHeading3.createCell(tmp).setCellValue("Contributions During Plan Year " + PensionableSalary.get(h + 1) + "." + StartMonth + "." + StartDay + " to " + getDate(PensionableSalary.get(h + 1), StartMonth, StartDay));
                rowHeading3.createCell(newIndex++).setCellValue(getDate(PensionableSalary.get(h), StartMonth, StartDay));

                if (h == PensionableSalary.size() - 2) //when it gets to end of loop
                {
            /*        for (int k = 0; k < (PensionableSalary.size()) * 2; k++) {
                        //Members Age Col
                        if (k == (PensionableSalary.size() * 2) - 1) break;
                        rowHeading2.createCell(Contindex++).setCellValue("Employees Basic");
                        rowHeading2.createCell(Contindex++).setCellValue("Employees' Optional");
                        rowHeading2.createCell(Contindex++).setCellValue("Employers' Required");
                        rowHeading2.createCell(Contindex++).setCellValue("Employers' Optional");
                        //    rowHeading2.createCell(Contindex++).setCellValue("Fees");
                    }*/
                    for (int k = 0; k < PensionableSalary.size(); k++) {
                        //Members Age Col

                        rowHeading2.createCell(Contindex++).setCellValue("Employees Basic");
                        rowHeading2.createCell(Contindex++).setCellValue("Employees' Optional");
                        rowHeading2.createCell(Contindex++).setCellValue("Employers' Required");
                        rowHeading2.createCell(Contindex++).setCellValue("Employers' Optional");

                        if (k == PensionableSalary.size() - 1) break;
                        rowHeading2.createCell(Contindex++).setCellValue("Employees Basic");
                        rowHeading2.createCell(Contindex++).setCellValue("Employees' Optional");
                        rowHeading2.createCell(Contindex++).setCellValue("Employers' Required");
                        rowHeading2.createCell(Contindex++).setCellValue("Employers' Optional");
                        //    if(k>0) {

//int tempInd = Contindex +
                        rowHeading2.createCell(Contindex++).setCellValue("FEES");

                    }

                    rowHeading3.createCell(tmp++).setCellValue("Vested Balances at Date of Refund as Computed by GFRAM");
                    rowHeading3.createCell(tmp += 2).setCellValue("Amounts Refunded to Member Upon Termination, based n Data Submitted");
                    rowHeading3.createCell(tmp += 2).setCellValue("Under/Over Payment at Date of Refund");
                    rowHeading2.createCell(Contindex++).setCellValue("Employees' Basic Contributions plus Credited Interest");
                    rowHeading2.createCell(Contindex++).setCellValue("Employees Optional Plus Credited Interest");
                    rowHeading2.createCell(Contindex++).setCellValue("Employer's Contributions Plus Credited Interest");

                    rowHeading2.createCell(Contindex++).setCellValue("Employees' Basic Contributions plus Credited Interest");
                    rowHeading2.createCell(Contindex++).setCellValue("Employees Optional Plus Credited Interest");

                    rowHeading2.createCell(Contindex++).setCellValue("Employees' Basic Contributions plus Credited Interest");
                    rowHeading2.createCell(Contindex++).setCellValue("Employees Optional Plus Credited Interest");
                    rowHeading2.createCell(Contindex++).setCellValue("Vesting %");
                    rowHeading2.createCell(Contindex++).setCellValue("Er Non-Vested Bal as at "+EndYear+"."+EndMonth+"."+EndYear);
                    break;
                }


                StartCol = tmp;
                LastCol = tmp + 3;
                sheet2.addMergedRegion(new CellRangeAddress(5, 5, StartCol, LastCol));
                newIndex += 4; //jump over 4 cell

            }
            int r = 7;
            for (MemberInfo M : model.findAll()) {
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
            for (int x = 0; x < Contindex; x++) {
                //   sheet.autoSizeColumn(x);
                sheet2.autoSizeColumn(x);

            }


            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(
                    //  new File("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Template_Terminee_Sheet.xlsx"));
                    new File(workingDir + "\\Template_Terminee_Sheet.xlsx"));
            workbook.write(out);
            out.close();
            workbook.close();
            System.out.println("Template_Terminee_Sheet.xlsx written successfully");
        } catch (NoSuchFileException e1) {
            JOptionPane.showMessageDialog(null, "Please ensure the Plan Requirements are set, then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void Create_Template_Active_Terminee_Sheet(String workingDir) {
        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet ActiveSheet = workbook.createSheet("Actives");
        XSSFSheet TermineeSheet = workbook.createSheet("Terminees");

        XSSFRow ActiveHeading = ActiveSheet.createRow(0);
        XSSFRow TermineeHeading = TermineeSheet.createRow(0);

        ActiveHeading.createCell(0).setCellValue("KEY");
        ActiveHeading.createCell(1).setCellValue("CORP");
        ActiveHeading.createCell(2).setCellValue("LAST NAME");
        ActiveHeading.createCell(3).setCellValue("FIRST NAME");
        ActiveHeading.createCell(4).setCellValue("SEX");
        ActiveHeading.createCell(5).setCellValue("BIRTHDATE");
        ActiveHeading.createCell(6).setCellValue("PLAN ENTRY");
        ActiveHeading.createCell(7).setCellValue("EMP.DATE");
        ActiveHeading.createCell(8).setCellValue("STATUS DATE");
        ActiveHeading.createCell(9).setCellValue("STATUS");

        TermineeHeading.createCell(0).setCellValue("KEY");
        TermineeHeading.createCell(1).setCellValue("CORP");
        TermineeHeading.createCell(2).setCellValue("LAST NAME");
        TermineeHeading.createCell(3).setCellValue("FIRST NAME");
        TermineeHeading.createCell(4).setCellValue("SEX");
        TermineeHeading.createCell(5).setCellValue("BIRTHDATE");
        TermineeHeading.createCell(6).setCellValue("PLAN ENTRY");
        TermineeHeading.createCell(7).setCellValue("EMP.DATE");
        TermineeHeading.createCell(8).setCellValue("STATUS DATE");
        TermineeHeading.createCell(9).setCellValue("STATUS");
        TermineeHeading.createCell(10).setCellValue("DATE OF REFUND");

        XSSFCellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontName(XSSFFont.DEFAULT_FONT_NAME);
        font.setFontHeightInPoints((short) 14);
        font.setBold(true);
        //   font.setColor(Color.GREEN);
        //     font.setColor(HSSFColor.GREEN);
        style.setFont(font);

        for (int j = 0; j < 11; j++) {
            if (j < 10) ActiveHeading.getCell(j).setCellStyle(style);

            TermineeHeading.getCell(j).setCellStyle(style);

            //     ActiveHeading.setFillForegroundColor(XSSFColor.GREY_25_PERCENT.index);
            //      csFirstRow.setFillPattern(CellStyle.SOLID_FOREGROUND);
        }


//autofit
        for (int x = 0; x < ActiveSheet.getRow(0).getPhysicalNumberOfCells(); x++) {
            //   sheet.autoSizeColumn(x);
            ActiveSheet.autoSizeColumn(x);

        }

        for (int x = 0; x < TermineeSheet.getRow(0).getPhysicalNumberOfCells(); x++) {
            TermineeSheet.autoSizeColumn(x);
            //    ActiveSheet.autoSizeColumn(x);
        }


        try {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(
                    //  new File("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Template_Terminee_Sheet.xlsx"));
                    new File(workingDir + "\\Template_Separated.xlsx"));
            workbook.write(out);
            out.close();
            workbook.close();
        } catch (NoSuchFileException e1) {
            JOptionPane.showMessageDialog(null, "Please ensure the Plan Requirements are set, then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println("Template_Separated.xlsx Created Sucessfully");
    }

    //create table templates
    public static void Create_Template_Inc_Exp_Sheet(String StartDate, String EndDate, String PensionPlanName, String workingDir) throws IOException {
        try {

            String str[] = StartDate.split("/");
            int StartMonth = Integer.parseInt(str[0]);
            int StartDay = Integer.parseInt(str[1]);
            int StartYear = Integer.parseInt(str[2]);

            String str2[] = EndDate.split("/");
            int EndMonth = Integer.parseInt(str2[0]);
            int EndDay = Integer.parseInt(str2[1]);
            int EndYear = Integer.parseInt(str2[2]);

            SimpleDateFormat df = new SimpleDateFormat("yyyy.MM.dd");
            // Date startDate = df.parse("2012.01.01");
            Date startDate = df.parse(StartYear + "." + StartMonth + "." + StartDay);
            Date endDate = df.parse(EndYear + "." + EndMonth + "." + EndDay);

            int years = getDiffYears(startDate, endDate);
            System.out.println(years);
            years += 1;
            System.out.println(years);


            XSSFWorkbook workbook = new XSSFWorkbook();
            //       MemberModel model = new MemberModel();


            XSSFSheet sheet2 = workbook.createSheet("Income & Expenditure Analysis");

            XSSFRow row1 = sheet2.createRow(0);
            row1.createCell(0).setCellValue("ANALYSIS OF INCOME AND EXPENDITURE OVER THE PERIOD " + StartYear + "." + StartMonth + "." + " " + StartDay + " TO " + EndYear + "." + EndMonth + "." + EndDay);
            sheet2.addMergedRegion(new CellRangeAddress(0, 0, 0, years));

            XSSFRow row2 = sheet2.createRow(1);
            row2.createCell(0).setCellValue("FOR THE ACTUARIAL FUNDING VALUATION AS AT " + EndYear + "." + EndMonth + "." + EndDay + " OF THE " + PensionPlanName);
            sheet2.addMergedRegion(new CellRangeAddress(1, 1, 0, years));

            XSSFRow row3 = sheet2.createRow(3);
            row3.createCell(1).setCellValue("AUDITED FINANCIAL STATEMENTS");
            sheet2.addMergedRegion(new CellRangeAddress(3, 3, 1, years));

            ArrayList<Integer> yearList = new ArrayList<Integer>();

            System.out.println("years:" + years);
            for (int h = 0; h < years; h++) {
                yearList.add(h, StartYear + h);
            }


            XSSFRow yearHeadings = sheet2.createRow(4);
            XSSFRow dollarHeadings = sheet2.createRow(5);
            for (int t = 0; t < years; t++) {


                yearHeadings.createCell(1 + t).setCellValue((yearList.get(t)) + "." + StartMonth + "." + StartDay + " to " + getDate(yearList.get(t), StartMonth, StartDay));
                dollarHeadings.createCell(1 + t).setCellValue("$");

                if (t == years - 1) {
                    yearHeadings.createCell(years + 1).setCellValue("Consolidated " + StartYear + "." + StartMonth + "." + StartDay + " to " + EndYear + "." + EndMonth + "." + EndDay);
                    //    sheet2.addMergedRegion(new CellRangeAddress(3, 3, 1, years));
                    dollarHeadings.createCell(1 + years).setCellValue("$");
                }

            }
            // yearHeadings.createCell(years+1).setCellValue("Consolidated "+StartYear+"." + StartMonth + "." + StartDay + " to "+EndYear+"."+EndMonth+"."+EndDay);

            XSSFRow fieldRow1 = sheet2.createRow(6);
            fieldRow1.createCell(0).setCellValue("FUND AT BEGINNING OF PERIOD");

            XSSFRow fieldRow2 = sheet2.createRow(7);
            fieldRow2.createCell(0).setCellValue("Prior Year Adjustment");

            XSSFRow fieldRow3 = sheet2.createRow(8);
            fieldRow3.createCell(0).setCellValue("INCOME");

            XSSFRow fieldRow4 = sheet2.createRow(9);
            fieldRow4.createCell(0).setCellValue("Employees' Contributions (Basic and Optional)");

            XSSFRow fieldRow5 = sheet2.createRow(10);
            fieldRow5.createCell(0).setCellValue("Employer's Contributions");

            XSSFRow fieldRow6 = sheet2.createRow(11);
            fieldRow6.createCell(0).setCellValue("Interest/Dividend Income");

            XSSFRow fieldRow7 = sheet2.createRow(12);
            fieldRow7.createCell(0).setCellValue("Net Realized Gain/(Loss) on Investments");

            XSSFRow fieldRow8 = sheet2.createRow(13);
            fieldRow8.createCell(0).setCellValue("Net Unrealized Gain/(Loss) on Investments");

            XSSFRow fieldRow9 = sheet2.createRow(14);
            fieldRow9.createCell(0).setCellValue("Total Income");


            XSSFRow fieldRow10 = sheet2.createRow(16);
            fieldRow10.createCell(0).setCellValue("EXPENDITURE");

            XSSFRow fieldRow11 = sheet2.createRow(17);
            fieldRow11.createCell(0).setCellValue("Refunds of Contributions Upon Termination/Death");


            XSSFRow fieldRow12 = sheet2.createRow(18);
            fieldRow12.createCell(0).setCellValue("Employees' Basic");

            XSSFRow fieldRow13 = sheet2.createRow(19);
            fieldRow13.createCell(0).setCellValue("Employees' Optional");

            XSSFRow fieldRow14 = sheet2.createRow(20);
            fieldRow14.createCell(0).setCellValue("Employer's Required");

            XSSFRow fieldRow15 = sheet2.createRow(21);
            fieldRow15.createCell(0).setCellValue("Purchase of Immediate Pensions");

            XSSFRow fieldRow16 = sheet2.createRow(22);
            fieldRow16.createCell(0).setCellValue("Purchase of Deferred  Pensions");

            XSSFRow fieldRow17 = sheet2.createRow(23);
            fieldRow17.createCell(0).setCellValue("Lump Sum to Retirees");

            XSSFRow fieldRow18 = sheet2.createRow(24);
            fieldRow18.createCell(0).setCellValue("Monthly Pensions Paid to Pensioners");

            XSSFRow fieldRow19 = sheet2.createRow(25);
            fieldRow19.createCell(0).setCellValue("Amounts Used to Purchase  Immediate/Deferred Annuities");

            XSSFRow fieldRow20 = sheet2.createRow(26);
            fieldRow20.createCell(0).setCellValue("Transfer ");

            XSSFRow fieldRow21 = sheet2.createRow(27);
            fieldRow21.createCell(0).setCellValue("Administration Fees");

            XSSFRow fieldRow22 = sheet2.createRow(28);
            fieldRow22.createCell(0).setCellValue("Investment Management Fees");

            XSSFRow fieldRow23 = sheet2.createRow(29);
            fieldRow23.createCell(0).setCellValue("Fees for Professional Services (Actuarial, Audit, Legal, etc.)");

            XSSFRow fieldRow24 = sheet2.createRow(30);
            fieldRow24.createCell(0).setCellValue("Other Expenses");

            XSSFRow fieldRow25 = sheet2.createRow(31);
            fieldRow25.createCell(0).setCellValue("Total Expenditure");

            XSSFRow fieldRow26 = sheet2.createRow(32);
            fieldRow26.createCell(0).setCellValue("NET INCOME");

            XSSFRow fieldRow27 = sheet2.createRow(33);
            fieldRow27.createCell(0).setCellValue("FUND AT END OF PERIOD");

            XSSFRow fieldRow28 = sheet2.createRow(34);
            fieldRow28.createCell(0).setCellValue("Administrative and Other Expenses");

            XSSFRow fieldRow29 = sheet2.createRow(35);
            fieldRow29.createCell(0).setCellValue("Investment Expenses");


            XSSFRow fieldRow30 = sheet2.createRow(36);
            fieldRow30.createCell(0).setCellValue("Total Expenses");

            XSSFRow fieldRow31 = sheet2.createRow(37);
            fieldRow31.createCell(0).setCellValue("Total Investment Income");

            XSSFRow fieldRow32 = sheet2.createRow(39);
            fieldRow32.createCell(0).setCellValue("Gross Fund Yield - GFY (% p.a.)");

            XSSFRow fieldRow33 = sheet2.createRow(40);
            fieldRow33.createCell(0).setCellValue("Adjusted Fund Yield - AFY (% p.a.)");

            XSSFRow fieldRow34 = sheet2.createRow(41);
            fieldRow34.createCell(0).setCellValue("Net Fund Yield - NFY (% p.a.)");


            XSSFRow fieldRow35 = sheet2.createRow(43);
            fieldRow35.createCell(0).setCellValue("Plan Year Inflation (% p.a.)");

            XSSFRow fieldRow36 = sheet2.createRow(44);
            fieldRow36.createCell(0).setCellValue("Real Adjusted Fund Yield (% p.a.)");

            //   XSSFRow[] fieldRow = new XSSFRow[10];
            for (int x = 0; x <= yearList.size(); x++) {
                //   sheet.autoSizeColumn(x);
                sheet2.autoSizeColumn(x);

            }

            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File(workingDir + "\\Template_Inc_Exp_Sheet.xlsx"));
            workbook.write(out);
            out.close();
            workbook.close();
            System.out.println("Template_Inc_Exp_Sheet.xlsx written successfully");
        } catch (NoSuchFileException e1) {
            JOptionPane.showMessageDialog(null, "Please ensure the Plan Requirements are set, then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void Create_Template_Balance_Sheet(String StartDate, String EndDate, String PensionPlanName, String workingDir) throws IOException, ParseException {
        try {
            String str2[] = EndDate.split("/");
            int EndMonth = Integer.parseInt(str2[0]);
            int EndDay = Integer.parseInt(str2[1]);
            int EndYear = Integer.parseInt(str2[2]);


            XSSFWorkbook workbook = new XSSFWorkbook();
            //       MemberModel model = new MemberModel();

            XSSFSheet sheet = workbook.createSheet("Val Balance Sheet");

            XSSFRow row = sheet.createRow(0);
            row.createCell(1).setCellValue(PensionPlanName);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 4));

            row = sheet.createRow(1);
            row.createCell(1).setCellValue("FUNDING VALUATION BALANCE SHEET AS AT  " + EndYear + "." + EndMonth + "." + EndDay);
            sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 4));

            row = sheet.createRow(2);
            row.createCell(0).setCellValue("A");
            row.createCell(1).setCellValue("LIABILITY");
            //    sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));


            // row = sheet.createRow(2);
            row.createCell(3).setCellValue("AMOUNT ($)");
            //   sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));

            row = sheet.createRow(4);
            row.createCell(0).setCellValue("1");
            row.createCell(1).setCellValue("BENEFITS TO ACTIVE MEMBERS: ");
            //  sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));


            row = sheet.createRow(5);
            row.createCell(0).setCellValue("(a)");
            row.createCell(1).setCellValue("Related to Employees' Basic Contributions Plus Credited Interest");
            //   sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));

            row = sheet.createRow(6);
            row.createCell(0).setCellValue("(b)");
            row.createCell(1).setCellValue("Related to Employees' Optional Contributions Plus Credited Interest ");

            row = sheet.createRow(7);
            row.createCell(0).setCellValue("(c)");
            row.createCell(1).setCellValue("Related to Employer's Contributions Paid into the Fund Plus Credited ");

            row = sheet.createRow(8);
            row.createCell(0).setCellValue("(d)");
            row.createCell(1).setCellValue("Sub-Total (Accrued Benefits to Actives) = Sum of 1(a) to 1(c) ");

            row = sheet.createRow(10);
            row.createCell(0).setCellValue("2");
            row.createCell(1).setCellValue("REFUNDS OUTSTANDING TO UNCLAIMED MEMBERS ");

            row = sheet.createRow(11);
            row.createCell(0).setCellValue("3");
            row.createCell(1).setCellValue("REFUNDS OUTSTANDING TO TERMINATED MEMBERS");

            row = sheet.createRow(12);
            row.createCell(0).setCellValue("4");
            row.createCell(1).setCellValue("ACCRUED BENEFITS TO RETIRED MEMBERS");


            row = sheet.createRow(13);
            row.createCell(0).setCellValue("5");
            row.createCell(1).setCellValue("ACCRUED BENEFITS TO DEFERRED VESTED PENSIONERS");


            row = sheet.createRow(15);
            row.createCell(0).setCellValue("6");
            row.createCell(1).setCellValue("TOTAL ACTUARIAL LIABILITY [1(d) + 2 + 3 + 4 + 5]");

            row = sheet.createRow(17);
            row.createCell(0).setCellValue("B");
            row.createCell(1).setCellValue("MARKET VALUE OF NET ASSETS ");

            row = sheet.createRow(19);
            row.createCell(0).setCellValue("C");
            row.createCell(1).setCellValue("ACTUARIAL SURPLUS/(DEFICIT) = B - A(6)");

            row = sheet.createRow(21);
            row.createCell(0).setCellValue("D");
            row.createCell(1).setCellValue("SOLVENCY LEVEL = [ B/A(6) ] *100");
            for (int x = 0; x < 4; x++) {
                //   sheet.autoSizeColumn(x);
                sheet.autoSizeColumn(x);

            }

            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File(workingDir + "\\Template_Balance_Sheet.xlsx"));
            workbook.write(out);
            out.close();
            workbook.close();
            System.out.println("Template_Inc_Exp_Sheet.xlsx written successfully");
        } catch (NoSuchFileException e1) {
            JOptionPane.showMessageDialog(null, "Please ensure the Plan Requirements are set, then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void Create_Template_Summary_of_Active_Memberships(String StartDate, String EndDate, String PensionPlanName, String workingDir) throws IOException {
        try {
            String str2[] = EndDate.split("/");
            int EndMonth = Integer.parseInt(str2[0]);
            int EndDay = Integer.parseInt(str2[1]);
            int EndYear = Integer.parseInt(str2[2]);

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("Template Summary of Active Membership");

            XSSFRow row = sheet.createRow(0);
            row.createCell(0).setCellValue(PensionPlanName);
          //  sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 4));

            row = sheet.createRow(1);
            row.createCell(0).setCellValue("Table 3.1 Summary of Active Membership Statistics as at "+EndYear+"."+EndMonth+"."+EndDay);
        //    sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 4));

            row = sheet.createRow(2);
            row.createCell(0).setCellValue("Statistic");


            row = sheet.getRow(2);
            row.createCell(1).setCellValue("Males");


            row = sheet.getRow(2);
            row.createCell(2).setCellValue("Females");

            row = sheet.getRow(2);
            row.createCell(3).setCellValue("Both Sexes");

            row = sheet.createRow(3);
            row.createCell(0).setCellValue("Head Count (#) ");

            row = sheet.createRow(4);
            row.createCell(0).setCellValue("Average Age (Years)");

            row = sheet.createRow(5);
            row.createCell(0).setCellValue("Average Pensionable Service (Years)");

            row = sheet.createRow(6);
            row.createCell(0).setCellValue("Avg. Pensionable Salary Earned in the Plan Year Ended "+EndYear+"."+EndMonth+"."+EndDay+" ($p.a.) ");

            for (int x = 0; x <3; x++) {
                //   sheet.autoSizeColumn(x);
                sheet.autoSizeColumn(x);

            }



            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File(workingDir + "\\Template_Summary_of_Active_Membership.xlsx"));
            workbook.write(out);
            out.close();
            workbook.close();
            System.out.println("Template_Summary_of_Active_Membership.xlsx written successfully");
        } catch (NoSuchFileException e1) {
            JOptionPane.showMessageDialog(null, "Please ensure the Plan Requirements are set, then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    public static void Create_Template_Movement_in_Active_Memberships(String StartDate, String EndDate, String PensionPlanName, String workingDir) throws IOException{
   try{
        String str[] = StartDate.split("/");
        int StartMonth = Integer.parseInt(str[0]);
        int StartDay = Integer.parseInt(str[1]);
        int StartYear = Integer.parseInt(str[2]);

        String str2[] = EndDate.split("/");
        int EndMonth = Integer.parseInt(str2[0]);
        int EndDay = Integer.parseInt(str2[1]);
        int EndYear = Integer.parseInt(str2[2]);

        SimpleDateFormat df = new SimpleDateFormat("yyyy.MM.dd");
        // Date startDate = df.parse("2012.01.01");
        Date startDate = df.parse(StartYear + "." + StartMonth + "." + StartDay);
        Date endDate = df.parse(EndYear + "." + EndMonth + "." + EndDay);

        int years = getDiffYears(startDate, endDate);
        years+=1;

       XSSFWorkbook workbook = new XSSFWorkbook();
       XSSFSheet sheet = workbook.createSheet("Template Movement in Active Membership");

       XSSFRow row = sheet.createRow(0);
       row.createCell(0).setCellValue(PensionPlanName);

       row = sheet.createRow(1);
       row.createCell(0).setCellValue("Table 3.2 Movements in Active Membership from "+StartYear+"."+StartMonth+"."+StartDay+" to "+EndYear+"."+EndMonth+"."+EndDay);

       row = sheet.createRow(2);
       row.createCell(1).setCellValue("Males");
       row.createCell(2).setCellValue("Females");
       row.createCell(3).setCellValue("Both Sexes");

       row = sheet.createRow(3);
       row.createCell(0).setCellValue("Active Members as at "+StartYear+"."+StartMonth+"."+StartDay);

       row = sheet.createRow(4);
       row.createCell(0).setCellValue("Plus: New Entrants");

       row = sheet.createRow(6);
       row.createCell(0).setCellValue("Less: ");

       row = sheet.createRow(7);
       row.createCell(0).setCellValue("Terminations");

       row = sheet.createRow(8);
       row.createCell(0).setCellValue("Retirements");

       row = sheet.createRow(9);
       row.createCell(0).setCellValue("Deaths");

       row = sheet.createRow(10);
       row.createCell(0).setCellValue("Unclaimed");

       row = sheet.createRow(12);
       row.createCell(0).setCellValue("Active Members as at "+EndYear+"."+EndMonth+"."+EndDay);

       for (int x = 0; x <4; x++) {
            //   sheet.autoSizeColumn(x);
            sheet.autoSizeColumn(x);
        }

        //Write the workbook in file system
        FileOutputStream out = new FileOutputStream(new File(workingDir + "\\Template_Movements_in_Active_Membership.xlsx"));
        workbook.write(out);
        out.close();
        workbook.close();
        System.out.println("Template_Movements_in_Active_Membership.xlsx written successfully");
    } catch (NoSuchFileException e1) {
        JOptionPane.showMessageDialog(null, "Please ensure the Plan Requirements are set, then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
    } catch (Exception e) {
        e.printStackTrace();
    }
    }

}
import com.sun.org.apache.regexp.internal.RE;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Locale;

import static java.util.Calendar.DATE;
import static java.util.Calendar.MONTH;
import static java.util.Calendar.YEAR;

/**
 * Created by Ayshahvez on 12/6/2016.
 */
public class ValidationChecks {

    public String getResult() {
        return result;
    }

    public void setResult(String result) {
        this.result = result;
    }

    public String result;

    public ArrayList Check_For_Duplicates(String workingDir) throws IOException {

        ArrayList<String> list = new ArrayList<>();
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("Notice: The Duplicate Check Validates the list of Members for Repeated Records");
        boolean once = true;

        // FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Valuation Data.xlsx");
        //   FileInputStream fileInputStream = new FileInputStream(filePathValData);
        FileInputStream fileInputStream = new FileInputStream(workingDir + "\\HAS Input Sheet.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet DemoSheet = workbook.getSheet("Actives at End of Plan Yr 2015");

        //  int NoMembers = DemoSheet.getPhysicalNumberOfRows();
        int numOfMembers = Utility.getNumberOfMembersInSheet(DemoSheet);

        //  numOfMembers +=1;

        SimpleDateFormat dF = new SimpleDateFormat("dd-MMM-yy");

        int check = 0;
        int FindIt = 0;

        for (int readFromRow = 7, row = 0; row < numOfMembers; readFromRow++, row++) {

            XSSFRow DemoRow = DemoSheet.getRow(readFromRow);

            XSSFCell cellA1_EM = DemoRow.getCell((short) 0);  //employee number
            if (cellA1_EM == null) {
                cellA1_EM = DemoRow.createCell(0);
            }
            String result = cellA1_EM.getStringCellValue();
            String a1Val_EM = result.replaceAll("[-]", "");

            XSSFCell cellC1_LN = DemoRow.getCell((short) 1);   //last name
            String b1Val_LN = cellC1_LN.getStringCellValue();

            XSSFCell cellD1_FN = DemoRow.getCell((short) 2);   //first name
            String c1Val_FN = cellD1_FN.getStringCellValue();

            XSSFCell cellF1_DOB = DemoRow.getCell((short) 5);   //DOB
            Date f1Val_DOB = cellF1_DOB.getDateCellValue();

            FindIt = 0;
            check = 0;
            boolean findIt = false;

            for (int readFromRow2 = 7, row2 = 0; row2 < numOfMembers; readFromRow2++, row2++) {

                XSSFRow DemoRow2 = DemoSheet.getRow(readFromRow2);


                XSSFCell cellA1_EM2 = DemoRow2.getCell(0);  //employee number
                if (cellA1_EM2 == null) {
                    cellA1_EM2 = DemoRow2.createCell(0);
                    cellA1_EM.setCellValue("");
                }
                String result2 = cellA1_EM2.getStringCellValue();
                String a1Val_EM2 = result2.replaceAll("[-]", "");

                XSSFCell cellC1_LN2 = DemoRow2.getCell((short) 1);   //last name
                String b1Val_LN2 = cellC1_LN2.getStringCellValue();

                XSSFCell cellD1_FN2 = DemoRow2.getCell((short) 2);   //first name
                String c1Val_FN2 = cellD1_FN2.getStringCellValue();

                XSSFCell cell_DOB = DemoRow2.getCell((short) 5);   //DOB
                Date f1_DOB = cell_DOB.getDateCellValue();

                XSSFCell cellG1 = DemoRow2.getCell((short) 7); //plan entry
                // String g1Val =datetemp.format( cellG1.getDateCellValue());
                Date g1Val = cellG1.getDateCellValue();

                XSSFCell cellH1 = DemoRow2.getCell((short) 6); //emp date
                //    String h1Val = datetemp.format( cellH1.getDateCellValue());
                Date h1Val = cellH1.getDateCellValue();

                //    XSSFCell cellI1 = DemoRow2.getCell((short) 8);  //status date
                // String i1Val = datetemp.format( cellI1.getDateCellValue());
                //   Date i1Val = cellI1.getDateCellValue();

                //   XSSFCell cellJ1 = DemoRow2.getCell((short) 9); //status
                //  String j1Val = cellJ1.getStringCellValue();


                if (a1Val_EM.equals(a1Val_EM2) && b1Val_LN.equals(b1Val_LN2) && c1Val_FN.equals(c1Val_FN2) || c1Val_FN.equals(c1Val_FN2) && b1Val_LN.equals(b1Val_LN2) && f1Val_DOB.equals(f1_DOB) || a1Val_EM.equals(a1Val_EM2)) {
                    FindIt++;
                    //  findIt=true;


                    if (once == true && FindIt > 1) {
                        stringBuilder.append("\n\nThe Process has found these members to be as repeated records\n\n");
                        once = false;
                    }


                    if (FindIt > 1) {
                        check++;
                        //   System.out.println("Please Contact Administrator");
                        System.out.println("Employee ID: " + a1Val_EM);
                        System.out.println("Last Name: " + b1Val_LN);
                        System.out.println("First Name: " + c1Val_FN);
                        System.out.println("DOB: " + dF.format(f1Val_DOB));

                        System.out.println();
                        FindIt = 0;
                        //  break;
                        //    stringBuilder.append("This Member is Duplicate Record, Please Contact Administrator\n"+"\n");
                        //  stringBuilder.append("Record: " + row + "\n");
                        stringBuilder.append("Employee ID: " + a1Val_EM + "\n");
                        stringBuilder.append("Last Name: " + b1Val_LN + "\n");
                        stringBuilder.append("First Name: " + c1Val_FN + "\n");
                        stringBuilder.append("Date of Birth: " + dF.format(f1Val_DOB) + "\n");

                        stringBuilder.append("------------------------------------------\n");
                        //   list.add(a1Val_EM+","+b1Val_LN+","+c1Val_FN+","+dF.format(f1Val_DOB));
                        list.add(a1Val_EM + "," + b1Val_LN + "," + c1Val_FN + "," + dF.format(f1Val_DOB) + "," + dF.format(h1Val) + "," + dF.format(g1Val));
                    }
                }


            }


        }
        System.out.println("Notice: The Duplicate check process has now been completed");
        if (check == 0 && once)
            stringBuilder.append("\n\nNotice: There were no Duplicate records found in this list of Members");
        stringBuilder.append("\n\nNotice: The Duplicate check process has now been completed");
        //      return String.valueOf(stringBuilder + "\n");
        this.setResult(String.valueOf(stringBuilder));
        return list;
    }

    public ArrayList Check_FivePercent_PS(String PensionPlanStartDate, String PensionPlanEndDate, String workindDir) throws IOException {
        ArrayList<String> list = new ArrayList<>();
        int Check = 0;
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("Notice: Results of the Pensionable Salary Check Process:\n");

        DecimalFormat dF = new DecimalFormat("#.####");//#.##
        SimpleDateFormat dateF = new SimpleDateFormat("dd-MMM-yy");

        //OPEN ACTIVE SHEET
        FileInputStream fileInputStream = new FileInputStream(workindDir + "\\Valuation Data.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet DemoSheet = workbook.getSheet("DEMO");

        FileInputStream fs = null;
        try {
            fs = new FileInputStream(workindDir + "\\Pensionable Salary - HAS.xlsx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        XSSFWorkbook WB = new XSSFWorkbook(fs);

        String SD[] = PensionPlanStartDate.split("/");
        int startMonth = Integer.parseInt(SD[0]);
        int startDay = Integer.parseInt(SD[1]);
        int startYear = Integer.parseInt(SD[2]);

        String ED[] = PensionPlanEndDate.split("/");
        int endMonth = Integer.parseInt(ED[0]);
        int endDay = Integer.parseInt(ED[1]);
        int endYear = Integer.parseInt(ED[2]);

        int EndYear = endYear;
        int EndMonth = endMonth;
        int EndDay = endDay;

        int StartYear = startYear;
        int StartMonth = startMonth;
        int StartDay = startDay;

        DateFormat df = new SimpleDateFormat("yyyy.MM.dd");
        Date startDate = null;
        Date endDate = null;
        try {
            startDate = df.parse(StartYear + "." + StartMonth + "." + StartDay);
            endDate = df.parse(EndYear + "." + EndMonth + "." + EndDay);
        } catch (ParseException e) {
            e.printStackTrace();
        }

        int years = Utility.getDiffYears(startDate, endDate);
        years += 1;
        FileInputStream fsRecon = null;
        try {
            fsRecon = new FileInputStream(workindDir + "\\Valuation Data.xlsx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        XSSFWorkbook WBRecon = new XSSFWorkbook(fsRecon);
        int counter = 1;


        for (int x = 0; x < years; x++) {

            Calendar cal = Calendar.getInstance();
            cal.set(StartYear, StartMonth, StartDay);
            SimpleDateFormat sdf = new SimpleDateFormat("yy"); // Just the year, with 2 digits
            String formattedDate = sdf.format(cal.getTime());


            String Recon = ("Recon " + formattedDate);


//GET RECON SHEET
            XSSFSheet Reconsheet = WBRecon.getSheet(Recon);
            System.out.println(formattedDate);
            //   int CountReconRow = Reconsheet.getPhysicalNumberOfRows();


            XSSFSheet PensionabeSalarySheet = WB.getSheet("Sheet1");

            int CountPSRow = PensionabeSalarySheet.getPhysicalNumberOfRows();
            int CountReconRow = Reconsheet.getPhysicalNumberOfRows();
            int numOfDemoMembers = DemoSheet.getPhysicalNumberOfRows(); // get the last row number

            XSSFRow[] rowR = new XSSFRow[numOfDemoMembers];
            Cell cellR = null;

            //      System.out.println("Number in Recon Sheet: " + CountReconRow);
            //   System.out.println("Number in Demo sheet: " + numOfDemoMembers);
            //  System.out.println("Number in PS: " + CountPSRow);

            for (int row = 0; row < numOfDemoMembers; row++) {

                int temp = row;

                if (temp == 0) {
                    row = 1;  //start reading from second row
                }

                int Row = row;

                XSSFRow DemoRow = DemoSheet.getRow(row);


                XSSFCell cellA1_EM = DemoRow.getCell((short) 0);  //employee number
                if (cellA1_EM == null) {
                    cellA1_EM = DemoRow.createCell(0);
                }
                String result = cellA1_EM.getStringCellValue();
                String a1Val_EM = result.replaceAll("[-]", "");

                XSSFCell cellC1_LN = DemoRow.getCell((short) 2);   //last name
                String b1Val_LN = cellC1_LN.getStringCellValue();

                XSSFCell cellD1_FN = DemoRow.getCell((short) 3);   //first name
                String c1Val_FN = cellD1_FN.getStringCellValue();

                //  XSSFCell cellD1 = DemoRow.getCell((short) 3);
                // String d1Val = cellD1.getStringCellValue();

                XSSFCell cellF1_DOB = DemoRow.getCell((short) 5);   //DOB
                Date f1Val_DOB = cellF1_DOB.getDateCellValue();


                XSSFCell cellG1 = DemoRow.getCell((short) 6); //plan entry
                // String g1Val =datetemp.format( cellG1.getDateCellValue());
                Date g1Val = cellG1.getDateCellValue();

                XSSFCell cellH1 = DemoRow.getCell((short) 7); //emp date
                //    String h1Val = datetemp.format( cellH1.getDateCellValue());
                Date h1Val = cellH1.getDateCellValue();

                XSSFCell cellstatusDate = DemoRow.getCell((short) 8);  //status date
                // String i1Val = datetemp.format( cellI1.getDateCellValue());
                Date i1Val = cellstatusDate.getDateCellValue();

                XSSFCell cellJ1 = DemoRow.getCell((short) 9); //status
                String j1Val = cellJ1.getStringCellValue();


                rowR[Row] = DemoSheet.getRow(counter++);

                //LOOP THROUGH PENSIONABLE SALARY SHEET
                for (int y = 0; y < CountPSRow; y++) {
                    int temp2 = y;

                    if (temp2 == 0) {
                        y = 1;
                    }

                    XSSFRow PSRow = PensionabeSalarySheet.getRow(y);


                    XSSFCell A1_EM = PSRow.getCell(0);  //employee number
                    if (A1_EM == null) {
                        A1_EM = PSRow.createCell(0);
                    }
                    String a1_EM = A1_EM.getStringCellValue();


                    XSSFCell B1_LN = PSRow.getCell(1);  //LNAME
                    if (B1_LN == null) {
                        B1_LN = PSRow.createCell(1);

                    }
                    String b1 = B1_LN.getStringCellValue();

                    XSSFCell C1_FN = PSRow.getCell(2);  //LNAME
                    if (C1_FN == null) {
                        C1_FN = PSRow.createCell(2);
                    }
                    String c1_FN = C1_FN.getStringCellValue();


                    XSSFCell F1_Year = PSRow.getCell((short) 5);
                    //Update the value of cell
                    if (F1_Year == null) {
                        F1_Year = PSRow.createCell(5);
                    }
                    double F1 = F1_Year.getNumericCellValue();

                    double F1_Yr = 0;
                    XSSFCell[] cellI1 = new XSSFCell[years];
                    double[] d = new double[years];


                    if (a1Val_EM.equals(a1_EM) && !a1Val_EM.equals("ASSE88888")) {

                        for (int g = 5, j = 0; g < 17; g++, j++) {
                            cellI1[j] = PSRow.getCell(g);
                            F1_Yr = cellI1[j].getNumericCellValue();
                            //  d[j] = Double.parseDouble(F1_Yr);
                            d[j] = F1_Yr;
                        }

                        //CHECK IN RECON SHEET FOR EMPLOYEE CONTRIBUTION
                        for (int v = 6; v <= CountReconRow; v++) {
                            XSSFRow reconRow = Reconsheet.getRow(v); //RECEN

                            XSSFCell A1 = reconRow.getCell(0);  //employee number
                            if (A1 == null) {
                                A1 = reconRow.createCell(0);
                            }
                            String a1 = A1.getStringCellValue();

                            XSSFCell B1 = reconRow.getCell(1);  //FNAME
                            if (B1 == null) {
                                B1 = reconRow.createCell(1);

                            }
                            String Cb1 = B1.getStringCellValue();

                            XSSFCell C1 = reconRow.getCell(2);  //LNAME
                            if (C1 == null) {
                                C1 = reconRow.createCell(2);

                            }
                            String c1 = C1.getStringCellValue();


                            //RECON SHEET
                            XSSFCell H1 = reconRow.getCell((short) 7);
                            //Update the value of cell
                            if (H1 == null) {
                                H1 = reconRow.createCell(7);
                            }
                            double h1 = H1.getNumericCellValue();

                            if (a1.equals(a1Val_EM)) {
                                double check = 0.05 * d[x];
                                check = Double.parseDouble(dF.format(check));
                                if (check != h1) {
                                    Check++;

                                    String test = "";

                                    stringBuilder.append("\nEmployee ID: " + a1 + "\n");
                                    stringBuilder.append("Last Name: " + c1 + "\n");
                                    stringBuilder.append("First Name: " + Cb1 + "\n");
                                    stringBuilder.append("Pensionable Salary: $" + dF.format(d[x]) + "\n");
                                    stringBuilder.append("Employee Basic: $" + h1 + "\n");

                                    list.add(a1Val_EM + "," + c1 + "," + Cb1 + "," + dateF.format(f1Val_DOB) + "," + dateF.format(h1Val) + "," + dateF.format(g1Val) + "," + dateF.format(i1Val) + ',' + j1Val);
                                    if (check == h1) {
                                        test = "true";
                                        System.out.println("TEST: Pensionable Salary: $" + dF.format(d[x]));
                                        System.out.println("Result: 5% of " + dF.format(d[x] + " is " + check));
                                        stringBuilder.append("Result: Contribution is " + dF.format((h1 / d[x]) * 100) + "% of " + dF.format(d[x]) + "\n");
                                        System.out.println("Decision: " + test);

                                    } else {
                                        test = "false";

                                        System.out.println("Result: 5% of " + dF.format(d[x]) + " is not " + h1 + "\n");
                                        stringBuilder.append("Result: " + Cb1 + " Contribution of: $" + h1 + " is " + dF.format(h1 / d[x]) + "% of $" + dF.format(d[x]) + "\n");
                                        stringBuilder.append("---------------------------------------------------------------------------------------------------------------------------------------------------\n");
                                        System.out.println("Decision: Please contact administrator");
                                    }

                                    System.out.println();


                                }

                            }

                        }


                    }


                }

            }

            //   WriteAt+=8;
            StartYear++; //comment out years if COMMENTED
            counter = 1;

        }// END OF LOOP YEARS
        //   stringBuilder.append("-------------------------------------------------------\n");
        if (Check == 0)
            stringBuilder.append("\nNotice: There were no Discrepancies found among the Members' Pensionable Salary and their Contributions throughout the Review Period");
        stringBuilder.append("\n\nNotice: The Pensionable Check Process has now been completed.\n");
        //  return String.valueOf(stringBuilder);

        this.setResult(String.valueOf(stringBuilder));
        return list;
    }

    public ArrayList Check_Plan_EntryDate_empDATE(String workindDir) throws IOException {
        ArrayList<String> list = new ArrayList<>();

        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("Notice: The Plan Entry Check Validates whether a Member has enrolled in the Plan before Employment Date \n");

        FileInputStream fileInputStream = new FileInputStream(workindDir + "\\Valuation Data.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet DemoSheet = workbook.getSheet("DEMO");

        int check = 0;
        int NoMembers = DemoSheet.getPhysicalNumberOfRows();

        for (int row = 0; row < NoMembers; row++) {
            int temp = row;

            if (temp == 0) {
                row = 1;  //start reading from second row
            }


            XSSFRow DemoRow = DemoSheet.getRow(row);


            XSSFCell cellA1_EM = DemoRow.getCell((short) 0);  //employee number
            if (cellA1_EM == null) {
                cellA1_EM = DemoRow.createCell(0);
            }
            String result = cellA1_EM.getStringCellValue();
            String cellEmployeeNumber = result.replaceAll("[-]", "");

            XSSFCell cellC1_LN = DemoRow.getCell((short) 2);   //last name
            String cellLastName = cellC1_LN.getStringCellValue();

            XSSFCell cellD1_FN = DemoRow.getCell((short) 3);   //first name
            String cellFirstName = cellD1_FN.getStringCellValue();

            XSSFCell cellF1_DOB = DemoRow.getCell((short) 5);   //DOB
            Date cellDOB = cellF1_DOB.getDateCellValue();

            XSSFCell cellG1_PE = DemoRow.getCell((short) 6);   //plan entry
            Date cellPlanEntry = cellG1_PE.getDateCellValue();

            XSSFCell cellH1_PE = DemoRow.getCell((short) 7);   //emp date
            Date cellEmploymentDate = cellH1_PE.getDateCellValue();

            XSSFCell cellI1 = DemoRow.getCell((short) 8);  //status date
            // String i1Val = datetemp.format( cellI1.getDateCellValue());
            Date cellStatusDate = cellI1.getDateCellValue();

            XSSFCell cellJ1 = DemoRow.getCell((short) 9); //status
            String cellStatus = cellJ1.getStringCellValue();


            SimpleDateFormat dF = new SimpleDateFormat();
            dF.applyPattern("dd-MMM-yy");

            if (cellPlanEntry.before(cellEmploymentDate)) {
/*                System.out.println("Employee ID: " + cellEmployeeNumber);
                System.out.println("Last Name: " + cellLastName);
                System.out.println("First Name: " + cellFirstName);
                System.out.println("Plan Entry: " + dF.format(cellPlanEntry));
                System.out.println("Employment Date: " + dF.format(cellEmploymentDate));
                System.out.println("Please Contact Administrator");*/

                stringBuilder.append("\nResult: This member's Plan Entry Date at " + dF.format(cellPlanEntry) + " is before their Employment Date at " + dF.format(cellEmploymentDate) + "\n");
                stringBuilder.append("Employee ID: " + cellEmployeeNumber + "\n");
                stringBuilder.append("Last Name: " + cellLastName + "\n");
                stringBuilder.append("First Name: " + cellFirstName + "\n");
                stringBuilder.append("Plan Entry: " + dF.format(cellPlanEntry) + "\n");
                stringBuilder.append("Employment Date: " + dF.format(cellEmploymentDate) + "\n");
                stringBuilder.append("----------------------------------------------------------------------------------------------------------\n");
                list.add(cellEmployeeNumber + "," + cellLastName + "," + cellFirstName + "," + dF.format(cellDOB) + "," + dF.format(cellEmploymentDate) + "," + dF.format(cellPlanEntry) + "," + dF.format(cellStatusDate) + ',' + cellStatus);
                System.out.println();
                check++;
            }


        }
        System.out.println("Notice: The Plan Entry Check Process has now been completed.");
        if (check == 0)
            stringBuilder.append("\nNotice: There were no Discrepancies found with the Members' Emp. Date in respect to their Plane Entry Date\n");
        stringBuilder.append("\nNotice: The Plan Entry Check Process has now been completed.\n");

        this.setResult(String.valueOf(stringBuilder));
        return list;
    }

    public ArrayList Check_DateofBirth(String workindDir) throws IOException {
        SimpleDateFormat dF = new SimpleDateFormat();
        dF.applyPattern("dd-MMM-yy");
        ArrayList<String> list = new ArrayList<>();
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("Notice: The Date of Birth Check Validates the Members' were born before their Employment Date and Plan Entry Date\n");
        FileInputStream fileInputStream = new FileInputStream(workindDir + "\\HAS Input Sheet.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet DemoSheet = workbook.getSheet("Actives at End of Plan Yr 2014");

        //   int NoMembers = DemoSheet.getPhysicalNumberOfRows();
        int NoMembers = DemoSheet.getLastRowNum();
        //   NoMembers+=1;
        int check = 0;
        int readFromRow = 8;
        try {
            XSSFRow DemoRow;
            System.out.println("NoMembers->" + NoMembers);
            for (int row = 7; row < 31; row++) {

                //  int temp = row;

                //     if (temp == 0) {
                //     row = 1;  //start reading from second row
                //    }
                DemoRow = DemoSheet.getRow(row);


                XSSFCell cellA1_EM = DemoRow.getCell((short) 0);  //employee number
                if (cellA1_EM == null) {
                    cellA1_EM = DemoRow.createCell(0);
                }
                String result = cellA1_EM.getStringCellValue();
                String cellEmployeeNumber = result.replaceAll("[-]", "");


                XSSFCell cellC1_LN = DemoRow.getCell((short) 1);   //last name
                String cellLastName = cellC1_LN.getStringCellValue();

                XSSFCell cellD1_FN = DemoRow.getCell((short) 2);   //first name
                String cellFirstName = cellD1_FN.getStringCellValue();

                XSSFCell cellF1_DOB = DemoRow.getCell((short) 5);   //DOB
                Date cellDOB = cellF1_DOB.getDateCellValue();

                XSSFCell cellH1_PE = DemoRow.getCell((short) 6);   //emp date
                Date cellEmploymentDate = cellH1_PE.getDateCellValue();

                XSSFCell cellG1_PE = DemoRow.getCell((short) 7);   //plan entry
                Date cellPlanEntry = cellG1_PE.getDateCellValue();

                System.out.println("Employee ID: " + cellEmployeeNumber);
                System.out.println("Last Name: " + cellLastName);
                System.out.println("First Name: " + cellFirstName);
                System.out.println("DOB: " + dF.format(cellDOB));

                //    XSSFCell cellI1 = DemoRow.getCell((short) 8);  //status date
                // String i1Val = datetemp.format( cellI1.getDateCellValue());
                //    Date cellStatusDate = cellI1.getDateCellValue();

                //   XSSFCell cellJ1 = DemoRow.getCell((short) 9); //status
                //  String   cellStatus = cellJ1.getStringCellValue();


                if (cellDOB.after(cellPlanEntry) || cellDOB.after(cellEmploymentDate) /*|| cellDOB.after(cellStatusDate)*/) {
                    String choice = null;

                    if (cellDOB.after(cellPlanEntry)) {
                        choice = "Plan Entry Date ";
                    }
                    if (cellDOB.after(cellEmploymentDate)) {
                        choice = "Employment Date ";
                    }
                    //     if (cellDOB.after(cellStatusDate)) {
                    //      choice = "Status Date ";
                    //  }
                    if (cellDOB.after(cellPlanEntry) && cellDOB.after(cellEmploymentDate)) {
                        choice = "Plan Entry Date & Employment Date ";
                    }
                    //  if (cellDOB.after(cellStatusDate) && cellDOB.after(cellPlanEntry)) {
                    //       choice = "Plan Entry Date & Status Date ";
                    //  }

                    //   if (cellDOB.after(cellEmploymentDate) && cellDOB.after(cellStatusDate)) {
                    //       choice = "Employment Date & Status Date ";
                    //    }
                    //   if (cellDOB.after(cellEmploymentDate) && cellDOB.after(cellStatusDate) && cellDOB.after(cellPlanEntry)) {
                    //       choice = "Employment Date & Status Date & Plan Entry Date ";
                    //  }
/*
                System.out.println("Employee ID: " + a1Val_EM);
                System.out.println("Last Name: " + b1Val_LN);
                System.out.println("First Name: " + c1Val_FN);
                System.out.println("Date of Birth: " + dF.format(f1Val_DOB));
                System.out.println("Plan Entry: " + dF.format(g1Val_PE));
                System.out.println("Employment Date: " + dF.format(h1Val_EM));
                System.out.println("Status Date: " + dF.format(i1Val_SD));
                System.out.println("Decision: Contact the administrator");
                System.out.println();*/

                    //  stringBuilder.append("\nResult: This member's date of birth is not before date of employment and enrollment to the plan"+"\n");
                    stringBuilder.append("\nResult: This member's date of birth is not before the " + choice + "\n");
                    stringBuilder.append("Employee ID: " + cellEmployeeNumber + "\n");
                    stringBuilder.append("Last Name: " + cellLastName + "\n");
                    stringBuilder.append("First Name: " + cellFirstName + "\n");
                    stringBuilder.append("Date of Birth: " + dF.format(cellDOB) + "\n");
                    stringBuilder.append("Plan Entry: " + dF.format(cellPlanEntry) + "\n");
                    stringBuilder.append("Employment Date: " + dF.format(cellEmploymentDate) + "\n");
                    // stringBuilder.append("Status Date: " + dF.format(cellStatusDate) + "\n");
                    stringBuilder.append("--------------------------------------------------------------------------------------------------------------\n");
                    //   list.add(cellEmployeeNumber+","+cellLastName+","+cellFirstName+","+dF.format(cellDOB)+","+dF.format(cellEmploymentDate)+","+dF.format(cellPlanEntry)+","+dF.format(cellStatusDate)+','+cellStatus);
                    list.add(cellEmployeeNumber + "," + cellLastName + "," + cellFirstName + "," + dF.format(cellDOB) + "," + dF.format(cellEmploymentDate) + "," + dF.format(cellPlanEntry));
                    check++;

                }
//readFromRow+=1;
            }
        } catch (NullPointerException npe) {
            npe.printStackTrace();
        }
        System.out.println("Notice: The Date of Birth Check Process has now been completed.");
        if (check == 0)
            stringBuilder.append("\nNotice: There were no Discrepancies found with the Members' Date of Birth in this list of Members");
        stringBuilder.append("\n\nNotice: The Date of Birth Check Process has now been completed.\n");

        //   return String.valueOf(stringBuilder);
        this.setResult(String.valueOf(stringBuilder));

        return list;
    }

    public ArrayList Check_Age(String workingDir, int age) throws IOException {
        ArrayList<String> list = new ArrayList<>();
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("Notice: The Age Check Validates that all Employees Are Eligible to be enrolled in The Plan\n\n");
//        FileInputStream fileInputStream = new FileInputStream(filePathValData);
        //    FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Valuation Data.xlsx");
        FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Valuation Data.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet DemoSheet = workbook.getSheet("DEMO");
        int check = 0;
        // int NoMembers = DemoSheet.getPhysicalNumberOfRows();
        int NoMembers = Utility.getNumberOfMembersInSheet(DemoSheet);


        for (int row = 0; row < NoMembers; row++) {
            int temp = row;

            if (temp == 0) {
                row = 1;  //start reading from second row
            }

            XSSFRow DemoRow = DemoSheet.getRow(row);

            XSSFCell cellA1_EM = DemoRow.getCell((short) 0);  //employee number
            if (cellA1_EM == null) {
                cellA1_EM = DemoRow.createCell(0);
            }
            String result = cellA1_EM.getStringCellValue();
            String cellEmployeeNumber = result.replaceAll("[-]", "");

            XSSFCell cellC1_LN = DemoRow.getCell((short) 2);   //last name
            String cellLastName = cellC1_LN.getStringCellValue();

            XSSFCell cellD1_FN = DemoRow.getCell((short) 3);   //first name
            String cellFirstName = cellD1_FN.getStringCellValue();

            XSSFCell cellF1_DOB = DemoRow.getCell((short) 5);   //DOB
            Date cellDOB = cellF1_DOB.getDateCellValue();

            XSSFCell cellG1_PE = DemoRow.getCell((short) 6);   //plan entry
            Date cellPlanEntry = cellG1_PE.getDateCellValue();

            XSSFCell cellH1_PE = DemoRow.getCell((short) 7);   //emp date
            Date cellEmploymentDate = cellH1_PE.getDateCellValue();

            XSSFCell cellI1 = DemoRow.getCell((short) 8);  //status date
            // String i1Val = datetemp.format( cellI1.getDateCellValue());
            Date cellStatusDate = cellI1.getDateCellValue();

            XSSFCell cellJ1 = DemoRow.getCell((short) 9); //status
            String cellStatus = cellJ1.getStringCellValue();


            SimpleDateFormat dF = new SimpleDateFormat();

            dF.applyPattern("dd-MMM-yy");

            int memberAge = Utility.getAge(cellDOB, cellPlanEntry);
            // int memberAge = Utility.calculateAge(birthDate, planEntryDate);

            if (memberAge < age || memberAge > 70) {
/*                System.out.println("Employee ID: " + a1Val_EM);
                System.out.println("Last Name: " + b1Val_LN);
                System.out.println("First Name: " + c1Val_FN);
                System.out.println("Date of Birth: " + dF.format(g1Val_DOB));*/
                //  System.out.println("Employment Date: " + dF.format(h1Val_PE));
                // System.out.println("Result: This member age is "+age +"as at their Plan Entry date " + dF.format(cellPlantry));
                System.out.println();

                //for gui
                stringBuilder.append("Result: This member age is " + memberAge + " as at their Plan Entry date " + dF.format(cellPlanEntry) + "\n");
                stringBuilder.append("Employee ID: " + cellEmployeeNumber + "\n");
                stringBuilder.append("Last Name: " + cellLastName + "\n");
                stringBuilder.append("First Name: " + cellFirstName + "\n");
                stringBuilder.append("Date of Birth: " + dF.format(cellDOB) + "\n");
                stringBuilder.append("Plan Entry: " + dF.format(cellPlanEntry) + "\n");
                stringBuilder.append("---------------------------------------------------------------------------------\n");


                list.add(cellEmployeeNumber + "," + cellLastName + "," + cellFirstName + "," + dF.format(cellDOB) + "," + dF.format(cellEmploymentDate) + "," + dF.format(cellPlanEntry) + "," + dF.format(cellStatusDate) + ',' + cellStatus);
                check++;
            }

        }
        System.out.println("Notice: The Age Check Process has now been completed.");
        if (check == 0)
            stringBuilder.append("Notice: All Members' Ages have been checked in respect to their Plan Entry Date, No Discrepancy was found\n");
        stringBuilder.append("\nNotice: The Age Check Process has now been completed.\n");

        this.setResult(String.valueOf(stringBuilder));
        return list;
    }


    ///NEW CHECKS FOR NEW FORMAT
    public ArrayList check_FivePercent_PensionableSalary(String PensionPlanStartDate, String PensionPlanEndDate, String workindDir) throws IOException {
        ArrayList<String> list = new ArrayList<>();
        StringBuilder stringBuilder = new StringBuilder();

        DecimalFormat dF = new DecimalFormat("#.####");//#.##
        SimpleDateFormat dateF = new SimpleDateFormat("dd-MMM-yy");

        String SD[] = PensionPlanStartDate.split("/");
        int startMonth = Integer.parseInt(SD[0]);
        int startDay = Integer.parseInt(SD[1]);
        int startYear = Integer.parseInt(SD[2]);

        String ED[] = PensionPlanEndDate.split("/");
        int endMonth = Integer.parseInt(ED[0]);
        int endDay = Integer.parseInt(ED[1]);
        int endYear = Integer.parseInt(ED[2]);

        int EndYear = endYear;
        int EndMonth = endMonth;
        int EndDay = endDay;

        int StartYear = startYear;
        int StartMonth = startMonth;
        int StartDay = startDay;

        DateFormat df = new SimpleDateFormat("yyyy.MM.dd");
        Date startDate = null;
        Date endDate = null;
        try {
            startDate = df.parse(StartYear + "." + StartMonth + "." + StartDay);
            endDate = df.parse(EndYear + "." + EndMonth + "." + EndDay);
        } catch (ParseException e) {
            e.printStackTrace();
        }

        int years = Utility.getDiffYears(startDate, endDate);
        years += 1;

        //OPEN ACTIVE SHEET
        FileInputStream fileInputStream = new FileInputStream(workindDir + "\\Input Sheet.xlsx");
        XSSFWorkbook workbookInputSheet = new XSSFWorkbook(fileInputStream);

        for (int x = 0; x < years; x++) {
            int Check = 0;
            Calendar cal = Calendar.getInstance();
            cal.set(StartYear, StartMonth, StartDay);
            SimpleDateFormat year = new SimpleDateFormat("yyyy"); // Just the year, with 2 digits
            String getYear = year.format(cal.getTime());


            String Recon = ("Actives at End of Plan Yr " + getYear);
            stringBuilder.append("Notice: Results of the Pensionable Salary Data Quality Check for " + Recon + "\n");
//GET RECON SHEET
            XSSFSheet reconSheet = workbookInputSheet.getSheet(Recon);
            int numofActives = Utility.getNumberOfMembersInSheet(reconSheet);

//loop through each member in current active sheet

            for (int readFromRow = 7, rowIterator = 0; rowIterator < numofActives; readFromRow++, rowIterator++) {

                XSSFRow getRow = reconSheet.getRow(readFromRow);

                //get employee id
                XSSFCell reconCellA = getRow.getCell(0);  //employee number
                if (reconCellA == null) {
                    reconCellA = getRow.createCell((short) 0);
                    reconCellA.setCellValue("");
                }
                String resultRecon = reconCellA.getStringCellValue();
                String ReconcellEmployeeID = resultRecon.replaceAll("[-]", "");

                //get LAST NAME
                XSSFCell ReconCellB = getRow.getCell((short) 1);  //last name
                if (ReconCellB == null) {
                    ReconCellB = getRow.createCell((short) 1);
                    ReconCellB.setCellValue("");
                }
                String ReconCellLastName = ReconCellB.getStringCellValue();

                //get FIRST NAME
                XSSFCell ReconCellC = getRow.getCell((short) 2);  //first name
                if (ReconCellC == null) {
                    ReconCellC = getRow.createCell((short) 2);
                    ReconCellC.setCellValue("");
                }
                String ReconCellFirstName = ReconCellC.getStringCellValue();

                //get DOB
                XSSFCell ReconCellF = getRow.getCell((short) 5);  //first name
                if (ReconCellF == null) {
                    ReconCellF = getRow.createCell((short) 5);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellDOB = ReconCellF.getDateCellValue();

                //get date of employment
                XSSFCell ReconCellG = getRow.getCell((short) 6);  //first name
                if (ReconCellG == null) {
                    ReconCellG = getRow.createCell((short) 6);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCelldateofEmployment = ReconCellG.getDateCellValue();

                //get DOB
                XSSFCell ReconCellH = getRow.getCell((short) 7);  //first name
                if (ReconCellH == null) {
                    ReconCellH = getRow.createCell((short) 7);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellDateofEnrolment = ReconCellH.getDateCellValue();


                //GET PENSIONABLE SALARY
                XSSFCell ReconCellL = getRow.getCell((short) 11);
                if (ReconCellL == null) {
                    ReconCellL = getRow.createCell((short) 11);
                    ReconCellL.setCellValue(0.00);
                }
                double ReconCell_PensionableSalary = ReconCellL.getNumericCellValue();


                //employee basic contribution
                XSSFCell ReconCellM = getRow.getCell((short) 12);
                if (ReconCellM == null) {
                    ReconCellM = getRow.createCell((short) 12);
                    ReconCellM.setCellValue(0.00);
                }
                double ReconCell_EmployeeBasic_Contribution = ReconCellM.getNumericCellValue();

                double check = 0.05 * ReconCell_PensionableSalary;
                check = Double.parseDouble(dF.format(check));

                if (check != ReconCell_EmployeeBasic_Contribution) {
                    Check++;

                    String test = "";

                    stringBuilder.append("\nEmployee ID: " + ReconcellEmployeeID + "\n");
                    stringBuilder.append("Last Name: " + ReconCellLastName + "\n");
                    stringBuilder.append("First Name: " + ReconCellFirstName + "\n");
                    stringBuilder.append("Pensionable Salary: $" + dF.format(ReconCell_PensionableSalary) + "\n");
                    stringBuilder.append("Employee Basic: $" + ReconCell_EmployeeBasic_Contribution + "\n");

                    list.add(ReconcellEmployeeID + "," + ReconCellLastName + "," + ReconCellFirstName + "," + dateF.format(ReconCellDOB) + "," + dateF.format(ReconCelldateofEmployment) + "," + dateF.format(ReconCellDateofEnrolment));

                    test = "false";
                    stringBuilder.append("Result: " + ReconCellLastName + " Contribution of: $" + ReconCell_EmployeeBasic_Contribution + " is " + dF.format((ReconCell_EmployeeBasic_Contribution / ReconCell_PensionableSalary) * 100) + "% of $" + dF.format(ReconCell_PensionableSalary) + "\n");
                }

            }//end of looping through current active sheet
            if (Check == 0)
                stringBuilder.append("\nNotice: There were no Discrepancies found among the Members' Pensionable Salary and their Contributions for " + Recon);
            //    stringBuilder.append("\n\nNotice: The Pensionable Check Process for "+Recon+"\n\n");
            stringBuilder.append("\n----------------------------------------------------------------------------------------------------------------------------------------------------------------------------\n\n");
            this.setResult(String.valueOf(stringBuilder));
            StartYear++;

        }//end of year looping through years

        return list;

    }

    public ArrayList check_For_Duplicates(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {
        ArrayList<String> list = new ArrayList<>();
        StringBuilder stringBuilder = new StringBuilder();
        //   stringBuilder.append("Notice: The Duplicate Check Validates the list of Members for Repeated Records");


        SimpleDateFormat dateF = new SimpleDateFormat("dd-MMM-yy");

        String SD[] = PensionPlanStartDate.split("/");
        int startMonth = Integer.parseInt(SD[0]);
        int startDay = Integer.parseInt(SD[1]);
        int startYear = Integer.parseInt(SD[2]);

        String ED[] = PensionPlanEndDate.split("/");
        int endMonth = Integer.parseInt(ED[0]);
        int endDay = Integer.parseInt(ED[1]);
        int endYear = Integer.parseInt(ED[2]);

        int EndYear = endYear;
        int EndMonth = endMonth;
        int EndDay = endDay;

        int StartYear = startYear;
        int StartMonth = startMonth;
        int StartDay = startDay;

        DateFormat df = new SimpleDateFormat("yyyy.MM.dd");
        Date startDate = null;
        Date endDate = null;
        try {
            startDate = df.parse(StartYear + "." + StartMonth + "." + StartDay);
            endDate = df.parse(EndYear + "." + EndMonth + "." + EndDay);
        } catch (ParseException e) {
            e.printStackTrace();
        }

        int years = Utility.getDiffYears(startDate, endDate);
        years += 2;

        //OPEN ACTIVE SHEET
        FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Input Sheet.xlsx");
        XSSFWorkbook workbookInputSheet = new XSSFWorkbook(fileInputStream);

        String Recon = null;
        int numofActives;
        int startRow;
        for (int x = 0; x < years; x++) {

            Calendar cal = Calendar.getInstance();
            cal.set(StartYear, StartMonth, StartDay);
            SimpleDateFormat year = new SimpleDateFormat("yyyy"); // Just the year, with 2 digits
            String getYear = year.format(cal.getTime());

            if (x < (years - 1)) {
                Recon = ("Actives at End of Plan Yr " + getYear);
                XSSFSheet reconSheet = workbookInputSheet.getSheet(Recon);
                numofActives = Utility.getNumberOfMembersInSheet(reconSheet);
                startRow = 7;
            } else {
                Recon = ("Terminated up to " + EndYear + "." + EndMonth + "." + EndDay);
                XSSFSheet reconSheet = workbookInputSheet.getSheet(Recon);
                numofActives = Utility.getNumberOfTermineeMembersInSheet(reconSheet);
                startRow = 11;
            }

            stringBuilder.append("Notice: Results of the Duplicate Quality Check for " + Recon + "\n");
//GET RECON SHEET
            XSSFSheet reconSheet = workbookInputSheet.getSheet(Recon);


//loop through each member in current active sheet
            int check = 0;
            int FindIt = 0;
            boolean once = true;
            System.out.println("Recon " + Recon);
            System.out.println("startRow " + startRow);
            System.out.println("numofActives " + numofActives);
            for (int readFromRow = startRow, rowIterator = 0; rowIterator < numofActives; readFromRow++, rowIterator++) {

                XSSFRow getRow = reconSheet.getRow(readFromRow);

                //get employee id
                XSSFCell reconCellA = getRow.getCell(0);  //employee number
                if (reconCellA == null) {
                    reconCellA = getRow.createCell((short) 0);
                    reconCellA.setCellValue("");
                }
                String resultRecon = reconCellA.getStringCellValue();
                String ReconcellEmployeeID = resultRecon.replaceAll("[-]", "");

                //get LAST NAME
                XSSFCell ReconCellB = getRow.getCell((short) 1);  //last name
                if (ReconCellB == null) {
                    ReconCellB = getRow.createCell((short) 1);
                    ReconCellB.setCellValue("");
                }
                String ReconCellLastName = ReconCellB.getStringCellValue();

                //get FIRST NAME
                XSSFCell ReconCellC = getRow.getCell((short) 2);  //first name
                if (ReconCellC == null) {
                    ReconCellC = getRow.createCell((short) 2);
                    ReconCellC.setCellValue("");
                }
                String ReconCellFirstName = ReconCellC.getStringCellValue();

                //get DOB
                XSSFCell ReconCellF = getRow.getCell((short) 5);
                if (ReconCellF == null) {
                    ReconCellF = getRow.createCell((short) 5);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellDOB = ReconCellF.getDateCellValue();

                //get date of employment
                XSSFCell ReconCellG = getRow.getCell((short) 6);
                if (ReconCellG == null) {
                    ReconCellG = getRow.createCell((short) 6);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCelldateofEmployment = ReconCellG.getDateCellValue();

                //get DOB
                XSSFCell ReconCellH = getRow.getCell((short) 7);  //first name
                if (ReconCellH == null) {
                    ReconCellH = getRow.createCell((short) 7);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellDateofEnrolment = ReconCellH.getDateCellValue();

                FindIt = 0;
                for (int readFromRow2 = startRow, rowIterator2 = 0; rowIterator2 < numofActives; readFromRow2++, rowIterator2++) {

                    XSSFRow getRow2 = reconSheet.getRow(readFromRow2);

                    //get employee id
                    XSSFCell reconCellA2 = getRow2.getCell(0);  //employee number
                    if (reconCellA2 == null) {
                        reconCellA2 = getRow2.createCell((short) 0);
                        reconCellA2.setCellValue("");
                    }
                    String resultRecon2 = reconCellA2.getStringCellValue();
                    String ReconcellEmployeeID2 = resultRecon2.replaceAll("[-]", "");

                    //get LAST NAME
                    XSSFCell ReconCellB2 = getRow2.getCell((short) 1);  //last name
                    if (ReconCellB2 == null) {
                        ReconCellB2 = getRow2.createCell((short) 1);
                        ReconCellB2.setCellValue("");
                    }
                    String ReconCellLastName2 = ReconCellB2.getStringCellValue();

                    //get FIRST NAME
                    XSSFCell ReconCellC2 = getRow2.getCell((short) 2);  //first name
                    if (ReconCellC2 == null) {
                        ReconCellC2 = getRow2.createCell((short) 2);
                        ReconCellC2.setCellValue("");
                    }
                    String ReconCellFirstName2 = ReconCellC2.getStringCellValue();

                    //get DOB
                    XSSFCell ReconCellF2 = getRow2.getCell((short) 5);
                    if (ReconCellF2 == null) {
                        ReconCellF2 = getRow2.createCell((short) 5);
                        // ReconCellF.setCellValue("");
                    }
                    Date ReconCellDOB2 = ReconCellF2.getDateCellValue();

                    if (ReconcellEmployeeID.equals(ReconcellEmployeeID2) && ReconCellLastName.equals(ReconCellLastName2) && ReconCellFirstName.equals(ReconCellFirstName2) || ReconCellFirstName.equals(ReconCellFirstName2) && ReconCellLastName.equals(ReconCellLastName2) && ReconCellDOB.equals(ReconCellDOB2) || ReconcellEmployeeID.equals(ReconcellEmployeeID2)) {
                        FindIt++;
                        //  findIt=true;


                        if (once && FindIt > 1) {
                            stringBuilder.append("\n\nThe Process has found these members to be as repeated records\n\n");
                            once = false;
                        }


                        if (FindIt > 1) {
                            check++;
                            FindIt = 0;

                            stringBuilder.append("Employee ID: " + ReconcellEmployeeID + "\n");
                            stringBuilder.append("Last Name: " + ReconCellLastName + "\n");
                            stringBuilder.append("First Name: " + ReconCellFirstName + "\n");
                            stringBuilder.append("Date of Birth: " + dateF.format(ReconCellDOB) + "\n\n");
                            list.add(ReconcellEmployeeID + "," + ReconCellLastName + "," + ReconCellFirstName + "," + dateF.format(ReconCellDOB) + "," + dateF.format(ReconCelldateofEmployment) + "," + dateF.format(ReconCellDateofEnrolment));
                            break;
                        }

                    }


                }//end of looping through current active sheet

            }
            if (check == 0 && once)
                stringBuilder.append("\n\nNotice: There were no Duplicate records found in this list of Members for " + Recon);
            stringBuilder.append("\n----------------------------------------------------------------------------------------------------------------------------------------------------------------------------\n\n");
            this.setResult(String.valueOf(stringBuilder));
            StartYear++;
        }//end of year looping through years
        return list;
    }

    public ArrayList check_Age(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir, int age) throws IOException {
        ArrayList<String> list = new ArrayList<>();
        StringBuilder stringBuilder = new StringBuilder();

        DecimalFormat dF = new DecimalFormat("#.####");//#.##
        SimpleDateFormat dateF = new SimpleDateFormat("dd-MMM-yy");


        String SD[] = PensionPlanStartDate.split("/");
        int startMonth = Integer.parseInt(SD[0]);
        int startDay = Integer.parseInt(SD[1]);
        int startYear = Integer.parseInt(SD[2]);

        String ED[] = PensionPlanEndDate.split("/");
        int endMonth = Integer.parseInt(ED[0]);
        int endDay = Integer.parseInt(ED[1]);
        int endYear = Integer.parseInt(ED[2]);

        int EndYear = endYear;
        int EndMonth = endMonth;
        int EndDay = endDay;

        int StartYear = startYear;
        int StartMonth = startMonth;
        int StartDay = startDay;

        DateFormat df = new SimpleDateFormat("yyyy.MM.dd");
        Date startDate = null;
        Date endDate = null;
        try {
            startDate = df.parse(StartYear + "." + StartMonth + "." + StartDay);
            endDate = df.parse(EndYear + "." + EndMonth + "." + EndDay);
        } catch (ParseException e) {
            e.printStackTrace();
        }

        int years = Utility.getDiffYears(startDate, endDate);
        years += 2;

        //OPEN ACTIVE SHEET
        FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Input Sheet.xlsx");
        XSSFWorkbook workbookInputSheet = new XSSFWorkbook(fileInputStream);
        String Recon = null;
        int numofActives;
        int startRow;

        for (int x = 0; x < years; x++) {
            int check = 0;
            Calendar cal = Calendar.getInstance();
            cal.set(StartYear, StartMonth, StartDay);
            SimpleDateFormat year = new SimpleDateFormat("yyyy"); // Just the year, with 2 digits
            String getYear = year.format(cal.getTime());


            if (x < (years - 1)) {
                Recon = ("Actives at End of Plan Yr " + getYear);
                XSSFSheet reconSheet = workbookInputSheet.getSheet(Recon);
                numofActives = Utility.getNumberOfMembersInSheet(reconSheet);
                startRow = 7;
            } else {
                Recon = ("Terminated up to " + EndYear + "." + EndMonth + "." + EndDay);
                XSSFSheet reconSheet = workbookInputSheet.getSheet(Recon);
                numofActives = Utility.getNumberOfTermineeMembersInSheet(reconSheet);
                startRow = 11;
            }


            stringBuilder.append("Notice: Results of the Age Data Quality Check for " + Recon + "\n");
//GET RECON SHEET
            XSSFSheet reconSheet = workbookInputSheet.getSheet(Recon);


//loop through each member in current active sheet
            System.out.println("Recon " + Recon);
            System.out.println("startRow " + startRow);
            System.out.println("numofActives " + numofActives);
            for (int readFromRow = startRow, rowIterator = 0; rowIterator < numofActives; readFromRow++, rowIterator++) {

                XSSFRow getRow = reconSheet.getRow(readFromRow);

                //get employee id
                XSSFCell reconCellA = getRow.getCell(0);  //employee number
                if (reconCellA == null) {
                    reconCellA = getRow.createCell((short) 0);
                    reconCellA.setCellValue("");
                }
                String resultRecon = reconCellA.getStringCellValue();
                String ReconcellEmployeeID = resultRecon.replaceAll("[-]", "");

                //get LAST NAME
                XSSFCell ReconCellB = getRow.getCell((short) 1);  //last name
                if (ReconCellB == null) {
                    ReconCellB = getRow.createCell((short) 1);
                    ReconCellB.setCellValue("");
                }
                String ReconCellLastName = ReconCellB.getStringCellValue();

                //get FIRST NAME
                XSSFCell ReconCellC = getRow.getCell((short) 2);  //first name
                if (ReconCellC == null) {
                    ReconCellC = getRow.createCell((short) 2);
                    ReconCellC.setCellValue("");
                }
                String ReconCellFirstName = ReconCellC.getStringCellValue();

                //get DOB
                XSSFCell ReconCellF = getRow.getCell((short) 5);  //first name
                if (ReconCellF == null) {
                    ReconCellF = getRow.createCell((short) 5);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellDOB = ReconCellF.getDateCellValue();

                //get date of employment
                XSSFCell ReconCellG = getRow.getCell((short) 6);  //first name
                if (ReconCellG == null) {
                    ReconCellG = getRow.createCell((short) 6);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCelldateofEmployment = ReconCellG.getDateCellValue();

                //get DOB
                XSSFCell ReconCellH = getRow.getCell((short) 7);  //first name
                if (ReconCellH == null) {
                    ReconCellH = getRow.createCell((short) 7);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellDateofEnrolment = ReconCellH.getDateCellValue();

                System.out.println("ReconCellDOB" + ReconCellDOB);
                System.out.println("ReconCellDateofEnrolment" + ReconCellDateofEnrolment);

                int memberAge = Utility.getAge(ReconCellDOB, ReconCellDateofEnrolment);

                if (memberAge < age || memberAge > 70) {
                    stringBuilder.append("\nResult: This member's age is " + memberAge + " as at their Plan Entry date " + dateF.format(ReconCellDateofEnrolment) + "\n");
                    stringBuilder.append("Employee ID: " + ReconcellEmployeeID + "\n");
                    stringBuilder.append("Last Name: " + ReconCellLastName + "\n");
                    stringBuilder.append("First Name: " + ReconCellFirstName + "\n");
                    stringBuilder.append("Date of Birth: " + dateF.format(ReconCellDOB) + "\n");
                    stringBuilder.append("Plan Entry: " + dateF.format(ReconCellDateofEnrolment) + "\n");

                    list.add(ReconcellEmployeeID + "," + ReconCellLastName + "," + ReconCellFirstName + "," + dateF.format(ReconCellDOB) + "," + dateF.format(ReconCelldateofEmployment) + "," + dateF.format(ReconCellDateofEnrolment));
                    check++;
                }


            }
            if (check == 0)
                stringBuilder.append("\nNotice: There were no Discrepancies found among the Members' Ages for " + Recon);
            stringBuilder.append("\n----------------------------------------------------------------------------------------------------------------------------------------------------------------------------\n\n");
            this.setResult(String.valueOf(stringBuilder));
            StartYear++;
        }
        this.setResult(String.valueOf(stringBuilder));
        return list;
    }

    public ArrayList check_DateofBirth(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {
        ArrayList<String> list = new ArrayList<>();
        StringBuilder stringBuilder = new StringBuilder();
        SimpleDateFormat dateF = new SimpleDateFormat("dd-MMM-yy");

        String SD[] = PensionPlanStartDate.split("/");
        int startMonth = Integer.parseInt(SD[0]);
        int startDay = Integer.parseInt(SD[1]);
        int startYear = Integer.parseInt(SD[2]);

        String ED[] = PensionPlanEndDate.split("/");
        int endMonth = Integer.parseInt(ED[0]);
        int endDay = Integer.parseInt(ED[1]);
        int endYear = Integer.parseInt(ED[2]);

        int EndYear = endYear;
        int EndMonth = endMonth;
        int EndDay = endDay;

        int StartYear = startYear;
        int StartMonth = startMonth;
        int StartDay = startDay;

        DateFormat df = new SimpleDateFormat("yyyy.MM.dd");
        Date startDate = null;
        Date endDate = null;
        try {
            startDate = df.parse(StartYear + "." + StartMonth + "." + StartDay);
            endDate = df.parse(EndYear + "." + EndMonth + "." + EndDay);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        String Recon = null;
        int numofActives;
        int startRow;
        XSSFSheet reconSheet;

        int years = Utility.getDiffYears(startDate, endDate);
        years += 2;//

        //OPEN ACTIVE SHEET
        FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Input Sheet.xlsx");
        XSSFWorkbook workbookInputSheet = new XSSFWorkbook(fileInputStream);

        for (int x = 0; x < years; x++) {
            int Check = 0;
            Calendar cal = Calendar.getInstance();
            cal.set(StartYear, StartMonth, StartDay);
            SimpleDateFormat year = new SimpleDateFormat("yyyy"); // Just the year, with 2 digits
            String getYear = year.format(cal.getTime());


            if (x < (years - 1)) {
                Recon = ("Actives at End of Plan Yr " + getYear);
                reconSheet = workbookInputSheet.getSheet(Recon);
                numofActives = Utility.getNumberOfMembersInSheet(reconSheet);
                startRow = 7;
            } else {
                Recon = ("Terminated up to " + EndYear + "." + EndMonth + "." + EndDay);
                reconSheet = workbookInputSheet.getSheet(Recon);
                numofActives = Utility.getNumberOfTermineeMembersInSheet(reconSheet);
                startRow = 11;
            }

            stringBuilder.append("Notice: Results of the Date of Birth Data Quality Check for " + Recon + "\n");


//loop through each member in current active sheet

            for (int readFromRow = startRow, rowIterator = 0; rowIterator < numofActives; readFromRow++, rowIterator++) {

                XSSFRow getRow = reconSheet.getRow(readFromRow);

                //get employee id
                XSSFCell reconCellA = getRow.getCell(0);  //employee number
                if (reconCellA == null) {
                    reconCellA = getRow.createCell((short) 0);
                    reconCellA.setCellValue("");
                }
                String resultRecon = reconCellA.getStringCellValue();
                String ReconcellEmployeeID = resultRecon.replaceAll("[-]", "");

                //get LAST NAME
                XSSFCell ReconCellB = getRow.getCell((short) 1);  //last name
                if (ReconCellB == null) {
                    ReconCellB = getRow.createCell((short) 1);
                    ReconCellB.setCellValue("");
                }
                String ReconCellLastName = ReconCellB.getStringCellValue();

                //get FIRST NAME
                XSSFCell ReconCellC = getRow.getCell((short) 2);  //first name
                if (ReconCellC == null) {
                    ReconCellC = getRow.createCell((short) 2);
                    ReconCellC.setCellValue("");
                }
                String ReconCellFirstName = ReconCellC.getStringCellValue();

                //get DOB
                XSSFCell ReconCellF = getRow.getCell((short) 5);  //first name
                if (ReconCellF == null) {
                    ReconCellF = getRow.createCell((short) 5);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellDOB = ReconCellF.getDateCellValue();

                //get date of employment
                XSSFCell ReconCellG = getRow.getCell((short) 6);  //first name
                if (ReconCellG == null) {
                    ReconCellG = getRow.createCell((short) 6);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCelldateofEmployment = ReconCellG.getDateCellValue();

                //get DOB
                XSSFCell ReconCellH = getRow.getCell((short) 7);  //first name
                if (ReconCellH == null) {
                    ReconCellH = getRow.createCell((short) 7);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellDateofEnrolment = ReconCellH.getDateCellValue();


                if (ReconCellDOB.after(ReconCellDateofEnrolment) || ReconCellDOB.after(ReconCelldateofEmployment)) {
                    String choice = null;

                    if (ReconCellDOB.after(ReconCellDateofEnrolment)) {
                        choice = "Plan Entry Date ";
                    }

                    if (ReconCellDOB.after(ReconCelldateofEmployment)) {
                        choice = "Employment Date ";
                    }

                    if (ReconCellDOB.after(ReconCellDateofEnrolment) && ReconCellDOB.after(ReconCelldateofEmployment)) {
                        choice = "Plan Entry Date & Employment Date ";
                    }

                    stringBuilder.append("\nResult: This member's date of birth is not before the " + choice + "\n");
                    stringBuilder.append("Employee ID: " + ReconcellEmployeeID + "\n");
                    stringBuilder.append("Last Name: " + ReconCellLastName + "\n");
                    stringBuilder.append("First Name: " + ReconCellFirstName + "\n");
                    stringBuilder.append("Date of Birth: " + dateF.format(ReconCellDOB) + "\n");
                    stringBuilder.append("Plan Entry: " + dateF.format(ReconCellDateofEnrolment) + "\n");
                    stringBuilder.append("Employment Date: " + dateF.format(ReconCelldateofEmployment) + "\n");

                    //   list.add(cellEmployeeNumber+","+cellLastName+","+cellFirstName+","+dF.format(cellDOB)+","+dF.format(cellEmploymentDate)+","+dF.format(cellPlanEntry)+","+dF.format(cellStatusDate)+','+cellStatus);
                    list.add(ReconcellEmployeeID + "," + ReconCellLastName + "," + ReconCellFirstName + "," + dateF.format(ReconCellDOB) + "," + dateF.format(ReconCelldateofEmployment) + "," + dateF.format(ReconCellDateofEnrolment));
                    Check++;

                }


            }//end of looping through current active sheet
            if (Check == 0)
                stringBuilder.append("\nNotice: There were no Discrepancies found with the Members' Date of Birth in this list of Members for " + Recon);
            //    stringBuilder.append("\n\nNotice: The Pensionable Check Process for "+Recon+"\n\n");
            stringBuilder.append("\n----------------------------------------------------------------------------------------------------------------------------------------------------------------------------\n\n");
            this.setResult(String.valueOf(stringBuilder));
            StartYear++;
        }//end of year looping through years
        return list;
    }

    public ArrayList check_Plan_EntryDate_empDATE(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {
        ArrayList<String> list = new ArrayList<>();
        StringBuilder stringBuilder = new StringBuilder();
        SimpleDateFormat dateF = new SimpleDateFormat("dd-MMM-yy");

        String SD[] = PensionPlanStartDate.split("/");
        int startMonth = Integer.parseInt(SD[0]);
        int startDay = Integer.parseInt(SD[1]);
        int startYear = Integer.parseInt(SD[2]);

        String ED[] = PensionPlanEndDate.split("/");
        int endMonth = Integer.parseInt(ED[0]);
        int endDay = Integer.parseInt(ED[1]);
        int endYear = Integer.parseInt(ED[2]);

        int EndYear = endYear;
        int EndMonth = endMonth;
        int EndDay = endDay;

        int StartYear = startYear;
        int StartMonth = startMonth;
        int StartDay = startDay;

        DateFormat df = new SimpleDateFormat("yyyy.MM.dd");
        Date startDate = null;
        Date endDate = null;
        try {
            startDate = df.parse(StartYear + "." + StartMonth + "." + StartDay);
            endDate = df.parse(EndYear + "." + EndMonth + "." + EndDay);
        } catch (ParseException e) {
            e.printStackTrace();
        }

        int years = Utility.getDiffYears(startDate, endDate);
        years += 2;

        //OPEN ACTIVE SHEET
        FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Input Sheet.xlsx");
        XSSFWorkbook workbookInputSheet = new XSSFWorkbook(fileInputStream);
        String Recon = null;
        int numofActives;
        int startRow;
        XSSFSheet reconSheet;
        for (int x = 0; x < years; x++) {
            int Check = 0;
            Calendar cal = Calendar.getInstance();
            cal.set(StartYear, StartMonth, StartDay);
            SimpleDateFormat year = new SimpleDateFormat("yyyy"); // Just the year, with 2 digits
            String getYear = year.format(cal.getTime());

            if (x < (years - 1)) {
                Recon = ("Actives at End of Plan Yr " + getYear);
                reconSheet = workbookInputSheet.getSheet(Recon);
                numofActives = Utility.getNumberOfMembersInSheet(reconSheet);
                startRow = 7;
            } else {
                Recon = ("Terminated up to " + EndYear + "." + EndMonth + "." + EndDay);
                reconSheet = workbookInputSheet.getSheet(Recon);
                numofActives = Utility.getNumberOfTermineeMembersInSheet(reconSheet);
                startRow = 11;
            }


            stringBuilder.append("Notice: Results of the Plan Entry Data Quality Check for " + Recon + "\n");

//loop through each member in current active sheet

            for (int readFromRow = startRow, rowIterator = 0; rowIterator < numofActives; readFromRow++, rowIterator++) {

                XSSFRow getRow = reconSheet.getRow(readFromRow);

                //get employee id
                XSSFCell reconCellA = getRow.getCell(0);  //employee number
                if (reconCellA == null) {
                    reconCellA = getRow.createCell((short) 0);
                    reconCellA.setCellValue("");
                }
                String resultRecon = reconCellA.getStringCellValue();
                String ReconcellEmployeeID = resultRecon.replaceAll("[-]", "");

                //get LAST NAME
                XSSFCell ReconCellB = getRow.getCell((short) 1);  //last name
                if (ReconCellB == null) {
                    ReconCellB = getRow.createCell((short) 1);
                    ReconCellB.setCellValue("");
                }
                String ReconCellLastName = ReconCellB.getStringCellValue();

                //get FIRST NAME
                XSSFCell ReconCellC = getRow.getCell((short) 2);  //first name
                if (ReconCellC == null) {
                    ReconCellC = getRow.createCell((short) 2);
                    ReconCellC.setCellValue("");
                }
                String ReconCellFirstName = ReconCellC.getStringCellValue();

                //get DOB
                XSSFCell ReconCellF = getRow.getCell((short) 5);  //first name
                if (ReconCellF == null) {
                    ReconCellF = getRow.createCell((short) 5);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellDOB = ReconCellF.getDateCellValue();

                //get date of employment
                XSSFCell ReconCellG = getRow.getCell((short) 6);  //first name
                if (ReconCellG == null) {
                    ReconCellG = getRow.createCell((short) 6);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCelldateofEmployment = ReconCellG.getDateCellValue();

                //get DOB
                XSSFCell ReconCellH = getRow.getCell((short) 7);  //first name
                if (ReconCellH == null) {
                    ReconCellH = getRow.createCell((short) 7);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellDateofEnrolment = ReconCellH.getDateCellValue();


                if (ReconCellDateofEnrolment.before(ReconCelldateofEmployment)) {

                    stringBuilder.append("\nResult: This member's Plan Entry Date at " + dateF.format(ReconCellDateofEnrolment) + " is before their Employment Date at " + dateF.format(ReconCelldateofEmployment) + "\n");
                    stringBuilder.append("Employee ID: " + ReconcellEmployeeID + "\n");
                    stringBuilder.append("Last Name: " + ReconCellLastName + "\n");
                    stringBuilder.append("First Name: " + ReconCellFirstName + "\n");
                    stringBuilder.append("Date of Birth: " + dateF.format(ReconCellDOB) + "\n");
                    stringBuilder.append("Plan Entry: " + dateF.format(ReconCellDateofEnrolment) + "\n");
                    stringBuilder.append("Employment Date: " + dateF.format(ReconCelldateofEmployment) + "\n");

                    //   list.add(cellEmployeeNumber+","+cellLastName+","+cellFirstName+","+dF.format(cellDOB)+","+dF.format(cellEmploymentDate)+","+dF.format(cellPlanEntry)+","+dF.format(cellStatusDate)+','+cellStatus);
                    list.add(ReconcellEmployeeID + "," + ReconCellLastName + "," + ReconCellFirstName + "," + dateF.format(ReconCellDOB) + "," + dateF.format(ReconCelldateofEmployment) + "," + dateF.format(ReconCellDateofEnrolment));
                    Check++;
                }


            }//end of looping through current active sheet
            if (Check == 0)
                stringBuilder.append("\nNotice: There were no Discrepancies found with the Members' Employment Date in respect to their Plane Entry Date for " + Recon);
            //    stringBuilder.append("\n\nNotice: The Pensionable Check Process for "+Recon+"\n\n");
            stringBuilder.append("\n----------------------------------------------------------------------------------------------------------------------------------------------------------------------------\n\n");
            this.setResult(String.valueOf(stringBuilder));
            StartYear++;
        }//end of year looping through years
        return list;
    }

}
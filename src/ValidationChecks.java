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

Utility utility = new Utility();

    public String Check_For_Duplicates(String workingDir) throws IOException {
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("Notice: Results of the Duplicate Check Process:");
        boolean once=true;

                // FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Hose Valuation Data (Actuary's copy).xlsx");
     //   FileInputStream fileInputStream = new FileInputStream(filePathValData);
        FileInputStream fileInputStream = new FileInputStream(workingDir +"\\Hose Valuation Data (Actuary's copy).xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet DemoSheet = workbook.getSheet("DEMO");

        int NoMembers = DemoSheet.getPhysicalNumberOfRows();
SimpleDateFormat dF = new SimpleDateFormat("dd-MMM-yy");
//dF.format("dd-MMM-yy");
        for (int row=0;row<NoMembers;row++){
            int temp = row;

            if (temp == 0) {
                row= 1;  //start reading from second row
            }

          //  int Row = row;

            XSSFRow DemoRow = DemoSheet.getRow(row);


            XSSFCell cellA1_EM = DemoRow.getCell((short) 0);  //employee number
            if(cellA1_EM==null){
                cellA1_EM = DemoRow.createCell(0);
            }
            String result = cellA1_EM.getStringCellValue();
            String a1Val_EM = result.replaceAll("[-]","");

            XSSFCell cellC1_LN = DemoRow.getCell((short) 2);   //last name
            String b1Val_LN = cellC1_LN.getStringCellValue();

            XSSFCell cellD1_FN= DemoRow.getCell((short) 3);   //first name
            String c1Val_FN = cellD1_FN.getStringCellValue();

            XSSFCell cellF1_DOB= DemoRow.getCell((short) 5);   //first name
            Date f1Val_DOB = cellF1_DOB.getDateCellValue();


            int FindIt = 0;
           boolean findIt = false;

            for (int row2=0;row2<NoMembers;row2++){
                int temp2 = row2;

                if (temp2 == 0) {
                    row2= 1;  //start reading from second row
                }

                //  int Row = row;

                XSSFRow DemoRow2 = DemoSheet.getRow(row2);


                XSSFCell cellA1_EM2 = DemoRow2.getCell((short) 0);  //employee number
                if(cellA1_EM2==null){
                    cellA1_EM2 = DemoRow2.createCell(0);
                }
                String result2 = cellA1_EM2.getStringCellValue();
                String a1Val_EM2 = result2.replaceAll("[-]","");

                XSSFCell cellC1_LN2 = DemoRow2.getCell((short) 2);   //last name
                String b1Val_LN2 = cellC1_LN2.getStringCellValue();

                XSSFCell cellD1_FN2= DemoRow2.getCell((short) 3);   //first name
                String c1Val_FN2 = cellD1_FN2.getStringCellValue();

                if(a1Val_EM.equals(a1Val_EM2)&& b1Val_LN.equals(b1Val_LN2) && c1Val_FN.equals(c1Val_FN2)) {
                    FindIt++;
                  //  findIt=true;


                    if(once==true  && FindIt>1){
                        stringBuilder.append("\n\nThe Process has found these members to be as repeated records\n\n");
                        once=false;
                    }


                    if (FindIt > 1) {



                     //   System.out.println("Please Contact Administrator");
                        System.out.println("Employee ID: " + a1Val_EM);
                        System.out.println("Last Name: " + b1Val_LN);
                        System.out.println("First Name: " + c1Val_FN);
                        System.out.println("DOB: " + dF.format(f1Val_DOB));

                        System.out.println();
                        FindIt = 0;
                        //  break;
                    //    stringBuilder.append("This Member is Duplicate Record, Please Contact Administrator\n"+"\n");
                        stringBuilder.append("Employee ID: " + a1Val_EM+"\n");
                        stringBuilder.append("Last Name: " + b1Val_LN+"\n");
                        stringBuilder.append("First Name: " + c1Val_FN+"\n");
                        stringBuilder.append("DOB: " + dF.format(f1Val_DOB)+"\n");

                        stringBuilder.append("------------------------------------------\n");
                    }
                }


            }


        }
        System.out.println("Notice: The Duplicate check process has now been completed");
        stringBuilder.append("\n\nNotice: The Duplicate check process has now been completed");
        return String.valueOf(stringBuilder+"\n");
    }

    public String Check_FivePercent_PS(String workindDir) throws IOException {

        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("Notice: Results of the Pensionable Salary Check Process:\n\n");

        DecimalFormat dF = new DecimalFormat("#.##");//#.##

        //OPEN ACTIVE SHEET
        FileInputStream fileInputStream = new FileInputStream(workindDir+"\\Hose Valuation Data (Actuary's copy).xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet DemoSheet = workbook.getSheet("DEMO");

        FileInputStream fs = null;
        try {
            fs = new FileInputStream(workindDir + "\\Pensionable Salary - HAS.xlsx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        XSSFWorkbook WB = new XSSFWorkbook(fs);

        int EndYear = 2015;
        int EndMonth = 12;
        int EndDay = 31;

        int StartYear = 2004;
        int StartMonth = 01;
        int StartDay = 01;
        int WriteAt =26;

        DateFormat df = new SimpleDateFormat("yyyy.MM.dd");
        Date startDate = null;
        Date endDate = null;
        try {
            startDate = df.parse(StartYear + "." + StartMonth + "." + StartDay);
            endDate = df.parse(EndYear + "." + EndMonth + "." + EndDay);
        } catch (ParseException e) {
            e.printStackTrace();
        }

        int years = utility.getDiffYears(startDate, endDate);
        FileInputStream fsRecon = null;
        try {
            fsRecon = new FileInputStream(workindDir+"\\Hose Valuation Data (Actuary's copy).xlsx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        XSSFWorkbook WBRecon = new XSSFWorkbook(fsRecon);
        int counter=1;


        for (int x = 0; x <= years; x++) {

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

            for (int row = 0;  row < numOfDemoMembers; row++) {

                    int temp = row;

                    if (temp == 0) {
                        row= 1;  //start reading from second row
                    }

                int Row = row;

                XSSFRow DemoRow = DemoSheet.getRow(row);


                XSSFCell cellA1_EM = DemoRow.getCell((short) 0);  //employee number
                if(cellA1_EM==null){
                    cellA1_EM = DemoRow.createCell(0);
                }
                String result = cellA1_EM.getStringCellValue();
                String a1Val_EM = result.replaceAll("[-]","");

                XSSFCell cellC1_LN = DemoRow.getCell((short) 2);   //last name
                String b1Val_LN = cellC1_LN.getStringCellValue();

                XSSFCell cellD1_FN= DemoRow.getCell((short) 3);   //first name
                String c1Val_FN = cellD1_FN.getStringCellValue();

              //  XSSFCell cellD1 = DemoRow.getCell((short) 3);
               // String d1Val = cellD1.getStringCellValue();

                rowR[Row] = DemoSheet.getRow(counter++);


                //LOOP THROUGH PENSIONABLE SALARY SHEET
                for (int y = 0; y < CountPSRow; y++) {
                    int temp2=y;

                    if(temp2==0){
                        y=1;
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
                    XSSFCell[] cellI1 = new XSSFCell[12];
                    double [] d = new double[12];

              //      for (int r=0;ReconRow)


                    if (a1Val_EM.equals(a1_EM) && !a1Val_EM.equals("ASSE88888")) {

                        for (int g=5,j=0;g<17;g++,j++){
                            cellI1[j] = PSRow.getCell(g);
                            F1_Yr = cellI1[j].getNumericCellValue();
                          //  d[j] = Double.parseDouble(F1_Yr);
                          d[j] = F1_Yr;
                        }


                        //CHECK IN RECON SHEET FOR EMPLOYEE CONTRIBUTION

                        for(int v=6; v<=CountReconRow;v++){
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


                            if(a1.equals(a1Val_EM) ){
                                double check = 0.05 *d[x];
                                check = Double.parseDouble(dF.format(check));
if (check!=h1) {
    //    ArrayList val = new ArrayList();

    System.out.println("Emplyee Number: " + a1);
    System.out.println("Last Name: " + Cb1);
    System.out.println("First Name: " + c1);
    System.out.println("Pensionable Salary: " + dF.format(d[x]));
    System.out.println("Employee Basic: " + h1);

    //   System.out.print(" I1: " + i1);
    //   System.out.print(" J1: " + j1);
    String test = "";

    stringBuilder.append("\nEmployee ID: " + a1+"\n");
    stringBuilder.append("Last Name: " + Cb1+"\n");
    stringBuilder.append("First Name: " + c1+"\n");
    stringBuilder.append("Pensionable Salary: " + dF.format(d[x]) +"\n");
    stringBuilder.append("Employee Basic: " + h1+"\n");

    if (check == h1) {
        test = "true";
        System.out.println("TEST: Pensionable Salary: " + dF.format(d[x]));
        System.out.println("Result: 5% of " +  dF.format(d[x] + " is " + check));
        stringBuilder.append("Result: 5% of " +  dF.format(d[x] + " is " + check+"\n"));
        System.out.println("Decision: " + test);

    } else {
        test = "false";

        System.out.println("Result: 5% of " +  dF.format(d[x])+ " is not " + h1+"\n");
        stringBuilder.append("Result: 5% of " +  dF.format(d[x])+ " is not " + h1+"\n");
        System.out.println("Decision: Please contact administrator");
    }

    System.out.println();


    stringBuilder.append("-------------------------------------------------------\n");
    stringBuilder.append("\nNotice: The Pensionable Check Process has now been completed.\n");
      }

                            }

                        }


                    }


                }

            }

         //   WriteAt+=8;
            StartYear++; //comment out years if COMMENTED
            counter=1;

        }// END OF LOOP YEARS
        return String.valueOf(stringBuilder);
    }

    public String Check_Plan_EntryDate_empDATE(String workindDir) throws IOException {
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("Notice: Results of the Plan Entry Check Process: \n");

        FileInputStream fileInputStream = new FileInputStream(workindDir + "\\Hose Valuation Data (Actuary's copy).xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet DemoSheet = workbook.getSheet("DEMO");

        int NoMembers = DemoSheet.getPhysicalNumberOfRows();

        for (int row=0;row<NoMembers;row++){
            int temp = row;

            if (temp == 0) {
                row= 1;  //start reading from second row
            }


            XSSFRow DemoRow = DemoSheet.getRow(row);


            XSSFCell cellA1_EM = DemoRow.getCell((short) 0);  //employee number
            if(cellA1_EM==null){
                cellA1_EM = DemoRow.createCell(0);
            }
            String result = cellA1_EM.getStringCellValue();
            String a1Val_EM = result.replaceAll("[-]","");

            XSSFCell cellC1_LN = DemoRow.getCell((short) 2);   //last name
            String b1Val_LN = cellC1_LN.getStringCellValue();

            XSSFCell cellD1_FN= DemoRow.getCell((short) 3);   //first name
            String c1Val_FN = cellD1_FN.getStringCellValue();

            XSSFCell cellG1_PE= DemoRow.getCell((short) 6);   //plan entry
            Date g1Val_PE = cellG1_PE.getDateCellValue();


            XSSFCell cellH1_PE= DemoRow.getCell((short) 7);   //emp date
            Date h1Val_PE = cellH1_PE.getDateCellValue();

            SimpleDateFormat dF = new SimpleDateFormat();
            dF.applyPattern("dd-MMM-yy");

            if(g1Val_PE.before(h1Val_PE)){
                System.out.println("Employee ID: " + a1Val_EM);
                System.out.println("Last Name: " + b1Val_LN);
                System.out.println("First Name: " + c1Val_FN);
                System.out.println("Plan Entry: " +dF.format(g1Val_PE));
                System.out.println("Employment Date: " + dF.format(h1Val_PE));
                System.out.println("Please Contact Administrator");

                stringBuilder.append("\nResult: This member cannot enter the plan before Employment date at "+dF.format(h1Val_PE)+"\n");
                stringBuilder.append("Employee ID: " + a1Val_EM+"\n");
                stringBuilder.append("Last Name: " + b1Val_LN+"\n");
                stringBuilder.append("First Name: " + c1Val_FN+"\n");
                stringBuilder.append("Plan Entry: " +dF.format(g1Val_PE)+"\n");
                stringBuilder.append("Employment Date: " + dF.format(h1Val_PE)+"\n");

                stringBuilder.append("----------------------------------------------------------------------\n");
                System.out.println();
            }


        }
        System.out.println("Notice: The Plan Entry Check Process has now been completed.");
        stringBuilder.append("\nNotice: The Plan Entry Check Process has now been completed.\n");

        return String.valueOf(stringBuilder);
    }


    public String Check_DateofBirth(String workindDir) throws IOException {
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("Notice: Results of the Date of Birth Check Process: \n");
        FileInputStream fileInputStream = new FileInputStream(workindDir + "\\Hose Valuation Data (Actuary's copy).xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet DemoSheet = workbook.getSheet("DEMO");

        int NoMembers = DemoSheet.getPhysicalNumberOfRows();

        for (int row=0;row<NoMembers;row++){
            int temp = row;

            if (temp == 0) {
                row= 1;  //start reading from second row
            }


            XSSFRow DemoRow = DemoSheet.getRow(row);


            XSSFCell cellA1_EM = DemoRow.getCell((short) 0);  //employee number
            if(cellA1_EM==null){
                cellA1_EM = DemoRow.createCell(0);
            }
            String result = cellA1_EM.getStringCellValue();
            String a1Val_EM = result.replaceAll("[-]","");

            XSSFCell cellC1_LN = DemoRow.getCell((short) 2);   //last name
            String b1Val_LN = cellC1_LN.getStringCellValue();
        //   b1Val_LN = result.replaceAll("[-]","");

            XSSFCell cellD1_FN= DemoRow.getCell((short) 3);   //first name
            String c1Val_FN = cellD1_FN.getStringCellValue();

            XSSFCell cellF1_DOB= DemoRow.getCell((short) 5);   //dob
            Date f1Val_DOB= cellF1_DOB.getDateCellValue();

            XSSFCell cellG1_PE= DemoRow.getCell((short) 6);   //plan entry
            Date g1Val_PE= cellG1_PE.getDateCellValue();


            XSSFCell cellH1_EM= DemoRow.getCell((short) 7);   //emp date
            Date h1Val_EM = cellH1_EM.getDateCellValue();

            XSSFCell cellI1_SD= DemoRow.getCell((short) 8);   //emp date
            Date i1Val_SD = cellI1_SD.getDateCellValue();

            SimpleDateFormat dF = new SimpleDateFormat();
            dF.applyPattern("dd-MMM-yy");

            if(f1Val_DOB.after(g1Val_PE) || f1Val_DOB.after(h1Val_EM)|| f1Val_DOB.after(h1Val_EM)|| f1Val_DOB.after(i1Val_SD)){
                System.out.println("Employee ID: " + a1Val_EM);
                System.out.println("Last Name: " + b1Val_LN);
                System.out.println("First Name: " + c1Val_FN);
                System.out.println("Date of Birth: " +dF.format(f1Val_DOB));
                System.out.println("Plan Entry: " +dF.format(g1Val_PE));
                System.out.println("Employment Date: " + dF.format(h1Val_EM));
                System.out.println("Status Date: " + dF.format(i1Val_SD));
                System.out.println("Decision: Contact the administrator");
                System.out.println();

                stringBuilder.append("\nResult: This member's date of birth must be before date of employment and enrollment to the plan"+"\n");
                stringBuilder.append("Employee ID: " + a1Val_EM+"\n");
                stringBuilder.append("Last Name: " + b1Val_LN+"\n");
                stringBuilder.append("First Name: " + c1Val_FN+"\n");
                stringBuilder.append("Date of Birth: " + dF.format(f1Val_DOB)+"\n");
                stringBuilder.append("Plan Entry: " + dF.format(g1Val_PE)+"\n");
                stringBuilder.append("Employment Date: " + dF.format(h1Val_EM)+"\n");
                stringBuilder.append("Status Date: " + dF.format(i1Val_SD)+"\n");
                stringBuilder.append("--------------------------------------------------------------------------------------------------------------\n");

            }


        }

        System.out.println("Notice: The Date of Birth Check Process has now been completed.");
        stringBuilder.append("\nNotice: The Date of Birth Check Process has now been completed.\n");

        return String.valueOf(stringBuilder);
    }

    public String Check_Age(String workingDir) throws IOException {

        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("Notice: Results of the Age Check Process:\n\n");
//        FileInputStream fileInputStream = new FileInputStream(filePathValData);
    //    FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Hose Valuation Data (Actuary's copy).xlsx");
        FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Hose Valuation Data (Actuary's copy).xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet DemoSheet = workbook.getSheet("DEMO");

        int NoMembers = DemoSheet.getPhysicalNumberOfRows();

        for (int row=0;row<NoMembers;row++){
            int temp = row;

            if (temp == 0) {
                row= 1;  //start reading from second row
            }


            XSSFRow DemoRow = DemoSheet.getRow(row);


            XSSFCell cellA1_EM = DemoRow.getCell((short) 0);  //employee number
            if(cellA1_EM==null){
                cellA1_EM = DemoRow.createCell(0);
            }
            String result = cellA1_EM.getStringCellValue();
            String a1Val_EM = result.replaceAll("[-]","");

            XSSFCell cellC1_LN = DemoRow.getCell((short) 2);   //last name
            String b1Val_LN = cellC1_LN.getStringCellValue();

            XSSFCell cellD1_FN= DemoRow.getCell((short) 3);   //first name
            String c1Val_FN = cellD1_FN.getStringCellValue();

            XSSFCell cellG1_DOB= DemoRow.getCell((short) 5);   //dob
            Date g1Val_DOB= cellG1_DOB.getDateCellValue();


            XSSFCell cellH1_PE= DemoRow.getCell((short) 7);   //emp date
            Date h1Val_PE = cellH1_PE.getDateCellValue();

            SimpleDateFormat dF = new SimpleDateFormat();
            dF.applyPattern("dd-MMM-yy");

            int age = utility.getAge(g1Val_DOB);

            if(age<15){
                System.out.println("Employee ID: " + a1Val_EM);
                System.out.println("Last Name: " + b1Val_LN);
                System.out.println("First Name: " + c1Val_FN);
                System.out.println("Date of Birth: " +dF.format(g1Val_DOB));
              //  System.out.println("Employment Date: " + dF.format(h1Val_PE));
                System.out.println("Result: This member age is "+age);
                System.out.println();

                //for gui
                stringBuilder.append("\nResult: This member's age is "+age +"\n");
                stringBuilder.append("Employee ID: " + a1Val_EM+"\n");
                stringBuilder.append("Last Name: " + b1Val_LN+"\n");
                stringBuilder.append("First Name: " + c1Val_FN+"\n");
                stringBuilder.append("DOB: " + dF.format(g1Val_DOB)+"\n");

                stringBuilder.append("-------------------------------------------------------\n");
            }


        }
        System.out.println("Notice: The Age Check Process has now been completed.");
        stringBuilder.append("\nNotice: The Age Check Process has now been completed.\n");

        return String.valueOf(stringBuilder);
    }

}

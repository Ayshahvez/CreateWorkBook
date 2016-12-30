import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.*;
//import java.lang.reflect.Array;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.*;

/**
 * Created by Ayshahvez on 11/10/2016.
 */

public class ExcelReader {

static Utility utility = new Utility();

    //NO FEES
    public String Create_Actives_Sheet(String workingDir) throws IndexOutOfBoundsException {
       StringBuilder stringBuilder = new StringBuilder();
        SimpleDateFormat dF1 = new SimpleDateFormat();
        dF1.applyPattern("dd-MMM-yy");
        try {
           // FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Seperated Members.xlsx");
            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Seperated Members.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet worksheet = workbook.getSheet("Actives");
            //   XSSFSheet sheet = workbook.getSheetAt(0);

            int rowCount=worksheet.getPhysicalNumberOfRows();
           // System.out.println(rowCount);


            int num = rowCount;
            //  int noOfColumns = sheet.getRow(num).getLastCellNum();

            FileInputStream fileR = new FileInputStream(workingDir + "\\Template_Active_Sheet.xlsx");
            XSSFWorkbook workbookR = new XSSFWorkbook(fileR);
            XSSFSheet sheetR = workbookR.getSheetAt(0);

            SimpleDateFormat datetemp = new SimpleDateFormat("dd-MMM-yy");
            // Date cellValue = datetemp.parse("1994-01-01");

            XSSFRow[] rowR = new XSSFRow[num];
            Cell cellR = null;

            ArrayList arraylist = new ArrayList();
            ArrayList PensionableSalary = new ArrayList();
            int counter = 7;
            int index =0;
           // int TermineesStartRow = 1;

            for (int row = 0; row < num; row++) {
                int Row = row;

                if (Row == 0) {
                    row = 1;  //start reading from second row
                }
                XSSFRow row1 = worksheet.getRow(row);

                XSSFCell cellA1 = row1.getCell((short) 0);  //employee number
                String result = cellA1.getStringCellValue();
                String a1Val = result.replaceAll("[-]", "");

                XSSFCell cellB1 = row1.getCell((short) 1);
                String b1Val = cellB1.getStringCellValue();

                XSSFCell cellC1 = row1.getCell((short) 2);    //last name
                String c1Val = cellC1.getStringCellValue();

                XSSFCell cellD1 = row1.getCell((short) 3); //first name
                String d1Val = cellD1.getStringCellValue();

                XSSFCell cellE1 = row1.getCell((short) 4);  //sex
                String e1Val = cellE1.getStringCellValue();

                XSSFCell cellF1 = row1.getCell((short) 5); //dob
                //  Date f1Vall = cellF1.getDateCellValue();
                //String f1Val =datetemp.format( cellF1.getDateCellValue());
                String f1Val = cellF1.getStringCellValue();


                XSSFCell cellG1 = row1.getCell((short) 6); //plan entry
                // String g1Val =datetemp.format( cellG1.getDateCellValue());
                String g1Val = cellG1.getStringCellValue();

                XSSFCell cellH1 = row1.getCell((short) 7); //emp date
                //    String h1Val = datetemp.format( cellH1.getDateCellValue());
                String h1Val = cellH1.getStringCellValue();

                XSSFCell cellI1 = row1.getCell((short) 8);  //status date
                // String i1Val = datetemp.format( cellI1.getDateCellValue());
                String i1Val = cellI1.getStringCellValue();

                XSSFCell cellJ1 = row1.getCell((short) 9); //status
                String j1Val = cellJ1.getStringCellValue();

                if (j1Val.equals("ACTIVE") && !(d1Val.equals("KEY"))) {
                //    System.out.print("A1: " + a1Val);
                    //     System.out.print(" B1: " + b1Val);
            /*        System.out.print(" C1: " + c1Val);
                    System.out.print(" D1: " + d1Val);
                    System.out.print(" E1: " + e1Val);
                    System.out.print(" F1: " + f1Val);
                    System.out.print(" G1: " + g1Val);
                    System.out.print(" H1: " + h1Val);
                    System.out.print(" I1: " + i1Val);
                    System.out.print(" J1: " + j1Val);*/

                    stringBuilder.append("Employee ID: " + a1Val+"\n");
                    stringBuilder.append("Last Name: " + c1Val+"\n");
                    stringBuilder.append("First Name: " + d1Val+"\n");
                    stringBuilder.append("Date of Birth: " + f1Val+"\n");
                    stringBuilder.append("Status Date: "+i1Val+"\n");
                    stringBuilder.append("Status: "+j1Val+"\n");
                    stringBuilder.append("-------------------------------------------------------\n");


                    System.out.println();

                    arraylist.add(0, a1Val);  //employee number
                    arraylist.add(1, c1Val);   //LAST NAME
                    arraylist.add(2, d1Val);
                    arraylist.add(3, e1Val);
                    arraylist.add(4, " ");
                    arraylist.add(5, f1Val);
                    arraylist.add(6, h1Val);
                    arraylist.add(7, g1Val);
                    //  arraylist.add(8, i1Val);
                    // arraylist.add(9, j1Val);


                    rowR[Row] = sheetR.createRow(counter++);
                    // }
                    for (int Col = 0; Col < 8; Col++) {
                        //Update the value of cell
                        //   cellR = rowR[Row].getCell(Col);
                        //   if (cellR == null) {
                        cellR = rowR[Row].createCell(Col);
                        //   }
                        cellR.setCellValue(String.valueOf(arraylist.get(Col)));
                    }
                    //PENSIONABLE SALARY*****************************
                    FileInputStream F = new FileInputStream(workingDir + "\\Pensionable Salary - HAS.xlsx");  //pensionable salary sheet
                    XSSFWorkbook workbookF = new XSSFWorkbook(F);
                    XSSFSheet sheetF = workbookF.getSheetAt(0);
                    int PensionableRow = sheetF.getPhysicalNumberOfRows();
                    DecimalFormat dF = new DecimalFormat("#.##");//#.##

                    for (int y = 0; y <= PensionableRow; y++) {
                        int temp = y;

                        if (temp == 0) {
                            y = 1;  //start reading from second row
                        }
                        XSSFRow rowF = sheetF.getRow(y);

                        XSSFCell cellAA1 = rowF.getCell(0);  //employee number
                        String aa1Val = cellAA1.getStringCellValue();

                        XSSFCell cellBB1 = rowF.getCell((short) 1);  //Last name
                        String bb1Val = cellBB1.getStringCellValue();

                        XSSFCell cellCC1 = rowF.getCell((short) 2);  //First Name
                        String cc1Val = cellCC1.getStringCellValue();

                        XSSFCell cellDD1 = rowF.getCell((short) 5);
                        Double dd1Val = cellDD1.getNumericCellValue();

                        XSSFCell cellEE1 = rowF.getCell((short) 6);
                        Double ee1Val = cellEE1.getNumericCellValue();

                        XSSFCell cellFF1 = rowF.getCell((short) 7);
                        Double ff1Val = cellFF1.getNumericCellValue();

                        XSSFCell cellGG1 = rowF.getCell((short) 8);
                        Double gg1Val = cellGG1.getNumericCellValue();

                        XSSFCell cellHH1 = rowF.getCell((short) 9);
                        Double hh1Val = cellHH1.getNumericCellValue();

                        XSSFCell cellII1 = rowF.getCell((short) 10);
                        Double ii1Val = cellII1.getNumericCellValue();

                        XSSFCell cellJJ1 = rowF.getCell((short) 11);
                        Double jj1Val = cellJJ1.getNumericCellValue();

                        XSSFCell cellKK1 = rowF.getCell((short) 12);
                        Double kk1Val = cellKK1.getNumericCellValue();

                        XSSFCell cellLL1 = rowF.getCell((short) 13);
                        Double ll1Val = cellLL1.getNumericCellValue();

                        XSSFCell cellMM1 = rowF.getCell((short) 14);
                        if (cellMM1 == null) {
                            cellMM1 = rowF.createCell(14);
                        }
                        Double mm1Val = cellMM1.getNumericCellValue();

                        XSSFCell cellNN1 = rowF.getCell((short) 15);
                        if (cellNN1 == null) {
                            cellNN1 = rowF.createCell(15);
                        }
                        Double nn1Val = cellNN1.getNumericCellValue();


                        XSSFCell cellOO1 = rowF.getCell((short) 16);
                        if (cellOO1 == null) {
                            cellOO1 = rowF.createCell(16);
                        }
                        Double oo1Val = cellOO1.getNumericCellValue();

                        //   XSSFCell cellMM1 = rowF.getCell((short) 17);
                        //   Double mm1Val = cellMM1.getNumericCellValue();

                        PensionableSalary.add(0, dd1Val);
                        PensionableSalary.add(1, ee1Val);
                        PensionableSalary.add(2, ff1Val);
                        PensionableSalary.add(3, gg1Val);
                        PensionableSalary.add(4, hh1Val);
                        PensionableSalary.add(5, ii1Val);
                        PensionableSalary.add(6, jj1Val);
                        PensionableSalary.add(7, kk1Val);
                        PensionableSalary.add(8, ll1Val);
                        PensionableSalary.add(9, mm1Val);
                        PensionableSalary.add(10, nn1Val);
                        PensionableSalary.add(11, oo1Val);

                        if (aa1Val.equals(a1Val)) {

                            for (int k = 8, l = 0; l < 12; k++, l++) {

                                cellR = rowR[Row].createCell(k);
                                cellR.setCellValue(String.valueOf(dF.format(PensionableSalary.get(l))));
                            }
                            break;
                        }
                    }

                }

                //*************************MEMBERS AGE AS AT**********************
                String dateString1 = f1Val;
                Date date = null;
                String dateString2 = null;
                try {
                    date = new SimpleDateFormat("dd-MMM-yy").parse(dateString1);
                    dateString2 = new SimpleDateFormat("dd-MM-yyyy").format(date);
                } catch (ParseException e) {
                    e.printStackTrace();
                }

                String str[] = dateString2.split("-");
                int day = Integer.parseInt(str[0]);
                int month = Integer.parseInt(str[1]);
                int year = Integer.parseInt(str[2]);

                LocalDate birthDate = LocalDate.of(year, month, day);
                LocalDate endDate = LocalDate.of(2015, 12, 31);

                cellR = rowR[Row].createCell(20);
                cellR.setCellValue(Utility.calculateAge(birthDate, endDate));

//********PENSIONABLE SERVICE***********************
                String dateString = g1Val;
                Date date1 = null;
                String dateString22 = null;
                try {
                    date1 = new SimpleDateFormat("dd-MMM-yy").parse(dateString);
                    dateString22 = new SimpleDateFormat("dd-MM-yyyy").format(date1);
                } catch (ParseException e) {
                    e.printStackTrace();
                }

                 //  System.out.println(dateString22); // 2011-04-16

                String str1[] = dateString22.split("-");
                int day1 = Integer.parseInt(str1[0]);
                int month1 = Integer.parseInt(str1[1]);
                int year1 = Integer.parseInt(str1[2]);
                DateFormat df = new SimpleDateFormat("yyyy.MM.dd");

                Date startDate1 = null;
                Date endDate1 = null;
                try {
                    startDate1 = df.parse(year1 + "." + month1 + "." + day1);
                    endDate1 = df.parse(2015 + "." + 12 + "." + 31);
                } catch (ParseException e) {
                    e.printStackTrace();
                }
                DecimalFormat dF = new DecimalFormat("#.##");//#.##

                cellR = rowR[Row].createCell(21);
                cellR.setCellValue(dF.format((Utility.betweenDates(startDate1, endDate1)/365.25)));

            }
            FileOutputStream outFile = new FileOutputStream(new File(workingDir +"\\Actives_Sheet.xlsx"));
            workbookR.write(outFile);
            fileR.close();
            outFile.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        catch(NullPointerException e){
            e.printStackTrace();

        }
        return String.valueOf(stringBuilder);
    }// end of create active sheet

    public void Create_Activee_Contribution(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {
        /*
        FileInputStream fileR = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Actives_Sheet.xlsx");
        XSSFWorkbook workbookR = new XSSFWorkbook(fileR);
        XSSFSheet CopyFromSheet = workbookR.getSheetAt(0);
        */

        DecimalFormat dF = new DecimalFormat("#.##");//#.##

        //OPEN ACTIVE SHEET
        FileInputStream fileInputStream = new FileInputStream(workingDir +"\\Actives_Sheet.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet ActiveSheet = workbook.getSheet("Actives");

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

        // System.out.println(StartYear + "." + StartMonth + "." + StartDay);
        //  System.out.println(EndYear + "." + EndMonth + "." + EndDay);

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

        int years = Utility.getDiffYears(startDate, endDate);

        int counter=7;
        for (int x = 0; x <= years; x++) {

            Calendar cal = Calendar.getInstance();
            cal.set(StartYear, StartMonth, StartDay);
            SimpleDateFormat sdf = new SimpleDateFormat("yy"); // Just the year, with 2 digits
            String formattedDate = sdf.format(cal.getTime());

            // String StartDate = StartDay + "-" + StartMonth + "-" + StartYear; //"01-01-05";
            // String EndDate = EndDay + "-" + EndMonth + "-" + StartYear;//"31-12-05";
            String Recon = ("Recon " + formattedDate);


            FileInputStream fs = null;
            try {
                fs = new FileInputStream(workingDir+"\\Valuation Data.xlsx");
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            XSSFWorkbook WB = new XSSFWorkbook(fs);

//GET RECON SHEET
            XSSFSheet Reconsheet = WB.getSheet(Recon);
            System.out.println(formattedDate);
            int CountReconRow = Reconsheet.getPhysicalNumberOfRows();



            int numOfActives = ActiveSheet.getLastRowNum()+1; // get the last row number
            // System.out.println(rowCount);
            //     System.out.println("Recon"+ Recon + "count"+CountReconRow +"Actives" +numOfActives);
            XSSFRow[] rowR = new XSSFRow[numOfActives];
            Cell cellR = null;

/*if(Recon.equals(15)){
    numOfActives=35;
}*/
            int Crow = 8;
            for (int row = 7, I=8; row < numOfActives; row++,I++) {

                int Row = row;

                XSSFRow ActiveRow = ActiveSheet.getRow(row);

                XSSFCell cellA1 = ActiveRow.getCell((short) 0);  //employee number
                String result = cellA1.getStringCellValue();
                String a1Val = result.replaceAll("[-]","");

                XSSFCell cellB1 = ActiveRow.getCell((short) 1);   //last name
                String b1Val = cellB1.getStringCellValue();

                XSSFCell cellC1 = ActiveRow.getCell((short) 2);   //first name
                String c1Val = cellC1.getStringCellValue();

                XSSFCell cellD1 = ActiveRow.getCell((short) 3);
                String d1Val = cellD1.getStringCellValue();

                String i1Val = null;
                XSSFCell[] cellI1 = new XSSFCell[12];
                double [] d = new double[12];
                for (int g=8,j=0;g<20;g++,j++){
                    cellI1[j] = ActiveRow.getCell(g);
                    i1Val = cellI1[j].getStringCellValue();
                    d[j] = Double.parseDouble(i1Val);
                }


                //    Crow++;

                rowR[Row] = ActiveSheet.getRow(counter++);

                for (int y = 6; y < CountReconRow; y++) {

                    XSSFRow reconRow = Reconsheet.getRow(y);

                    XSSFCell A1 = reconRow.getCell(0);  //employee number
                    if (A1 == null) {
                        A1 = reconRow.createCell(0);
                    }
                    String a1 = A1.getStringCellValue();


                    XSSFCell B1 = reconRow.getCell(1);  //FNAME
                    if (B1 == null) {
                        B1 = reconRow.createCell(1);

                    }
                    String b1 = B1.getStringCellValue();

                    XSSFCell C1 = reconRow.getCell(2);  //LNAME
                    if (C1 == null) {
                        C1 = reconRow.createCell(2);

                    }
                    String c1 = C1.getStringCellValue();


                    XSSFCell H1 = reconRow.getCell((short) 7);
                    //Update the value of cell
                    if (H1 == null) {
                        H1 = reconRow.createCell(7);
                    }
                    double h1 = H1.getNumericCellValue();

                    XSSFCell I1 = reconRow.getCell((short) 8);
                    //Update the value of cell
                    if (I1 == null) {
                        I1 = reconRow.createCell(8);
                    }
                    double i1 = I1.getNumericCellValue();

                    XSSFCell J1 = reconRow.getCell((short) 9);
                    //Update the value of cell
                    if (J1 == null) {
                        J1 = reconRow.createCell(9);
                    }
                    double j1 = J1.getNumericCellValue();

                    if (a1.equals(a1Val) ) {

                        ArrayList val = new ArrayList();

                        System.out.print("A1: " + a1);
                        System.out.print(" D1: " + b1);
                        System.out.print(" C1: " + c1);
                        System.out.print("PS: " + i1Val);
                        System.out.print(" H1: " + h1);
                        System.out.print(" I1: " + i1);
                        System.out.print(" J1: " + j1);
                        String test="";
                        double check = 0.05 *d[x];
                        check= Double.parseDouble(dF.format(check));
                        if (check == h1) {
                            test = "true";

                        } else {
                            test = "false";
                        }
                        System.out.print(" TEST: Pensionable Salary"+d[x]+"result "+check+test);

                        System.out.println();

                        val.add(0, h1);
                        val.add(1, i1);
                        val.add(2, j1);
                        val.add(3, 0.00);


                        for (int b = 0; b < val.size(); b++) {
                            cellR = rowR[Row].createCell(WriteAt + b);
                            cellR.setCellValue((Double)val.get(b));
                        }
                        break;

                    }

                }



            }

            WriteAt+=8;
            StartYear++; //comment out years if COMMENTED
            counter=7;

        }// END OF LOOP YEARS

        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Updated_Actives_Sheet.xlsx"));
        workbook.write(outFile);
        fileInputStream.close();
        outFile.close();
    }

    public void Create_Active_Acc_Balances(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {
        DecimalFormat dF = new DecimalFormat("#.##");//#.##
        int readCol=26;//start to read from Col 26, which is starting contribution column
        int YearCol = readCol;//as we are in the same year, we should always be reading from the correct column
        int Write_Coloumn =30;//start to write at column 30 which are the accumullated cells

        //OPEN ACTIVE SHEET
        FileInputStream fileInputStream = new FileInputStream(workingDir +"\\Updated_Actives_Sheet.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet ActiveSheet = workbook.getSheet("Actives");

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


        ArrayList<Double> val = new ArrayList();
        int years = Utility.getDiffYears(startDate, endDate);//gets number of years
        years+=1;
        int numOfActives = ActiveSheet.getLastRowNum()+1;//gets number of active members
        //get the interest rates values
        double[] interestValues = new double [years];
        interestValues= getInterestRates(workingDir,years);


        //variables to hold calculated new Accumulated balances
        double[] newAccEmployeeBalance= new double[numOfActives];
        double[] newAccEmployeeOptional= new double[numOfActives];
        double[] newAccEmployerRequired=new double[numOfActives];
        double[] newAccEmployerOptional=new double[numOfActives];

//run for the appropiate number of years
        double CellAccEmployeeBasic=0;
        double CellAccEmployeeOptional=0;
        double CellAccEmployerRequired=0;
        double CellAccEmployerOptional=0;

        //get the initial Accumulated Balances
        double [] CellAccEmployeeBasic0 = new double[numOfActives];
        double[] CellAccEmployeeOptional0 = new double[numOfActives];
        double[] CellAccEmployerRequired0 = new double[numOfActives];
        double[] CellAccEmployerOptional0 = new double[numOfActives];


        for(int row=7,I=0;row<numOfActives;row++,I++){ //get the initial accumulated balances

            XSSFRow ActiveRow = ActiveSheet.getRow(row);

            XSSFCell[] cellAccEmployeeBasic0 = new XSSFCell[numOfActives];
            cellAccEmployeeBasic0[I] = ActiveRow.getCell( 22);  //employee number
            if(cellAccEmployeeBasic0[I]==null){
                cellAccEmployeeBasic0[I] = ActiveRow.createCell(22);
                cellAccEmployeeBasic0[I].setCellValue(0);
            }
            CellAccEmployeeBasic0[I]= cellAccEmployeeBasic0[I].getNumericCellValue();

            XSSFCell[] cellAccEmployeeOptional0 = new XSSFCell[numOfActives];
            cellAccEmployeeOptional0[I] = ActiveRow.getCell(23);  //employee number
            if(cellAccEmployeeOptional0[I]==null){
                cellAccEmployeeOptional0[I] = ActiveRow.createCell(23);
                cellAccEmployeeOptional0[I].setCellValue(0);
            }
            CellAccEmployeeOptional0[I]= cellAccEmployeeOptional0[I].getNumericCellValue();


            XSSFCell[] cellAccEmployerRequired0 = new XSSFCell[numOfActives];
            cellAccEmployerRequired0[I] = ActiveRow.getCell(24); //employee number
            if(cellAccEmployerRequired0[I]==null){
                cellAccEmployerRequired0[I] = ActiveRow.createCell(24);
                cellAccEmployerRequired0[I].setCellValue(0);
            }
            CellAccEmployerRequired0[I] = cellAccEmployerRequired0[I].getNumericCellValue();

            XSSFCell[] cellAccEmployerOptional0 = new XSSFCell[numOfActives];
            cellAccEmployerOptional0[I] = ActiveRow.getCell(25); //employee number
            if(cellAccEmployerOptional0[I]==null){
                cellAccEmployerOptional0[I] = ActiveRow.createCell(25);
                cellAccEmployerOptional0[I].setCellValue(0);
            }
            CellAccEmployerOptional0[I] = cellAccEmployerOptional0[I].getNumericCellValue();
        }//end of loop to get inital accumulated balances
        XSSFRow ActiveRow = null;
        Cell cellR;
        //MAIN PROCESSING-to get acc Balances and cont Balances
        for (int x = 0; x < years; x++) { //run for the appropiate number of years
            //  System.out.println(numOfActives);


            for (int row = 7, I = 0; row < numOfActives; row++, I++) { //run for appropiate number of total active members
                readCol=YearCol;//need to ensure to start reading from same Column in same year

                ActiveRow = ActiveSheet.getRow(row);

                if (x == 0) {//get the accumulated balances just for the 1st year
                    CellAccEmployeeBasic = CellAccEmployeeBasic0[I];
                    CellAccEmployeeOptional = CellAccEmployeeOptional0[I];
                    CellAccEmployerRequired = CellAccEmployerRequired0[I];
                    CellAccEmployerOptional = CellAccEmployerOptional0[I];
                }
                else{ //get accumulated balances for every year after 1st year
                    CellAccEmployeeBasic =  newAccEmployeeBalance[I];
                    CellAccEmployeeOptional = newAccEmployeeOptional[I];
                    CellAccEmployerRequired = newAccEmployerRequired[I];
                    CellAccEmployerOptional = newAccEmployerOptional[I];
                }

                //get the CONTRIBUTIONS starting from column 26
                XSSFCell cellConEmployeeBasic = ActiveRow.getCell(readCol);
                if (cellConEmployeeBasic == null) {
                    cellConEmployeeBasic = ActiveRow.createCell(readCol);
                    cellConEmployeeBasic.setCellValue(0);
                }
                double CellConEmployeeBasic = cellConEmployeeBasic.getNumericCellValue();
                readCol+=1;

                XSSFCell cellConEmployeeOptional = ActiveRow.getCell(readCol);
                if (cellConEmployeeOptional == null) {
                    cellConEmployeeOptional = ActiveRow.createCell(readCol);
                    cellConEmployeeOptional.setCellValue(0);
                }
                double CellConEmployeeOptional = cellConEmployeeOptional.getNumericCellValue();
                readCol+=1;


                XSSFCell cellConEmployerRequired = ActiveRow.getCell(readCol);
                if (cellConEmployerRequired == null) {
                    cellConEmployerRequired = ActiveRow.createCell(readCol);
                    cellConEmployerRequired.setCellValue(0);
                }
                double CellConEmployerRequired = cellConEmployerRequired.getNumericCellValue();
                readCol+=1;


                XSSFCell cellConEmployerOptional = ActiveRow.getCell(readCol);
                if (cellConEmployerOptional == null) {
                    cellConEmployerOptional = ActiveRow.createCell(readCol);
                    cellConEmployerOptional.setCellValue(0);
                }
                double CellConEmployerOptional = cellConEmployerOptional.getNumericCellValue();
                readCol+=1;

/*

                XSSFCell cellFees = ActiveRow.getCell(readCol);
                if (cellFees == null) {
                    cellFees = ActiveRow.createCell(readCol);
                    cellFees.setCellValue(0);
                }
                double CellFees = cellFees.getNumericCellValue();*/


                int year = 2004;
                year+=x;
                System.out.println();
                System.out.println("Year: " +year+"Row: "+row);
                System.out.println("Acc "+ CellAccEmployeeBasic +"Con"+CellConEmployeeBasic);
                System.out.println("Acc "+ CellAccEmployeeOptional +"Con"+CellConEmployeeOptional);
                System.out.println("Acc "+ CellAccEmployerRequired +"Con"+CellConEmployerRequired);
                System.out.println("Acc "+ CellAccEmployerOptional +"Con"+CellConEmployerOptional);
                System.out.println(interestValues[x]);
                //FORMULA CALCULATIONS
                newAccEmployeeBalance[I] = ((CellAccEmployeeBasic * (1+interestValues[x])) + (CellConEmployeeBasic * (1+(interestValues[x]*0.5))));//CellAccEmployeeBasic * (1 + 1) + CellConEmployeeBasic * (1 + 1 * 0.5);
                newAccEmployeeOptional[I] =((CellAccEmployeeOptional * (1+interestValues[x])) + (CellConEmployeeOptional * (1+(interestValues[x]*0.5))));//CellAccEmployeeOptional +  CellConEmployeeOptional; CellAccEmployeeOptional * (1 + 1) + CellConEmployeeOptional * (1 + 1 * 0.5);
                newAccEmployerRequired[I] = ((CellAccEmployerRequired * (1+interestValues[x])) + (CellConEmployerRequired *(1+(interestValues[x]*0.5))));//CellAccEmployerRequired * (1 + 1) + CellConEmployerRequired * (1 + 1 * 0.5) + CellFees * (1 + 1 * 0.5);
                newAccEmployerOptional[I] =((CellAccEmployerOptional * (1+interestValues[x])) + (CellConEmployerOptional * (1+(interestValues[x]*0.5))));//CellAccEmployerOptional * (1 + 1) + CellConEmployerOptional * (1 + 1 * 0.5);

                val.add(0, newAccEmployeeBalance[I]);
                val.add(1, newAccEmployeeOptional[I]);
                val.add(2, newAccEmployerRequired[I]);
                val.add(3, newAccEmployerOptional[I]);

//write the calculated accumulated balances to the sheet; start to write at column 31
                for (int b = 0; b < 4; b++) {
                    cellR = ActiveRow.createCell(Write_Coloumn + b);
                    cellR.setCellValue(dF.format(val.get(b)));
                }//end of loop

            }//end of looping through each member

            //MOVING THE INDEXES
            YearCol=readCol+4;//move over by 4 columns to get next set of contributions for next year
            readCol=YearCol;//give the readCol 5 so that, it can always read the correct set of contributions of that same year
            Write_Coloumn += 8;//8 no fees  || 9 fees
            StartYear++;//increment Start year by one until we reach end year

            //when we are at end of year...we should write account balance as at end date
            if(x==(years-1)) {
                for (int I = 0, row=7; row < numOfActives; I++, row++) {
                    ActiveRow=ActiveSheet.getRow(row);
                    cellR = ActiveRow.createCell(readCol);
                    cellR.setCellValue(  newAccEmployeeBalance[I] + newAccEmployeeOptional[I] + newAccEmployerRequired[I]+newAccEmployerOptional[I]);
                }
            }
        }// END OF LOOP YEARS

        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Accumulated_Actives_Sheet.xlsx"));
        workbook.write(outFile);
        fileInputStream.close();
        outFile.close();
    }//end of function Create Active Accumulated Balances

    public String Create_Terminee_Sheet(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IndexOutOfBoundsException {
        StringBuilder stringBuilder = new StringBuilder();

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
 /*
        int EndYear = 2015;
        int EndMonth = 12;
        int EndDay = 31;

        int StartYear = 2004;
        int StartMonth = 01;
        int StartDay = 01;
*/

        DateFormat df2 = new SimpleDateFormat("yyyy.MM.dd");
        Date startDate = null;
        Date endDatee = null;
        try {
            startDate = df2.parse(StartYear + "." + StartMonth + "." + StartDay);
            endDatee = df2.parse(EndYear + "." + EndMonth + "." + EndDay);
        }catch(ParseException e){
e.printStackTrace();
        }

        DecimalFormat dF = new DecimalFormat("#.##");//#.##
        DateFormat df = new SimpleDateFormat("dd-MMM-yy");
        SimpleDateFormat datetemp = new SimpleDateFormat("dd-MMM-yy");
        SimpleDateFormat datetemp2 = new SimpleDateFormat("dd-MM-yy");
        Date beginDate = null;
        Date endDate = null;

        try {
            FileInputStream fileInputStream = new FileInputStream(workingDir+"\\Seperated Members.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet worksheet = workbook.getSheet("Terminees");
            //   XSSFSheet sheet = workbook.getSheetAt(0);

            int rowCount=worksheet.getPhysicalNumberOfRows();
            // System.out.println(rowCount);

            int num = rowCount;
            //  int noOfColumns = sheet.getRow(num).getLastCellNum();

            FileInputStream fileR = new FileInputStream(workingDir+"\\Template_Terminee_Sheet.xlsx");
            XSSFWorkbook workbookR = new XSSFWorkbook(fileR);
            XSSFSheet sheetR = workbookR.getSheetAt(0);


            // Date cellValue = datetemp.parse("1994-01-01");

            XSSFRow[] rowR = new XSSFRow[num];
            Cell cellR = null;

            ArrayList arraylist = new ArrayList();
            ArrayList PensionableSalary = new ArrayList();
            int counter = 7;
            int index =0;

         int years= Utility.getAge(startDate,endDatee);
         years+=1;//always add 1 to computer language

           for(int u=0; u<years;u++) {
               try {
                   beginDate = datetemp2.parse(StartDay+"-"+StartMonth+"-"+StartYear);//"01-Jan-05"
                   endDate = datetemp.parse(Utility.getEndDate(StartYear,01,01));
               } catch (ParseException e) {
                   e.printStackTrace();
               }

               for (int row = 0; row < num; row++) {

                   int Row = row;

                   if (Row == 0) {
                       row = 1;  //start reading from second row
                   }
                   XSSFRow row1 = worksheet.getRow(row);

                   XSSFCell cellA1 = row1.getCell((short) 0);  //employee number
                   String result = cellA1.getStringCellValue();
                   String a1Val = result.replaceAll("[-]", "");

                   XSSFCell cellB1 = row1.getCell((short) 1);
                   String b1Val = cellB1.getStringCellValue();

                   XSSFCell cellC1 = row1.getCell((short) 2);    //last name
                   String c1Val = cellC1.getStringCellValue();

                   XSSFCell cellD1 = row1.getCell((short) 3); //first name
                   String d1Val = cellD1.getStringCellValue();

                   XSSFCell cellE1 = row1.getCell((short) 4);  //sex
                   String e1Val = cellE1.getStringCellValue();

                   XSSFCell cellF1 = row1.getCell((short) 5); //dob
                   //  Date f1Vall = cellF1.getDateCellValue();
                   //String f1Val =datetemp.format( cellF1.getDateCellValue());
                   String f1Val = cellF1.getStringCellValue();


                   XSSFCell cellG1 = row1.getCell((short) 6); //plan entry
                   // String g1Val =datetemp.format( cellG1.getDateCellValue());
                   String g1Val = cellG1.getStringCellValue();

                   XSSFCell cellH1 = row1.getCell((short) 7); //emp date
                   //    String h1Val = datetemp.format( cellH1.getDateCellValue());
                   String h1Val = cellH1.getStringCellValue();

                   XSSFCell cellI1 = row1.getCell((short) 8);  //status date

//                    String i1Val = datetemp.format( cellI1.getDateCellValue());
                   String i1Val = cellI1.getStringCellValue();

                   XSSFCell cellJ1 = row1.getCell((short) 9); //status
                   String j1Val = cellJ1.getStringCellValue();

                   Date k1Val = null;
                 //  try {
                       XSSFCell cellK1 = row1.getCell((short) 10); //date of refund
                  // if (cellK1 == null){
                 //      cellK1 = row1.createCell(10);
                    //  cellK1.setCellValue(11/11/1111);
                 //  }

                       k1Val = cellK1.getDateCellValue();
                  // }catch (NullPointerException e){

                  // }
                   if (j1Val.equals("DEATH") ||j1Val.equals("RETIREMENT") || j1Val.equals("TERMINATED") || j1Val.equals("DEFERRED") && !(d1Val.equals("KEY"))) {
                       //    System.out.print("A1: " + a1Val);
                       //    System.out.print(" B1: " + b1Val);
                /*   System.out.print(" C1: " + c1Val);
                    System.out.print(" D1: " + d1Val);
                    System.out.print(" E1: " + e1Val);
                    System.out.print(" F1: " + f1Val);
                    System.out.print(" G1: " + g1Val);
                    System.out.print(" H1: " + h1Val);
                    System.out.print(" I1: " + i1Val);
                    System.out.print(" J1: " + j1Val);
*/

                       stringBuilder.append("Employee ID: " + a1Val+"\n");
                       stringBuilder.append("Last Name: " + c1Val+"\n");
                       stringBuilder.append("First Name: " + d1Val+"\n");
                       stringBuilder.append("Date of Birth: " + f1Val+"\n");
                       stringBuilder.append("Status Date: "+i1Val+"\n");
                       stringBuilder.append("Status: "+j1Val+"\n");
                       stringBuilder.append("-------------------------------------------------------\n");
                       //      System.out.println();

                       Date statusDate = null;
                       try {
                           statusDate = datetemp.parse(i1Val);
                       } catch (ParseException e) {
                           e.printStackTrace();
                       }


                       if(Row==0){
                           System.out.println(StartYear+ " Terminations");
                          // rowR[Row] = sheetR.getRow(counter);
                        //   cellR.setCellValue(StartYear+ " Terminations");
                       //   if(u>0){
                            XSSFRow r = sheetR.createRow(counter);
                           Cell cell= r.createCell(0);
                           cell.setCellValue(StartYear+ " Terminations");
                              counter+=1;
                      //    }
                       }
                       //    System.out.print("status: " + statusDate + " beginDate " + beginDate + " endDate" + endDate);
                       if (statusDate.after(beginDate) && statusDate.before(endDate)) {

                           rowR[Row] = sheetR.createRow(counter++);
                           //  System.out.println("status: " + statusDate + " beginDate " + beginDate + " endDate" + endDate);
                           System.out.print(" A1: " + a1Val);
                           System.out.print(" C1: " + c1Val);
                           System.out.print(" D1: " + d1Val);
                           System.out.print(" E1: " + e1Val);
                           System.out.print(" F1: " + f1Val);
                           System.out.print(" G1: " + g1Val);
                           System.out.print(" H1: " + h1Val);
                           System.out.print(" I1: " + i1Val);
                           System.out.print(" J1: " + j1Val);
                         // System.out.print(" K1: " + datetemp.format(k1Val));
                           System.out.println();


                           //DATE PROCESSING

                         //  String dateString = i1Val;
                         //  Date date1 = null;
                           String PlanEntry = g1Val;
                           Date date2 = null;
                           try {
                           //    date1 = new SimpleDateFormat("dd-MMM-yy").parse(dateString);
                               date2 = new SimpleDateFormat("dd-MMM-yy").parse(PlanEntry);
                           } catch (ParseException e) {
                               e.printStackTrace();
                           }

                           String str[] = i1Val.split("-");
                           int year = Integer.parseInt(str[2]);
                           String j = str[2];

                           String str22[] = g1Val.split("-");
                           //int year = Integer.parseInt(str[2]);
                           String e = str22[2];

                           //
String J= "01-Jan-"+j; //start of plan year of termination
String K ="31-Dec-"+j;//end of plan year of termination
String L = "31-Dec-"+e;//end of plan year of enrolment


                           Date dateJ = null;  //start of plan year of termination
                           Date dateK = null;//end of plan year of termination
                           Date dateL = null;//end of plan year of enrolment
                           try {
                                   dateJ = new SimpleDateFormat("dd-MMM-yy").parse(J);
                               dateK = new SimpleDateFormat("dd-MMM-yy").parse(K);
                               dateL = new SimpleDateFormat("dd-MMM-yy").parse(L);


                           } catch (ParseException E) {
                               E.printStackTrace();
                           }


                           arraylist.add(0, a1Val);  //employee number
                           arraylist.add(1, c1Val);   //LAST NAME
                           arraylist.add(2, d1Val);//FN
                           arraylist.add(3, e1Val);//sex
                           arraylist.add(4, j1Val);//type of term
                           arraylist.add(5, f1Val);//dob
                           arraylist.add(6, h1Val); //doh
                           arraylist.add(7, g1Val);//doe
                           arraylist.add(8, i1Val);//DOT
                           arraylist.add(9, J); //start of plan year of termination
                           arraylist.add(10,K); //end of plan year of termination
                           arraylist.add(11,L); //end of plan year of enrolment
                           arraylist.add(12, datetemp.format(k1Val)); //date of refund
                           arraylist.add(13,dF.format(Utility.betweenDates(date2,k1Val)/365.25)); //doe to dor
                          arraylist.add(14,dF.format(Utility.betweenDates(dateJ,k1Val)/365.25)); //start of plan year of temrination to dor
                           arraylist.add(15,dF.format(Utility.betweenDates(date2,dateL)/365.25)); //period from doe to end of plan year of enrolment
                           arraylist.add(16,dF.format(Utility.betweenDates(k1Val,dateK)/365.25)); //period from doe to end of plan year of enrolment
                           arraylist.add(17,dF.format(Utility.betweenDates(date2,k1Val)/365.25)); //doe to dor

                           // arraylist.add(9, j1Val);
                           for (int Col = 0, temp = 0; Col < 18; Col++, temp++) {
                               //Update the value of cell
                               //   cellR = rowR[Row].getCell(Col);
                               //   if (cellR == null) {

                               cellR = rowR[Row].createCell(Col);
                               //   }
                               cellR.setCellValue(String.valueOf(arraylist.get(Col)));
                           }
                       }

                   }


               }
            StartYear++;
           }//end years loop

            FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Terminees_Sheet.xlsx"));
            workbookR.write(outFile);
            fileR.close();
            outFile.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        catch(NullPointerException e){
            e.printStackTrace();

        }
        return String.valueOf(stringBuilder);
    }

    public void Create_Terminee_Contribution(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {

        DecimalFormat dF = new DecimalFormat("#.##");//#.##

        //OPEN ACTIVE SHEET
        FileInputStream fileInputStream = new FileInputStream(workingDir+"\\Terminees_Sheet.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet TermineeSheet = workbook.getSheet("Terminees");



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


    /*    int EndYear = 2015;
        int EndMonth = 12;
        int EndDay = 31;

        int StartYear = 2004;
        int StartMonth = 01;
        int StartDay = 01;*/

        int WriteAt = 23;

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
        years+=1;

        int counter=7;
        for (int x = 0; x < years; x++) {

            Calendar cal = Calendar.getInstance();
            cal.set(StartYear, StartMonth, StartDay);
            SimpleDateFormat sdf = new SimpleDateFormat("yy"); // Just the year, with 2 digits
            String formattedDate = sdf.format(cal.getTime());

            // String StartDate = StartDay + "-" + StartMonth + "-" + StartYear; //"01-01-05";
            // String EndDate = EndDay + "-" + EndMonth + "-" + StartYear;//"31-12-05";
            String Recon = ("Recon " + formattedDate);


            FileInputStream fs = null;
            try {
                fs = new FileInputStream(workingDir+ "\\Valuation Data.xlsx");
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            XSSFWorkbook WB = new XSSFWorkbook(fs);

//GET RECON SHEET
            XSSFSheet Reconsheet = WB.getSheet(Recon);
            System.out.println(formattedDate);
            int CountReconRow = Reconsheet.getPhysicalNumberOfRows();

            int numOfTerminee = TermineeSheet.getLastRowNum()+1; // get the last row number
            // System.out.println(rowCount);
            //     System.out.println("Recon"+ Recon + "count"+CountReconRow +"Actives" +numOfActives);
            XSSFRow[] rowR = new XSSFRow[numOfTerminee];
            Cell cellR = null;

         //   int Crow = 8;
            for (int row = 7; row < numOfTerminee; row++) {

                int Row = row;

                XSSFRow TermineeRow = TermineeSheet.getRow(row);

                XSSFCell cellA1 = TermineeRow.getCell((short) 0);  //employee number
                String result = cellA1.getStringCellValue();
                String a1Val = result.replaceAll("[-]","");

                XSSFCell cellB1 = TermineeRow.getCell((short) 1);   //last name
                if(cellB1==null){
                    cellB1 = TermineeRow.createCell(1);
                }
                String b1Val = cellB1.getStringCellValue();

                XSSFCell cellC1 = TermineeRow.getCell((short) 2);   //first name
                if(cellC1==null){
                    cellC1=TermineeRow.createCell(2);
                }
                String c1Val = cellC1.getStringCellValue();

                rowR[Row] = TermineeSheet.getRow(counter++);

                for (int y = 6; y < CountReconRow; y++) {

                    XSSFRow reconRow = Reconsheet.getRow(y);

                    XSSFCell A1 = reconRow.getCell(0);  //employee number
                    if (A1 == null) {
                        A1 = reconRow.createCell(0);
                    }
                    String a1 = A1.getStringCellValue();


                    XSSFCell B1 = reconRow.getCell(1);  //FNAME
                    if (B1 == null) {
                        B1 = reconRow.createCell(1);

                    }
                    String b1 = B1.getStringCellValue();

                    XSSFCell C1 = reconRow.getCell(2);  //LNAME
                    if (C1 == null) {
                        C1 = reconRow.createCell(2);

                    }
                    String c1 = C1.getStringCellValue();


                    XSSFCell H1 = reconRow.getCell((short) 7);
                    //Update the value of cell
                    if (H1 == null) {
                        H1 = reconRow.createCell(7);
                    }
                    double h1 = H1.getNumericCellValue();  //empoyee basic

                    XSSFCell I1 = reconRow.getCell((short) 8);
                    //Update the value of cell
                    if (I1 == null) {
                        I1 = reconRow.createCell(8);
                    }
                    double i1 = I1.getNumericCellValue();

                    XSSFCell J1 = reconRow.getCell((short) 9);
                    //Update the value of cell
                    if (J1 == null) {
                        J1 = reconRow.createCell(9);
                    }
                    double j1 = J1.getNumericCellValue();  //employee optional

                    //get withdrawal amount
                    XSSFCell cellL = reconRow.getCell((short) 11);
                    //Update the value of cell
                    if (cellL == null) {
                        cellL= reconRow.createCell(11);
                    }

                    double CellL = cellL.getNumericCellValue();  //employee basic withdrawal

                    //get withdrawal amount
                    XSSFCell cellM = reconRow.getCell((short) 12);
                    //Update the value of cell
                    if (cellM == null) {
                        cellM= reconRow.createCell(12);
                    }

                    double CellM = cellM.getNumericCellValue();  //employee basic withdrawal


                    if (a1Val.equals(a1) ) {
                        int AmtRefundedindex = 18 + (8*years+8); //move to the amount refunded column 8*years for no fees || 9*years for fees

                        ArrayList val = new ArrayList();
                        ArrayList val2 = new ArrayList();

              /*          System.out.print("A1: " + a1);
                        System.out.print(" D1: " + b1);
                        System.out.print(" C1: " + c1);
                    //    System.out.print("PS: " + i1Val);
                        System.out.print(" H1: " + h1);
                        System.out.print(" I1: " + i1);
                        System.out.print(" J1: " + j1);

                        System.out.println();
*/
                        val.add(0, h1);
                        val.add(1, i1);
                        val.add(2, j1);
                        val.add(3, 0.00);

                        val2.add(0,CellL);//employee basic withdrawal
                        val2.add(1,CellM);//employee optional withdrawal

                        for (int b = 0; b < val.size(); b++) {
                            cellR = rowR[Row].createCell(WriteAt + b);
                            cellR.setCellValue((Double)val.get(b));
                        }

                        for(int r=0;r<val2.size();r++){
                            //       System.out.println("index" + index);
                            //   System.out.println(" years: " + years);
                            cellR = rowR[Row].createCell(AmtRefundedindex);
                            cellR.setCellValue((Double)val2.get(0));

                            cellR = rowR[Row].createCell(AmtRefundedindex+r);
                            cellR.setCellValue((Double)val2.get(1));

                            //UNDER / OVER COLUMNS
               /*             cellR = rowR[Row].createCell(UnderOverIndex);
                            cellR.setCellValue((Double)val2.get(0));

                            cellR = rowR[Row].createCell(UnderOverIndex+r);
                            cellR.setCellValue((Double)val2.get(1));*/
                        }
                        break;

                    }

                }


            }
            WriteAt+=8;
            StartYear++; //comment out years if COMMENTED
            counter=7;

        }// END OF LOOP YEARS

        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Updated_Terminee_Sheet.xlsx"));
        workbook.write(outFile);
        fileInputStream.close();
        outFile.close();
    }

    public void Create_Terminee_Acc_Balances(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {
        DecimalFormat dF = new DecimalFormat("#.##");//#.##
        int readCol=23;//start to read from Col 26, which is starting contribution column
        int YearCol = readCol;//as we are in the same year, we should always be reading from the correct column
        int Write_Coloumn=27;//start to write at column 31 which are the accumullated cells

        //OPEN ACTIVE SHEET
        FileInputStream fileInputStream = new FileInputStream(workingDir +"\\Updated_Terminee_Sheet.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet TermineeSheet = workbook.getSheet("Terminees");

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


        ArrayList<Double> val = new ArrayList();
        int years = Utility.getDiffYears(startDate, endDate);//gets number of years
        years+=1;
        int numOfActives = TermineeSheet.getLastRowNum()+1;//gets number of active members
        //get the interest rates values
        double[] interestValues = new double [years];
        interestValues= getInterestRates(workingDir,years);

        //variables to hold calculated new Accumulated balances
        double[] newAccEmployeeBalance= new double[numOfActives];
        double[] newAccEmployeeOptional= new double[numOfActives];
        double[] newAccEmployerRequired=new double[numOfActives];
        double[] newAccEmployerOptional=new double[numOfActives];

//run for the appropiate number of years
        double CellAccEmployeeBasic=0;
        double CellAccEmployeeOptional=0;
        double CellAccEmployerRequired=0;
        double CellAccEmployerOptional=0;

        //get the initial Accumulated Balances
        double [] CellAccEmployeeBasic0 = new double[numOfActives];
        double[] CellAccEmployeeOptional0 = new double[numOfActives];
        double[] CellAccEmployerRequired0 = new double[numOfActives];
        double[] CellAccEmployerOptional0 = new double[numOfActives];


        for(int row=7,I=0;row<numOfActives;row++,I++){ //get the initial accumulated balances

            XSSFRow ActiveRow = TermineeSheet.getRow(row);

            XSSFCell[] cellAccEmployeeBasic0 = new XSSFCell[numOfActives];
            cellAccEmployeeBasic0[I] = ActiveRow.getCell( 19);  //employee number
            if(cellAccEmployeeBasic0[I]==null){
                cellAccEmployeeBasic0[I] = ActiveRow.createCell(19);
                cellAccEmployeeBasic0[I].setCellValue(0);
            }
            CellAccEmployeeBasic0[I]= cellAccEmployeeBasic0[I].getNumericCellValue();

            XSSFCell[] cellAccEmployeeOptional0 = new XSSFCell[numOfActives];
            cellAccEmployeeOptional0[I] = ActiveRow.getCell(20);  //employee number
            if(cellAccEmployeeOptional0[I]==null){
                cellAccEmployeeOptional0[I] = ActiveRow.createCell(20);
                cellAccEmployeeOptional0[I].setCellValue(0);
            }
            CellAccEmployeeOptional0[I]= cellAccEmployeeOptional0[I].getNumericCellValue();


            XSSFCell[] cellAccEmployerRequired0 = new XSSFCell[numOfActives];
            cellAccEmployerRequired0[I] = ActiveRow.getCell(21); //employee number
            if(cellAccEmployerRequired0[I]==null){
                cellAccEmployerRequired0[I] = ActiveRow.createCell(21);
                cellAccEmployerRequired0[I].setCellValue(0);
            }
            CellAccEmployerRequired0[I] = cellAccEmployerRequired0[I].getNumericCellValue();

            XSSFCell[] cellAccEmployerOptional0 = new XSSFCell[numOfActives];
            cellAccEmployerOptional0[I] = ActiveRow.getCell(22); //employee number
            if(cellAccEmployerOptional0[I]==null){
                cellAccEmployerOptional0[I] = ActiveRow.createCell(22);
                cellAccEmployerOptional0[I].setCellValue(0);
            }
            CellAccEmployerOptional0[I] = cellAccEmployerOptional0[I].getNumericCellValue();
        }//end of loop to get inital accumulated balances

        //MAIN PROCESSING-to get acc Balances and cont Balances
        Date BD= null;
        SimpleDateFormat datetemp = new SimpleDateFormat("dd-MMM-yy");
        SimpleDateFormat datetemp2 = new SimpleDateFormat("dd-MM-yy");
        for (int x = 0; x < years; x++) { //run for the appropiate number of years

            try {
                BD = datetemp2.parse(StartDay+"-"+StartMonth+"-"+StartYear);//"01-Jan-05"
                //      endDate = datetemp.parse(Utility.getEndDate(StartYear,01,01));
            } catch (ParseException e) {
                e.printStackTrace();
            }
            Cell cellR;

            for (int row = 7, I = 0; row < numOfActives; row++, I++) { //run for appropiate number of total active members
                readCol=YearCol;//need to ensure to start reading from same Column in same year

                XSSFRow ActiveRow = TermineeSheet.getRow(row);

                if (x == 0) {//get the accumulated balances just for the 1st year
                    CellAccEmployeeBasic = CellAccEmployeeBasic0[I];
                    CellAccEmployeeOptional = CellAccEmployeeOptional0[I];
                    CellAccEmployerRequired = CellAccEmployerRequired0[I];
                    CellAccEmployerOptional = CellAccEmployerOptional0[I];
                }
                else{ //get accumulated balances for every year after 1st year
                    CellAccEmployeeBasic =  newAccEmployeeBalance[I];
                    CellAccEmployeeOptional = newAccEmployeeOptional[I];
                    CellAccEmployerRequired = newAccEmployerRequired[I];
                    CellAccEmployerOptional = newAccEmployerOptional[I];
                }

                //get the CONTRIBUTIONS starting from column 26
                XSSFCell cellConEmployeeBasic = ActiveRow.getCell(readCol);
                if (cellConEmployeeBasic == null) {
                    cellConEmployeeBasic = ActiveRow.createCell(readCol);
                    cellConEmployeeBasic.setCellValue(0);
                }
                double CellConEmployeeBasic = cellConEmployeeBasic.getNumericCellValue();
                readCol+=1;

                XSSFCell cellConEmployeeOptional = ActiveRow.getCell(readCol);
                if (cellConEmployeeOptional == null) {
                    cellConEmployeeOptional = ActiveRow.createCell(readCol);
                    cellConEmployeeOptional.setCellValue(0);
                }
                double CellConEmployeeOptional = cellConEmployeeOptional.getNumericCellValue();
                readCol+=1;


                XSSFCell cellConEmployerRequired = ActiveRow.getCell(readCol);
                if (cellConEmployerRequired == null) {
                    cellConEmployerRequired = ActiveRow.createCell(readCol);
                    cellConEmployerRequired.setCellValue(0);
                }
                double CellConEmployerRequired = cellConEmployerRequired.getNumericCellValue();
                readCol+=1;

                XSSFCell cellConEmployerOptional = ActiveRow.getCell(readCol);
                if (cellConEmployerOptional == null) {
                    cellConEmployerOptional = ActiveRow.createCell(readCol);
                    cellConEmployerOptional.setCellValue(0);
                }
                double CellConEmployerOptional = cellConEmployerOptional.getNumericCellValue();
                readCol+=1;



                //GET DATE OF TERMINATION
                XSSFCell cellDOT = ActiveRow.getCell(8);
                if (cellDOT == null) {
                    cellDOT = ActiveRow.createCell(8);
                    cellDOT.setCellValue("01-Jan-01");
                }
                String CellDOT = cellDOT.getStringCellValue();

                //end of year of termination
                XSSFCell cellD = ActiveRow.getCell(10);
                if (cellD == null) {
                    cellD = ActiveRow.createCell(10);
                    cellD.setCellValue("01-Jan-01");
                }
                String CellD = cellD.getStringCellValue();//end of year of termination


                Date statusDate = null;
                Date EndDateofTermination = null;
                try {
                    statusDate = datetemp.parse(CellDOT);
                    EndDateofTermination=datetemp.parse(CellD);
                } catch (ParseException e) {
                    e.printStackTrace();
                }



          /*      int year = 2004;
                year+=x;
                System.out.println();
                System.out.println("Year: " +year+"Row: "+row);
                System.out.println("Acc "+ CellAccEmployeeBasic +"Con"+CellConEmployeeBasic);
                System.out.println("Acc "+ CellAccEmployeeOptional +"Con"+CellConEmployeeOptional);
                System.out.println("Acc "+ CellAccEmployerRequired +"Con"+CellConEmployerRequired);
                System.out.println("Acc "+ CellAccEmployerOptional +"Con"+CellConEmployerOptional);
                System.out.println(interestValues[x]);*/
                //FORMULA CALCULATIONS
                if (statusDate.after(BD) && statusDate.before(EndDateofTermination)) {
                newAccEmployeeBalance[I] = ((CellAccEmployeeBasic * (1+interestValues[x])) + (CellConEmployeeBasic * (1+(interestValues[x]*0.5))));//CellAccEmployeeBasic * (1 + 1) + CellConEmployeeBasic * (1 + 1 * 0.5);
                newAccEmployeeOptional[I] =((CellAccEmployeeOptional * (1+interestValues[x])) + (CellConEmployeeOptional * (1+(interestValues[x]*0.5))));//CellAccEmployeeOptional +  CellConEmployeeOptional; CellAccEmployeeOptional * (1 + 1) + CellConEmployeeOptional * (1 + 1 * 0.5);
                newAccEmployerRequired[I] = ((CellAccEmployerRequired * (1+interestValues[x])) + (CellConEmployerRequired *(1+(interestValues[x]*0.5))));//CellAccEmployerRequired * (1 + 1) + CellConEmployerRequired * (1 + 1 * 0.5) + CellFees * (1 + 1 * 0.5);
                newAccEmployerOptional[I] =((CellAccEmployerOptional * (1+interestValues[x])) + (CellConEmployerOptional * (1+(interestValues[x]*0.5))));//CellAccEmployerOptional * (1 + 1) + CellConEmployerOptional * (1 + 1 * 0.5);

                val.add(0, newAccEmployeeBalance[I]);
                val.add(1, newAccEmployeeOptional[I]);
                val.add(2, newAccEmployerRequired[I]);
                val.add(3, newAccEmployerOptional[I]);

//write the calculated accumulated balances to the sheet; start to write at column 31
                for (int b = 0; b < 4; b++) {
                    cellR = ActiveRow.createCell(Write_Coloumn + b);
                    cellR.setCellValue(dF.format(val.get(b)));
                }//end of loop
                }
            }//end of looping through each member

            //MOVING THE INDEXES
            YearCol=readCol+4;//move over by 5 columns to get next set of contributions for next year
            readCol=YearCol;//give the readCol 5 so that, it can always read the correct set of contributions of that same year
            Write_Coloumn += 8;//8 no fees  || 9 fees
            StartYear++;//increment Start year by one until we reach end year

            //when we are at end of year...we should write account balance as at end date
            if(x==(years-1)) {
                int AmtRefundedindex = readCol+3;
                int UnderOverIndex = readCol+5;
                //    readCol+=5;//move over by 3 columns to get to amount refunded

                for (int I = 0, row=7; row < numOfActives; I++, row++) {
                    XSSFRow ActiveRow = TermineeSheet.getRow(row);

                    //get the employee basic refund
                    XSSFCell cellEBAmtRefunded = ActiveRow.getCell(readCol+3);
                    if (cellEBAmtRefunded == null) {
                        cellEBAmtRefunded = ActiveRow.createCell(readCol+3);
                        //   cellEBAmtRefunded.setCellValue(0);
                    }
                    double CellEBAmtRefunded = cellEBAmtRefunded.getNumericCellValue();

                    //   AmtRefundedindex+=1;//move over to amount refunded for employee optional
                    //get the employee optional refund
                    XSSFCell cellEOAmtRefunded = ActiveRow.getCell(AmtRefundedindex+1);
                    if (cellEOAmtRefunded == null) {
                        cellEOAmtRefunded = ActiveRow.createCell(AmtRefundedindex+1);
                        //   cellEOAmtRefunded.setCellValue(0);
                    }
                    double CellEOAmtRefunded = cellEOAmtRefunded.getNumericCellValue();

                    //WRITE TO UNDER/OVER CELLS
                    // for(int col=0;col<2;col++) {
                    //write to employee basic of under/over column
                    cellR = ActiveRow.createCell(UnderOverIndex);
                    double resultEB =CellEBAmtRefunded-newAccEmployeeBalance[I];
                    cellR.setCellValue(resultEB);

                    //write to employee optional of under/over column
                    cellR = ActiveRow.createCell(UnderOverIndex+1);
                    double resultEO =CellEOAmtRefunded-newAccEmployeeOptional[I];
                    cellR.setCellValue(resultEO);
                    //    }
                }

            }

        }// END OF LOOP YEARS

        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Accumulated_Terminee_Sheet.xlsx"));
        workbook.write(outFile);
        fileInputStream.close();
        outFile.close();
    }//end of function Create Active Accumulated Balances

    //FEES
    public void Create_Fees_Activee_Contribution(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {
        /*
        FileInputStream fileR = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Actives_Sheet.xlsx");
        XSSFWorkbook workbookR = new XSSFWorkbook(fileR);
        XSSFSheet CopyFromSheet = workbookR.getSheetAt(0);
        */

        DecimalFormat dF = new DecimalFormat("#.##");//#.##

        //OPEN ACTIVE SHEET
        FileInputStream fileInputStream = new FileInputStream(workingDir +"\\Actives_Sheet.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet ActiveSheet = workbook.getSheet("Actives");

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

        // System.out.println(StartYear + "." + StartMonth + "." + StartDay);
        //  System.out.println(EndYear + "." + EndMonth + "." + EndDay);

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

        int years = Utility.getDiffYears(startDate, endDate);

        int counter=7;
        for (int x = 0; x <= years; x++) {


            Calendar cal = Calendar.getInstance();
            cal.set(StartYear, StartMonth, StartDay);
            SimpleDateFormat sdf = new SimpleDateFormat("yy"); // Just the year, with 2 digits
            String formattedDate = sdf.format(cal.getTime());

            // String StartDate = StartDay + "-" + StartMonth + "-" + StartYear; //"01-01-05";
            // String EndDate = EndDay + "-" + EndMonth + "-" + StartYear;//"31-12-05";
            String Recon = ("Recon " + formattedDate);


            FileInputStream fs = null;
            try {
                fs = new FileInputStream(workingDir+"\\Valuation Data.xlsx");
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            XSSFWorkbook WB = new XSSFWorkbook(fs);

//GET RECON SHEET
            XSSFSheet Reconsheet = WB.getSheet(Recon);
            System.out.println(formattedDate);
            int CountReconRow = Reconsheet.getPhysicalNumberOfRows();



            int numOfActives = ActiveSheet.getLastRowNum()+1; // get the last row number
            // System.out.println(rowCount);
            //     System.out.println("Recon"+ Recon + "count"+CountReconRow +"Actives" +numOfActives);
            XSSFRow[] rowR = new XSSFRow[numOfActives];
            Cell cellR = null;

/*if(Recon.equals(15)){
    numOfActives=35;
}*/
            int Crow = 8;
            for (int row = 7, I=8; row < numOfActives; row++,I++) {

                int Row = row;

                XSSFRow ActiveRow = ActiveSheet.getRow(row);

                XSSFCell cellA1 = ActiveRow.getCell((short) 0);  //employee number
                String result = cellA1.getStringCellValue();
                String a1Val = result.replaceAll("[-]","");

                XSSFCell cellB1 = ActiveRow.getCell((short) 1);   //last name
                String b1Val = cellB1.getStringCellValue();

                XSSFCell cellC1 = ActiveRow.getCell((short) 2);   //first name
                String c1Val = cellC1.getStringCellValue();

                XSSFCell cellD1 = ActiveRow.getCell((short) 3);
                String d1Val = cellD1.getStringCellValue();

                String i1Val = null;
                XSSFCell[] cellI1 = new XSSFCell[12];
                double [] d = new double[12];
                for (int g=8,j=0;g<20;g++,j++){
                    cellI1[j] = ActiveRow.getCell(g);
                    i1Val = cellI1[j].getStringCellValue();
                    d[j] = Double.parseDouble(i1Val);
                }


                //    Crow++;

                rowR[Row] = ActiveSheet.getRow(counter++);

                for (int y = 6; y < CountReconRow; y++) {

                    XSSFRow reconRow = Reconsheet.getRow(y);

                    XSSFCell A1 = reconRow.getCell(0);  //employee number
                    if (A1 == null) {
                        A1 = reconRow.createCell(0);
                    }
                    String a1 = A1.getStringCellValue();


                    XSSFCell B1 = reconRow.getCell(1);  //FNAME
                    if (B1 == null) {
                        B1 = reconRow.createCell(1);

                    }
                    String b1 = B1.getStringCellValue();

                    XSSFCell C1 = reconRow.getCell(2);  //LNAME
                    if (C1 == null) {
                        C1 = reconRow.createCell(2);

                    }
                    String c1 = C1.getStringCellValue();


                    XSSFCell H1 = reconRow.getCell((short) 7);
                    //Update the value of cell
                    if (H1 == null) {
                        H1 = reconRow.createCell(7);
                    }
                    double h1 = H1.getNumericCellValue();

                    XSSFCell I1 = reconRow.getCell((short) 8);
                    //Update the value of cell
                    if (I1 == null) {
                        I1 = reconRow.createCell(8);
                    }
                    double i1 = I1.getNumericCellValue();

                    XSSFCell J1 = reconRow.getCell((short) 9);
                    //Update the value of cell
                    if (J1 == null) {
                        J1 = reconRow.createCell(9);
                    }
                    double j1 = J1.getNumericCellValue();

                    XSSFCell cellFee = reconRow.getCell((short) 17);
                    //Update the value of cell
                    if (cellFee == null) {
                        cellFee = reconRow.createCell(17);
                    }
                    double CellFeeVal = cellFee.getNumericCellValue();

                    if (a1.equals(a1Val) ) {

                        ArrayList<Double> val = new ArrayList();

                        System.out.print("A1: " + a1);
                        System.out.print(" D1: " + b1);
                        System.out.print(" C1: " + c1);
                        System.out.print("PS: " + i1Val);
                        System.out.print(" H1: " + h1);
                        System.out.print(" I1: " + i1);
                        System.out.print(" J1: " + j1);
                        String test="";
                        double check = 0.05 *d[x];
                        check= Double.parseDouble(dF.format(check));
                        if (check == h1) {
                            test = "true";

                        } else {
                            test = "false";
                        }
                        System.out.print(" TEST: Pensionable Salary"+d[x]+"result "+check+test);

                        System.out.println();

                        val.add(0, h1); //employee basic
                        val.add(1, i1);  //employee optional
                        val.add(2, j1); //employer required
                        val.add(3, 0.00); //employer optional
                        val.add(4,CellFeeVal); //fees


                        for (int b = 0; b < val.size(); b++) {
                            cellR = rowR[Row].createCell(WriteAt + b);
                            cellR.setCellValue((Double)val.get(b));
                        }

                                                                           /*    cellR = rowR[Row].createCell(26);
                                                                                cellR.setCellValue(h1);

                                                                                cellR = rowR[Row].createCell(27);
                                                                                cellR.setCellValue(i1);

                                                                                cellR = rowR[Row].createCell(28);
                                                                                cellR.setCellValue(j1);

                                                                                cellR = rowR[Row].createCell(29);
                                                                                cellR.setCellValue(0);*/

                        break;

                    }


                    //   counter=7;
                }



            }

            WriteAt+=9;
            StartYear++; //comment out years if COMMENTED
            counter=7;

        }// END OF LOOP YEARS


        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Updated_Actives_Sheet.xlsx"));
        workbook.write(outFile);
        fileInputStream.close();
        outFile.close();
    }

    public void Create_Fees_Active_Acc_Balances(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {
        DecimalFormat dF = new DecimalFormat("#.##");//#.##
        int readCol=26;//start to read from Col 26, which is starting contribution column
        int YearCol = readCol;//as we are in the same year, we should always be reading from the correct column
        int Write_Coloumn =31;//start to write at column 31 which are the accumullated cells

        //OPEN ACTIVE SHEET
        FileInputStream fileInputStream = new FileInputStream(workingDir +"\\Updated_Actives_Sheet.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet ActiveSheet = workbook.getSheet("Actives");

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


        ArrayList<Double> val = new ArrayList();
        int years = Utility.getDiffYears(startDate, endDate);//gets number of years
        years+=1;
        int numOfActives = ActiveSheet.getLastRowNum()+1;//gets number of active members
        //get the interest rates values
        double[] interestValues = new double [years];
        interestValues= getInterestRates(workingDir,years);


        //variables to hold calculated new Accumulated balances
        double[] newAccEmployeeBalance= new double[numOfActives];
        double[] newAccEmployeeOptional= new double[numOfActives];
        double[] newAccEmployerRequired=new double[numOfActives];
        double[] newAccEmployerOptional=new double[numOfActives];

//run for the appropiate number of years
        double CellAccEmployeeBasic=0;
        double CellAccEmployeeOptional=0;
        double CellAccEmployerRequired=0;
        double CellAccEmployerOptional=0;

        //get the initial Accumulated Balances
        double [] CellAccEmployeeBasic0 = new double[numOfActives];
        double[] CellAccEmployeeOptional0 = new double[numOfActives];
        double[] CellAccEmployerRequired0 = new double[numOfActives];
        double[] CellAccEmployerOptional0 = new double[numOfActives];


        for(int row=7,I=0;row<numOfActives;row++,I++){ //get the initial accumulated balances

            XSSFRow ActiveRow = ActiveSheet.getRow(row);

            XSSFCell[] cellAccEmployeeBasic0 = new XSSFCell[numOfActives];
            cellAccEmployeeBasic0[I] = ActiveRow.getCell( 22);  //employee number
            if(cellAccEmployeeBasic0[I]==null){
                cellAccEmployeeBasic0[I] = ActiveRow.createCell(22);
                cellAccEmployeeBasic0[I].setCellValue(0);
            }
            CellAccEmployeeBasic0[I]= cellAccEmployeeBasic0[I].getNumericCellValue();

            XSSFCell[] cellAccEmployeeOptional0 = new XSSFCell[numOfActives];
            cellAccEmployeeOptional0[I] = ActiveRow.getCell(23);  //employee number
            if(cellAccEmployeeOptional0[I]==null){
                cellAccEmployeeOptional0[I] = ActiveRow.createCell(23);
                cellAccEmployeeOptional0[I].setCellValue(0);
            }
            CellAccEmployeeOptional0[I]= cellAccEmployeeOptional0[I].getNumericCellValue();


            XSSFCell[] cellAccEmployerRequired0 = new XSSFCell[numOfActives];
            cellAccEmployerRequired0[I] = ActiveRow.getCell(24); //employee number
            if(cellAccEmployerRequired0[I]==null){
                cellAccEmployerRequired0[I] = ActiveRow.createCell(24);
                cellAccEmployerRequired0[I].setCellValue(0);
            }
            CellAccEmployerRequired0[I] = cellAccEmployerRequired0[I].getNumericCellValue();

            XSSFCell[] cellAccEmployerOptional0 = new XSSFCell[numOfActives];
            cellAccEmployerOptional0[I] = ActiveRow.getCell(25); //employee number
            if(cellAccEmployerOptional0[I]==null){
                cellAccEmployerOptional0[I] = ActiveRow.createCell(25);
                cellAccEmployerOptional0[I].setCellValue(0);
            }
            CellAccEmployerOptional0[I] = cellAccEmployerOptional0[I].getNumericCellValue();
        }//end of loop to get inital accumulated balances

        //MAIN PROCESSING-to get acc Balances and cont Balances
        for (int x = 0; x < years; x++) { //run for the appropiate number of years

            //  System.out.println(numOfActives);
            Cell cellR;

            for (int row = 7, I = 0; row < numOfActives; row++, I++) { //run for appropiate number of total active members
                readCol=YearCol;//need to ensure to start reading from same Column in same year

                XSSFRow ActiveRow = ActiveSheet.getRow(row);

                if (x == 0) {//get the accumulated balances just for the 1st year
                    CellAccEmployeeBasic = CellAccEmployeeBasic0[I];
                    CellAccEmployeeOptional = CellAccEmployeeOptional0[I];
                    CellAccEmployerRequired = CellAccEmployerRequired0[I];
                    CellAccEmployerOptional = CellAccEmployerOptional0[I];
                }
                else{ //get accumulated balances for every year after 1st year
                    CellAccEmployeeBasic =  newAccEmployeeBalance[I];
                    CellAccEmployeeOptional = newAccEmployeeOptional[I];
                    CellAccEmployerRequired = newAccEmployerRequired[I];
                    CellAccEmployerOptional = newAccEmployerOptional[I];
                }

                //get the CONTRIBUTIONS starting from column 26
                XSSFCell cellConEmployeeBasic = ActiveRow.getCell(readCol);
                if (cellConEmployeeBasic == null) {
                    cellConEmployeeBasic = ActiveRow.createCell(readCol);
                    cellConEmployeeBasic.setCellValue(0);
                }
                double CellConEmployeeBasic = cellConEmployeeBasic.getNumericCellValue();
                readCol+=1;

                XSSFCell cellConEmployeeOptional = ActiveRow.getCell(readCol);
                if (cellConEmployeeOptional == null) {
                    cellConEmployeeOptional = ActiveRow.createCell(readCol);
                    cellConEmployeeOptional.setCellValue(0);
                }
                double CellConEmployeeOptional = cellConEmployeeOptional.getNumericCellValue();
                readCol+=1;


                XSSFCell cellConEmployerRequired = ActiveRow.getCell(readCol);
                if (cellConEmployerRequired == null) {
                    cellConEmployerRequired = ActiveRow.createCell(readCol);
                    cellConEmployerRequired.setCellValue(0);
                }
                double CellConEmployerRequired = cellConEmployerRequired.getNumericCellValue();
                readCol+=1;

                XSSFCell cellConEmployerOptional = ActiveRow.getCell(readCol);
                if (cellConEmployerOptional == null) {
                    cellConEmployerOptional = ActiveRow.createCell(readCol);
                    cellConEmployerOptional.setCellValue(0);
                }
                double CellConEmployerOptional = cellConEmployerOptional.getNumericCellValue();
                readCol+=1;



                XSSFCell cellFees = ActiveRow.getCell(readCol);
                if (cellFees == null) {
                    cellFees = ActiveRow.createCell(readCol);
                    cellFees.setCellValue(0);
                }
                double CellFees = cellFees.getNumericCellValue();


          /*      int year = 2004;
                year+=x;
                System.out.println();
                System.out.println("Year: " +year+"Row: "+row);
                System.out.println("Acc "+ CellAccEmployeeBasic +"Con"+CellConEmployeeBasic);
                System.out.println("Acc "+ CellAccEmployeeOptional +"Con"+CellConEmployeeOptional);
                System.out.println("Acc "+ CellAccEmployerRequired +"Con"+CellConEmployerRequired);
                System.out.println("Acc "+ CellAccEmployerOptional +"Con"+CellConEmployerOptional);
                System.out.println(interestValues[x]);*/
                //FORMULA CALCULATIONS
                newAccEmployeeBalance[I] = ((CellAccEmployeeBasic * (1+interestValues[x])) + (CellConEmployeeBasic * (1+(interestValues[x]*0.5))));//CellAccEmployeeBasic * (1 + 1) + CellConEmployeeBasic * (1 + 1 * 0.5);
                newAccEmployeeOptional[I] =((CellAccEmployeeOptional * (1+interestValues[x])) + (CellConEmployeeOptional * (1+(interestValues[x]*0.5))));//CellAccEmployeeOptional +  CellConEmployeeOptional; CellAccEmployeeOptional * (1 + 1) + CellConEmployeeOptional * (1 + 1 * 0.5);
                newAccEmployerRequired[I] = ((CellAccEmployerRequired * (1+interestValues[x])) + (CellConEmployerRequired *(1+(interestValues[x]*0.5))) + (CellFees * (1+(interestValues[x]*0.5))));//CellAccEmployerRequired * (1 + 1) + CellConEmployerRequired * (1 + 1 * 0.5) + CellFees * (1 + 1 * 0.5);
                newAccEmployerOptional[I] =((CellAccEmployerOptional * (1+interestValues[x])) + (CellConEmployerOptional * (1+(interestValues[x]*0.5))));//CellAccEmployerOptional * (1 + 1) + CellConEmployerOptional * (1 + 1 * 0.5);

                val.add(0, newAccEmployeeBalance[I]);
                val.add(1, newAccEmployeeOptional[I]);
                val.add(2, newAccEmployerRequired[I]);
                val.add(3, newAccEmployerOptional[I]);

//write the calculated accumulated balances to the sheet; start to write at column 31
                for (int b = 0; b < 4; b++) {
                    cellR = ActiveRow.createCell(Write_Coloumn + b);
                    cellR.setCellValue(dF.format(val.get(b)));
                }//end of loop

            }//end of looping through each member

            //MOVING THE INDEXES
            YearCol=readCol+5;//move over by 5 columns to get next set of contributions for next year
            readCol=YearCol;//give the readCol 5 so that, it can always read the correct set of contributions of that same year
            Write_Coloumn += 9;//8 no fees  || 9 fees
            StartYear++;//increment Start year by one until we reach end year

            //when we are at end of year...we should write account balance as at end date
            if(x==(years-1)) {
                for (int I = 0, row=7; row < numOfActives; I++, row++) {
                    XSSFRow  ActiveRow=ActiveSheet.getRow(row);
                    cellR = ActiveRow.createCell(readCol);
                    double AccountBalance =  newAccEmployeeBalance[I] + newAccEmployeeOptional[I] + newAccEmployerRequired[I]+newAccEmployerOptional[I];
                    cellR.setCellValue(AccountBalance);
                }
            }

        }// END OF LOOP YEARS

        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Accumulated_Actives_Sheet.xlsx"));
        workbook.write(outFile);
        fileInputStream.close();
        outFile.close();
    }//end of function Create Active Accumulated Balances

    public void Create_Fees_Terminee_Acc_Balances(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {
        DecimalFormat dF = new DecimalFormat("#.##");//#.##
        int readCol=23;//start to read from Col 26, which is starting contribution column
        int YearCol = readCol;//as we are in the same year, we should always be reading from the correct column
        int Write_Coloumn=28;//start to write at column 31 which are the accumullated cells

        //OPEN ACTIVE SHEET
        FileInputStream fileInputStream = new FileInputStream(workingDir +"\\Updated_Terminee_Sheet.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet TermineeSheet = workbook.getSheet("Terminees");

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
        Date endDate2 = null;
        try {
            startDate = df.parse(StartYear + "." + StartMonth + "." + StartDay);
            endDate = df.parse(EndYear + "." + EndMonth + "." + EndDay);
        //    endDate2 = df.parse(2009 + "." +12+ "." + 31);
        } catch (ParseException e) {
            e.printStackTrace();
        }


        ArrayList<Double> val = new ArrayList();
        int years = Utility.getDiffYears(startDate, endDate);//gets number of years
        years+=1;
        int numOfActives = TermineeSheet.getLastRowNum()+1;//gets number of active members
        //get the interest rates values
        double[] interestValues = new double [years];
        interestValues= getInterestRates(workingDir,years);

        //variables to hold calculated new Accumulated balances
        double[] newAccEmployeeBalance= new double[numOfActives];
        double[] newAccEmployeeOptional= new double[numOfActives];
        double[] newAccEmployerRequired=new double[numOfActives];
        double[] newAccEmployerOptional=new double[numOfActives];

//run for the appropiate number of years
        double CellAccEmployeeBasic=0;
        double CellAccEmployeeOptional=0;
        double CellAccEmployerRequired=0;
        double CellAccEmployerOptional=0;

        //get the initial Accumulated Balances
        double [] CellAccEmployeeBasic0 = new double[numOfActives];
        double[] CellAccEmployeeOptional0 = new double[numOfActives];
        double[] CellAccEmployerRequired0 = new double[numOfActives];
        double[] CellAccEmployerOptional0 = new double[numOfActives];


        for(int row=7,I=0;row<numOfActives;row++,I++){ //get the initial accumulated balances

            XSSFRow ActiveRow = TermineeSheet.getRow(row);

            XSSFCell[] cellAccEmployeeBasic0 = new XSSFCell[numOfActives];
            cellAccEmployeeBasic0[I] = ActiveRow.getCell( 19);  //employee number
            if(cellAccEmployeeBasic0[I]==null){
                cellAccEmployeeBasic0[I] = ActiveRow.createCell(19);
                cellAccEmployeeBasic0[I].setCellValue(0);
            }
            CellAccEmployeeBasic0[I]= cellAccEmployeeBasic0[I].getNumericCellValue();

            XSSFCell[] cellAccEmployeeOptional0 = new XSSFCell[numOfActives];
            cellAccEmployeeOptional0[I] = ActiveRow.getCell(20);  //employee number
            if(cellAccEmployeeOptional0[I]==null){
                cellAccEmployeeOptional0[I] = ActiveRow.createCell(20);
                cellAccEmployeeOptional0[I].setCellValue(0);
            }
            CellAccEmployeeOptional0[I]= cellAccEmployeeOptional0[I].getNumericCellValue();


            XSSFCell[] cellAccEmployerRequired0 = new XSSFCell[numOfActives];
            cellAccEmployerRequired0[I] = ActiveRow.getCell(21); //employee number
            if(cellAccEmployerRequired0[I]==null){
                cellAccEmployerRequired0[I] = ActiveRow.createCell(21);
                cellAccEmployerRequired0[I].setCellValue(0);
            }
            CellAccEmployerRequired0[I] = cellAccEmployerRequired0[I].getNumericCellValue();

            XSSFCell[] cellAccEmployerOptional0 = new XSSFCell[numOfActives];
            cellAccEmployerOptional0[I] = ActiveRow.getCell(22); //employee number
            if(cellAccEmployerOptional0[I]==null){
                cellAccEmployerOptional0[I] = ActiveRow.createCell(22);
                cellAccEmployerOptional0[I].setCellValue(0);
            }
            CellAccEmployerOptional0[I] = cellAccEmployerOptional0[I].getNumericCellValue();
        }//end of loop to get inital accumulated balances

        //MAIN PROCESSING-to get acc Balances and cont Balances
        Date BD= null;
        SimpleDateFormat datetemp = new SimpleDateFormat("dd-MMM-yy");
        SimpleDateFormat datetemp2 = new SimpleDateFormat("dd-MM-yy");
        //int AmtRefundedindex = 18 + (9*years+8); //move to the amount refunded column
        for (int x = 0; x < years; x++) { //run for the appropiate number of years 0 up to whatever year

            try {
               BD = datetemp2.parse(StartDay+"-"+StartMonth+"-"+StartYear);//"01-Jan-05"
          //      endDate = datetemp.parse(Utility.getEndDate(StartYear,01,01));
            } catch (ParseException e) {
                e.printStackTrace();
            }

            Cell cellR;
            for (int row = 7, I = 0; row < numOfActives; row++, I++) { //run for appropiate number of total active members
                readCol=YearCol;//need to ensure to start reading from same Column in same year

                XSSFRow ActiveRow = TermineeSheet.getRow(row);

                if (x == 0) {//get the accumulated balances just for the 1st year
                    CellAccEmployeeBasic = CellAccEmployeeBasic0[I];
                    CellAccEmployeeOptional = CellAccEmployeeOptional0[I];
                    CellAccEmployerRequired = CellAccEmployerRequired0[I];
                    CellAccEmployerOptional = CellAccEmployerOptional0[I];
                }
                else{ //get accumulated balances for every year after 1st year
                    CellAccEmployeeBasic =  newAccEmployeeBalance[I];
                    CellAccEmployeeOptional = newAccEmployeeOptional[I];
                    CellAccEmployerRequired = newAccEmployerRequired[I];
                    CellAccEmployerOptional = newAccEmployerOptional[I];
                }

                //get the CONTRIBUTIONS starting from column 23
                XSSFCell cellConEmployeeBasic = ActiveRow.getCell(readCol);
                if (cellConEmployeeBasic == null) {
                    cellConEmployeeBasic = ActiveRow.createCell(readCol);
                    cellConEmployeeBasic.setCellValue(0);
                }
                double CellConEmployeeBasic = cellConEmployeeBasic.getNumericCellValue();
                readCol+=1;

                XSSFCell cellConEmployeeOptional = ActiveRow.getCell(readCol);
                if (cellConEmployeeOptional == null) {
                    cellConEmployeeOptional = ActiveRow.createCell(readCol);
                    cellConEmployeeOptional.setCellValue(0);
                }
                double CellConEmployeeOptional = cellConEmployeeOptional.getNumericCellValue();
                readCol+=1;

                XSSFCell cellConEmployerRequired = ActiveRow.getCell(readCol);
                if (cellConEmployerRequired == null) {
                    cellConEmployerRequired = ActiveRow.createCell(readCol);
                    cellConEmployerRequired.setCellValue(0);
                }
                double CellConEmployerRequired = cellConEmployerRequired.getNumericCellValue();
                readCol+=1;

                XSSFCell cellConEmployerOptional = ActiveRow.getCell(readCol);
                if (cellConEmployerOptional == null) {
                    cellConEmployerOptional = ActiveRow.createCell(readCol);
                    cellConEmployerOptional.setCellValue(0);
                }
                double CellConEmployerOptional = cellConEmployerOptional.getNumericCellValue();
                readCol+=1;


                XSSFCell cellFees = ActiveRow.getCell(readCol);
                if (cellFees == null) {
                    cellFees = ActiveRow.createCell(readCol);
                    cellFees.setCellValue(0);
                }
                double CellFees = cellFees.getNumericCellValue();

                //GET DATE OF TERMINATION
                XSSFCell cellDOT = ActiveRow.getCell(8);
                if (cellDOT == null) {
                    cellDOT = ActiveRow.createCell(8);
                    cellDOT.setCellValue("01-Jan-01");
                }
               String CellDOT = cellDOT.getStringCellValue();

                //end of year of termination
                XSSFCell cellD = ActiveRow.getCell(10);
                if (cellD == null) {
                    cellD = ActiveRow.createCell(10);
                    cellD.setCellValue("01-Jan-01");
                }
                String CellD = cellD.getStringCellValue();//end of year of termination


                Date statusDate = null;
                Date EndDateofTermination = null;
                try {
                    statusDate = datetemp.parse(CellDOT);
                    EndDateofTermination=datetemp.parse(CellD);
                } catch (ParseException e) {
                    e.printStackTrace();
                }


                //FORMULA CALCULATIONS
                if (statusDate.after(BD) && statusDate.before(EndDateofTermination)) {
                    newAccEmployeeBalance[I] = ((CellAccEmployeeBasic * (1 + interestValues[x])) + (CellConEmployeeBasic * (1 + (interestValues[x] * 0.5))));//CellAccEmployeeBasic * (1 + 1) + CellConEmployeeBasic * (1 + 1 * 0.5);
                    newAccEmployeeOptional[I] = ((CellAccEmployeeOptional * (1 + interestValues[x])) + (CellConEmployeeOptional * (1 + (interestValues[x] * 0.5))));//CellAccEmployeeOptional +  CellConEmployeeOptional; CellAccEmployeeOptional * (1 + 1) + CellConEmployeeOptional * (1 + 1 * 0.5);
                    newAccEmployerRequired[I] = ((CellAccEmployerRequired * (1 + interestValues[x])) + (CellConEmployerRequired * (1 + (interestValues[x] * 0.5))) + (CellFees * (1 + (interestValues[x] * 0.5))));//CellAccEmployerRequired * (1 + 1) + CellConEmployerRequired * (1 + 1 * 0.5) + CellFees * (1 + 1 * 0.5);
                    newAccEmployerOptional[I] = ((CellAccEmployerOptional * (1 + interestValues[x])) + (CellConEmployerOptional * (1 + (interestValues[x] * 0.5))));//CellAccEmployerOptional * (1 + 1) + CellConEmployerOptional * (1 + 1 * 0.5);
          /*      }else{
                    newAccEmployeeBalance[I] = 0;//CellAccEmployeeBasic * (1 + 1) + CellConEmployeeBasic * (1 + 1 * 0.5);
                    newAccEmployeeOptional[I] = 0;//CellAccEmployeeOptional +  CellConEmployeeOptional; CellAccEmployeeOptional * (1 + 1) + CellConEmployeeOptional * (1 + 1 * 0.5);
                    newAccEmployerRequired[I] = 0;//CellAccEmployerRequired * (1 + 1) + CellConEmployerRequired * (1 + 1 * 0.5) + CellFees * (1 + 1 * 0.5);
                    newAccEmployerOptional[I]=0;
                }*/
                    val.add(0, newAccEmployeeBalance[I]);
                    val.add(1, newAccEmployeeOptional[I]);
                    val.add(2, newAccEmployerRequired[I]);
                    val.add(3, newAccEmployerOptional[I]);

//write the calculated accumulated balances to the sheet; start to write at column 31
                    for (int b = 0; b < 4; b++) {
                        cellR = ActiveRow.createCell(Write_Coloumn + b);
                        cellR.setCellValue(dF.format(val.get(b)));
                    }//end of loop

                }
            }//end of looping through each member

            //MOVING THE INDEXES
            YearCol=readCol+5;//move over by 5 columns to get next set of contributions for next year
            readCol=YearCol;//give the readCol 5 so that, it can always read the correct set of contributions of that same year
            Write_Coloumn += 9;//8 no fees  || 9 fees
            StartYear++;//increment Start year by one until we reach end year

            //when we are at end of year...we should write account balance as at end date
         if(x==(years-1)) {
             int AmtRefundedindex = readCol+3;
             int UnderOverIndex = readCol+5;
           //    readCol+=5;//move over by 3 columns to get to amount refunded

                for (int I = 0, row=7; row < numOfActives; I++, row++) {
                    XSSFRow ActiveRow = TermineeSheet.getRow(row);

                    //get the employee basic refund
                    XSSFCell cellEBAmtRefunded = ActiveRow.getCell(readCol+3);
                if (cellEBAmtRefunded == null) {
                    cellEBAmtRefunded = ActiveRow.createCell(readCol+3);
                 //   cellEBAmtRefunded.setCellValue(0);
                }
                double CellEBAmtRefunded = cellEBAmtRefunded.getNumericCellValue();

             //   AmtRefundedindex+=1;//move over to amount refunded for employee optional
                    //get the employee optional refund
               XSSFCell cellEOAmtRefunded = ActiveRow.getCell(AmtRefundedindex+1);
                if (cellEOAmtRefunded == null) {
                    cellEOAmtRefunded = ActiveRow.createCell(AmtRefundedindex+1);
                 //   cellEOAmtRefunded.setCellValue(0);
                }
                double CellEOAmtRefunded = cellEOAmtRefunded.getNumericCellValue();

               //WRITE TO UNDER/OVER CELLS
                   // for(int col=0;col<2;col++) {
                                //write to employee basic of under/over column
                        cellR = ActiveRow.createCell(UnderOverIndex);
                        double resultEB =CellEBAmtRefunded-newAccEmployeeBalance[I];
                        cellR.setCellValue(resultEB);

                    //write to employee optional of under/over column
                        cellR = ActiveRow.createCell(UnderOverIndex+1);
                        double resultEO =CellEOAmtRefunded-newAccEmployeeOptional[I];
                        cellR.setCellValue(resultEO);
                //    }
                    }

            }

        }// END OF LOOP YEARS
        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Accumulated_Terminee_Sheet.xlsx"));
        workbook.write(outFile);
        fileInputStream.close();
        outFile.close();
    }//end of function Create Active Accumulated Balances

    public void Create_Fees_Terminee_Contribution(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {

        DecimalFormat dF = new DecimalFormat("#.##");//#.##

        //OPEN ACTIVE SHEET
        FileInputStream fileInputStream = new FileInputStream(workingDir+"\\Terminees_Sheet.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet TermineeSheet = workbook.getSheet("Terminees");



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


    /*    int EndYear = 2015;
        int EndMonth = 12;
        int EndDay = 31;

        int StartYear = 2004;
        int StartMonth = 01;
        int StartDay = 01;*/

        int WriteAt = 23;

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
        years+=1;

        int counter=7;
        for (int x = 0; x < years; x++) {

            Calendar cal = Calendar.getInstance();
            cal.set(StartYear, StartMonth, StartDay);
            SimpleDateFormat sdf = new SimpleDateFormat("yy"); // Just the year, with 2 digits
            String formattedDate = sdf.format(cal.getTime());

            // String StartDate = StartDay + "-" + StartMonth + "-" + StartYear; //"01-01-05";
            // String EndDate = EndDay + "-" + EndMonth + "-" + StartYear;//"31-12-05";
            String Recon = ("Recon " + formattedDate);


            FileInputStream fs = null;
            try {
                fs = new FileInputStream(workingDir+ "\\Valuation Data.xlsx");
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            XSSFWorkbook WB = new XSSFWorkbook(fs);

//GET RECON SHEET
            XSSFSheet Reconsheet = WB.getSheet(Recon);
            System.out.println(formattedDate);
            int CountReconRow = Reconsheet.getPhysicalNumberOfRows();

            int numOfTerminee = TermineeSheet.getLastRowNum()+1; // get the last row number
            // System.out.println(rowCount);
            //     System.out.println("Recon"+ Recon + "count"+CountReconRow +"Actives" +numOfActives);
            XSSFRow[] rowR = new XSSFRow[numOfTerminee];
            Cell cellR = null;

            int Crow = 8;
            for (int row = 7; row < numOfTerminee; row++) {

                int Row = row;

                XSSFRow TermineeRow = TermineeSheet.getRow(row);

                XSSFCell cellA1 = TermineeRow.getCell((short) 0);  //employee number
                String result = cellA1.getStringCellValue();
                String a1Val = result.replaceAll("[-]","");

                XSSFCell cellB1 = TermineeRow.getCell((short) 1);   //last name
                if(cellB1==null){
                    cellB1 = TermineeRow.createCell(1);
                }
                String b1Val = cellB1.getStringCellValue();

                XSSFCell cellC1 = TermineeRow.getCell((short) 2);   //first name
                if(cellC1==null){
                    cellC1=TermineeRow.createCell(2);
                }
                String c1Val = cellC1.getStringCellValue();

                rowR[Row] = TermineeSheet.getRow(counter++);

                for (int y = 6; y < CountReconRow; y++) {

                    XSSFRow reconRow = Reconsheet.getRow(y);

                    XSSFCell A1 = reconRow.getCell(0);  //employee number
                    if (A1 == null) {
                        A1 = reconRow.createCell(0);
                    }
                    String a1 = A1.getStringCellValue();


                    XSSFCell B1 = reconRow.getCell(1);  //FNAME
                    if (B1 == null) {
                        B1 = reconRow.createCell(1);

                    }
                    String b1 = B1.getStringCellValue();

                    XSSFCell C1 = reconRow.getCell(2);  //LNAME
                    if (C1 == null) {
                        C1 = reconRow.createCell(2);

                    }
                    String c1 = C1.getStringCellValue();


                    XSSFCell H1 = reconRow.getCell((short) 7);
                    //Update the value of cell
                    if (H1 == null) {
                        H1 = reconRow.createCell(7);
                    }
                    double h1 = H1.getNumericCellValue();  //empoyee basic

                    XSSFCell I1 = reconRow.getCell((short) 8);
                    //Update the value of cell
                    if (I1 == null) {
                        I1 = reconRow.createCell(8);
                    }
                    double i1 = I1.getNumericCellValue();

                    XSSFCell J1 = reconRow.getCell((short) 9);
                    //Update the value of cell
                    if (J1 == null) {
                        J1 = reconRow.createCell(9);
                    }
                    double j1 = J1.getNumericCellValue();  //employee optional

                    //get withdrawal amount
                    XSSFCell cellL = reconRow.getCell((short) 11);
                    //Update the value of cell
                    if (cellL == null) {
                        cellL= reconRow.createCell(11);
                    }

                    double CellL = cellL.getNumericCellValue();  //employee basic withdrawal

                    //get withdrawal amount
                    XSSFCell cellM = reconRow.getCell((short) 12);
                    //Update the value of cell
                    if (cellM == null) {
                        cellM= reconRow.createCell(12);
                    }

                    double CellM = cellM.getNumericCellValue();  //employee basic withdrawal

                    XSSFCell cellFee = reconRow.getCell((short) 17);
                    //Update the value of cell
                    if (cellFee == null) {
                        cellFee = reconRow.createCell(17);
                    }
                    double CellFeeVal = cellFee.getNumericCellValue();


                    if (a1Val.equals(a1) ) {
                int AmtRefundedindex = 18 + (9*years+8); //move to the amount refunded column
                      //  int UnderOverIndex = AmtRefundedindex+2;


                       ArrayList val = new ArrayList();
                        ArrayList val2 = new ArrayList();
/*
                        System.out.print("A1: " + a1);
                        System.out.print(" D1: " + b1);
                        System.out.print(" C1: " + c1);
                        //    System.out.print("PS: " + i1Val);
                        System.out.print(" H1: " + h1);
                        System.out.print(" I1: " + i1);
                        System.out.print(" J1: " + j1);

                        System.out.println();*/

                        val.add(0, h1); //employee basic contribution
                        val.add(1, i1);//employee optional contribution
                        val.add(2, j1);//employer required contribution
                        val.add(3, 0.00);//employer optional contribution
                        val.add(4,CellFeeVal);//fee

                        val2.add(0,CellL);//employee basic withdrawal
                        val2.add(1,CellM);//employee optional withdrawal


                        for (int b = 0; b < val.size(); b++) {
                            cellR = rowR[Row].createCell(WriteAt + b);
                            cellR.setCellValue((Double)val.get(b));
                        }

                        for(int r=0;r<val2.size();r++){
                     //       System.out.println("index" + index);
                         //   System.out.println(" years: " + years);
                            cellR = rowR[Row].createCell(AmtRefundedindex);
                            cellR.setCellValue((Double)val2.get(0));

                            cellR = rowR[Row].createCell(AmtRefundedindex+r);
                            cellR.setCellValue((Double)val2.get(1));

                            //UNDER / OVER COLUMNS
               /*             cellR = rowR[Row].createCell(UnderOverIndex);
                            cellR.setCellValue((Double)val2.get(0));

                            cellR = rowR[Row].createCell(UnderOverIndex+r);
                            cellR.setCellValue((Double)val2.get(1));*/
                        }

           /*             if(x==(years-1)){
                            cellR = rowR[Row].createCell(WriteAt + b);
                            cellR.setCellValue((Double)val.get(b));
                        }*/
                        break;

                    }
                }

            }
            WriteAt+=9;//becuse of fees, we add extra column
            StartYear++; //comment out years if COMMENTED
            counter=7;

        }// END OF LOOP YEARS

        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Updated_Terminee_Sheet.xlsx"));
        workbook.write(outFile);
        fileInputStream.close();
        outFile.close();
    }

    public String View_Actives_Members(String workingDir, String endDate) throws IndexOutOfBoundsException {

        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("The following is a list of Active Members present as at " + endDate + " \n\n");
        SimpleDateFormat dF1 = new SimpleDateFormat();
        dF1.applyPattern("dd-MMM-yy");
        //   ArrayList<tableWindow.Fields> list = null;
        ArrayList list = null;
        try {
            // FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Seperated Members.xlsx");
            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Seperated Members.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet worksheet = workbook.getSheet("Actives");
            //   XSSFSheet sheet = workbook.getSheetAt(0);

            int rowCount = worksheet.getPhysicalNumberOfRows();


            SimpleDateFormat datetemp = new SimpleDateFormat("dd-MMM-yy");


            for (int row = 0; row < rowCount; row++) {
                int Row = row;

                if (Row == 0) {
                    row = 1;  //start reading from second row
                }
                XSSFRow row1 = worksheet.getRow(row);

                XSSFCell cellA1 = row1.getCell((short) 0);  //employee number
                String result = cellA1.getStringCellValue();
                String a1Val = result.replaceAll("[-]", "");

                XSSFCell cellB1 = row1.getCell((short) 1);
                String b1Val = cellB1.getStringCellValue();

                XSSFCell cellC1 = row1.getCell((short) 2);    //last name
                String c1Val = cellC1.getStringCellValue();

                XSSFCell cellD1 = row1.getCell((short) 3); //first name
                String d1Val = cellD1.getStringCellValue();

                XSSFCell cellE1 = row1.getCell((short) 4);  //sex
                String e1Val = cellE1.getStringCellValue();

                XSSFCell cellF1 = row1.getCell((short) 5); //dob
                //  Date f1Vall = cellF1.getDateCellValue();
                //String f1Val =datetemp.format( cellF1.getDateCellValue());
                String f1Val = cellF1.getStringCellValue();


                XSSFCell cellG1 = row1.getCell((short) 6); //plan entry
                // String g1Val =datetemp.format( cellG1.getDateCellValue());
                String g1Val = cellG1.getStringCellValue();

                XSSFCell cellH1 = row1.getCell((short) 7); //emp date
                //    String h1Val = datetemp.format( cellH1.getDateCellValue());
                String h1Val = cellH1.getStringCellValue();

                XSSFCell cellI1 = row1.getCell((short) 8);  //status date
                // String i1Val = datetemp.format( cellI1.getDateCellValue());
                String i1Val = cellI1.getStringCellValue();

                XSSFCell cellJ1 = row1.getCell((short) 9); //status
                String j1Val = cellJ1.getStringCellValue();

                if (j1Val.equals("ACTIVE") && !(d1Val.equals("KEY"))) {

                    System.out.print("A1: " + a1Val);
                    System.out.print(" B1: " + b1Val);
                    System.out.print(" C1: " + c1Val);
                    System.out.print(" D1: " + d1Val);
                    System.out.print(" E1: " + e1Val);
                    System.out.print(" F1: " + f1Val);
                    System.out.print(" G1: " + g1Val);
                    System.out.print(" H1: " + h1Val);
                    System.out.print(" I1: " + i1Val);
                    System.out.print(" J1: " + j1Val);

                    stringBuilder.append("Employee ID: " + a1Val + "\n ");
                    stringBuilder.append("Last Name: " + c1Val + "\n");
                    stringBuilder.append("First Name: " + d1Val + "\n");
                    stringBuilder.append("Date of Birth: " + f1Val + "\n");
                    stringBuilder.append("Employment Date: " + h1Val + "\n");
                    stringBuilder.append("Plan Entry Date: " + g1Val + "\n");
                    stringBuilder.append("Status Date: " + i1Val + "\n");
                    stringBuilder.append("Status: " + j1Val + "\n");
                    stringBuilder.append("-------------------------------------------------------\n");
                    list = new ArrayList<>();
                    list.add(a1Val);
                    list.add(c1Val);
                    list.add(d1Val);
                    list.add(f1Val);
                    // list = new ArrayList<tableWindow.Fields>();
                    //  tableWindow.Fields fields1 = new tableWindow.Fields(a1Val, c1Val, d1Val, f1Val);
                    //  list.add(fields1);
                    //  tableWindow tableWindow = null;
                    //  tableWindow.addRow(list);
                    System.out.println();
                    Color LINES = new Color(105, 105, 107);
//new ResultsWindow().appendToPane(new ResultsWindow(), stringBuilder+ "\n", LINES, true);
                }

            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (NullPointerException e) {
            e.printStackTrace();

        }
          return String.valueOf(stringBuilder);
       // return list;
    }// end of view active sheet

    public String View_Retired_Members(String workingDir, String endDate) throws IndexOutOfBoundsException {
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("The following is a list of Retired Members present as at "+endDate+" \n\n");
        SimpleDateFormat dF1 = new SimpleDateFormat();
        dF1.applyPattern("dd-MMM-yy");
        try {
            // FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Seperated Members.xlsx");
            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Seperated Members.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet worksheet = workbook.getSheet("Terminees");
            //   XSSFSheet sheet = workbook.getSheetAt(0);

            int rowCount = worksheet.getPhysicalNumberOfRows();


            SimpleDateFormat datetemp = new SimpleDateFormat("dd-MMM-yy");

            for (int row = 0; row < rowCount; row++) {
                int Row = row;

                if (Row == 0) {
                    row = 1;  //start reading from second row
                }
                XSSFRow row1 = worksheet.getRow(row);

                XSSFCell cellA1 = row1.getCell((short) 0);  //employee number
                String result = cellA1.getStringCellValue();
                String a1Val = result.replaceAll("[-]", "");

                XSSFCell cellB1 = row1.getCell((short) 1);
                String b1Val = cellB1.getStringCellValue();

                XSSFCell cellC1 = row1.getCell((short) 2);    //last name
                String c1Val = cellC1.getStringCellValue();

                XSSFCell cellD1 = row1.getCell((short) 3); //first name
                String d1Val = cellD1.getStringCellValue();

                XSSFCell cellE1 = row1.getCell((short) 4);  //sex
                String e1Val = cellE1.getStringCellValue();

                XSSFCell cellF1 = row1.getCell((short) 5); //dob
                //  Date f1Vall = cellF1.getDateCellValue();
                //String f1Val =datetemp.format( cellF1.getDateCellValue());
                String f1Val = cellF1.getStringCellValue();


                XSSFCell cellG1 = row1.getCell((short) 6); //plan entry
                // String g1Val =datetemp.format( cellG1.getDateCellValue());
                String g1Val = cellG1.getStringCellValue();

                XSSFCell cellH1 = row1.getCell((short) 7); //emp date
                //    String h1Val = datetemp.format( cellH1.getDateCellValue());
                String h1Val = cellH1.getStringCellValue();

                XSSFCell cellI1 = row1.getCell((short) 8);  //status date
                // String i1Val = datetemp.format( cellI1.getDateCellValue());
                String i1Val = cellI1.getStringCellValue();

                XSSFCell cellJ1 = row1.getCell((short) 9); //status
                String j1Val = cellJ1.getStringCellValue();

                if ( (j1Val.equals("RETIREMENT")) && !(d1Val.equals("KEY"))){
                    System.out.print("A1: " + a1Val);
                    System.out.print(" B1: " + b1Val);
                    System.out.print(" C1: " + c1Val);
                    System.out.print(" D1: " + d1Val);
                    System.out.print(" E1: " + e1Val);
                    System.out.print(" F1: " + f1Val);
                    System.out.print(" G1: " + g1Val);
                    System.out.print(" H1: " + h1Val);
                    System.out.print(" I1: " + i1Val);
                    System.out.print(" J1: " + j1Val);

                    stringBuilder.append("Employee ID: " + a1Val + "\n");
                    stringBuilder.append("Last Name: " + c1Val + "\n");
                    stringBuilder.append("First Name: " + d1Val + "\n");
                    stringBuilder.append("Date of Birth: " + f1Val + "\n");
                    stringBuilder.append("Employment Date: " + h1Val + "\n");
                    stringBuilder.append("Plan Entry Date: " + g1Val + "\n");
                    stringBuilder.append("Status Date: " + i1Val + "\n");
                    stringBuilder.append("Status: " + j1Val + "\n");
                    stringBuilder.append("-------------------------------------------------------\n");

                    System.out.println();

                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (NullPointerException e) {
            e.printStackTrace();

        }
        return String.valueOf(stringBuilder);
    }// end of view active sheet

    public String View_Terminee_Members(String workingDir, String endDate) throws IndexOutOfBoundsException {
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("The following is a list of Terminee Members present as at "+endDate+" \n\n");
        SimpleDateFormat dF1 = new SimpleDateFormat();
        dF1.applyPattern("dd-MMM-yy");
        try {
            // FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Seperated Members.xlsx");
            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Seperated Members.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet worksheet = workbook.getSheet("Terminees");
            //   XSSFSheet sheet = workbook.getSheetAt(0);

            int rowCount = worksheet.getPhysicalNumberOfRows();


            SimpleDateFormat datetemp = new SimpleDateFormat("dd-MMM-yy");

            for (int row = 0; row < rowCount; row++) {
                int Row = row;

                if (Row == 0) {
                    row = 1;  //start reading from second row
                }
                XSSFRow row1 = worksheet.getRow(row);

                XSSFCell cellA1 = row1.getCell((short) 0);  //employee number
                String result = cellA1.getStringCellValue();
                String a1Val = result.replaceAll("[-]", "");

                XSSFCell cellB1 = row1.getCell((short) 1);
                String b1Val = cellB1.getStringCellValue();

                XSSFCell cellC1 = row1.getCell((short) 2);    //last name
                String c1Val = cellC1.getStringCellValue();

                XSSFCell cellD1 = row1.getCell((short) 3); //first name
                String d1Val = cellD1.getStringCellValue();

                XSSFCell cellE1 = row1.getCell((short) 4);  //sex
                String e1Val = cellE1.getStringCellValue();

                XSSFCell cellF1 = row1.getCell((short) 5); //dob
                //  Date f1Vall = cellF1.getDateCellValue();
                //String f1Val =datetemp.format( cellF1.getDateCellValue());
                String f1Val = cellF1.getStringCellValue();


                XSSFCell cellG1 = row1.getCell((short) 6); //plan entry
                // String g1Val =datetemp.format( cellG1.getDateCellValue());
                String g1Val = cellG1.getStringCellValue();

                XSSFCell cellH1 = row1.getCell((short) 7); //emp date
                //    String h1Val = datetemp.format( cellH1.getDateCellValue());
                String h1Val = cellH1.getStringCellValue();

                XSSFCell cellI1 = row1.getCell((short) 8);  //status date
                // String i1Val = datetemp.format( cellI1.getDateCellValue());
                String i1Val = cellI1.getStringCellValue();

                XSSFCell cellJ1 = row1.getCell((short) 9); //status
                String j1Val = cellJ1.getStringCellValue();

                if ((j1Val.equals("DEFERRED")||j1Val.equals("RETIREMENT")|| j1Val.equals("DEATH")||j1Val.equals("TERMINATED")) && !(d1Val.equals("KEY"))) {
                    System.out.print("A1: " + a1Val);
                    System.out.print(" B1: " + b1Val);
                    System.out.print(" C1: " + c1Val);
                    System.out.print(" D1: " + d1Val);
                    System.out.print(" E1: " + e1Val);
                    System.out.print(" F1: " + f1Val);
                    System.out.print(" G1: " + g1Val);
                    System.out.print(" H1: " + h1Val);
                    System.out.print(" I1: " + i1Val);
                    System.out.print(" J1: " + j1Val);

                    stringBuilder.append("Employee ID: " + a1Val + "\n");
                    stringBuilder.append("Last Name: " + c1Val + "\n");
                    stringBuilder.append("First Name: " + d1Val + "\n");
                    stringBuilder.append("Date of Birth: " + f1Val + "\n");
                    stringBuilder.append("Employment Date: " + h1Val + "\n");
                    stringBuilder.append("Plan Entry Date: " + g1Val + "\n");
                    stringBuilder.append("Status Date: " + i1Val + "\n");
                    stringBuilder.append("Status: " + j1Val + "\n");
                    stringBuilder.append("-------------------------------------------------------\n");

                    System.out.println();

                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (NullPointerException e) {
            e.printStackTrace();

        }
        return String.valueOf(stringBuilder);
    }// end of view active sheet

    public String View_Deceased_Members(String workingDir, String endDate) throws IndexOutOfBoundsException {
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("The following is a list of Deceased Members present as at "+endDate+" \n\n");
        SimpleDateFormat dF1 = new SimpleDateFormat();
        dF1.applyPattern("dd-MMM-yy");
        try {
            // FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Seperated Members.xlsx");
            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Seperated Members.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet worksheet = workbook.getSheet("Terminees");
            //   XSSFSheet sheet = workbook.getSheetAt(0);

            int rowCount = worksheet.getPhysicalNumberOfRows();


            SimpleDateFormat datetemp = new SimpleDateFormat("dd-MMM-yy");

            for (int row = 0; row < rowCount; row++) {
                int Row = row;

                if (Row == 0) {
                    row = 1;  //start reading from second row
                }
                XSSFRow row1 = worksheet.getRow(row);

                XSSFCell cellA1 = row1.getCell((short) 0);  //employee number
                String result = cellA1.getStringCellValue();
                String a1Val = result.replaceAll("[-]", "");

                XSSFCell cellB1 = row1.getCell((short) 1);
                String b1Val = cellB1.getStringCellValue();

                XSSFCell cellC1 = row1.getCell((short) 2);    //last name
                String c1Val = cellC1.getStringCellValue();

                XSSFCell cellD1 = row1.getCell((short) 3); //first name
                String d1Val = cellD1.getStringCellValue();

                XSSFCell cellE1 = row1.getCell((short) 4);  //sex
                String e1Val = cellE1.getStringCellValue();

                XSSFCell cellF1 = row1.getCell((short) 5); //dob
                //  Date f1Vall = cellF1.getDateCellValue();
                //String f1Val =datetemp.format( cellF1.getDateCellValue());
                String f1Val = cellF1.getStringCellValue();


                XSSFCell cellG1 = row1.getCell((short) 6); //plan entry
                // String g1Val =datetemp.format( cellG1.getDateCellValue());
                String g1Val = cellG1.getStringCellValue();

                XSSFCell cellH1 = row1.getCell((short) 7); //emp date
                //    String h1Val = datetemp.format( cellH1.getDateCellValue());
                String h1Val = cellH1.getStringCellValue();

                XSSFCell cellI1 = row1.getCell((short) 8);  //status date
                // String i1Val = datetemp.format( cellI1.getDateCellValue());
                String i1Val = cellI1.getStringCellValue();

                XSSFCell cellJ1 = row1.getCell((short) 9); //status
                String j1Val = cellJ1.getStringCellValue();

                if ( (j1Val.equals("DEATH")) && !(d1Val.equals("KEY"))){
                    System.out.print("A1: " + a1Val);
                    System.out.print(" B1: " + b1Val);
                    System.out.print(" C1: " + c1Val);
                    System.out.print(" D1: " + d1Val);
                    System.out.print(" E1: " + e1Val);
                    System.out.print(" F1: " + f1Val);
                    System.out.print(" G1: " + g1Val);
                    System.out.print(" H1: " + h1Val);
                    System.out.print(" I1: " + i1Val);
                    System.out.print(" J1: " + j1Val);

                    stringBuilder.append("Employee ID: " + a1Val + "\n");
                    stringBuilder.append("Last Name: " + c1Val + "\n");
                    stringBuilder.append("First Name: " + d1Val + "\n");
                    stringBuilder.append("Date of Birth: " + f1Val + "\n");
                    stringBuilder.append("Employment Date: " + h1Val + "\n");
                    stringBuilder.append("Plan Entry Date: " + g1Val + "\n");
                    stringBuilder.append("Status Date: " + i1Val + "\n");
                    stringBuilder.append("Status: " + j1Val + "\n");
                    stringBuilder.append("-------------------------------------------------------\n");

                    System.out.println();

                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (NullPointerException e) {
            e.printStackTrace();

        }
        return String.valueOf(stringBuilder);
    }// end of view active sheet

    public String View_Deferred_Members(String workingDir, String endDate) throws IndexOutOfBoundsException {
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("The following is a list of Deferred Members present as at "+endDate+" \n\n");
        SimpleDateFormat dF1 = new SimpleDateFormat();
        dF1.applyPattern("dd-MMM-yy");
        try {
            // FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Seperated Members.xlsx");
            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Seperated Members.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet worksheet = workbook.getSheet("Terminees");
            //   XSSFSheet sheet = workbook.getSheetAt(0);

            int rowCount = worksheet.getPhysicalNumberOfRows();


            SimpleDateFormat datetemp = new SimpleDateFormat("dd-MMM-yy");

            for (int row = 0; row < rowCount; row++) {
                int Row = row;

                if (Row == 0) {
                    row = 1;  //start reading from second row
                }
                XSSFRow row1 = worksheet.getRow(row);

                XSSFCell cellA1 = row1.getCell((short) 0);  //employee number
                String result = cellA1.getStringCellValue();
                String a1Val = result.replaceAll("[-]", "");

                XSSFCell cellB1 = row1.getCell((short) 1);
                String b1Val = cellB1.getStringCellValue();

                XSSFCell cellC1 = row1.getCell((short) 2);    //last name
                String c1Val = cellC1.getStringCellValue();

                XSSFCell cellD1 = row1.getCell((short) 3); //first name
                String d1Val = cellD1.getStringCellValue();

                XSSFCell cellE1 = row1.getCell((short) 4);  //sex
                String e1Val = cellE1.getStringCellValue();

                XSSFCell cellF1 = row1.getCell((short) 5); //dob
                //  Date f1Vall = cellF1.getDateCellValue();
                //String f1Val =datetemp.format( cellF1.getDateCellValue());
                String f1Val = cellF1.getStringCellValue();


                XSSFCell cellG1 = row1.getCell((short) 6); //plan entry
                // String g1Val =datetemp.format( cellG1.getDateCellValue());
                String g1Val = cellG1.getStringCellValue();

                XSSFCell cellH1 = row1.getCell((short) 7); //emp date
                //    String h1Val = datetemp.format( cellH1.getDateCellValue());
                String h1Val = cellH1.getStringCellValue();

                XSSFCell cellI1 = row1.getCell((short) 8);  //status date
                // String i1Val = datetemp.format( cellI1.getDateCellValue());
                String i1Val = cellI1.getStringCellValue();

                XSSFCell cellJ1 = row1.getCell((short) 9); //status
                String j1Val = cellJ1.getStringCellValue();

                if ( (j1Val.equals("DEFERRED")) && !(d1Val.equals("KEY"))){
                    System.out.print("A1: " + a1Val);
                    System.out.print(" B1: " + b1Val);
                    System.out.print(" C1: " + c1Val);
                    System.out.print(" D1: " + d1Val);
                    System.out.print(" E1: " + e1Val);
                    System.out.print(" F1: " + f1Val);
                    System.out.print(" G1: " + g1Val);
                    System.out.print(" H1: " + h1Val);
                    System.out.print(" I1: " + i1Val);
                    System.out.print(" J1: " + j1Val);

                    stringBuilder.append("Employee ID: " + a1Val + "\n");
                    stringBuilder.append("Last Name: " + c1Val + "\n");
                    stringBuilder.append("First Name: " + d1Val + "\n");
                    stringBuilder.append("Date of Birth: " + f1Val + "\n");
                    stringBuilder.append("Employment Date: " + h1Val + "\n");
                    stringBuilder.append("Plan Entry Date: " + g1Val + "\n");
                    stringBuilder.append("Status Date: " + i1Val + "\n");
                    stringBuilder.append("Status: " + j1Val + "\n");
                    stringBuilder.append("-------------------------------------------------------\n");

                    System.out.println();

                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (NullPointerException e) {
            e.printStackTrace();

        }
        return String.valueOf(stringBuilder);
    }// end of view active sheet

    public String View_Terminated_Members(String workingDir, String endDate) throws IndexOutOfBoundsException {
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("The following is a list of Terminated Members present as at "+endDate+" \n\n");
        SimpleDateFormat dF1 = new SimpleDateFormat();
        dF1.applyPattern("dd-MMM-yy");
        try {
            // FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Seperated Members.xlsx");
            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Seperated Members.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet worksheet = workbook.getSheet("Terminees");
            //   XSSFSheet sheet = workbook.getSheetAt(0);

            int rowCount = worksheet.getPhysicalNumberOfRows();


            SimpleDateFormat datetemp = new SimpleDateFormat("dd-MMM-yy");

            for (int row = 0; row < rowCount; row++) {
                int Row = row;

                if (Row == 0) {
                    row = 1;  //start reading from second row
                }
                XSSFRow row1 = worksheet.getRow(row);

                XSSFCell cellA1 = row1.getCell((short) 0);  //employee number
                String result = cellA1.getStringCellValue();
                String a1Val = result.replaceAll("[-]", "");

                XSSFCell cellB1 = row1.getCell((short) 1);
                String b1Val = cellB1.getStringCellValue();

                XSSFCell cellC1 = row1.getCell((short) 2);    //last name
                String c1Val = cellC1.getStringCellValue();

                XSSFCell cellD1 = row1.getCell((short) 3); //first name
                String d1Val = cellD1.getStringCellValue();

                XSSFCell cellE1 = row1.getCell((short) 4);  //sex
                String e1Val = cellE1.getStringCellValue();

                XSSFCell cellF1 = row1.getCell((short) 5); //dob
                //  Date f1Vall = cellF1.getDateCellValue();
                //String f1Val =datetemp.format( cellF1.getDateCellValue());
                String f1Val = cellF1.getStringCellValue();


                XSSFCell cellG1 = row1.getCell((short) 6); //plan entry
                // String g1Val =datetemp.format( cellG1.getDateCellValue());
                String g1Val = cellG1.getStringCellValue();

                XSSFCell cellH1 = row1.getCell((short) 7); //emp date
                //    String h1Val = datetemp.format( cellH1.getDateCellValue());
                String h1Val = cellH1.getStringCellValue();

                XSSFCell cellI1 = row1.getCell((short) 8);  //status date
                // String i1Val = datetemp.format( cellI1.getDateCellValue());
                String i1Val = cellI1.getStringCellValue();

                XSSFCell cellJ1 = row1.getCell((short) 9); //status
                String j1Val = cellJ1.getStringCellValue();

                if ( (j1Val.equals("TERMINATED")) && !(d1Val.equals("KEY"))){
                    System.out.print("A1: " + a1Val);
                    System.out.print(" B1: " + b1Val);
                    System.out.print(" C1: " + c1Val);
                    System.out.print(" D1: " + d1Val);
                    System.out.print(" E1: " + e1Val);
                    System.out.print(" F1: " + f1Val);
                    System.out.print(" G1: " + g1Val);
                    System.out.print(" H1: " + h1Val);
                    System.out.print(" I1: " + i1Val);
                    System.out.print(" J1: " + j1Val);

                    stringBuilder.append("Employee ID: " + a1Val + "\n");
                    stringBuilder.append("Last Name: " + c1Val + "\n");
                    stringBuilder.append("First Name: " + d1Val + "\n");
                    stringBuilder.append("Date of Birth: " + f1Val + "\n");
                    stringBuilder.append("Employment Date: " + h1Val + "\n");
                    stringBuilder.append("Plan Entry Date: " + g1Val + "\n");
                    stringBuilder.append("Status Date: " + i1Val + "\n");
                    stringBuilder.append("Status: " + j1Val + "\n");
                    stringBuilder.append("-------------------------------------------------------\n");

                    System.out.println();

                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (NullPointerException e) {
            e.printStackTrace();

        }
        return String.valueOf(stringBuilder);
    }// end of view

    public double[] getInterestRates(String workingDir, int numOfYears) throws IOException {
    //    ArrayList list = new ArrayList();
//System.out.print(numOfYears);
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(workingDir + "\\Interest Rates.xlsx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet =workbook.getSheet("rates");
XSSFCell [] cell = new XSSFCell[numOfYears];
double [] values = new double[numOfYears];

        for(int row=1,I=0;I<numOfYears;row++,I++) {
            XSSFRow interestRow = sheet.getRow(row);

            cell[I] = interestRow.getCell(1);
        //    list.add(cell[I].getNumericCellValue(),I);
            values[I]=cell[I].getNumericCellValue();

        }

            return values;
    }

    public void Results(){

        DecimalFormat dF = new DecimalFormat("#");//#.##


        int EndYear = 2015;
        int EndMonth = 12;
        int EndDay=31;

        int StartYear=2004;
        int StartMonth=01;
        int StartDay= 01;

        DateFormat df = new SimpleDateFormat("yyyy.MM.dd");
        Date startDate = null;
        Date endDate = null;
        try {
            startDate = df.parse(StartYear+"."+StartMonth+"."+StartDay);
            endDate = df.parse(EndYear+"."+EndMonth+"."+EndDay);
        } catch (ParseException e) {
            e.printStackTrace();
        }


        int years= Utility.getDiffYears(startDate,endDate);

        for(int x=0;x<=years;x++) {
            Calendar cal = Calendar.getInstance();
            cal.set(StartYear, StartMonth, StartDay);
            SimpleDateFormat sdf = new SimpleDateFormat("yy"); // Just the year, with 2 digits
            String formattedDate = sdf.format(cal.getTime());

            String StartDate = StartDay + "-" + StartMonth + "-" + StartYear; //"01-01-05";
            String EndDate = EndDay + "-" + EndMonth + "-" + StartYear;//"31-12-05";
            String Recon = ("Recon "+formattedDate);
            System.out.println(StartYear+ " Activees Sum: " +   Double.valueOf(dF.format(ActiveSumReader(StartDate, EndDate, Recon))));
            System.out.println(StartYear+ " Terminees Sum: " + Double.valueOf(dF.format(TermineeSumReader(StartDate, EndDate, Recon))));
            StartYear++;
        }
    }

    public double TermineeSumReader(String StartDate, String EndDate, String Recon) throws IndexOutOfBoundsException {

        double TActiveSum = 0;
        try {
            FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Valuation Data.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

            XSSFSheet Demosheet = workbook.getSheet("DEMO");
            XSSFSheet Reconsheet = workbook.getSheet(Recon);
            String T = EndDate;
            String start = StartDate;

            int DEMOrowCount = Demosheet.getPhysicalNumberOfRows();
            int ReconrowCount = Reconsheet.getPhysicalNumberOfRows() + 1;


            double ActiveSum = 0;
            TActiveSum = 0;
            double TermineeSum = 0;

            int count = 0;
            for (int row = 0; row < DEMOrowCount; row++) {  //start looping through demo records
                int Row = row;

                if (Row == 0) {
                    row = 1;  //start reading from second row
                }

                XSSFRow row1 = Demosheet.getRow(row);

                XSSFCell cellA1 = row1.getCell((short) 0);
                String result = cellA1.getStringCellValue();

                String a1Val = result.replaceAll("[-]", "");

                XSSFCell cellC1 = row1.getCell((short) 2);
                String c1Val = cellC1.getStringCellValue();

                XSSFCell cellD1 = row1.getCell((short) 3);
                String d1Val = cellD1.getStringCellValue();

                XSSFCell cellI1 = row1.getCell((short) 8);
                Date celli1 = cellI1.getDateCellValue();
                SimpleDateFormat datetemp = new SimpleDateFormat("dd-MMM-yy");
                SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yy");

                //Date cellValue = datetemp.parse(i1Val);
                String i1Val = sdf.format(celli1);
                Date date1 = null;
                Date date2 = null;
                Date Startdate = null;

                try {
                    date1 = sdf.parse(i1Val);
                    date2 = sdf.parse(T);
                    Startdate = sdf.parse(start);
                } catch (ParseException e) {
                    e.printStackTrace();
                }

                XSSFCell cellJ1 = row1.getCell((short) 9);
                String j1Val = cellJ1.getStringCellValue();


                //  count = 0;

                //***********TERMINATIONS********************************
                for (int Col = 6; Col < ReconrowCount; Col++) {
                    XSSFRow Reconrow = Reconsheet.getRow(Col);

                    XSSFCell ReconCellA1 = Reconrow.getCell((short) 0);
                    //Update the value of cell
                    if (ReconCellA1 == null) {
                        ReconCellA1 = Reconrow.createCell(0);
                    }
                    String ReconA1Val = ReconCellA1.getStringCellValue();

                    XSSFCell ReconCellB1 = Reconrow.getCell((short) 1);
                    //Update the value of cell
                    if (ReconCellB1 == null) {
                        ReconCellB1 = Reconrow.createCell(1);
                    }
                    String ReconB1Val = ReconCellB1.getStringCellValue();

                    XSSFCell ReconCellC1 = Reconrow.getCell((short) 2);
                    //Update the value of cell
                    if (ReconCellC1 == null) {
                        ReconCellC1 = Reconrow.createCell(2);
                    }
                    String ReconC1Val = ReconCellC1.getStringCellValue();

                    XSSFCell ReconCellT1 = Reconrow.getCell((short) 19);
                    //Update the value of cell
                    if (ReconCellT1 == null) {
                        ReconCellT1 = Reconrow.createCell(19);
                    }
                    double ReconT1Val = ReconCellT1.getNumericCellValue();


                    XSSFCell ReconCellU1 = Reconrow.getCell((short) 20);
                    //Update the value of cell
                    if (ReconCellU1 == null) {
                        ReconCellU1 = Reconrow.createCell(20);
                    }
                    //  else{
                    double ReconU1Val = ReconCellU1.getNumericCellValue();
                    //   }

                    XSSFCell ReconCellV1 = Reconrow.getCell((short) 21);
                    //Update the value of cell
                    if (ReconCellV1 == null) {
                        ReconCellV1 = Reconrow.createCell(21);
                    }

                    double ReconV1Val = ReconCellV1.getNumericCellValue();


                    if (j1Val.equals("RETIREMENT") || j1Val.equals("DEATH") && (date1.after(Startdate) && date1.before(date2)) && a1Val.equals(ReconA1Val)) {
                        j1Val = "TERMINATED";
                   /*       System.out.print("A1: " + a1Val);
                System.out.print(" C1: " + c1Val);
                System.out.print(" D1: " + d1Val);
                System.out.print(" J1: " + j1Val);
                System.out.println();*/
                        //  break;
                    }

                    //  if (c1Val.equals(ReconC1Val) && d1Val.equals(ReconB1Val)) {
                    if (j1Val.equals("TERMINATED") && !a1Val.equals("ASSE88888") && a1Val.equals(ReconA1Val) && c1Val.equals(ReconC1Val) && d1Val.equals(ReconB1Val) && (date1.after(Startdate) && date1.before(date2))) {

                      /*  //  System.out.println("******Actives******");
                        System.out.print("A1: " + ReconA1Val);
                        System.out.print(" C1: " + ReconB1Val);
                        System.out.print(" D1: " + ReconC1Val);
                        System.out.print(" T1: " + ReconT1Val);
                        System.out.print(" U1: " + ReconU1Val);
                        System.out.print(" V1: " + ReconV1Val);
                        System.out.println();*/
                        TActiveSum += ReconT1Val + ReconU1Val + ReconV1Val;

                        //    }
                    }
                    //   Roww++;
                    count++;
                }

            }
            //  System.out.println("Terminee Sum: " + TActiveSum);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (NullPointerException e) {
            e.printStackTrace();

        }
        return TActiveSum;
    }

    public double ActiveSumReader(String StartDate, String EndDate, String Recon) throws IndexOutOfBoundsException {

        double ActiveSum = 0;
        try {
            FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Valuation Data.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

            XSSFSheet Demosheet = workbook.getSheet("DEMO");
            XSSFSheet Reconsheet = workbook.getSheet(Recon);
            String T = EndDate;
            String start = StartDate;


            int DEMOrowCount = Demosheet.getPhysicalNumberOfRows();
            int ReconrowCount = Reconsheet.getPhysicalNumberOfRows() + 1;


            ActiveSum = 0;
            //    double TActiveSum = 0;
            double TermineeSum = 0;

            int count = 0;
            for (int row = 0; row < DEMOrowCount; row++) {
                int Row = row;

                if (Row == 0) {
                    row = 1;  //start reading from second row
                }
                XSSFRow row1 = Demosheet.getRow(row);

                XSSFCell cellA1 = row1.getCell((short) 0);
                String result = cellA1.getStringCellValue();

                String a1Val = result.replaceAll("[-]", "");


                XSSFCell cellC1 = row1.getCell((short) 2);
                String c1Val = cellC1.getStringCellValue();

                XSSFCell cellD1 = row1.getCell((short) 3);
                String d1Val = cellD1.getStringCellValue();

                XSSFCell cellI1 = row1.getCell((short) 8);
                Date celli1 = cellI1.getDateCellValue();
                SimpleDateFormat datetemp = new SimpleDateFormat("dd-MMM-yy");
                SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yy");

                //Date cellValue = datetemp.parse(i1Val);
                String i1Val = sdf.format(celli1);
                Date date1 = null;
                Date date2 = null;
                Date Startdate = null;

                try {
                    date1 = sdf.parse(i1Val);
                    date2 = sdf.parse(T);
                    Startdate = sdf.parse(start);
                } catch (ParseException e) {
                    e.printStackTrace();
                }

                XSSFCell cellJ1 = row1.getCell((short) 9);
                String j1Val = cellJ1.getStringCellValue();

                String temp = j1Val;

              /*  System.out.print("A1: " + a1Val);
                System.out.print(" C1: " + c1Val);
                System.out.print(" D1: " + d1Val);
                System.out.print(" J1: " + j1Val);
                System.out.println();*/

                // int Roww = 6;

                count = 0;

                for (int Col = 6; Col < ReconrowCount; Col++) {
                    XSSFRow Reconrow = Reconsheet.getRow(Col);

                    XSSFCell ReconCellA1 = Reconrow.getCell((short) 0);
                    //Update the value of cell
                    if (ReconCellA1 == null) {
                        ReconCellA1 = Reconrow.createCell(0);
                    }
                    String ReconA1Val = ReconCellA1.getStringCellValue();

                    XSSFCell ReconCellB1 = Reconrow.getCell((short) 1);
                    //Update the value of cell
                    if (ReconCellB1 == null) {
                        ReconCellB1 = Reconrow.createCell(1);
                    }
                    String ReconB1Val = ReconCellB1.getStringCellValue();

                    XSSFCell ReconCellC1 = Reconrow.getCell((short) 2);
                    //Update the value of cell
                    if (ReconCellC1 == null) {
                        ReconCellC1 = Reconrow.createCell(2);
                    }
                    String ReconC1Val = ReconCellC1.getStringCellValue();

                    XSSFCell ReconCellT1 = Reconrow.getCell((short) 19);
                    //Update the value of cell
                    if (ReconCellT1 == null) {
                        ReconCellT1 = Reconrow.createCell(19);
                    }
                    double ReconT1Val = ReconCellT1.getNumericCellValue();


                    XSSFCell ReconCellU1 = Reconrow.getCell((short) 20);
                    //Update the value of cell
                    if (ReconCellU1 == null) {
                        ReconCellU1 = Reconrow.createCell(20);
                    }
                    //  else{
                    double ReconU1Val = ReconCellU1.getNumericCellValue();
                    //   }

                    XSSFCell ReconCellV1 = Reconrow.getCell((short) 21);
                    //Update the value of cell
                    if (ReconCellV1 == null) {
                        ReconCellV1 = Reconrow.createCell(21);
                    }

                    double ReconV1Val = ReconCellV1.getNumericCellValue();


                    if (j1Val.equals("DEATH") || j1Val.equals("DEFERRED") || j1Val.equals("RETIREMENT") || j1Val.equals("TERMINATED") && date1.after(date2) || date1.equals(date2)) {
                        j1Val = "ACTIVE";
                    }

                    //  if (c1Val.equals(ReconC1Val) && d1Val.equals(ReconB1Val)) {
                    if (j1Val.equals("ACTIVE") && !a1Val.equals("ASSE88888") && a1Val.equals(ReconA1Val) && c1Val.equals(ReconC1Val) && d1Val.equals(ReconB1Val) && !(date1.after(Startdate) && date1.before(date2))) {

                        //  System.out.println("******Actives******");
                      /*  System.out.print("A1: " + ReconA1Val);
                        System.out.print(" C1: " + ReconB1Val);
                        System.out.print(" D1: " + ReconC1Val);
                        System.out.print(" T1: " + ReconT1Val);
                        System.out.print(" U1: " + ReconU1Val);
                        System.out.print(" V1: " + ReconV1Val);
                        System.out.println();*/
                        ActiveSum += ReconT1Val + ReconU1Val + ReconV1Val;

                        //    }

                    }
                    //   Roww++;
                    // count++;
                }

                j1Val = temp;
                /*/*//***********TERMINATIONS********************************
                 for (int Col = 6; Col < ReconrowCount; Col++) {
                 XSSFRow Reconrow = Reconsheet.getRow(Col);

                 XSSFCell ReconCellA1 = Reconrow.getCell((short) 0);
                 //Update the value of cell
                 if (ReconCellA1 == null) {
                 ReconCellA1 = Reconrow.createCell(0);
                 }
                 String ReconA1Val = ReconCellA1.getStringCellValue();

                 XSSFCell ReconCellB1 = Reconrow.getCell((short) 1);
                 //Update the value of cell
                 if (ReconCellB1 == null) {
                 ReconCellB1 = Reconrow.createCell(1);
                 }
                 String ReconB1Val = ReconCellB1.getStringCellValue();

                 XSSFCell ReconCellC1 = Reconrow.getCell((short) 2);
                 //Update the value of cell
                 if (ReconCellC1 == null) {
                 ReconCellC1 = Reconrow.createCell(2);
                 }
                 String ReconC1Val = ReconCellC1.getStringCellValue();

                 XSSFCell ReconCellT1 = Reconrow.getCell((short) 19);
                 //Update the value of cell
                 if (ReconCellT1 == null) {
                 ReconCellT1 = Reconrow.createCell(19);
                 }
                 double ReconT1Val = ReconCellT1.getNumericCellValue();


                 XSSFCell ReconCellU1 = Reconrow.getCell((short) 20);
                 //Update the value of cell
                 if (ReconCellU1 == null) {
                 ReconCellU1 = Reconrow.createCell(20);
                 }
                 //  else{
                 double ReconU1Val = ReconCellU1.getNumericCellValue();
                 //   }

                 XSSFCell ReconCellV1 = Reconrow.getCell((short) 21);
                 //Update the value of cell
                 if (ReconCellV1 == null) {
                 ReconCellV1 = Reconrow.createCell(21);
                 }

                 double ReconV1Val = ReconCellV1.getNumericCellValue();


                 if (j1Val.equals("RETIREMENT")|| j1Val.equals("DEATH") && (date1.after(Startdate)&&date1.before(date2)) && a1Val.equals(ReconA1Val) ) {
                 j1Val = "TERMINATED";
                 *//*       System.out.print("A1: " + a1Val);
                System.out.print(" C1: " + c1Val);
                System.out.print(" D1: " + d1Val);
                System.out.print(" J1: " + j1Val);
                System.out.println();*//*
                        //  break;
                    }

                    //  if (c1Val.equals(ReconC1Val) && d1Val.equals(ReconB1Val)) {
                    if (j1Val.equals("TERMINATED") && !a1Val.equals("ASSE88888") && a1Val.equals(ReconA1Val) && c1Val.equals(ReconC1Val) && d1Val.equals(ReconB1Val)&& (date1.after(Startdate)&&date1.before(date2))) {

                        //  System.out.println("******Actives******");
                        System.out.print("A1: " + ReconA1Val);
                        System.out.print(" C1: " + ReconB1Val);
                        System.out.print(" D1: " + ReconC1Val);
                        System.out.print(" T1: " + ReconT1Val);
                        System.out.print(" U1: " + ReconU1Val);
                        System.out.print(" V1: " + ReconV1Val);
                        System.out.println();
                        TActiveSum += ReconT1Val + ReconU1Val + ReconV1Val;

                        //    }
                    }
                    //   Roww++;
                    count++;
                }
*/

            }
            //  System.out.println("Active Sum: " + ActiveSum);
            // System.out.println("Terminee Sum: " + TActiveSum);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (NullPointerException e) {
            e.printStackTrace();

        }
        return ActiveSum;
    }

    //seperate actives
    public String Separate_Actives_Terminees(String workingDir) throws IndexOutOfBoundsException {
        // public String Separate_Actives_Terminees(String filePathValData,String filePathOutputTemplate, String WorkingDir) throws IndexOutOfBoundsException {
        StringBuilder stringBuilder = new StringBuilder();
        FileOutputStream outFile = null;
        FileInputStream fileInputStream = null;
        FileInputStream fileR = null;
        try {
            //  FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Valuation Data.xlsx");
            //   FileInputStream fileInputStream = new FileInputStream("C:\\Users\\akonowalchuk\\GFRAM\\Valuation Data.xlsx");
            //     FileInputStream fileInputStream = new FileInputStream(filePathValData);
            fileInputStream = new FileInputStream(workingDir + "\\Valuation Data.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet worksheet = workbook.getSheet("DEMO");
            //   XSSFSheet sheet = workbook.getSheetAt(0);

            int rowCount = worksheet.getPhysicalNumberOfRows();
            System.out.println(rowCount);


            int num = rowCount;
            //  int noOfColumns = sheet.getRow(num).getLastCellNum();

            //   FileInputStream fileR = new FileInputStream("C:\\Users\\akonowalchuk\\GFRAM\\template.xlsx");

            //  FileInputStream fileR = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\template.xlsx");
            //    FileInputStream fileR = new FileInputStream(filePathOutputTemplate);
            fileR = new FileInputStream(workingDir + "\\Template_Separated.xlsx");
            // fileR = new FileInputStream(workingDir + "\\template.xlsx");
            XSSFWorkbook workbookR = new XSSFWorkbook(fileR);
            XSSFSheet sheetR = workbookR.getSheetAt(0);

            ////////////////////////
            XSSFSheet sheetW = workbookR.getSheet("Terminees");
            ////////////////////////

            SimpleDateFormat datetemp = new SimpleDateFormat("dd-MMM-yy");
            // Date cellValue = null;

            //    cellValue = datetemp.parse("1994-01-01");

            XSSFRow[] rowR = new XSSFRow[num];
            Cell cellR = null;

            ArrayList arraylist = new ArrayList();
            int counter = 1;
            int TermineesStartRow = 1;
            for (int row = 0; row < num; row++) {
                int Row = row;

                if (Row == 0) {
                    row = 1;  //start reading from second row
                }
                XSSFRow row1 = worksheet.getRow(row);

                XSSFCell cellA1 = row1.getCell((short) 0);
                String a1Val = cellA1.getStringCellValue();

                XSSFCell cellB1 = row1.getCell((short) 1);
                String b1Val = cellB1.getStringCellValue();

                XSSFCell cellC1 = row1.getCell((short) 2);
                String c1Val = cellC1.getStringCellValue();

                XSSFCell cellD1 = row1.getCell((short) 3);
                String d1Val = cellD1.getStringCellValue();

                XSSFCell cellE1 = row1.getCell((short) 4);
                String e1Val = cellE1.getStringCellValue();
                String f1Val = null;
                try {
                    XSSFCell cellF1 = row1.getCell((short) 5);
                    //  Date f1Vall = cellF1.getDateCellValue();
                    f1Val = datetemp.format(cellF1.getDateCellValue());
                } catch (NullPointerException N) {

                }

                XSSFCell cellG1 = row1.getCell((short) 6);
                String g1Val = datetemp.format(cellG1.getDateCellValue());

                XSSFCell cellH1 = row1.getCell((short) 7);
                String h1Val = datetemp.format(cellH1.getDateCellValue());

                //   Date cellValue = null;
                XSSFCell cellI1 = row1.getCell((short) 8);
        /*        try {
                    cellValue = datetemp.parse(String.valueOf(cellI1.getDateCellValue()));
                } catch (ParseException e) {
                    e.printStackTrace();
                }*/
                String i1Val = datetemp.format(cellI1.getDateCellValue());

                XSSFCell cellJ1 = row1.getCell((short) 9);
                String j1Val = cellJ1.getStringCellValue();
                j1Val= j1Val.toUpperCase();

                if (j1Val.equals("ACTIVE") && !(d1Val.equals("KEY"))) {
                    System.out.print("A1: " + a1Val);
                    System.out.print(" B1: " + b1Val);
                    System.out.print(" C1: " + c1Val);
                    System.out.print(" D1: " + d1Val);
                    System.out.print(" E1: " + e1Val);
                    System.out.print(" F1: " + f1Val);
                    System.out.print(" G1: " + g1Val);
                    System.out.print(" H1: " + h1Val);
                    System.out.print(" I1: " + i1Val);
                    System.out.print(" J1: " + j1Val);

                    stringBuilder.append("Employee Number: " + a1Val + "\n");
                    stringBuilder.append("Last Name: " + c1Val + "\n");
                    stringBuilder.append("First Name: " + d1Val + "\n");
                    stringBuilder.append("Date of Birth: " + f1Val + "\n");
                    stringBuilder.append("Plan Entry: " + g1Val + "\n");
                    stringBuilder.append("Employment Date: " + h1Val + "\n");
                    stringBuilder.append("Status Date: " + i1Val + "\n");
                    stringBuilder.append("Status: " + j1Val + "\n");
                    stringBuilder.append("----------------------------------\n");

                    System.out.println();

                    arraylist.add(0, a1Val);
                    arraylist.add(1, b1Val);
                    arraylist.add(2, c1Val);
                    arraylist.add(3, d1Val);
                    arraylist.add(4, e1Val);
                    arraylist.add(5, f1Val);
                    arraylist.add(6, g1Val);
                    arraylist.add(7, h1Val);
                    arraylist.add(8, i1Val);
                    arraylist.add(9, j1Val);


                    rowR[Row] = sheetR.createRow(counter++);
                    // }
                    for (int Col = 0; Col < 10; Col++) {
                        //Update the value of cell
                        //   cellR = rowR[Row].getCell(Col);
                        //   if (cellR == null) {
                        cellR = rowR[Row].createCell(Col);
                        //   }
/*                        if(arraylist.get(Col) instanceof Date)
                            cellR.setCellValue((Date)arraylist.get(Col));
                        else if(arraylist.get(Col) instanceof Boolean)
                            cellR.setCellValue((Boolean)arraylist.get(Col));
                        else if(arraylist.get(Col) instanceof String)
                            cellR.setCellValue((String)arraylist.get(Col));
                        else if(arraylist.get(Col) instanceof Double)
                            cellR.setCellValue((Double)arraylist.get(Col));*/
                       cellR.setCellValue(String.valueOf(arraylist.get(Col)));
                    }

                }
                //  int c=0;
                if (j1Val.equals("TERMINATED") || (j1Val.equals("DEATH") || j1Val.equals("DEFERRED") || j1Val.equals("RETIREMENT"))) {
                    System.out.print("A1: " + a1Val);
                    System.out.print(" B1: " + b1Val);
                    System.out.print(" C1: " + c1Val);
                    System.out.print(" D1: " + d1Val);
                    System.out.print(" E1: " + e1Val);
                    System.out.print(" F1: " + f1Val);
                    System.out.print(" G1: " + g1Val);
                    System.out.print(" H1: " + h1Val);
                    System.out.print(" I1: " + i1Val);
                    System.out.print(" J1: " + j1Val);

                    stringBuilder.append("Employee Number: " + a1Val + "\n");
                    stringBuilder.append("Last Name: " + c1Val + "\n");
                    stringBuilder.append("First Name: " + d1Val + "\n");
                    stringBuilder.append("Date of Birth: " + f1Val + "\n");
                    stringBuilder.append("Plan Entry: " + g1Val + "\n");
                    stringBuilder.append("Employment Date: " + h1Val + "\n");
                    stringBuilder.append("Status Date: " + i1Val + "\n");
                    stringBuilder.append("Status: " + j1Val + "\n");
                    stringBuilder.append("----------------------------------\n");

                  /*  try {
                        Thread.sleep(5000);
                    } catch (InterruptedException e) {
                        e.printStackTrace();
                    }*/
                    //   }
                    System.out.println();

                    arraylist.add(0, a1Val);
                    arraylist.add(1, b1Val);
                    arraylist.add(2, c1Val);
                    arraylist.add(3, d1Val);
                    arraylist.add(4, e1Val);
                    arraylist.add(5, f1Val);
                    arraylist.add(6, g1Val);
                    arraylist.add(7, h1Val);
                    arraylist.add(8, i1Val);
                    arraylist.add(9, j1Val);

                    //  setList(arraylist);

                    //   int Row = row;

                    rowR[Row] = sheetW.createRow(TermineesStartRow++);
                    // }
                    for (int Col = 0; Col < 10; Col++) {
                        //Update the value of cell
                        //   cellR = rowR[Row].getCell(Col);
                        //   if (cellR == null) {
                        cellR = rowR[Row].createCell(Col);
                        //   }
  /*                      if(arraylist.get(Col) instanceof Date)
                            cellR.setCellValue((Date)arraylist.get(Col));
                        else if(arraylist.get(Col) instanceof Boolean)
                            cellR.setCellValue((Boolean)arraylist.get(Col));
                        else if(arraylist.get(Col) instanceof String)
                            cellR.setCellValue((String)arraylist.get(Col));
                        else if(arraylist.get(Col) instanceof Double)
                            cellR.setCellValue((Double)arraylist.get(Col));*/
                        cellR.setCellValue(String.valueOf(arraylist.get(Col)));
                    }

                }

                //    FileOutputStream outFile = new FileOutputStream(new File("C:\\Users\\akonowalchuk\\GFRAM\\Seperated Members.xlsx"));
            }
            //Auto size all the columns
            for (int x = 0; x < sheetW.getRow(0).getPhysicalNumberOfCells(); x++) {
                sheetW.autoSizeColumn(x);
                sheetR.autoSizeColumn(x);
            }
            outFile = new FileOutputStream(new File(workingDir + "\\Seperated Members.xlsx"));
            workbookR.write(outFile);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (NullPointerException e) {
            e.printStackTrace();

        }
        finally {
            try {
                fileR.close();
                outFile.close();
                fileInputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }

        }
        return String.valueOf(stringBuilder);
    }
}

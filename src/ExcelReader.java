import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import sun.rmi.runtime.Log;

import java.io.*;
import java.lang.reflect.Array;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.Period;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.*;

import static java.util.Calendar.DATE;
import static java.util.Calendar.MONTH;
import static java.util.Calendar.YEAR;

/**
 * Created by Ayshahvez on 11/10/2016.
 */

public class ExcelReader {

static Utility utility = new Utility();

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
            FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Hose Valuation Data (Actuary's copy).xlsx");
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
            FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Hose Valuation Data (Actuary's copy).xlsx");
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

    public String Separate_Actives_Terminees(String workingDir) throws IndexOutOfBoundsException {
        // public String Separate_Actives_Terminees(String filePathValData,String filePathOutputTemplate, String WorkingDir) throws IndexOutOfBoundsException {
        StringBuilder stringBuilder = new StringBuilder();
        FileOutputStream outFile = null;
        FileInputStream fileInputStream = null;
        FileInputStream fileR = null;
        try {
            //  FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Hose Valuation Data (Actuary's copy).xlsx");
            //   FileInputStream fileInputStream = new FileInputStream("C:\\Users\\akonowalchuk\\GFRAM\\Hose Valuation Data (Actuary's copy).xlsx");
            //     FileInputStream fileInputStream = new FileInputStream(filePathValData);
            fileInputStream = new FileInputStream(workingDir + "\\Hose Valuation Data (Actuary's copy).xlsx");
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
            fileR = new FileInputStream(workingDir + "\\template.xlsx");
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
                    for (int Col = 0; Col <= 9; Col++) {
                        //Update the value of cell
                        //   cellR = rowR[Row].getCell(Col);
                        //   if (cellR == null) {
                        cellR = rowR[Row].createCell(Col);
                        //   }
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
                    for (int Col = 0; Col <= 9; Col++) {
                        //Update the value of cell
                        //   cellR = rowR[Row].getCell(Col);
                        //   if (cellR == null) {
                        cellR = rowR[Row].createCell(Col);
                        //   }
                        cellR.setCellValue(String.valueOf(arraylist.get(Col)));
                    }

                }

                //    FileOutputStream outFile = new FileOutputStream(new File("C:\\Users\\akonowalchuk\\GFRAM\\Output.xlsx"));


            }
            outFile = new FileOutputStream(new File(workingDir + "\\Output.xlsx"));
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

    public String Create_Actives_Sheet(String workingDir) throws IndexOutOfBoundsException {
       StringBuilder stringBuilder = new StringBuilder();
        SimpleDateFormat dF1 = new SimpleDateFormat();
        dF1.applyPattern("dd-MMM-yy");
        try {
           // FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Output.xlsx");
            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Output.xlsx");
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
                    stringBuilder.append("DOB: " + f1Val+"\n");
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
/*
        DateFormat df = new SimpleDateFormat("yyyy.MM.dd");
        Date startDate = null;
        Date endDate = null;
        try {
            startDate = df.parse(StartYear + "." + StartMonth + "." + StartDay);
            endDate = df.parse(EndYear + "." + EndMonth + "." + EndDay);
        */
        DecimalFormat dF = new DecimalFormat("#.##");//#.##
        DateFormat df = new SimpleDateFormat("dd-MMM-yy");
        SimpleDateFormat datetemp = new SimpleDateFormat("dd-MMM-yy");
        SimpleDateFormat datetemp2 = new SimpleDateFormat("dd-MM-yy");
        Date beginDate = null;
        Date endDate = null;

        try {

            FileInputStream fileInputStream = new FileInputStream(workingDir+"\\Output.xlsx");
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
            // int TermineesStartRow = 1;
           for(int u=0; u< 12;u++) {
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
                       stringBuilder.append("DOB: " + f1Val+"\n");
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
                          System.out.print(" K1: " + datetemp.format(k1Val));
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

    public void Create_Activee_Contribution(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {
/*        FileInputStream fileR = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Actives_Sheet.xlsx");
        XSSFWorkbook workbookR = new XSSFWorkbook(fileR);
        XSSFSheet CopyFromSheet = workbookR.getSheetAt(0);*/

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
            fs = new FileInputStream(workingDir+"\\Hose Valuation Data (Actuary's copy).xlsx");
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
                                                                                val.add(3, 0);


                                                                                for (int b = 0; b < val.size(); b++) {
                                                                                    cellR = rowR[Row].createCell(WriteAt + b);
                                                                                    cellR.setCellValue(String.valueOf(val.get(b)));
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

            WriteAt+=8;
             StartYear++; //comment out years if COMMENTED
            counter=7;

       }// END OF LOOP YEARS

        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Updated_Actives_Sheet.xlsx"));
        workbook.write(outFile);
        fileInputStream.close();
        outFile.close();
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
                fs = new FileInputStream(workingDir+ "\\Hose Valuation Data (Actuary's copy).xlsx");
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

                    if (a1Val.equals(a1) ) {

                        ArrayList val = new ArrayList();

                        System.out.print("A1: " + a1);
                        System.out.print(" D1: " + b1);
                        System.out.print(" C1: " + c1);
                    //    System.out.print("PS: " + i1Val);
                        System.out.print(" H1: " + h1);
                        System.out.print(" I1: " + i1);
                        System.out.print(" J1: " + j1);
                 /*       String test="";
                        double check = 0.05 *d[x];
                        check= Double.parseDouble(dF.format(check));
                        if (check == h1) {
                            test = "true";

                        } else {
                            test = "false";
                        }
                        System.out.print(" TEST: Pensionable Salary"+d[x]+"result "+check+test);*/

                        System.out.println();

                        val.add(0, h1);
                        val.add(1, i1);
                        val.add(2, j1);
                        val.add(3, 0);


                        for (int b = 0; b < val.size(); b++) {
                            cellR = rowR[Row].createCell(WriteAt + b);
                            cellR.setCellValue(String.valueOf(val.get(b)));
                        }



                        break;

                    }


                    //   counter=7;
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

    public String View_Actives_Members(String workingDir, String endDate) throws IndexOutOfBoundsException {
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("The following is a list of Active Members present as at "+endDate+" \n\n");
        SimpleDateFormat dF1 = new SimpleDateFormat();
        dF1.applyPattern("dd-MMM-yy");
        try {
            // FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Output.xlsx");
            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Output.xlsx");
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

                    stringBuilder.append("Employee ID: " + a1Val + "\n");
                    stringBuilder.append("Last Name: " + c1Val + "\n");
                    stringBuilder.append("First Name: " + d1Val + "\n");
                    stringBuilder.append("DOB: " + f1Val + "\n");
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

    public String View_Retired_Members(String workingDir, String endDate) throws IndexOutOfBoundsException {
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("The following is a list of Retired Members present as at "+endDate+" \n\n");
        SimpleDateFormat dF1 = new SimpleDateFormat();
        dF1.applyPattern("dd-MMM-yy");
        try {
            // FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Output.xlsx");
            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Output.xlsx");
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
                    stringBuilder.append("DOB: " + f1Val + "\n");
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
            // FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Output.xlsx");
            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Output.xlsx");
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
                    stringBuilder.append("DOB: " + f1Val + "\n");
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
            // FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Output.xlsx");
            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Output.xlsx");
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
                    stringBuilder.append("DOB: " + f1Val + "\n");
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
            // FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Output.xlsx");
            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Output.xlsx");
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
                    stringBuilder.append("DOB: " + f1Val + "\n");
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
            // FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Output.xlsx");
            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Output.xlsx");
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
                    stringBuilder.append("DOB: " + f1Val + "\n");
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

}

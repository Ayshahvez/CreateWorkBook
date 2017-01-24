import com.sun.org.apache.regexp.internal.RE;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.*;

import javax.rmi.CORBA.Util;
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

    public String result=null;

    public void setResult(String x){
        this.result=x;
    }

    public String getResult(){
        return this.result;
    }

    static Utility utility = new Utility();

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

    public void WriteActivesTotalRow(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException{
        DecimalFormat dF = new DecimalFormat("#.##");//#.##
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
        years+=1;

        double value=0;

        try {
            FileInputStream fileR = new FileInputStream(workingDir + "\\Accumulated_Actives_Sheet.xlsx");
            XSSFWorkbook workbookR = new XSSFWorkbook(fileR);
            XSSFSheet ActiveSheet = workbookR.getSheetAt(0);
            int numOfActives = ActiveSheet.getLastRowNum() + 1;

//MAIN PROCESSING
            //WRITE PENSIONABLE SALARY SUM
            XSSFRow rowtotal = ActiveSheet.createRow(numOfActives + 1);
            rowtotal.createCell(2).setCellValue("TOTAL");
            for(int k=0;k<years;k++){
                //GET DATA FROM TEMPLATE
                for (int x = 0, row = 7; row < numOfActives; x++, row++) {

                    //get the prior year adjustment if any
                    XSSFRow Row = ActiveSheet.getRow(row);
                    XSSFCell cellValues = Row.getCell(8+k);
                    if (cellValues == null) {
                        cellValues = Row.createCell(8);
                        cellValues.setCellValue(0.00);
                    }
                    value += cellValues.getNumericCellValue();
                }
                rowtotal = ActiveSheet.getRow(numOfActives + 1);
                rowtotal.createCell(8+k).setCellValue(Double.parseDouble(dF.format(value)));
                value=0;
            }

            value=0;
            int index=(8+years)+2;
            int cellNumbers=(years*8)+5;// 9 for fees 8 for no fees; 5 to get the last 4 accumultion values and total for account balances
           // int feeIndex=index+8;

            //WRITE ACCUMULATION AND CONTRIBUTION SUMS
            for(int k=0;k<cellNumbers;k++) {

             //   if(index!=feeIndex){
                    //GET DATA FROM TEMPLATE
                    for (int x = 0, row = 7; row < numOfActives; x++, row++) {

                        //get the prior year adjustment if any
                        XSSFRow Row = ActiveSheet.getRow(row);
                        XSSFCell cellValues = Row.getCell(index);
                        if (cellValues == null) {
                            cellValues = Row.createCell(index);
                            cellValues.setCellValue(0.00);
                        }
                        value += cellValues.getNumericCellValue();
                    }

          /*      }
                else{
                    feeIndex+=9;
                }*/

                rowtotal = ActiveSheet.getRow(numOfActives + 1);
                rowtotal.createCell(index).setCellValue(Double.parseDouble(dF.format(value)));
                value=0;
                index+=1;
            }


            //write TTOAL FOR MALE AND FEMALE
            rowtotal = ActiveSheet.createRow(numOfActives + 3);
            rowtotal.createCell(2).setCellValue("Total Males");

            rowtotal = ActiveSheet.createRow(numOfActives + 4);
            rowtotal.createCell(2).setCellValue("Total Females");


            rowtotal = ActiveSheet.getRow(numOfActives + 3);
            rowtotal.createCell(7).setCellValue("Average for Males");

            rowtotal = ActiveSheet.getRow(numOfActives + 4);
            rowtotal.createCell(7).setCellValue("Average for Females");



//GET GENDER
            int maleCount=0;
            int femaleCount=0;
            for (int row = 7, I=8; row < numOfActives; row++,I++) {
                //GET THE FUND AT BEGINNING OF PERIOD IF ANY
                XSSFRow Row = ActiveSheet.getRow(row);

                XSSFCell cellGender = Row.getCell(3);
                if (cellGender == null) {
                    cellGender = Row.createCell(3);
                    cellGender.setCellValue("");
                }
                String memberGender= cellGender.getStringCellValue();
                memberGender=memberGender.toLowerCase();

                System.out.println("F: "+femaleCount);
                System.out.println("m: "+maleCount);
                if (memberGender.equals("f"))femaleCount++;
                if (memberGender.equals("m"))maleCount++;
            }

            rowtotal = ActiveSheet.getRow(numOfActives + 1);
            rowtotal.createCell(3).setCellValue(maleCount+femaleCount);

            rowtotal = ActiveSheet.getRow(numOfActives + 3);
            rowtotal.createCell(3).setCellValue(maleCount);

            rowtotal = ActiveSheet.getRow(numOfActives + 4);
            rowtotal.createCell(3).setCellValue(femaleCount);

            //AVERAGE PENSIONABLE SALARY FOR MALE AND FEMALE
            double femalePensionableSalary=0;
            double malePensionableSalary=0;
            double pensionableSalaryAmount=0;
            int moveIndex=8;
            for(int k=0;k<years+2;k++){
                //GET DATA FROM TEMPLATE
                for (int x = 0, row = 7; row < numOfActives; x++, row++) {

                    //GET THE FUND AT BEGINNING OF PERIOD IF ANY
                    XSSFRow Row = ActiveSheet.getRow(row);

                    XSSFCell cellGender = Row.getCell(3);
                    if (cellGender == null) {
                        cellGender = Row.createCell(3);
                        cellGender.setCellValue("");
                    }
                    String memberGender= cellGender.getStringCellValue();
                    memberGender=memberGender.toLowerCase();


                    //get the prior year adjustment if any
                    XSSFRow Row1 = ActiveSheet.getRow(row);
                    XSSFCell cellValues = Row1.getCell(8+k);
                    if (cellValues == null) {
                        cellValues = Row1.createCell(8);
                        cellValues.setCellValue(0.00);
                    }
                    pensionableSalaryAmount = cellValues.getNumericCellValue();
                    if (memberGender.equals("f")) femalePensionableSalary+=pensionableSalaryAmount;
                    if (memberGender.equals("m")) malePensionableSalary+=pensionableSalaryAmount;

                }
                rowtotal = ActiveSheet.getRow(numOfActives + 3);
                rowtotal.createCell(moveIndex).setCellValue(Double.parseDouble(dF.format(malePensionableSalary/maleCount)));

                rowtotal = ActiveSheet.getRow(numOfActives + 4);
                rowtotal.createCell(moveIndex).setCellValue(Double.parseDouble(dF.format(femalePensionableSalary/femaleCount)));
                femalePensionableSalary=0;
                malePensionableSalary=0;
                pensionableSalaryAmount=0;
                moveIndex+=1;
            }

            FileOutputStream outFile = new FileOutputStream(new File(workingDir +"\\Completed_Actives_Sheet.xlsx"));
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

    }

    public void WriteFeesActivesTotalRow(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException{

        DecimalFormat dF = new DecimalFormat("#.##");//#.##
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
        years+=1;

        double value=0;

        try {
            FileInputStream fileR = new FileInputStream(workingDir + "\\Accumulated_Actives_Sheet.xlsx");
            XSSFWorkbook workbookR = new XSSFWorkbook(fileR);
            XSSFSheet ActiveSheet = workbookR.getSheetAt(0);
            int numOfActives = ActiveSheet.getLastRowNum() + 1;
        //    int numOfActives = Utility.getNumberOfMembersInSheet(workbookR,ActiveSheet);

//MAIN PROCESSING
            //WRITE PENSIONABLE SALARY SUM
            XSSFRow rowtotal = ActiveSheet.createRow(numOfActives + 1);
            rowtotal.createCell(2).setCellValue("TOTAL");

            for(int k=0;k<years;k++){
            //GET DATA FROM TEMPLATE
            for (int x = 0, row = 7; row < numOfActives; x++, row++) {

                //get the prior year adjustment if any
                XSSFRow Row = ActiveSheet.getRow(row);
                XSSFCell cellValues = Row.getCell(8+k);
                if (cellValues == null) {
                    cellValues = Row.createCell(8);
                    cellValues.setCellValue(0.00);
                }
                value += cellValues.getNumericCellValue();
            }
            rowtotal = ActiveSheet.getRow(numOfActives + 1);
            rowtotal.createCell(8+k).setCellValue(Double.parseDouble(dF.format(value)));
            value=0;
        }

            value=0;
int index=(8+years)+2;
            int cellNumbers=(years*9)+5;// 9 for fees 8 for no fees; 5 to get the last 4 accumultion values and total for account balances
//int feeIndex=index+8;
        //WRITE ACCUMULATION AND CONTRIBUTION SUMS
            for(int k=0;k<cellNumbers;k++) {

           //        if(index!=feeIndex){
                //GET DATA FROM TEMPLATE
                for (int x = 0, row = 7; row < numOfActives; x++, row++) {

                    //get the prior year adjustment if any
                    XSSFRow Row = ActiveSheet.getRow(row);
                    XSSFCell cellValues = Row.getCell(index);
                    if (cellValues == null) {
                        cellValues = Row.createCell(index);
                        cellValues.setCellValue(0.00);
                    }
                    value += cellValues.getNumericCellValue();
                }

          //  }
        //    else{
             //          feeIndex+=9;
              //     }

                    rowtotal = ActiveSheet.getRow(numOfActives + 1);
                    rowtotal.createCell(index).setCellValue(Double.parseDouble(dF.format(value)));
                value=0;
                index+=1;
            }


            //write TTOAL FOR MALE AND FEMALE
          rowtotal = ActiveSheet.createRow(numOfActives + 3);
            rowtotal.createCell(2).setCellValue("Total Males");

            rowtotal = ActiveSheet.createRow(numOfActives + 4);
            rowtotal.createCell(2).setCellValue("Total Females");




            rowtotal = ActiveSheet.getRow(numOfActives + 3);
            rowtotal.createCell(7).setCellValue("Average for Males");

            rowtotal = ActiveSheet.getRow(numOfActives + 4);
            rowtotal.createCell(7).setCellValue("Average for Females");



//GET GENDER
            int maleCount=0;
            int femaleCount=0;
            for (int row = 7, I=0; row < numOfActives; row++,I++) {
                //GET THE FUND AT BEGINNING OF PERIOD IF ANY
                XSSFRow Row = ActiveSheet.getRow(row);

                XSSFCell cellGender = Row.getCell(3);
                if (cellGender == null) {
                    cellGender = Row.createCell(3);
                    cellGender.setCellValue("");
                }
               String memberGender= cellGender.getStringCellValue();
memberGender=memberGender.toLowerCase();

              //  System.out.println("F: "+femaleCount);
            //    System.out.println("m: "+maleCount);
                if (memberGender.equals("f"))femaleCount++;
                if (memberGender.equals("m"))maleCount++;
            }

            rowtotal = ActiveSheet.getRow(numOfActives + 1);
            rowtotal.createCell(3).setCellValue(maleCount+femaleCount);

            rowtotal = ActiveSheet.getRow(numOfActives + 3);
            rowtotal.createCell(3).setCellValue(maleCount);

            rowtotal = ActiveSheet.getRow(numOfActives + 4);
            rowtotal.createCell(3).setCellValue(femaleCount);

            //AVERAGE PENSIONABLE SALARY FOR MALE AND FEMALE
            double femalePensionableSalary=0;
            double malePensionableSalary=0;
            double pensionableSalaryAmount=0;
            int moveIndex=8;

            for(int k=0;k<years+2;k++){
                //GET DATA FROM TEMPLATE
                for (int x = 0, row = 7; row < numOfActives; x++, row++) {
                    System.out.println("x:" + x);
System.out.println("numOfActives:" + numOfActives);
                        //GET THE FUND AT BEGINNING OF PERIOD IF ANY
                        XSSFRow Row = ActiveSheet.getRow(row);

                        XSSFCell cellGender = Row.getCell(3);
                        if (cellGender == null) {
                            cellGender = Row.createCell(3);
                            cellGender.setCellValue("");
                        }
                        String memberGender= cellGender.getStringCellValue();
                        memberGender=memberGender.toLowerCase();


                    //get the prior year adjustment if any
                    XSSFRow Row1 = ActiveSheet.getRow(row);
                    XSSFCell cellValues = Row1.getCell(8+k);
                    if (cellValues == null) {
                        cellValues = Row1.createCell(8);
                        cellValues.setCellValue(0.00);
                    }
                    pensionableSalaryAmount = cellValues.getNumericCellValue();
                    if (memberGender.equals("f")) femalePensionableSalary+=pensionableSalaryAmount;
                    if (memberGender.equals("m")) malePensionableSalary+=pensionableSalaryAmount;

                }

                rowtotal = ActiveSheet.getRow(numOfActives + 3);
                rowtotal.createCell(moveIndex).setCellValue(Double.parseDouble(dF.format(malePensionableSalary/maleCount)));

                rowtotal = ActiveSheet.getRow(numOfActives + 4);
                rowtotal.createCell(moveIndex).setCellValue(Double.parseDouble(dF.format(femalePensionableSalary/femaleCount)));
                femalePensionableSalary=0;
                malePensionableSalary=0;
                pensionableSalaryAmount=0;
                moveIndex+=1;
            }

            FileOutputStream outFile = new FileOutputStream(new File(workingDir +"\\Completed_Actives_Sheet.xlsx"));
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

    }

    public void WriteTermineeTotalRow(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException{
        DecimalFormat dF = new DecimalFormat("#.##");//#.##
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
        years+=1;

        double value=0;

        try {
            FileInputStream fileR = new FileInputStream(workingDir + "\\Accumulated_Terminee_Sheet.xlsx");
            XSSFWorkbook workbookR = new XSSFWorkbook(fileR);
            XSSFSheet ActiveSheet = workbookR.getSheetAt(0);
            int numOfActives = ActiveSheet.getLastRowNum() + 1;

            XSSFRow rowtotal = ActiveSheet.createRow(numOfActives + 1);
            rowtotal.createCell(2).setCellValue("TOTAL");

            value=0;
            int index=19;
            int cellNumbers=(years*8)+4+9;// 9 for fees 8 for no fees; 4 for last 4 accumultion values and total for account balances
           // int feeIndex=index+8;

            //WRITE ACCUMULATION AND CONTRIBUTION SUMS
            for(int k=0;k<cellNumbers;k++) {

              //  if(index!=feeIndex){
                    //GET DATA FROM TEMPLATE
                    for (int x = 0, row = 7; row < numOfActives; x++, row++) {

                        //get the prior year adjustment if any
                        XSSFRow Row = ActiveSheet.getRow(row);
                        XSSFCell cellValues = Row.getCell(index);
                        if (cellValues == null) {
                            cellValues = Row.createCell(index);
                            cellValues.setCellValue(0.00);
                        }
                        value += cellValues.getNumericCellValue();
                    }

             //   }
             //   else{
             //       feeIndex+=9;
             //   }

                rowtotal = ActiveSheet.getRow(numOfActives + 1);
                rowtotal.createCell(index).setCellValue(Double.parseDouble(dF.format(value)));
                value=0;
                index+=1;
            }




            rowtotal = ActiveSheet.createRow(numOfActives + 3);
            rowtotal.createCell(2).setCellValue("Total Males");

            rowtotal = ActiveSheet.createRow(numOfActives + 4);
            rowtotal.createCell(2).setCellValue("Total Females");



//GET GENDER
            int maleCount=0;
            int femaleCount=0;
            for (int row = 7, I=8; row < numOfActives; row++,I++) {
                //GET THE FUND AT BEGINNING OF PERIOD IF ANY
                XSSFRow Row = ActiveSheet.getRow(row);

                XSSFCell cellGender = Row.getCell(3);
                if (cellGender == null) {
                    cellGender = Row.createCell(3);
                    cellGender.setCellValue("");
                }
                String memberGender= cellGender.getStringCellValue();
                memberGender=memberGender.toLowerCase();

                System.out.println("F: "+femaleCount);
                System.out.println("m: "+maleCount);
                if (memberGender.equals("f"))femaleCount++;
                if (memberGender.equals("m"))maleCount++;
            }

            //write TTOAL FOR MALE AND FEMALE
            rowtotal = ActiveSheet.getRow(numOfActives + 1);
            rowtotal.createCell(3).setCellValue(maleCount+femaleCount);


            rowtotal = ActiveSheet.getRow(numOfActives + 3);
            rowtotal.createCell(3).setCellValue(maleCount);

            rowtotal = ActiveSheet.getRow(numOfActives + 4);
            rowtotal.createCell(3).setCellValue(femaleCount);

            FileOutputStream outFile = new FileOutputStream(new File(workingDir +"\\Completed_Terminee_Sheet.xlsx"));
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

    }

    public void WriteFeesTermineeTotalRow(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException{
        DecimalFormat dF = new DecimalFormat("#.##");//#.##
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
        years+=1;

        double value=0;

        try {
            FileInputStream fileR = new FileInputStream(workingDir + "\\Accumulated_Terminee_Sheet.xlsx");
            XSSFWorkbook workbookR = new XSSFWorkbook(fileR);
            XSSFSheet ActiveSheet = workbookR.getSheetAt(0);
            int numOfActives = ActiveSheet.getLastRowNum() + 1;

//MAIN PROCESSING
            //WRITE PENSIONABLE SALARY SUM
            XSSFRow rowtotal = ActiveSheet.createRow(numOfActives + 1);
            rowtotal.createCell(2).setCellValue("TOTAL");
        /*    for(int k=0;k<years;k++){
                //GET DATA FROM TEMPLATE
                for (int x = 0, row = 7; row < numOfActives; x++, row++) {

                    //get the prior year adjustment if any
                    XSSFRow Row = ActiveSheet.getRow(row);
                    XSSFCell cellValues = Row.getCell(8+k);
                    if (cellValues == null) {
                        cellValues = Row.createCell(8);
                        cellValues.setCellValue(0.00);
                    }
                    value += cellValues.getNumericCellValue();
                }
                rowtotal = ActiveSheet.getRow(numOfActives + 1);
                rowtotal.createCell(8+k).setCellValue(Double.parseDouble(dF.format(value)));
                value=0;
            }*/

            value=0;
            int index=19;
            int cellNumbers=(years*9)+4+9;// 9 for fees 8 for no fees; 4 for last 4 accumultion values and total for account balances
            int feeIndex=index+8;

            //WRITE ACCUMULATION AND CONTRIBUTION SUMS
            for(int k=0;k<cellNumbers;k++) {

                if(index!=feeIndex){
                    //GET DATA FROM TEMPLATE
                    for (int x = 0, row = 7; row < numOfActives; x++, row++) {

                        //get the prior year adjustment if any
                        XSSFRow Row = ActiveSheet.getRow(row);
                        XSSFCell cellValues = Row.getCell(index);
                        if (cellValues == null) {
                            cellValues = Row.createCell(index);
                            cellValues.setCellValue(0.00);
                        }
                        value += cellValues.getNumericCellValue();
                    }

                }
                else{
                    feeIndex+=9;
                }

                rowtotal = ActiveSheet.getRow(numOfActives + 1);
                rowtotal.createCell(index).setCellValue(Double.parseDouble(dF.format(value)));
                value=0;
                index+=1;
            }


            //write TTOAL FOR MALE AND FEMALE
            rowtotal = ActiveSheet.createRow(numOfActives + 3);
            rowtotal.createCell(2).setCellValue("Total Males");

            rowtotal = ActiveSheet.createRow(numOfActives + 4);
            rowtotal.createCell(2).setCellValue("Total Females");

//GET GENDER
            int maleCount=0;
            int femaleCount=0;
            for (int row = 7, I=8; row < numOfActives; row++,I++) {
                //GET THE FUND AT BEGINNING OF PERIOD IF ANY
                XSSFRow Row = ActiveSheet.getRow(row);

                XSSFCell cellGender = Row.getCell(3);
                if (cellGender == null) {
                    cellGender = Row.createCell(3);
                    cellGender.setCellValue("");
                }
                String memberGender= cellGender.getStringCellValue();
                memberGender=memberGender.toLowerCase();

                System.out.println("F: "+femaleCount);
                System.out.println("m: "+maleCount);
                if (memberGender.equals("f"))femaleCount++;
                if (memberGender.equals("m"))maleCount++;
            }

            //write TTOAL FOR MALE AND FEMALE
            rowtotal = ActiveSheet.getRow(numOfActives + 1);
            rowtotal.createCell(3).setCellValue(maleCount+femaleCount);


            rowtotal = ActiveSheet.getRow(numOfActives + 3);
            rowtotal.createCell(3).setCellValue(maleCount);

            rowtotal = ActiveSheet.getRow(numOfActives + 4);
            rowtotal.createCell(3).setCellValue(femaleCount);

            FileOutputStream outFile = new FileOutputStream(new File(workingDir +"\\Completed_Terminee_Sheet.xlsx"));
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

    }

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
                                cellR.setCellValue(Double.parseDouble(dF.format(PensionableSalary.get(l))));
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
                cellR.setCellValue(Double.parseDouble(dF.format((Utility.betweenDates(startDate1, endDate1)/365.25))));

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

    public void Write_Members_Monetary_Values(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {

        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yy");

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

        int years = Utility.getDiffYears(startDate, endDate);//gets number of years
        years += 1;

        //get access to the actives template
        FileInputStream fileTemplate = new FileInputStream(workingDir + "\\Actives_Sheet.xlsx");
        XSSFWorkbook workbookTemplate = new XSSFWorkbook(fileTemplate);
        XSSFSheet Activesheet = workbookTemplate.getSheet("Actives");

        FileInputStream fileRecon = new FileInputStream(workingDir + "\\Input Sheet.xlsx");
        XSSFWorkbook workbookRecon = new XSSFWorkbook(fileRecon);

        //    XSSFSheet sheetTemplate = workbookTemplate.getSheet("Actives");
        ArrayList list = new ArrayList();

        //INDEXES
        int PensionableSalaryIndex=8;
        int membersAgeIndex=PensionableSalaryIndex+years;
        int PensionableServiceIndex = membersAgeIndex+1;
        int memberBasicStartIndex = PensionableServiceIndex+1;
        int memberVoluntaryStartIndex = memberBasicStartIndex+1;
        int employerBasicStartIndex = memberVoluntaryStartIndex+1;


        int memberBasicContribution_Index = PensionableServiceIndex+5;
        int memberVoluntaryContribution_Index = memberBasicContribution_Index+1;
        int employerContribution_Index = memberVoluntaryContribution_Index+1;
        int employerVoluntary_Index = employerContribution_Index+1;
        int Fees_Index = employerVoluntary_Index+1;

        for (int x = 0; x < years; x++) {

            Calendar cal = Calendar.getInstance();
            cal.set((StartYear), StartMonth, StartDay);
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy"); // Just the year, with 2 digits
            String formattedDate = sdf.format(cal.getTime());
            //  String formattedDate = "2014";
            String Recon = ("Actives at End of Plan Yr " + formattedDate);
         //   System.out.println("Actives at End of Plan Yr " + formattedDate);

//GET RECON SHEET
            XSSFSheet Reconsheet = workbookRecon.getSheet(Recon);
            int currentNumberofActiveMembers = Utility.getNumberOfMembersInSheet(Activesheet);//get curent members from recon row
            int currentNumberofReconMembers = Utility.getNumberOfMembersInSheet(Reconsheet);//get curent members from recon row

         //   XSSFRow[] rowWrite = new XSSFRow[currentNumberofActiveMembers];
            XSSFRow rowWrite = null;
            Cell cellWrite = null;
        //    System.out.println("currentNumberofActiveMembers" + currentNumberofActiveMembers);
            for (int row = 0, readFromRecon = 7; row < currentNumberofActiveMembers; row++, readFromRecon++) {

                XSSFRow rowPosition = Activesheet.getRow(readFromRecon);

                //get employee id
                XSSFCell cellA = rowPosition.getCell(0);  //employee number
                if (cellA == null) {
                    cellA = rowPosition.createCell((short) 0);
                    cellA.setCellValue("");
                }
                String result = cellA.getStringCellValue();
                String cellEmployeeID = result.replaceAll("[-]", "");

                //get LAST NAME
                XSSFCell cellB = rowPosition.getCell((short) 1);  //last name
                if (cellB == null) {
                    cellB = rowPosition.createCell((short) 1);
                    cellB.setCellValue("");
                }
                String cellLastName = cellB.getStringCellValue();

                //get FIRST NAME
                XSSFCell cellC = rowPosition.getCell((short) 2);  //first name
                if (cellC == null) {
                    cellC = rowPosition.createCell((short) 2);
                    cellC.setCellValue("");
                }
                String cellFirstName = cellC.getStringCellValue();

            //    System.out.println("currentNumberofReconMembers" + currentNumberofReconMembers);


                for (int readFromReconRow = 7, iterate = 0; iterate < currentNumberofReconMembers; readFromReconRow++,iterate++) {
                    XSSFRow reconRow = Reconsheet.getRow(readFromReconRow);
              //      System.out.println("readFromReconRow" + readFromReconRow);

                    //get employee id
                    XSSFCell reconCellA = reconRow.getCell(0);  //employee number
                    if (reconCellA == null) {
                        reconCellA = reconRow.createCell((short) 0);
                        reconCellA.setCellValue("");
                    }
                    String resultRecon = reconCellA.getStringCellValue();
                    String ReconcellEmployeeID = resultRecon.replaceAll("[-]", "");

                    //get LAST NAME
                    XSSFCell ReconCellB = reconRow.getCell((short) 1);  //last name
                    if (ReconCellB == null) {
                        ReconCellB = reconRow.createCell((short) 1);
                        ReconCellB.setCellValue("");
                    }
                    String ReconCellLastName = ReconCellB.getStringCellValue();

                    //get FIRST NAME
                    XSSFCell ReconCellC = reconRow.getCell((short) 2);  //first name
                    if (ReconCellC == null) {
                        ReconCellC = reconRow.createCell((short) 2);
                        ReconCellC.setCellValue("");
                    }
                    String ReconCellFirstName = ReconCellC.getStringCellValue();

                    //get DOB
                    XSSFCell ReconCellF = reconRow.getCell((short) 5);  //first name
                    if (ReconCellF == null) {
                        ReconCellF = reconRow.createCell((short) 5);
                       // ReconCellF.setCellValue("");
                    }
                    Date ReconCellDOB = ReconCellF.getDateCellValue();

                    //get DOB
                    XSSFCell ReconCellH = reconRow.getCell((short) 7);  //first name
                    if (ReconCellH == null) {
                        ReconCellH = reconRow.createCell((short) 7);
                        // ReconCellF.setCellValue("");
                    }
                    Date ReconCellDateofEnrolment = ReconCellH.getDateCellValue();

//GET PENSIONABLE SALARY
                    XSSFCell ReconCellL = reconRow.getCell((short) 11);  //first name
                    if (ReconCellL == null) {
                        ReconCellL = reconRow.createCell((short) 11);
                        ReconCellL.setCellValue(0.00);
                    }
                    double ReconCell_PensionableSalary = ReconCellL.getNumericCellValue();

                    //getINITIAL ACCUMUMATION AT START OF YEAR
                    XSSFCell ReconCellI = reconRow.getCell((short) 8);  //last name
                    if (ReconCellI == null) {
                        ReconCellI = reconRow.createCell((short) 8);
                        ReconCellI.setCellValue(0.00);
                    }
                    double ReconCell_MemberBasicContribution_atStart = ReconCellI.getNumericCellValue();

                    XSSFCell ReconCellJ= reconRow.getCell((short) 9);  //last name
                    if (ReconCellJ == null) {
                        ReconCellJ = reconRow.createCell((short) 9);
                        ReconCellJ.setCellValue(0.00);
                    }
                    double ReconCell_MemberVoluntaryContribution_atStart = ReconCellJ.getNumericCellValue();

                    XSSFCell ReconCellK= reconRow.getCell((short) 10);  //last name
                    if (ReconCellK == null) {
                        ReconCellK = reconRow.createCell((short) 10);
                        ReconCellK.setCellValue(0.00);
                    }
                    double ReconCell_EmployerContribution_atStart = ReconCellK.getNumericCellValue();

                    //employee basic contribution
                    XSSFCell ReconCellM= reconRow.getCell((short) 12);  //last name
                    if (ReconCellM == null) {
                        ReconCellM = reconRow.createCell((short) 12);
                        ReconCellM.setCellValue(0.00);
                    }
                    double ReconCell_EmployeeBasic_Contribution = ReconCellM.getNumericCellValue();

                    //employee voluntary contribution
                    XSSFCell ReconCellN= reconRow.getCell((short) 13);  //last name
                    if (ReconCellN == null) {
                        ReconCellN = reconRow.createCell((short) 13);
                        ReconCellN.setCellValue(0.00);
                    }
                    double ReconCell_EmployeeVoluntary_Contribution = ReconCellN.getNumericCellValue();


                    //employer contribution
                    XSSFCell ReconCellO= reconRow.getCell((short) 14);  //last name
                    if (ReconCellO == null) {
                        ReconCellO = reconRow.createCell((short) 14);
                        ReconCellO.setCellValue(0.00);
                    }
                    double ReconCell_Employer_Contribution = ReconCellO.getNumericCellValue();

//                    //FEES
//                    XSSFCell ReconCellY= reconRow.getCell((short) 24);  //last name
//                    if (ReconCellY == null) {
//                        ReconCellY = reconRow.createCell((short) 24);
//                        ReconCellY.setCellValue(0.00);
//                    }
//                    double ReconCell_Fees = ReconCellY.getNumericCellValue();
//


                    if(cellEmployeeID.equals(ReconcellEmployeeID)){

                        // write pensionable salary
                        rowWrite = Activesheet.getRow(readFromRecon);

                        cellWrite= rowWrite.createCell((PensionableSalaryIndex + x));
                        cellWrite.setCellValue(ReconCell_PensionableSalary);


                        //*************************MEMBERS AGE AS AT**********************

                        cellWrite = rowWrite.createCell(membersAgeIndex);
                        cellWrite.setCellValue(Utility.getDiffYears(ReconCellDOB, endDate));


//********PENSIONABLE SERVICE***********************
                        DecimalFormat dF = new DecimalFormat("#.##");//#.##
                        cellWrite = rowWrite.createCell(PensionableServiceIndex);
                        cellWrite.setCellValue(Double.parseDouble(dF.format((Utility.betweenDates(ReconCellDateofEnrolment, endDate)/365.25))));

                 //write memebr initial
                        if(x==0) {
                            cellWrite = rowWrite.createCell(memberBasicStartIndex);
                            cellWrite.setCellValue(ReconCell_MemberBasicContribution_atStart);

                            cellWrite = rowWrite.createCell(memberVoluntaryStartIndex);
                            cellWrite.setCellValue(ReconCell_MemberVoluntaryContribution_atStart);

                            cellWrite = rowWrite.createCell(employerBasicStartIndex);
                            cellWrite.setCellValue(ReconCell_EmployerContribution_atStart);


                        }

                        cellWrite = rowWrite.createCell((memberBasicContribution_Index));
                        cellWrite.setCellValue(ReconCell_EmployeeBasic_Contribution);

                          cellWrite = rowWrite.createCell(memberVoluntaryContribution_Index);
                           cellWrite.setCellValue(ReconCell_EmployeeVoluntary_Contribution);

                            cellWrite = rowWrite.createCell(employerContribution_Index);
                          cellWrite.setCellValue(ReconCell_Employer_Contribution);

                   /*     cellWrite = rowWrite.createCell(Fees_Index);
                        cellWrite.setCellValue(ReconCell_Fees);*/
                            }

                }//end of looping through each member in each year period

                //GET INITIAL BALANCES



            }
StartYear++;
            memberVoluntaryContribution_Index+=8;//9 for fees  || prob 8 for no fees
            memberBasicContribution_Index+=8;
            employerContribution_Index+=8;
            Fees_Index+=8;
        }//end of looping through the years


        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\temp2_Actives_Sheet.xlsx"));
        workbookTemplate.write(outFile);
        fileTemplate.close();
        fileRecon.close();
        outFile.close();
    }

    public void Write_Terminee_Members_Monetary_Fees_Values(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {

        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yy");

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

        int years = Utility.getDiffYears(startDate, endDate);//gets number of years
        years += 1;

        //get access to the actives template
        FileInputStream fileTemplate = new FileInputStream(workingDir + "\\Terminees_Sheet.xlsx");
        XSSFWorkbook workbookTemplate = new XSSFWorkbook(fileTemplate);
        XSSFSheet Templatesheet = workbookTemplate.getSheet("Terminees");

        FileInputStream fileRecon = new FileInputStream(workingDir + "\\Input Sheet.xlsx");
        XSSFWorkbook workbookRecon = new XSSFWorkbook(fileRecon);

        //    XSSFSheet sheetTemplate = workbookTemplate.getSheet("Actives");
        ArrayList list = new ArrayList();

        //INDEXES
     //  int PensionableSalaryIndex=8;
     //   int membersAgeIndex=PensionableSalaryIndex+years;
     //   int PensionableServiceIndex = membersAgeIndex+1;

        int memberBasicStartIndex = 19;//get 19 cell
        int memberVoluntaryStartIndex = memberBasicStartIndex+1;
        int employerBasicStartIndex = memberVoluntaryStartIndex+1;


        int memberBasicContribution_Index = memberBasicStartIndex+4;
        int memberVoluntaryContribution_Index = memberBasicContribution_Index+1;
        int employerContribution_Index = memberVoluntaryContribution_Index+1;
        int employerVoluntary_Index = employerContribution_Index+1;
        int Fees_Index = employerVoluntary_Index+1;


        for (int x = 0; x < years; x++) {

            Calendar cal = Calendar.getInstance();
            cal.set((StartYear), StartMonth, StartDay);
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy"); // Just the year, with 2 digits
            String formattedDate = sdf.format(cal.getTime());
            //  String formattedDate = "2014";
            String Recon = ("Actives at End of Plan Yr " + formattedDate);
            //   System.out.println("Actives at End of Plan Yr " + formattedDate);

//GET RECON SHEET
            XSSFSheet Reconsheet = workbookRecon.getSheet(Recon);
            int currentNumberofActiveMembers = Utility.getNumberOfMembersInSheet(Templatesheet);//get curent members from recon row
            int currentNumberofReconMembers = Utility.getNumberOfMembersInSheet(Reconsheet);//get curent members from recon row

            //   XSSFRow[] rowWrite = new XSSFRow[currentNumberofActiveMembers];
            XSSFRow rowWrite = null;
            Cell cellWrite = null;

            //iterate over members in terminee sheet
            for (int row = 0, readFromRecon = 7; row < currentNumberofActiveMembers; row++, readFromRecon++) {

                XSSFRow rowPosition = Templatesheet.getRow(readFromRecon);

                //get employee id
                XSSFCell cellA = rowPosition.getCell(0);  //employee number
                if (cellA == null) {
                    cellA = rowPosition.createCell((short) 0);
                    cellA.setCellValue("");
                }
                String result = cellA.getStringCellValue();
                String cellEmployeeID = result.replaceAll("[-]", "");

                //get LAST NAME
                XSSFCell cellB = rowPosition.getCell((short) 1);  //last name
                if (cellB == null) {
                    cellB = rowPosition.createCell((short) 1);
                    cellB.setCellValue("");
                }
                String cellLastName = cellB.getStringCellValue();

                //get FIRST NAME
                XSSFCell cellC = rowPosition.getCell((short) 2);  //first name
                if (cellC == null) {
                    cellC = rowPosition.createCell((short) 2);
                    cellC.setCellValue("");
                }
                String cellFirstName = cellC.getStringCellValue();

                //    System.out.println("currentNumberofReconMembers" + currentNumberofReconMembers);

//ITERATE OVER MEMBERS IN ACTIVE SHEET TO GET CONTRIBUTIONS
                for (int readFromReconRow = 7, iterate = 0; iterate < currentNumberofReconMembers; readFromReconRow++,iterate++) {
                    XSSFRow reconRow = Reconsheet.getRow(readFromReconRow);
                    //      System.out.println("readFromReconRow" + readFromReconRow);

                    //get employee id
                    XSSFCell reconCellA = reconRow.getCell(0);  //employee number
                    if (reconCellA == null) {
                        reconCellA = reconRow.createCell((short) 0);
                        reconCellA.setCellValue("");
                    }
                    String resultRecon = reconCellA.getStringCellValue();
                    String ReconcellEmployeeID = resultRecon.replaceAll("[-]", "");

                    //get LAST NAME
                    XSSFCell ReconCellB = reconRow.getCell((short) 1);  //last name
                    if (ReconCellB == null) {
                        ReconCellB = reconRow.createCell((short) 1);
                        ReconCellB.setCellValue("");
                    }
                    String ReconCellLastName = ReconCellB.getStringCellValue();

                    //get FIRST NAME
                    XSSFCell ReconCellC = reconRow.getCell((short) 2);  //first name
                    if (ReconCellC == null) {
                        ReconCellC = reconRow.createCell((short) 2);
                        ReconCellC.setCellValue("");
                    }
                    String ReconCellFirstName = ReconCellC.getStringCellValue();


                    //getINITIAL ACCUMUMATION AT START OF YEAR
                    XSSFCell ReconCellI = reconRow.getCell((short) 8);  //last name
                    if (ReconCellI == null) {
                        ReconCellI = reconRow.createCell((short) 8);
                        ReconCellI.setCellValue(0.00);
                    }
                    double ReconCell_MemberBasicContribution_atStart = ReconCellI.getNumericCellValue();

                    XSSFCell ReconCellJ= reconRow.getCell((short) 9);  //last name
                    if (ReconCellJ == null) {
                        ReconCellJ = reconRow.createCell((short) 9);
                        ReconCellJ.setCellValue(0.00);
                    }
                    double ReconCell_MemberVoluntaryContribution_atStart = ReconCellJ.getNumericCellValue();

                    XSSFCell ReconCellK= reconRow.getCell((short) 10);  //last name
                    if (ReconCellK == null) {
                        ReconCellK = reconRow.createCell((short) 10);
                        ReconCellK.setCellValue(0.00);
                    }
                    double ReconCell_EmployerContribution_atStart = ReconCellK.getNumericCellValue();


                    //employee basic contribution
                    XSSFCell ReconCellM= reconRow.getCell((short) 12);  //employee Basic Contribution
                    if (ReconCellM == null) {
                        ReconCellM = reconRow.createCell((short) 12);
                        ReconCellM.setCellValue(0.00);
                    }
                    double ReconCell_EmployeeBasic_Contribution = ReconCellM.getNumericCellValue();

                    //employee voluntary contribution
                    XSSFCell ReconCellN= reconRow.getCell((short) 13);  //employee optional contribution
                    if (ReconCellN == null) {
                        ReconCellN = reconRow.createCell((short) 13);
                        ReconCellN.setCellValue(0.00);
                    }
                    double ReconCell_EmployeeVoluntary_Contribution = ReconCellN.getNumericCellValue();


                    //employer contribution
                    XSSFCell ReconCellO= reconRow.getCell((short) 14);  //employer required
                    if (ReconCellO == null) {
                        ReconCellO = reconRow.createCell((short) 14);
                        ReconCellO.setCellValue(0.00);
                    }
                    double ReconCell_Employer_Contribution = ReconCellO.getNumericCellValue();

                    //FEES
                    XSSFCell ReconCellY= reconRow.getCell((short) 24);  //fees
                    if (ReconCellY == null) {
                        ReconCellY = reconRow.createCell((short) 24);
                        ReconCellY.setCellValue(0.00);
                    }
                    double ReconCell_Fees = ReconCellY.getNumericCellValue();



                    if(cellEmployeeID.equals(ReconcellEmployeeID)){

                        // write pensionable salary
                        rowWrite = Templatesheet.getRow(readFromRecon);


                        //write memebr initial
                        if(x==0) {
                            cellWrite = rowWrite.createCell(memberBasicStartIndex);
                            cellWrite.setCellValue(ReconCell_MemberBasicContribution_atStart);

                            cellWrite = rowWrite.createCell(memberVoluntaryStartIndex);
                            cellWrite.setCellValue(ReconCell_MemberVoluntaryContribution_atStart);

                            cellWrite = rowWrite.createCell(employerBasicStartIndex);
                            cellWrite.setCellValue(ReconCell_EmployerContribution_atStart);


                        }
//write members contributions
                        cellWrite = rowWrite.createCell((memberBasicContribution_Index));
                        cellWrite.setCellValue(ReconCell_EmployeeBasic_Contribution);

                        cellWrite = rowWrite.createCell(memberVoluntaryContribution_Index);
                        cellWrite.setCellValue(ReconCell_EmployeeVoluntary_Contribution);

                        cellWrite = rowWrite.createCell(employerContribution_Index);
                        cellWrite.setCellValue(ReconCell_Employer_Contribution);

                        //write members fees
                        cellWrite = rowWrite.createCell(Fees_Index);
                        cellWrite.setCellValue(ReconCell_Fees);
                    }

                }//end of looping through each member in each year period

                //GET INITIAL BALANCES

            }
            StartYear++;
            memberVoluntaryContribution_Index+=9;//9 for fees  || prob 8 for no fees
            memberBasicContribution_Index+=9;
            employerContribution_Index+=9;
            Fees_Index+=9;
        }//end of looping through the years


        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Updated_Terminee_Sheet.xlsx"));
        workbookTemplate.write(outFile);
        fileTemplate.close();
        fileRecon.close();
        outFile.close();
    }

    public void Write_Terminee_Members_Monetary_Values(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {

        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yy");

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

        int years = Utility.getDiffYears(startDate, endDate);//gets number of years
        years += 1;

        //get access to the actives template
        FileInputStream fileTemplate = new FileInputStream(workingDir + "\\Terminees_Sheet.xlsx");
        XSSFWorkbook workbookTemplate = new XSSFWorkbook(fileTemplate);
        XSSFSheet Templatesheet = workbookTemplate.getSheet("Terminees");

        FileInputStream fileRecon = new FileInputStream(workingDir + "\\Input Sheet.xlsx");
        XSSFWorkbook workbookRecon = new XSSFWorkbook(fileRecon);

        //    XSSFSheet sheetTemplate = workbookTemplate.getSheet("Actives");
        ArrayList list = new ArrayList();

        //INDEXES
        //  int PensionableSalaryIndex=8;
        //   int membersAgeIndex=PensionableSalaryIndex+years;
        //   int PensionableServiceIndex = membersAgeIndex+1;

        int memberBasicStartIndex = 19;//get 19 cell
        int memberVoluntaryStartIndex = memberBasicStartIndex+1;
        int employerBasicStartIndex = memberVoluntaryStartIndex+1;


        int memberBasicContribution_Index = memberBasicStartIndex+4;
        int memberVoluntaryContribution_Index = memberBasicContribution_Index+1;
        int employerContribution_Index = memberVoluntaryContribution_Index+1;
        int employerVoluntary_Index = employerContribution_Index+1;
        int Fees_Index = employerVoluntary_Index+1;


        for (int x = 0; x < years; x++) {

            Calendar cal = Calendar.getInstance();
            cal.set((StartYear), StartMonth, StartDay);
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy"); // Just the year, with 2 digits
            String formattedDate = sdf.format(cal.getTime());
            //  String formattedDate = "2014";
            String Recon = ("Actives at End of Plan Yr " + formattedDate);
            //   System.out.println("Actives at End of Plan Yr " + formattedDate);

//GET RECON SHEET
            XSSFSheet Reconsheet = workbookRecon.getSheet(Recon);
            int currentNumberofActiveMembers = Utility.getNumberOfMembersInSheet(Templatesheet);//get curent members from recon row
            int currentNumberofReconMembers = Utility.getNumberOfMembersInSheet(Reconsheet);//get curent members from recon row

            //   XSSFRow[] rowWrite = new XSSFRow[currentNumberofActiveMembers];
            XSSFRow rowWrite = null;
            Cell cellWrite = null;
            //    System.out.println("currentNumberofActiveMembers" + currentNumberofActiveMembers);
            for (int row = 0, readFromRecon = 7; row < currentNumberofActiveMembers; row++, readFromRecon++) {

                XSSFRow rowPosition = Templatesheet.getRow(readFromRecon);

                //get employee id
                XSSFCell cellA = rowPosition.getCell(0);  //employee number
                if (cellA == null) {
                    cellA = rowPosition.createCell((short) 0);
                    cellA.setCellValue("");
                }
                String result = cellA.getStringCellValue();
                String cellEmployeeID = result.replaceAll("[-]", "");

                //get LAST NAME
                XSSFCell cellB = rowPosition.getCell((short) 1);  //last name
                if (cellB == null) {
                    cellB = rowPosition.createCell((short) 1);
                    cellB.setCellValue("");
                }
                String cellLastName = cellB.getStringCellValue();

                //get FIRST NAME
                XSSFCell cellC = rowPosition.getCell((short) 2);  //first name
                if (cellC == null) {
                    cellC = rowPosition.createCell((short) 2);
                    cellC.setCellValue("");
                }
                String cellFirstName = cellC.getStringCellValue();

                //    System.out.println("currentNumberofReconMembers" + currentNumberofReconMembers);

//ITERATE OVER MEMBERS IN ACTIVE SHEET TO GET CONTRIBUTIONS
                for (int readFromReconRow = 7, iterate = 0; iterate < currentNumberofReconMembers; readFromReconRow++,iterate++) {
                    XSSFRow reconRow = Reconsheet.getRow(readFromReconRow);
                    //      System.out.println("readFromReconRow" + readFromReconRow);

                    //get employee id
                    XSSFCell reconCellA = reconRow.getCell(0);  //employee number
                    if (reconCellA == null) {
                        reconCellA = reconRow.createCell((short) 0);
                        reconCellA.setCellValue("");
                    }
                    String resultRecon = reconCellA.getStringCellValue();
                    String ReconcellEmployeeID = resultRecon.replaceAll("[-]", "");

                    //get LAST NAME
                    XSSFCell ReconCellB = reconRow.getCell((short) 1);  //last name
                    if (ReconCellB == null) {
                        ReconCellB = reconRow.createCell((short) 1);
                        ReconCellB.setCellValue("");
                    }
                    String ReconCellLastName = ReconCellB.getStringCellValue();

                    //get FIRST NAME
                    XSSFCell ReconCellC = reconRow.getCell((short) 2);  //first name
                    if (ReconCellC == null) {
                        ReconCellC = reconRow.createCell((short) 2);
                        ReconCellC.setCellValue("");
                    }
                    String ReconCellFirstName = ReconCellC.getStringCellValue();


                    //getINITIAL ACCUMUMATION AT START OF YEAR
                    XSSFCell ReconCellI = reconRow.getCell((short) 8);  //last name
                    if (ReconCellI == null) {
                        ReconCellI = reconRow.createCell((short) 8);
                        ReconCellI.setCellValue(0.00);
                    }
                    double ReconCell_MemberBasicContribution_atStart = ReconCellI.getNumericCellValue();

                    XSSFCell ReconCellJ= reconRow.getCell((short) 9);  //last name
                    if (ReconCellJ == null) {
                        ReconCellJ = reconRow.createCell((short) 9);
                        ReconCellJ.setCellValue(0.00);
                    }
                    double ReconCell_MemberVoluntaryContribution_atStart = ReconCellJ.getNumericCellValue();

                    XSSFCell ReconCellK= reconRow.getCell((short) 10);  //last name
                    if (ReconCellK == null) {
                        ReconCellK = reconRow.createCell((short) 10);
                        ReconCellK.setCellValue(0.00);
                    }
                    double ReconCell_EmployerContribution_atStart = ReconCellK.getNumericCellValue();


                    //employee basic contribution
                    XSSFCell ReconCellM= reconRow.getCell((short) 12);  //employee Basic Contribution
                    if (ReconCellM == null) {
                        ReconCellM = reconRow.createCell((short) 12);
                        ReconCellM.setCellValue(0.00);
                    }
                    double ReconCell_EmployeeBasic_Contribution = ReconCellM.getNumericCellValue();

                    //employee voluntary contribution
                    XSSFCell ReconCellN= reconRow.getCell((short) 13);  //employee optional contribution
                    if (ReconCellN == null) {
                        ReconCellN = reconRow.createCell((short) 13);
                        ReconCellN.setCellValue(0.00);
                    }
                    double ReconCell_EmployeeVoluntary_Contribution = ReconCellN.getNumericCellValue();


                    //employer contribution
                    XSSFCell ReconCellO= reconRow.getCell((short) 14);  //employer required
                    if (ReconCellO == null) {
                        ReconCellO = reconRow.createCell((short) 14);
                        ReconCellO.setCellValue(0.00);
                    }
                    double ReconCell_Employer_Contribution = ReconCellO.getNumericCellValue();

                    //employee basoc withdrawal

              /*      //FEES
                    XSSFCell ReconCellY= reconRow.getCell((short) 24);  //fees
                    if (ReconCellY == null) {
                        ReconCellY = reconRow.createCell((short) 24);
                        ReconCellY.setCellValue(0.00);
                    }
                    double ReconCell_Fees = ReconCellY.getNumericCellValue();
*/

                    if(cellEmployeeID.equals(ReconcellEmployeeID)){
                        // write pensionable salary
                        rowWrite = Templatesheet.getRow(readFromRecon);


                        //write memebr initial
                        if(x==0) {
                            cellWrite = rowWrite.createCell(memberBasicStartIndex);
                            cellWrite.setCellValue(ReconCell_MemberBasicContribution_atStart);

                            cellWrite = rowWrite.createCell(memberVoluntaryStartIndex);
                            cellWrite.setCellValue(ReconCell_MemberVoluntaryContribution_atStart);

                            cellWrite = rowWrite.createCell(employerBasicStartIndex);
                            cellWrite.setCellValue(ReconCell_EmployerContribution_atStart);


                        }
//write members contributions
                        cellWrite = rowWrite.createCell((memberBasicContribution_Index));
                        cellWrite.setCellValue(ReconCell_EmployeeBasic_Contribution);

                        cellWrite = rowWrite.createCell(memberVoluntaryContribution_Index);
                        cellWrite.setCellValue(ReconCell_EmployeeVoluntary_Contribution);

                        cellWrite = rowWrite.createCell(employerContribution_Index);
                        cellWrite.setCellValue(ReconCell_Employer_Contribution);

              /*          //write members fees
                        cellWrite = rowWrite.createCell(Fees_Index);
                        cellWrite.setCellValue(ReconCell_Fees);*/
                    }

                }//end of looping through each member in each year period

               //WHEN WE ARE AT END OF YEAR, WRITE AMOUNT WITHDRAWAL


            }
            StartYear++;
            memberVoluntaryContribution_Index+=8;//9 for fees  || prob 8 for no fees
            memberBasicContribution_Index+=8;
            employerContribution_Index+=8;
        //    Fees_Index+=9;
        }//end of looping through the years


        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Updated_Terminee_Sheet.xlsx"));
        workbookTemplate.write(outFile);
        fileTemplate.close();
        fileRecon.close();
        outFile.close();
    }

    public void Write_Members_Monetary_Fees_Values(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {

        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yy");

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

        int years = Utility.getDiffYears(startDate, endDate);//gets number of years
        years += 1;

        //get access to the actives template
        FileInputStream fileTemplate = new FileInputStream(workingDir + "\\Actives_Sheet.xlsx");
        XSSFWorkbook workbookTemplate = new XSSFWorkbook(fileTemplate);
        XSSFSheet Activesheet = workbookTemplate.getSheet("Actives");

        FileInputStream fileRecon = new FileInputStream(workingDir + "\\Input Sheet.xlsx");
        XSSFWorkbook workbookRecon = new XSSFWorkbook(fileRecon);

        //    XSSFSheet sheetTemplate = workbookTemplate.getSheet("Actives");
        ArrayList list = new ArrayList();

        //INDEXES
        int PensionableSalaryIndex=8;
        int membersAgeIndex=PensionableSalaryIndex+years;
        int PensionableServiceIndex = membersAgeIndex+1;
        int memberBasicStartIndex = PensionableServiceIndex+1;
        int memberVoluntaryStartIndex = memberBasicStartIndex+1;
        int employerBasicStartIndex = memberVoluntaryStartIndex+1;


        int memberBasicContribution_Index = PensionableServiceIndex+5;
        int memberVoluntaryContribution_Index = memberBasicContribution_Index+1;
        int employerContribution_Index = memberVoluntaryContribution_Index+1;
        int employerVoluntary_Index = employerContribution_Index+1;
        int Fees_Index = employerVoluntary_Index+1;

        for (int x = 0; x < years; x++) {

            Calendar cal = Calendar.getInstance();
            cal.set((StartYear), StartMonth, StartDay);
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy"); // Just the year, with 2 digits
            String formattedDate = sdf.format(cal.getTime());
            //  String formattedDate = "2014";
            String Recon = ("Actives at End of Plan Yr " + formattedDate);
            //   System.out.println("Actives at End of Plan Yr " + formattedDate);

//GET RECON SHEET
            XSSFSheet Reconsheet = workbookRecon.getSheet(Recon);
            int currentNumberofActiveMembers = Utility.getNumberOfMembersInSheet(Activesheet);//get curent members from recon row
            int currentNumberofReconMembers = Utility.getNumberOfMembersInSheet(Reconsheet);//get curent members from recon row

            //   XSSFRow[] rowWrite = new XSSFRow[currentNumberofActiveMembers];
            XSSFRow rowWrite = null;
            Cell cellWrite = null;
            //    System.out.println("currentNumberofActiveMembers" + currentNumberofActiveMembers);
            for (int row = 0, readFromRecon = 7; row < currentNumberofActiveMembers; row++, readFromRecon++) {

                XSSFRow rowPosition = Activesheet.getRow(readFromRecon);

                //get employee id
                XSSFCell cellA = rowPosition.getCell(0);  //employee number
                if (cellA == null) {
                    cellA = rowPosition.createCell((short) 0);
                    cellA.setCellValue("");
                }
                String result = cellA.getStringCellValue();
                String cellEmployeeID = result.replaceAll("[-]", "");

                //get LAST NAME
                XSSFCell cellB = rowPosition.getCell((short) 1);  //last name
                if (cellB == null) {
                    cellB = rowPosition.createCell((short) 1);
                    cellB.setCellValue("");
                }
                String cellLastName = cellB.getStringCellValue();

                //get FIRST NAME
                XSSFCell cellC = rowPosition.getCell((short) 2);  //first name
                if (cellC == null) {
                    cellC = rowPosition.createCell((short) 2);
                    cellC.setCellValue("");
                }
                String cellFirstName = cellC.getStringCellValue();

                //    System.out.println("currentNumberofReconMembers" + currentNumberofReconMembers);


                for (int readFromReconRow = 7, iterate = 0; iterate < currentNumberofReconMembers; readFromReconRow++,iterate++) {
                    XSSFRow reconRow = Reconsheet.getRow(readFromReconRow);
                    //      System.out.println("readFromReconRow" + readFromReconRow);

                    //get employee id
                    XSSFCell reconCellA = reconRow.getCell(0);  //employee number
                    if (reconCellA == null) {
                        reconCellA = reconRow.createCell((short) 0);
                        reconCellA.setCellValue("");
                    }
                    String resultRecon = reconCellA.getStringCellValue();
                    String ReconcellEmployeeID = resultRecon.replaceAll("[-]", "");

                    //get LAST NAME
                    XSSFCell ReconCellB = reconRow.getCell((short) 1);  //last name
                    if (ReconCellB == null) {
                        ReconCellB = reconRow.createCell((short) 1);
                        ReconCellB.setCellValue("");
                    }
                    String ReconCellLastName = ReconCellB.getStringCellValue();

                    //get FIRST NAME
                    XSSFCell ReconCellC = reconRow.getCell((short) 2);  //first name
                    if (ReconCellC == null) {
                        ReconCellC = reconRow.createCell((short) 2);
                        ReconCellC.setCellValue("");
                    }
                    String ReconCellFirstName = ReconCellC.getStringCellValue();

                    //get DOB
                    XSSFCell ReconCellF = reconRow.getCell((short) 5);  //first name
                    if (ReconCellF == null) {
                        ReconCellF = reconRow.createCell((short) 5);
                        // ReconCellF.setCellValue("");
                    }
                    Date ReconCellDOB = ReconCellF.getDateCellValue();

                    //get DOB
                    XSSFCell ReconCellH = reconRow.getCell((short) 7);  //first name
                    if (ReconCellH == null) {
                        ReconCellH = reconRow.createCell((short) 7);
                        // ReconCellF.setCellValue("");
                    }
                    Date ReconCellDateofEnrolment = ReconCellH.getDateCellValue();

//GET PENSIONABLE SALARY
                    XSSFCell ReconCellL = reconRow.getCell((short) 11);  //first name
                    if (ReconCellL == null) {
                        ReconCellL = reconRow.createCell((short) 11);
                        ReconCellL.setCellValue(0.00);
                    }
                    double ReconCell_PensionableSalary = ReconCellL.getNumericCellValue();

                    //getINITIAL ACCUMUMATION AT START OF YEAR
                    XSSFCell ReconCellI = reconRow.getCell((short) 8);  //last name
                    if (ReconCellI == null) {
                        ReconCellI = reconRow.createCell((short) 8);
                        ReconCellI.setCellValue(0.00);
                    }
                    double ReconCell_MemberBasicContribution_atStart = ReconCellI.getNumericCellValue();

                    XSSFCell ReconCellJ= reconRow.getCell((short) 9);  //last name
                    if (ReconCellJ == null) {
                        ReconCellJ = reconRow.createCell((short) 9);
                        ReconCellJ.setCellValue(0.00);
                    }
                    double ReconCell_MemberVoluntaryContribution_atStart = ReconCellJ.getNumericCellValue();

                    XSSFCell ReconCellK= reconRow.getCell((short) 10);  //last name
                    if (ReconCellK == null) {
                        ReconCellK = reconRow.createCell((short) 10);
                        ReconCellK.setCellValue(0.00);
                    }
                    double ReconCell_EmployerContribution_atStart = ReconCellK.getNumericCellValue();

                    //employee basic contribution
                    XSSFCell ReconCellM= reconRow.getCell((short) 12);  //last name
                    if (ReconCellM == null) {
                        ReconCellM = reconRow.createCell((short) 12);
                        ReconCellM.setCellValue(0.00);
                    }
                    double ReconCell_EmployeeBasic_Contribution = ReconCellM.getNumericCellValue();

                    //employee voluntary contribution
                    XSSFCell ReconCellN= reconRow.getCell((short) 13);  //last name
                    if (ReconCellN == null) {
                        ReconCellN = reconRow.createCell((short) 13);
                        ReconCellN.setCellValue(0.00);
                    }
                    double ReconCell_EmployeeVoluntary_Contribution = ReconCellN.getNumericCellValue();


                    //employer contribution
                    XSSFCell ReconCellO= reconRow.getCell((short) 14);  //last name
                    if (ReconCellO == null) {
                        ReconCellO = reconRow.createCell((short) 14);
                        ReconCellO.setCellValue(0.00);
                    }
                    double ReconCell_Employer_Contribution = ReconCellO.getNumericCellValue();

                    //FEES
                    XSSFCell ReconCellY= reconRow.getCell((short) 24);  //last name
                    if (ReconCellY == null) {
                        ReconCellY = reconRow.createCell((short) 24);
                        ReconCellY.setCellValue(0.00);
                    }
                    double ReconCell_Fees = ReconCellY.getNumericCellValue();



                    if(cellEmployeeID.equals(ReconcellEmployeeID)){

                        // write pensionable salary
                        rowWrite = Activesheet.getRow(readFromRecon);

                        cellWrite= rowWrite.createCell((PensionableSalaryIndex + x));
                        cellWrite.setCellValue(ReconCell_PensionableSalary);


                        //*************************MEMBERS AGE AS AT**********************

                        cellWrite = rowWrite.createCell(membersAgeIndex);
                        cellWrite.setCellValue(Utility.getDiffYears(ReconCellDOB, endDate));


//********PENSIONABLE SERVICE***********************
                        DecimalFormat dF = new DecimalFormat("#.##");//#.##
                        cellWrite = rowWrite.createCell(PensionableServiceIndex);
                        cellWrite.setCellValue(Double.parseDouble(dF.format((Utility.betweenDates(ReconCellDateofEnrolment, endDate)/365.25))));

                        //write memebr initial
                        if(x==0) {
                            cellWrite = rowWrite.createCell(memberBasicStartIndex);
                            cellWrite.setCellValue(ReconCell_MemberBasicContribution_atStart);

                            cellWrite = rowWrite.createCell(memberVoluntaryStartIndex);
                            cellWrite.setCellValue(ReconCell_MemberVoluntaryContribution_atStart);

                            cellWrite = rowWrite.createCell(employerBasicStartIndex);
                            cellWrite.setCellValue(ReconCell_EmployerContribution_atStart);


                        }

                        cellWrite = rowWrite.createCell((memberBasicContribution_Index));
                        cellWrite.setCellValue(ReconCell_EmployeeBasic_Contribution);

                        cellWrite = rowWrite.createCell(memberVoluntaryContribution_Index);
                        cellWrite.setCellValue(ReconCell_EmployeeVoluntary_Contribution);

                        cellWrite = rowWrite.createCell(employerContribution_Index);
                        cellWrite.setCellValue(ReconCell_Employer_Contribution);

                        cellWrite = rowWrite.createCell(Fees_Index);
                        cellWrite.setCellValue(ReconCell_Fees);
                    }

                }//end of looping through each member in each year period

                //GET INITIAL BALANCES



            }
            StartYear++;
            memberVoluntaryContribution_Index+=9;//9 for fees  || prob 8 for no fees
            memberBasicContribution_Index+=9;
            employerContribution_Index+=9;
            Fees_Index+=9;
        }//end of looping through the years


        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Actives_Sheet.xlsx"));
        workbookTemplate.write(outFile);
        fileTemplate.close();
        fileRecon.close();
        outFile.close();
    }

    public void Write_Members_To_Active_Sheet(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) {
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yy");

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

        int years = Utility.getDiffYears(startDate, endDate);//gets number of years
        years += 1;

        try {

            //get access to the DATA of the Active Members
            FileInputStream fileInputStreamWorkBookData = new FileInputStream(workingDir + "\\Input Sheet.xlsx");
            XSSFWorkbook workbookData = new XSSFWorkbook(fileInputStreamWorkBookData);

            XSSFSheet Initialsheet = workbookData.getSheet("Actives at End of Plan Yr "+StartYear);
            int InitialnumOfActives = Utility.getNumberOfMembersInSheet(Initialsheet);

            //get access to the actives template
            FileInputStream fileTemplate = new FileInputStream(workingDir + "\\Templates\\Template_Active_Sheet.xlsx");
            XSSFWorkbook workbookTemplate = new XSSFWorkbook(fileTemplate);
            XSSFSheet sheetTemplate = workbookTemplate.getSheet("Actives");

            XSSFCell[] pensionableSalary = new XSSFCell[years];
            XSSFRow[] writeRow = new XSSFRow[InitialnumOfActives];
            Cell writeCell;
            ArrayList coloumnData = new ArrayList();
            ArrayList newList = new ArrayList();
            System.out.println(InitialnumOfActives);

//1. Create the date cell style
            XSSFCreationHelper createHelper = workbookTemplate.getCreationHelper();
            XSSFCellStyle cellStyle         = workbookTemplate.createCellStyle();
            cellStyle.setDataFormat(
                    createHelper.createDataFormat().getFormat("dd-MMM-yy"));
            //MAIN PROCESSING

            String [] InitialcellEmployeeID = new String[InitialnumOfActives];
            String[] InitialcellLastName = new String[InitialnumOfActives];
            String[] InitialcellFirstName = new String [InitialnumOfActives];

int cumulativeRow=InitialnumOfActives+7;

            //get the initial members in first year period
            for (int row = 0, readFromRow = 7; row < InitialnumOfActives; row++, readFromRow++) {

                //get the current row position based on value of row
                XSSFRow rowPosition = Initialsheet.getRow(readFromRow);
           //     XSSFRow rowPosition = Reconsheet.getRow(readFromRow);

                //get employee id
                XSSFCell cellA = rowPosition.getCell(0);  //employee number
                if (cellA == null) {
                    cellA = rowPosition.createCell((short) 0);
                    cellA.setCellValue("");
                }
                String result = cellA.getStringCellValue();
                InitialcellEmployeeID[row] = result.replaceAll("[-]", "");

                //get LAST NAME
                XSSFCell cellB = rowPosition.getCell((short) 1);  //last name
                if (cellB == null) {
                    cellB = rowPosition.createCell((short) 1);
                    cellB.setCellValue("");
                }
                InitialcellLastName[row] = cellB.getStringCellValue();

                //get FIRST NAME
                XSSFCell cellC = rowPosition.getCell((short) 2);  //first name
                if (cellC == null) {
                    cellC = rowPosition.createCell((short) 2);
                    cellC.setCellValue("");
                }
                InitialcellFirstName[row] = cellC.getStringCellValue();


                //get GENDER
                XSSFCell cellD = rowPosition.getCell((short) 3);  //gender
                if (cellD == null) {
                    cellD = rowPosition.createCell((short) 3);
                    cellD.setCellValue("");
                }
                String cellGender = cellD.getStringCellValue();

                //get Marital Status
                XSSFCell cellE = rowPosition.getCell((short) 4);  //Marital Status
                if (cellE == null) {
                    cellE = rowPosition.createCell((short) 4);
                    cellE.setCellValue("");
                }
                String cellMaritalStatus = cellE.getStringCellValue();

                //get Date of birth
                XSSFCell cellF = rowPosition.getCell((short) 5);  // Date of birth
                //   if (cellF == null) {
                //       cellF = rowPosition.createCell((short) 5);
                //      cellF.setCellValue(new Date("01-Jan-01"));
                //  }
                Date cellDateofBirth = cellF.getDateCellValue();

                //get Date of Employment
                XSSFCell cellG = rowPosition.getCell((short) 6);  //Date of Employment
                // if (cellG == null) {
                //     cellG = rowPosition.createCell((short) 6);
                //      cellG.setCellValue(new Date("01-Jan-01"));
                //   }
                Date cellDateofEmployment = cellG.getDateCellValue();


                //get Date of Enrollment
                XSSFCell cellH = rowPosition.getCell((short) 7);  //Date of Employment
                //  if (cellH == null) {
                //      cellH = rowPosition.createCell((short) 7);
                //   cellH.setCellValue(new Date("01-Jan-01"));
                //    }
                Date cellDateofEnrolment = cellH.getDateCellValue();


                //when we are at the first year
             //   if(x==0) {
                    coloumnData.add(InitialcellEmployeeID[row]);
                    coloumnData.add(InitialcellLastName[row]);
                    coloumnData.add(InitialcellFirstName[row]);
                    coloumnData.add(cellGender);
                    coloumnData.add(cellMaritalStatus);
                    coloumnData.add(cellDateofBirth);
                    coloumnData.add(cellDateofEmployment);
                    coloumnData.add(cellDateofEnrolment);
                 //   coloumnData.add(Cell_PensionableSalary);

                    //   writeRow[row] = sheetTemplate.createRow( readFromRow);
                    XSSFRow WriteRow = sheetTemplate.createRow(readFromRow);

                //    System.out.println("createRow->" + readFromRow);

                //write first set of data to MAIN ROW
                    for (int colPosition = 0; colPosition < coloumnData.size(); colPosition++) {

                        writeCell = WriteRow.createCell(colPosition);

                      //  System.out.println("createCell->" + colPosition);

                        if (coloumnData.get(colPosition) instanceof Date) {
                            Date d = (Date) coloumnData.get(colPosition);
                            writeCell.setCellValue(d);
                            writeCell.setCellStyle(cellStyle);
                        }

   /*                 else if(coloumnData.get(colPosition) instanceof Boolean)
                        writeCell.setCellValue((Boolean)coloumnData.get(colPosition));*/
                        else if (coloumnData.get(colPosition) instanceof String) writeCell.setCellValue((String) coloumnData.get(colPosition));
                        else if (coloumnData.get(colPosition) instanceof Double) writeCell.setCellValue((Double) coloumnData.get(colPosition));

                    }
                    coloumnData.clear();

                }//end of loop running for 1st year


            boolean foundIt = false;
            //LOOP THROUGH THE OTHER YEARS
int InitialMembers = InitialnumOfActives;
int yearsRemaining = years-1;

                //GET OTHER ACTIVE MEMBERS FROM OTHER YEARS
                //   else{
                for (int x = 0; x < yearsRemaining; x++) {

                    int newStartYear = StartYear+1;

                    Calendar cal = Calendar.getInstance();
                    cal.set((newStartYear+x), StartMonth, StartDay);
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy"); // Just the year, with 2 digits
                    String formattedDate = sdf.format(cal.getTime());
                  //  String formattedDate = "2014";
                    String Recon = ("Actives at End of Plan Yr " + formattedDate);

//GET RECON SHEET
                    XSSFSheet Reconsheet = workbookData.getSheet(Recon);

                    int currentNumber = Utility.getNumberOfMembersInSheet(Reconsheet);//get curent members from recon row
                   String [] cellEmployeeID = new String[currentNumber];
                    String[] cellLastName = new String[currentNumber];
                    String[] cellFirstName = new String [currentNumber];

                    for (int row = 0, readFromRecon = 7; row < currentNumber; row++, readFromRecon++) {

                        XSSFRow rowPosition = Reconsheet.getRow(readFromRecon);

                        //get employee id
                        XSSFCell cellA = rowPosition.getCell(0);  //employee number
                        if (cellA == null) {
                            cellA = rowPosition.createCell((short) 0);
                            cellA.setCellValue("");
                        }
                        String result = cellA.getStringCellValue();
                      cellEmployeeID[row]= result.replaceAll("[-]", "");

                        //get LAST NAME
                        XSSFCell cellB = rowPosition.getCell((short) 1);  //last name
                        if (cellB == null) {
                            cellB = rowPosition.createCell((short) 1);
                            cellB.setCellValue("");
                        }
                      cellLastName[row]= cellB.getStringCellValue();

                        //get FIRST NAME
                        XSSFCell cellC = rowPosition.getCell((short) 2);  //first name
                        if (cellC == null) {
                            cellC = rowPosition.createCell((short) 2);
                            cellC.setCellValue("");
                        }
                     cellFirstName[row] = cellC.getStringCellValue();

                        //get GENDER
                        XSSFCell cellD = rowPosition.getCell((short) 3);  //gender
                        if (cellD == null) {
                            cellD = rowPosition.createCell((short) 3);
                            cellD.setCellValue("");
                        }
                        String cellGender = cellD.getStringCellValue();

                        //get Marital Status
                        XSSFCell cellE = rowPosition.getCell((short) 4);  //Marital Status
                        if (cellE == null) {
                            cellE = rowPosition.createCell((short) 4);
                            cellE.setCellValue("");
                        }
                        String cellMaritalStatus = cellE.getStringCellValue();

                        //get Date of birth
                        XSSFCell cellF = rowPosition.getCell((short) 5);  // Date of birth
                        //   if (cellF == null) {
                        //       cellF = rowPosition.createCell((short) 5);
                        //      cellF.setCellValue(new Date("01-Jan-01"));
                        //  }
                        Date cellDateofBirth = cellF.getDateCellValue();

                        //get Date of Employment
                        XSSFCell cellG = rowPosition.getCell((short) 6);  //Date of Employment
                        // if (cellG == null) {
                        //     cellG = rowPosition.createCell((short) 6);
                        //      cellG.setCellValue(new Date("01-Jan-01"));
                        //   }
                        Date cellDateofEmployment = cellG.getDateCellValue();


                        //get Date of Enrollment
                        XSSFCell cellH = rowPosition.getCell((short) 7);  //Date of Employment
                        //  if (cellH == null) {
                        //      cellH = rowPosition.createCell((short) 7);
                        //   cellH.setCellValue(new Date("01-Jan-01"));
                        //    }
                        Date cellDateofEnrolment = cellH.getDateCellValue();



                        boolean findIt=false;

                        //test the value from recon sheet against initial
                        for (int initial = 0 ; initial < InitialMembers; initial++) {


                            System.out.print("initial " + initial);

                            if (cellEmployeeID[row].equals(InitialcellEmployeeID[initial]) && cellLastName[row].equals(InitialcellLastName[initial] )) {
                    findIt=true;
                               break;
                            }
                        }

                        if (!findIt){
                            newList.add(cellEmployeeID[row]);
                            newList.add(cellLastName[row]);
                            newList.add(cellFirstName[row]);
                            newList.add(cellGender);
                            newList.add(cellMaritalStatus);
                            newList.add(cellDateofBirth);
                            newList.add(cellDateofEmployment);
                            newList.add(cellDateofEnrolment);

                            XSSFRow WriteRow = sheetTemplate.createRow(cumulativeRow++);

                            for (int colPosition = 0; colPosition < newList.size(); colPosition++) {

                                writeCell = WriteRow.createCell(colPosition);

                                if (newList.get(colPosition) instanceof Date) {
                                    Date d = (Date) newList.get(colPosition);
                                    writeCell.setCellValue(d);
                                    writeCell.setCellStyle(cellStyle);
                                } else if (newList.get(colPosition) instanceof String)
                                    writeCell.setCellValue((String) newList.get(colPosition));
                                else if (newList.get(colPosition) instanceof Double)
                                    writeCell.setCellValue((Double) newList.get(colPosition));
                            }
                            newList.clear();

                   }


                }//end loop years

InitialcellEmployeeID = new String[currentNumber];
                    InitialcellLastName =  new String[currentNumber];
                            InitialcellFirstName = new String[currentNumber];
                    for(int h=0;h<currentNumber;h++){
                        System.out.println("currentNumber"+currentNumber);
                        System.out.println("h"+h);
                        InitialcellEmployeeID[h]=cellEmployeeID[h];
                        InitialcellLastName[h] = cellLastName[h];
                        InitialcellFirstName[h]=cellFirstName[h];
                    }

                    InitialMembers = currentNumber;
            }//end of looping through each of members
                //write the data
                FileOutputStream outFile = new FileOutputStream(new File(workingDir + "\\Actives_Sheet.xlsx"));
                workbookTemplate.write(outFile);
                fileTemplate.close();
                outFile.close();

            } catch(FileNotFoundException e){
                e.printStackTrace();
            } catch(IOException e){
                e.printStackTrace();
            }
        catch(NullPointerException e){
                e.printStackTrace();

            }
        }

    public void Create_Activee_Contribution(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {

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

                //String i1Val = null;
                double i1Val = 0;
                XSSFCell[] cellI1 = new XSSFCell[12];
                double [] d = new double[12];
                for (int g=8,j=0;g<20;g++,j++){
                    cellI1[j] = ActiveRow.getCell(g);
                   // i1Val = cellI1[j].getStringCellValue();
                    i1Val = cellI1[j].getNumericCellValue();
                //    d[j] = Double.parseDouble(i1Val);
                    d[j] = i1Val;
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

        //OPEN ACTIVE SHEET
        FileInputStream fileInputStream = new FileInputStream(workingDir +"\\Accumulated_Actives_Sheet.xlsx");
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
      //  int numOfActives = ActiveSheet.getLastRowNum()+1;//gets number of active members
        int numOfActives = Utility.getNumberOfMembersInSheet(ActiveSheet);//gets number of active members
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

        //we are to get the indexes of the initial accumulation balances

        //indexes
        int initialEmployee_BasicAcc_Index= 8 + years + 2;
        int initialEmployee_VoluntaryAcc_Index=initialEmployee_BasicAcc_Index+1;
        int  initialEmployer_BasicAcc_Index=initialEmployee_VoluntaryAcc_Index+1;
        int initialEmployer_Optional_Index =  initialEmployer_BasicAcc_Index+1;



        int Contribution_MemberBasic_Index = initialEmployer_Optional_Index+1;
        int Contribution_MemberVoluntary_Index = Contribution_MemberBasic_Index+1;
        int Contribution_EmployerRequired_Index = Contribution_MemberVoluntary_Index+1;
        int Contribution_EmployerOptional_Index = Contribution_EmployerRequired_Index+1;

        int readCol=Contribution_MemberBasic_Index;//26;//start to read from Col 26, which is starting contribution column
        int YearCol = readCol;//as we are in the same year, we should always be reading from the correct column
        int Write_Coloumn =initialEmployee_BasicAcc_Index+8;//30;//start to write at column 30 which are the accumullated cells


        for(int row=7,I=0;I<numOfActives;row++,I++){ //get the initial accumulated balances

            XSSFRow ActiveRow = ActiveSheet.getRow(row);

            XSSFCell[] cellAccEmployeeBasic0 = new XSSFCell[numOfActives];
            cellAccEmployeeBasic0[I] = ActiveRow.getCell( initialEmployee_BasicAcc_Index);  //employee basic Start balances
            if(cellAccEmployeeBasic0[I]==null){
                cellAccEmployeeBasic0[I] = ActiveRow.createCell(initialEmployee_BasicAcc_Index);
                cellAccEmployeeBasic0[I].setCellValue(0.00);
            }
            CellAccEmployeeBasic0[I]= cellAccEmployeeBasic0[I].getNumericCellValue();

            XSSFCell[] cellAccEmployeeOptional0 = new XSSFCell[numOfActives];
            cellAccEmployeeOptional0[I] = ActiveRow.getCell(initialEmployee_VoluntaryAcc_Index);  //employee number
            if(cellAccEmployeeOptional0[I]==null){
                cellAccEmployeeOptional0[I] = ActiveRow.createCell(initialEmployee_VoluntaryAcc_Index);
                cellAccEmployeeOptional0[I].setCellValue(0.00);
            }
            CellAccEmployeeOptional0[I]= cellAccEmployeeOptional0[I].getNumericCellValue();


            XSSFCell[] cellAccEmployerRequired0 = new XSSFCell[numOfActives];
            cellAccEmployerRequired0[I] = ActiveRow.getCell(initialEmployer_BasicAcc_Index); //employee number
            if(cellAccEmployerRequired0[I]==null){
                cellAccEmployerRequired0[I] = ActiveRow.createCell(initialEmployer_BasicAcc_Index);
                cellAccEmployerRequired0[I].setCellValue(0.00);
            }
            CellAccEmployerRequired0[I] = cellAccEmployerRequired0[I].getNumericCellValue();

            XSSFCell[] cellAccEmployerOptional0 = new XSSFCell[numOfActives];
            cellAccEmployerOptional0[I] = ActiveRow.getCell(initialEmployer_Optional_Index); //employee number
            if(cellAccEmployerOptional0[I]==null){
                cellAccEmployerOptional0[I] = ActiveRow.createCell(initialEmployer_Optional_Index);
                cellAccEmployerOptional0[I].setCellValue(0.00);
            }
            CellAccEmployerOptional0[I] = cellAccEmployerOptional0[I].getNumericCellValue();
        }//end of loop to get inital accumulated balances
        XSSFRow ActiveRow = null;
        Cell cellR;

        //MAIN PROCESSING-to get acc Balances and cont Balances
        for (int x = 0; x < years; x++) { //run for the appropiate number of years
            //  System.out.println(numOfActives);


            for (int row = 7, I = 0; I < numOfActives; row++, I++) { //run for appropiate number of total active members
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
                    cellR.setCellValue(Double.parseDouble(dF.format(val.get(b))));
                }//end of loop

            }//end of looping through each member

            //MOVING THE INDEXES
            YearCol=readCol+4;//move over by 4 columns to get next set of contributions for next year
            readCol=YearCol;//give the readCol 5 so that, it can always read the correct set of contributions of that same year
            Write_Coloumn += 8;//8 no fees  || 9 fees
            StartYear++;//increment Start year by one until we reach end year

            //when we are at end of year...we should write account balance as at end date
            if(x==(years-1)) {
                for (int I = 0, row=7; I < numOfActives; I++, row++) {
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
            FileInputStream fileInputStream = new FileInputStream(workingDir+"\\Input Sheet.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet worksheet = workbook.getSheet("Terminated up to "+EndYear+"."+EndMonth+"."+EndDay);
            //   XSSFSheet sheet = workbook.getSheetAt(0);

          //  int rowCount=worksheet.getPhysicalNumberOfRows();
            int rowCount = Utility.getNumberOfTermineeMembersInSheet(worksheet);
            // System.out.println(rowCount);

            int num = rowCount;
            //  int noOfColumns = sheet.getRow(num).getLastCellNum();

            FileInputStream fileR = new FileInputStream(workingDir+"\\Templates\\Template_Terminee_Sheet.xlsx");
            XSSFWorkbook workbookR = new XSSFWorkbook(fileR);
            XSSFSheet sheetR = workbookR.getSheetAt(0);

            //1. Create the date cell style
            XSSFCreationHelper createHelper = workbookR.getCreationHelper();
            XSSFCellStyle cellStyle         = workbookR.createCellStyle();
            cellStyle.setDataFormat(
                    createHelper.createDataFormat().getFormat("dd-MMM-yy"));

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

               for (int row = 0,readFromRow=11; row < num; row++,readFromRow++) {

                   int Row = row;

                   XSSFRow row1 = worksheet.getRow(readFromRow);

                   XSSFCell cellA = row1.getCell((short) 0);  //employee number
                  if(cellA==null){
                      cellA=row1.createCell(0);
                      cellA.setCellValue("");
                  }
                   String result = cellA.getStringCellValue();
                   String cellEmployeeID = result.replaceAll("[-]", "");

                   XSSFCell cellB1 = row1.getCell((short) 1);
                   if(cellB1==null){
                       cellB1=row1.createCell(1);
                       cellB1.setCellValue("");
                   }
                  // String b1Val = cellB1.getStringCellValue();
                   String cellLastName = cellB1.getStringCellValue();


                   XSSFCell cellC = row1.getCell((short) 2);    //last name
                   if(cellC==null){
                       cellC=row1.createCell(2);
                       cellC.setCellValue("");
                   }
                   //   String c1Val = cellC1.getStringCellValue();
                   String cellFirstName = cellC.getStringCellValue();


                   XSSFCell cellD = row1.getCell((short) 3); //first name
                   if(cellD==null){
                       cellD=row1.createCell(3);
                       cellD.setCellValue("");
                   }
               //    String d1Val = cellD.getStringCellValue();
                   String cellGender = cellD.getStringCellValue();


                   XSSFCell cellE = row1.getCell((short) 4);  //sex
                   if(cellE==null){
                       cellE=row1.createCell(4);
                       cellE.setCellValue("");
                   }
            //       String e1Val = cellE.getStringCellValue();
                   String cellMaritalstatus = cellE.getStringCellValue();


                   XSSFCell cellF = row1.getCell((short) 5); //dob
                   if(cellF==null){
                       cellF=row1.createCell(5);
                       cellF.setCellValue(new Date("01-Jan-01"));
                   }
                  // String f1Val = cellF.getStringCellValue();
                   Date cellDOB= cellF.getDateCellValue();


                   XSSFCell cellG = row1.getCell((short) 6); //plan entry
                   if(cellG==null){
                       cellG=row1.createCell(6);
                      // cellG.setCellValue();
                   }
                   Date cellDateofEmployment = cellG.getDateCellValue();


                   XSSFCell cellH = row1.getCell((short) 7); //emp date
                   if(cellH==null){
                       cellH=row1.createCell(7);
                       cellH.setCellValue(new Date("01-Jan-01"));
                   }

                   Date cellPlanEntry = cellH.getDateCellValue();

                   XSSFCell cellI = row1.getCell((short) 8);  //status date
                   if(cellI==null){
                       cellI=row1.createCell(8);
                       cellI.setCellValue(new Date("01-Jan-01"));
                   }
                   Date cellStatusDate = cellI.getDateCellValue();

                   XSSFCell cellJ = row1.getCell((short) 9); //status
                   if(cellJ==null){
                       cellJ=row1.createCell(9);
                       cellJ.setCellValue("");
                   }
                   String cellStatus = cellJ.getStringCellValue();

                //   Date cellDateofRefund = null;
                       XSSFCell cellN = row1.getCell((short) 13); //date of refund
                   if(cellN==null){
                       cellN=row1.createCell(13);
                      cellN.setCellValue(new Date("01-Jan-01"));
                   }
                   Date  cellDateofRefund = cellN.getDateCellValue();

               //    if (cellTypeoftermination.equals("DEATH") ||j1Val.equals("RETIREMENT") || j1Val.equals("TERMINATED") || j1Val.equals("DEFERRED") && !(d1Val.equals("KEY"))) {


                       stringBuilder.append("Employee ID: " + cellEmployeeID+"\n");
                       stringBuilder.append("Last Name: " + cellLastName+"\n");
                       stringBuilder.append("First Name: " + cellFirstName+"\n");
                       stringBuilder.append("Date of Birth: " + datetemp.format(cellDOB)+"\n");
                       stringBuilder.append("Status Date: "+ datetemp.format(cellStatusDate)+"\n");
                       stringBuilder.append("Status: "+cellStatus+"\n");
                       stringBuilder.append("-------------------------------------------------------\n");
                       //      System.out.println();

            /*           Date statusDate = null;
                       try {
                    //       statusDate = datetemp.parse(i1Val);
                       } catch (ParseException e) {
                           e.printStackTrace();
                       }*/


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
                       if (cellStatusDate.after(beginDate) && cellStatusDate.before(endDate)) {

                           rowR[Row] = sheetR.createRow(counter++);
                           //  System.out.println("status: " + statusDate + " beginDate " + beginDate + " endDate" + endDate);
     /*                      System.out.print(" A1: " + a1Val);
                           System.out.print(" C1: " + c1Val);
                           System.out.print(" D1: " + d1Val);
                           System.out.print(" E1: " + e1Val);
                           System.out.print(" F1: " + f1Val);
                           System.out.print(" G1: " + g1Val);
                           System.out.print(" H1: " + h1Val);
                           System.out.print(" I1: " + i1Val);
                           System.out.print(" J1: " + j1Val);*/
                         // System.out.print(" K1: " + datetemp.format(k1Val));
                     //      System.out.println();


                           //DATE PROCESSING

                         //  String dateString = i1Val;
                         //  Date date1 = null;
                       //    String PlanEntry = g1Val;
         /*                  Date date2 = null;
                           try {
                           //    date1 = new SimpleDateFormat("dd-MMM-yy").parse(dateString);
                               date2 = new SimpleDateFormat("dd-MMM-yy").parse(PlanEntry);
                           } catch (ParseException e) {
                               e.printStackTrace();
                           }*/

                           String str[] = df.format(cellStatusDate).split("-");
                           int year = Integer.parseInt(str[2]);
                           String j = str[2];

                           String str22[] = datetemp.format(cellPlanEntry).split("-");
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


                           arraylist.add(0, cellEmployeeID);  //employee number
                           arraylist.add(1, cellLastName);   //LAST NAME
                           arraylist.add(2, cellFirstName);//FN
                           arraylist.add(3, cellGender);//sex
                           arraylist.add(4, cellStatus);//type of term
                           arraylist.add(5, cellDOB);//dob
                           arraylist.add(6, cellDateofEmployment); //doh
                           arraylist.add(7, cellPlanEntry);//doe
                           arraylist.add(8, cellStatusDate);//DOT
                           arraylist.add(9, dateJ); //start of plan year of termination
                           arraylist.add(10,dateK); //end of plan year of termination
                           arraylist.add(11,dateL); //end of plan year of enrolment
                           arraylist.add(12, cellDateofRefund); //date of refund
                           arraylist.add(13,Double.parseDouble(dF.format(Utility.betweenDates(cellPlanEntry,cellDateofRefund)/365.25))); //period from DOE to DOR
                          arraylist.add(14,Double.parseDouble((dF.format(Utility.betweenDates(dateJ,cellDateofRefund)/365.25)))); //PERIOD FROM start of plan year of temrination to dor
                           arraylist.add(15,Double.parseDouble(dF.format(Utility.betweenDates(cellPlanEntry,dateL)/365.25))); //period from doe to end of plan year of enrolment
                           arraylist.add(16,Double.parseDouble(dF.format(Utility.betweenDates(cellDateofRefund,dateK)/365.25))); //period from dor to end of plan year of termination
                           arraylist.add(17,Double.parseDouble(dF.format(Utility.betweenDates(cellPlanEntry,cellStatusDate)/365.25))); //Pensonable Service up to DOT

                           //check if member is vested or non vested
                          String checkVested = "Non-Vested";
                          double PS= Double.parseDouble(dF.format(Utility.betweenDates(cellPlanEntry,cellDateofRefund)/365.25));
                          if(PS>5) checkVested="Vested";
                           arraylist.add(18,checkVested); //Pensonable Service up to DOT


                           for (int Col = 0, temp = 0; Col <19; Col++, temp++) {
                               //Update the value of cell
                               //   cellR = rowR[Row].getCell(Col);
                               //   if (cellR == null) {

                               cellR = rowR[Row].createCell(Col);

                               if (arraylist.get(Col) instanceof Date) {
                                   Date d = (Date) arraylist.get(Col);
                                   cellR.setCellValue(d);
                                   cellR.setCellStyle(cellStyle);
                               } else if (arraylist.get(Col) instanceof String)
                                   cellR.setCellValue((String) arraylist.get(Col));
                               else if (arraylist.get(Col) instanceof Double)
                                   cellR.setCellValue((Double) arraylist.get(Col));
                             //  cellR.setCellValue(String.valueOf(arraylist.get(Col)));
                           }
                       }

              //     }

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
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy"); // Just the year, with 2 digits
            String formattedDate = sdf.format(cal.getTime());

            // String StartDate = StartDay + "-" + StartMonth + "-" + StartYear; //"01-01-05";
            // String EndDate = EndDay + "-" + EndMonth + "-" + StartYear;//"31-12-05";
            String Recon = ("Actives at End of Plan Yr " + formattedDate);


            FileInputStream fs = null;
            try {
                fs = new FileInputStream(workingDir+ "\\Input Sheet.xlsx");
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            XSSFWorkbook WB = new XSSFWorkbook(fs);

//GET RECON SHEET
            XSSFSheet Reconsheet = WB.getSheet(Recon);
            System.out.println(formattedDate);
        //    int CountReconRow = Reconsheet.getPhysicalNumberOfRows();
            int CountReconRow = Utility.getNumberOfMembersInSheet(Reconsheet);

          //  int numOfTerminee = TermineeSheet.getLastRowNum()+1; // get the last row number
            int numOfTerminee = Utility.getNumberOfMembersInSheet(TermineeSheet);
            // System.out.println(rowCount);
            //     System.out.println("Recon"+ Recon + "count"+CountReconRow +"Actives" +numOfActives);
            XSSFRow[] rowR = new XSSFRow[numOfTerminee];
            Cell cellR = null;

         //   int Crow = 8;
            for (int row = 7,u=0; u < numOfTerminee; row++,u++) {

                int Row = row;

                XSSFRow TermineeRow = TermineeSheet.getRow(row);

                XSSFCell cellA = TermineeRow.getCell((short) 0);  //employee number
                String result = cellA.getStringCellValue();
                String cellEmployeeID = result.replaceAll("[-]","");

                XSSFCell cellB = TermineeRow.getCell((short) 1);   //last name
                if(cellB==null){
                    cellB = TermineeRow.createCell(1);
                }
                String cellLastName = cellB.getStringCellValue();

                XSSFCell cellC = TermineeRow.getCell((short) 2);   //first name
                if(cellC==null){
                    cellC=TermineeRow.createCell(2);
                }
                String cellFirstname = cellC.getStringCellValue();

                rowR[Row] = TermineeSheet.getRow(counter++);

                for (int y = 7,j=0; j < CountReconRow; y++,j++) {

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


                    if (cellEmployeeID.equals(a1) ) {
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
   //     int numOfActives = TermineeSheet.getLastRowNum()+1;//gets number of active members
        int numOfActives = Utility.getNumberOfMembersInSheet(TermineeSheet);
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


        //indexes
        int initialEmployee_BasicAcc_Index=19;
        int initialEmployee_VoluntaryAcc_Index=initialEmployee_BasicAcc_Index+1;
        int  initialEmployer_BasicAcc_Index=initialEmployee_VoluntaryAcc_Index+1;
        int initialEmployer_Optional_Index =  initialEmployer_BasicAcc_Index+1;

        int Contribution_MemberBasic_Index = initialEmployer_Optional_Index+1;
        int Contribution_MemberVoluntary_Index = Contribution_MemberBasic_Index+1;
        int Contribution_EmployerRequired_Index = Contribution_MemberVoluntary_Index+1;
        int Contribution_EmployerOptional_Index = Contribution_EmployerRequired_Index+1;

        int readCol=Contribution_MemberBasic_Index;//26;//start to read from Col 26, which is starting contribution column
        int YearCol = readCol;//as we are in the same year, we should always be reading from the correct column
        int Write_Coloumn =initialEmployee_BasicAcc_Index+8;//30;//start to write at column 30 which are the accumullated cells



        for(int row=7,I=0;I<numOfActives;row++,I++){ //get the initial accumulated balances

            XSSFRow ActiveRow = TermineeSheet.getRow(row);

            XSSFCell[] cellAccEmployeeBasic0 = new XSSFCell[numOfActives];
            cellAccEmployeeBasic0[I] = ActiveRow.getCell( initialEmployee_BasicAcc_Index);  //employee number
            if(cellAccEmployeeBasic0[I]==null){
                cellAccEmployeeBasic0[I] = ActiveRow.createCell(initialEmployee_BasicAcc_Index);
                cellAccEmployeeBasic0[I].setCellValue(0.00);
            }
            CellAccEmployeeBasic0[I]= cellAccEmployeeBasic0[I].getNumericCellValue();

            XSSFCell[] cellAccEmployeeOptional0 = new XSSFCell[numOfActives];
            cellAccEmployeeOptional0[I] = ActiveRow.getCell(initialEmployee_VoluntaryAcc_Index);  //employee number
            if(cellAccEmployeeOptional0[I]==null){
                cellAccEmployeeOptional0[I] = ActiveRow.createCell(initialEmployee_VoluntaryAcc_Index);
                cellAccEmployeeOptional0[I].setCellValue(0.00);
            }
            CellAccEmployeeOptional0[I]= cellAccEmployeeOptional0[I].getNumericCellValue();


            XSSFCell[] cellAccEmployerRequired0 = new XSSFCell[numOfActives];
            cellAccEmployerRequired0[I] = ActiveRow.getCell(initialEmployer_BasicAcc_Index); //employee number
            if(cellAccEmployerRequired0[I]==null){
                cellAccEmployerRequired0[I] = ActiveRow.createCell(initialEmployer_BasicAcc_Index);
                cellAccEmployerRequired0[I].setCellValue(0.00);
            }
            CellAccEmployerRequired0[I] = cellAccEmployerRequired0[I].getNumericCellValue();

            XSSFCell[] cellAccEmployerOptional0 = new XSSFCell[numOfActives];
            cellAccEmployerOptional0[I] = ActiveRow.getCell(initialEmployer_Optional_Index); //employee number
            if(cellAccEmployerOptional0[I]==null){
                cellAccEmployerOptional0[I] = ActiveRow.createCell(initialEmployer_Optional_Index);
                cellAccEmployerOptional0[I].setCellValue(0.00);
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

            for (int row = 7, I = 0; I < numOfActives; row++, I++) { //run for appropiate number of total active members
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
                    cellConEmployeeBasic.setCellValue(0.00);
                }
                double CellConEmployeeBasic = cellConEmployeeBasic.getNumericCellValue();
                readCol+=1;

                XSSFCell cellConEmployeeOptional = ActiveRow.getCell(readCol);
                if (cellConEmployeeOptional == null) {
                    cellConEmployeeOptional = ActiveRow.createCell(readCol);
                    cellConEmployeeOptional.setCellValue(0.00);
                }
                double CellConEmployeeOptional = cellConEmployeeOptional.getNumericCellValue();
                readCol+=1;


                XSSFCell cellConEmployerRequired = ActiveRow.getCell(readCol);
                if (cellConEmployerRequired == null) {
                    cellConEmployerRequired = ActiveRow.createCell(readCol);
                    cellConEmployerRequired.setCellValue(0.00);
                }
                double CellConEmployerRequired = cellConEmployerRequired.getNumericCellValue();
                readCol+=1;

                XSSFCell cellConEmployerOptional = ActiveRow.getCell(readCol);
                if (cellConEmployerOptional == null) {
                    cellConEmployerOptional = ActiveRow.createCell(readCol);
                    cellConEmployerOptional.setCellValue(0.00);
                }
                double CellConEmployerOptional = cellConEmployerOptional.getNumericCellValue();
                readCol+=1;


                //GET DATE OF TERMINATION
                XSSFCell cellI = ActiveRow.getCell(8);
                if (cellI == null) {
                    cellI = ActiveRow.createCell(8);
                    cellI.setCellValue(new Date("01-Jan-01"));
                }
                Date cellDateofTermination = cellI.getDateCellValue();

                //end of year of termination
                XSSFCell cellK = ActiveRow.getCell(10);
                if (cellK == null) {
                    cellK = ActiveRow.createCell(10);
                    cellK.setCellValue(new Date ("01-Jan-01"));
                }
                Date cellEndofPlanYearofTermination = cellK.getDateCellValue();//end of year of termination


 /*               Date statusDate = null;
                Date EndDateofTermination = null;
                try {
                    statusDate = datetemp.parse(CellDOT);
                    EndDateofTermination=datetemp.parse(CellD);
                } catch (ParseException e) {
                    e.printStackTrace();
                }
*/


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
                if (cellDateofTermination.after(BD) && cellDateofTermination.before(cellEndofPlanYearofTermination)) {
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
                    cellR.setCellValue(Double.parseDouble(dF.format(val.get(b))));
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
                XSSFWorkbook inputSheetWorkbok = new XSSFWorkbook(new FileInputStream(new File(workingDir + "\\Input Sheet.xlsx")));
                XSSFSheet inputSheet = inputSheetWorkbok.getSheet("Terminated up to " + EndYear + "." + EndMonth + "." + EndDay);
                int numOfTerminees = Utility.getNumberOfTermineeMembersInSheet(inputSheet);

                double [] reconCell_EmployeeBasicRefund = new double[numOfActives];
                double [] reconCell_EmployeeOptionalRefund = new double[numOfActives];
                String[] ReconcellEmployeeID = new String[numOfActives];

                for(int readRow=11,g=0;g<numOfActives;readRow++,g++){
                    XSSFRow getRow = inputSheet.getRow(readRow);
                    //get employee id
                    XSSFCell reconCellA = getRow.getCell(0);  //employee number
                    if (reconCellA == null) {
                        reconCellA = getRow.createCell((short) 0);
                        reconCellA.setCellValue("");
                    }
                    String resultRecon = reconCellA.getStringCellValue();
                    ReconcellEmployeeID[g] = resultRecon.replaceAll("[-]", "");


                    XSSFCell reconCellK = getRow.getCell(10);
                    if (reconCellK == null) {
                        reconCellK = getRow.createCell((short) 10);
                        reconCellK.setCellValue("");
                    }
                    reconCell_EmployeeBasicRefund[g] = reconCellK.getNumericCellValue();

                    XSSFCell reconCellL = getRow.getCell(11);
                    if (reconCellL == null) {
                        reconCellL = getRow.createCell((short) 11);
                        reconCellL.setCellValue("");
                    }
                    reconCell_EmployeeOptionalRefund[g] = reconCellL.getNumericCellValue();

                }


                int employeesBasicRefundIndex = readCol+3;
                int employeesOptionalRefundIndex = employeesBasicRefundIndex+1;
                int UnderOverIndex = readCol+5;

                //    readCol+=5;//move over by 3 columns to get to amount refunded
                ArrayList netFundYieldRates = new ArrayList();

                for (int I = 0, row=7; I < numOfActives; I++, row++) {
                    XSSFRow ActiveRow = TermineeSheet.getRow(row);

                    XSSFCell cellA = ActiveRow.getCell(0);
                    if (cellA == null) {
                        cellA = ActiveRow.createCell(0);
                        cellA.setCellValue("");
                    }

                    String name = cellA.getStringCellValue();
                    String name2 = name.replaceAll("[-]", "");
                    //end of year of termination
                    XSSFCell cellK = ActiveRow.getCell(10);
                    if (cellK == null) {
                        cellK = ActiveRow.createCell(10);
                        cellK.setCellValue(new Date("01-Jan-01"));
                    }
                    Date cellEndofPlanYearofTerminaton = cellK.getDateCellValue();//end of year of termination

               /*     //get the employee basic refund
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
                    double CellEOAmtRefunded = cellEOAmtRefunded.getNumericCellValue();*/

                    //get non-Vested or Vested
                    XSSFCell cellType= ActiveRow.getCell(18);
                    if (cellType == null) {
                        cellType = ActiveRow.createCell(18);
                        //   cellEBAmtRefunded.setCellValue(0);
                    }
                    String CellType= cellType.getStringCellValue();



                    int f;
                    for (f = 0; f < numOfActives; f++) {
                        if (ReconcellEmployeeID[f].equals(name2)) {
                            cellR = ActiveRow.createCell(employeesBasicRefundIndex);
                            cellR.setCellValue(reconCell_EmployeeBasicRefund[f]);

                            cellR = ActiveRow.createCell(employeesOptionalRefundIndex);
                            cellR.setCellValue(reconCell_EmployeeOptionalRefund[f]);

                            //WRITE TO UNDER/OVER CELLS
                            // for(int col=0;col<2;col++) {
                            //write to employee basic of under/over column
                            cellR = ActiveRow.createCell(UnderOverIndex);
                            double resultEB = reconCell_EmployeeBasicRefund[f] - newAccEmployeeBalance[I];
                            cellR.setCellValue(resultEB);

                            //write to employee optional of under/over column
                            cellR = ActiveRow.createCell(UnderOverIndex + 1);
                            double resultEO = reconCell_EmployeeOptionalRefund[f] - newAccEmployeeOptional[I];
                            cellR.setCellValue(resultEO);
                            //    }
                            break;
                        }
                    }
                    //write to the vesting column
                    double vestSign=0;
                    cellR = ActiveRow.createCell(UnderOverIndex+2);

                    if(CellType.equals("Vested"))  vestSign=1;
                    cellR.setCellValue(vestSign);

                    //write to the Er Non-Vested Bal
                    cellR = ActiveRow.createCell(UnderOverIndex+3);

                    double erVal=1;
                    netFundYieldRates=getNetFundYieldRates(workingDir,years,cellEndofPlanYearofTerminaton);

                    Double[] values = new Double[netFundYieldRates.size()];

                    netFundYieldRates.toArray(values);

                    for(int a=0;a<values.length;a++){

                        // System.out.println("length" + netFundYieldRates.size());

                        erVal*= (1+values[a]);
                        //   System.out.println("year: "+ (a+2004) + "="+netFundYieldRates.get(a));
                    }

//erVal*= newAccEmployerRequired[I];
                    erVal= newAccEmployerRequired[I] *erVal;
                    erVal=erVal*  (1-vestSign);
                    cellR.setCellValue(erVal);
                    // erVal=1;
                }

            }//at end of year

        }// END OF LOOP YEARS

        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Accumulated_Terminee_Sheet.xlsx"));
        workbook.write(outFile);
        fileInputStream.close();
        outFile.close();
    }//end of function Create Active Accumulated Balances

    //FEES
    public void Create_Fees_Activee_Contribution(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {

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
                fs = new FileInputStream(workingDir+"\\Valuation Data.xlsx");
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            XSSFWorkbook WB = new XSSFWorkbook(fs);

//GET RECON SHEET
            XSSFSheet Reconsheet = WB.getSheet(Recon);
      //      System.out.println(formattedDate);
            int CountReconRow = Reconsheet.getPhysicalNumberOfRows();

            int numOfActives = ActiveSheet.getLastRowNum()+1; // get the last row number

            XSSFRow[] rowR = new XSSFRow[numOfActives];
            Cell cellR = null;

          //  int Crow = 8;
            for (int row = 7; row < numOfActives; row++) {

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



                        val.add(0, h1); //employee basic
                        val.add(1, i1);  //employee optional
                        val.add(2, j1); //employer required
                        val.add(3, 0.00); //employer optional
                        val.add(4,CellFeeVal); //fees

                        for (int b = 0; b < val.size(); b++) {
                            cellR = rowR[Row].createCell(WriteAt + b);
                            cellR.setCellValue((Double)val.get(b));
                        }

                        break;
                    }

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
        int numOfActives = Utility.getNumberOfMembersInSheet(ActiveSheet);//gets number of active members
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


        //indexes
        int initialEmployee_BasicAcc_Index= 8 + years + 2;
        int initialEmployee_VoluntaryAcc_Index=initialEmployee_BasicAcc_Index+1;
        int  initialEmployer_BasicAcc_Index=initialEmployee_VoluntaryAcc_Index+1;
        int initialEmployer_Optional_Index =  initialEmployer_BasicAcc_Index+1;

        int Contribution_MemberBasic_Index = initialEmployer_Optional_Index+1;
        int Contribution_MemberVoluntary_Index = Contribution_MemberBasic_Index+1;
        int Contribution_EmployerRequired_Index = Contribution_MemberVoluntary_Index+1;
        int Contribution_EmployerOptional_Index = Contribution_EmployerRequired_Index+1;

        int readCol=Contribution_MemberBasic_Index;//26;//start to read from Col 26, which is starting contribution column
        int YearCol = readCol;//as we are in the same year, we should always be reading from the correct column
        int Write_Coloumn =initialEmployee_BasicAcc_Index+9;//30;//start to write at column 30 which are the accumullated cells

/*
        int readCol=26;//start to read from Col 26, which is starting contribution column
        int YearCol = readCol;//as we are in the same year, we should always be reading from the correct column
        int Write_Coloumn =31;//start to write at column 31 which are the accumullated cells
*/

        for(int row=7,I=0;I<numOfActives;row++,I++){ //get the initial accumulated balances

            XSSFRow ActiveRow = ActiveSheet.getRow(row);

            XSSFCell[] cellAccEmployeeBasic0 = new XSSFCell[numOfActives];
            cellAccEmployeeBasic0[I] = ActiveRow.getCell( initialEmployee_BasicAcc_Index);  //employee number
            if(cellAccEmployeeBasic0[I]==null){
                cellAccEmployeeBasic0[I] = ActiveRow.createCell(initialEmployee_BasicAcc_Index);
                cellAccEmployeeBasic0[I].setCellValue(0.00);
            }
            CellAccEmployeeBasic0[I]= cellAccEmployeeBasic0[I].getNumericCellValue();

            XSSFCell[] cellAccEmployeeOptional0 = new XSSFCell[numOfActives];
            cellAccEmployeeOptional0[I] = ActiveRow.getCell(initialEmployee_VoluntaryAcc_Index);  //employee number
            if(cellAccEmployeeOptional0[I]==null){
                cellAccEmployeeOptional0[I] = ActiveRow.createCell(initialEmployee_VoluntaryAcc_Index);
                cellAccEmployeeOptional0[I].setCellValue(0.00);
            }
            CellAccEmployeeOptional0[I]= cellAccEmployeeOptional0[I].getNumericCellValue();


            XSSFCell[] cellAccEmployerRequired0 = new XSSFCell[numOfActives];
            cellAccEmployerRequired0[I] = ActiveRow.getCell(initialEmployer_BasicAcc_Index); //employee number
            if(cellAccEmployerRequired0[I]==null){
                cellAccEmployerRequired0[I] = ActiveRow.createCell(initialEmployer_BasicAcc_Index);
                cellAccEmployerRequired0[I].setCellValue(0.00);
            }
            CellAccEmployerRequired0[I] = cellAccEmployerRequired0[I].getNumericCellValue();

            XSSFCell[] cellAccEmployerOptional0 = new XSSFCell[numOfActives];
            cellAccEmployerOptional0[I] = ActiveRow.getCell(initialEmployer_Optional_Index); //employee number
            if(cellAccEmployerOptional0[I]==null){
                cellAccEmployerOptional0[I] = ActiveRow.createCell(initialEmployer_Optional_Index);
                cellAccEmployerOptional0[I].setCellValue(0.00);
            }
            CellAccEmployerOptional0[I] = cellAccEmployerOptional0[I].getNumericCellValue();
        }//end of loop to get inital accumulated balances

        //MAIN PROCESSING-to get acc Balances and cont Balances
        for (int x = 0; x < years; x++) { //run for the appropiate number of years

            //  System.out.println(numOfActives);
            Cell cellR;

            for (int row = 7, I = 0; I < numOfActives; row++, I++) { //run for appropiate number of total active members
                readCol=YearCol;//need to ensure to start reading from same Column in same year

                XSSFRow ActiveRow = ActiveSheet.getRow(row);

                if (x == 0) {//get the accumulated balances just for the 1st year
                    CellAccEmployeeBasic = CellAccEmployeeBasic0[I];
                    CellAccEmployeeOptional = CellAccEmployeeOptional0[I];
                    CellAccEmployerRequired = CellAccEmployerRequired0[I];
                    CellAccEmployerOptional = CellAccEmployerOptional0[I];
                }

                else
                {
                    //get accumulated balances for every year after 1st year
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


                //FORMULA CALCULATIONS
                newAccEmployeeBalance[I] = ((CellAccEmployeeBasic * (1+interestValues[x])) + (CellConEmployeeBasic * (1+(interestValues[x]*0.5))));//CellAccEmployeeBasic * (1 + 1) + CellConEmployeeBasic * (1 + 1 * 0.5);
                newAccEmployeeOptional[I] =((CellAccEmployeeOptional * (1+interestValues[x])) + (CellConEmployeeOptional * (1+(interestValues[x]*0.5))));//CellAccEmployeeOptional +  CellConEmployeeOptional; CellAccEmployeeOptional * (1 + 1) + CellConEmployeeOptional * (1 + 1 * 0.5);
                newAccEmployerRequired[I] = ((CellAccEmployerRequired * (1+interestValues[x])) + (CellConEmployerRequired *(1+(interestValues[x]*0.5))) + (CellFees * (1+(interestValues[x]*0.5))));//CellAccEmployerRequired * (1 + 1) + CellConEmployerRequired * (1 + 1 * 0.5) + CellFees * (1 + 1 * 0.5);
                newAccEmployerOptional[I] =((CellAccEmployerOptional * (1+interestValues[x])) + (CellConEmployerOptional * (1+(interestValues[x]*0.5))));//CellAccEmployerOptional * (1 + 1) + CellConEmployerOptional * (1 + 1 * 0.5);

                val.add(0, newAccEmployeeBalance[I]);
                val.add(1, newAccEmployeeOptional[I]);
                val.add(2, newAccEmployerRequired[I]);
                val.add(3, newAccEmployerOptional[I]);

//write the calculated accumulated balances to the sheet; start to write after the initial set of startbalances
                for (int b = 0; b < 4; b++) {
                    cellR = ActiveRow.createCell(Write_Coloumn + b);
                    cellR.setCellValue(Double.parseDouble(dF.format(val.get(b))));
                }//end of loop

            }//end of looping through each member

            //MOVING THE INDEXES
            YearCol=readCol+5;//move over by 5 columns to get next set of contributions for next year
            readCol=YearCol;//give the readCol 5 so that, it can always read the correct set of contributions of that same year
            Write_Coloumn += 9;//8 no fees  || 9 fees
            StartYear++;//increment Start year by one until we reach end year

            //when we are at end of year...we should write account balance as at end date
            if(x==(years-1)) {
                for (int I = 0, row=7; I < numOfActives; I++, row++) {
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
 /*       int readCol=23;//start to read from Col 26, which is starting contribution column
        int YearCol = readCol;//as we are in the same year, we should always be reading from the correct column
        int Write_Coloumn=28;//start to write at column 31 which are the accumullated cells
*/
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
     //   int numOfActives = TermineeSheet.getLastRowNum()+1;//gets number of active members
     int numOfActives = Utility.getNumberOfMembersInSheet(TermineeSheet);
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

        //indexes
        int initialEmployee_BasicAcc_Index=19;
        int initialEmployee_VoluntaryAcc_Index=initialEmployee_BasicAcc_Index+1;
        int  initialEmployer_BasicAcc_Index=initialEmployee_VoluntaryAcc_Index+1;
        int initialEmployer_Optional_Index =  initialEmployer_BasicAcc_Index+1;

        int Contribution_MemberBasic_Index = initialEmployer_Optional_Index+1;
        int Contribution_MemberVoluntary_Index = Contribution_MemberBasic_Index+1;
        int Contribution_EmployerRequired_Index = Contribution_MemberVoluntary_Index+1;
        int Contribution_EmployerOptional_Index = Contribution_EmployerRequired_Index+1;

        int readCol=Contribution_MemberBasic_Index;//26;//start to read from Col 26, which is starting contribution column
        int YearCol = readCol;//as we are in the same year, we should always be reading from the correct column
        int Write_Coloumn =initialEmployee_BasicAcc_Index+9;//30;//start to write at column 30 which are the accumullated cells


        for(int row=7,I=0;I<numOfActives;row++,I++){ //get the initial accumulated balances

            XSSFRow ActiveRow = TermineeSheet.getRow(row);

            XSSFCell[] cellAccEmployeeBasic0 = new XSSFCell[numOfActives];
            cellAccEmployeeBasic0[I] = ActiveRow.getCell( initialEmployee_BasicAcc_Index);  //employee number
            if(cellAccEmployeeBasic0[I]==null){
                cellAccEmployeeBasic0[I] = ActiveRow.createCell(initialEmployee_BasicAcc_Index);
                cellAccEmployeeBasic0[I].setCellValue(0.00);
            }
            CellAccEmployeeBasic0[I]= cellAccEmployeeBasic0[I].getNumericCellValue();

            XSSFCell[] cellAccEmployeeOptional0 = new XSSFCell[numOfActives];
            cellAccEmployeeOptional0[I] = ActiveRow.getCell(initialEmployee_VoluntaryAcc_Index);  //employee number
            if(cellAccEmployeeOptional0[I]==null){
                cellAccEmployeeOptional0[I] = ActiveRow.createCell(initialEmployee_VoluntaryAcc_Index);
                cellAccEmployeeOptional0[I].setCellValue(0.00);
            }
            CellAccEmployeeOptional0[I]= cellAccEmployeeOptional0[I].getNumericCellValue();


            XSSFCell[] cellAccEmployerRequired0 = new XSSFCell[numOfActives];
            cellAccEmployerRequired0[I] = ActiveRow.getCell(initialEmployer_BasicAcc_Index); //employee number
            if(cellAccEmployerRequired0[I]==null){
                cellAccEmployerRequired0[I] = ActiveRow.createCell(initialEmployer_BasicAcc_Index);
                cellAccEmployerRequired0[I].setCellValue(0.00);
            }
            CellAccEmployerRequired0[I] = cellAccEmployerRequired0[I].getNumericCellValue();

            XSSFCell[] cellAccEmployerOptional0 = new XSSFCell[numOfActives];
            cellAccEmployerOptional0[I] = ActiveRow.getCell(initialEmployer_Optional_Index); //employee number
            if(cellAccEmployerOptional0[I]==null){
                cellAccEmployerOptional0[I] = ActiveRow.createCell(initialEmployer_Optional_Index);
                cellAccEmployerOptional0[I].setCellValue(0.00);
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
            for (int row = 7, I = 0; I < numOfActives; row++, I++) { //run for appropiate number of total active members
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
                XSSFCell cellI = ActiveRow.getCell(8);
                if (cellI == null) {
                    cellI = ActiveRow.createCell(8);
                    cellI.setCellValue(new Date("01-Jan-01"));
                }
               Date cellDateofTermination = cellI.getDateCellValue();

                //end of year of termination
                XSSFCell cellK = ActiveRow.getCell(10);
                if (cellK == null) {
                    cellK = ActiveRow.createCell(10);
                    cellK.setCellValue(new Date("01-Jan-01"));
                }
                Date cellEndofPlanYearofTermination = cellK.getDateCellValue();//end of year of termination


                //FORMULA CALCULATIONS
                if (cellDateofTermination.after(BD) && cellDateofTermination.before(cellEndofPlanYearofTermination)) {
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
                        cellR.setCellValue(Double.parseDouble(dF.format(val.get(b))));
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
                 XSSFWorkbook inputSheetWorkbok = new XSSFWorkbook(new FileInputStream(new File(workingDir + "\\Input Sheet.xlsx")));
                 XSSFSheet inputSheet = inputSheetWorkbok.getSheet("Terminated up to " + EndYear + "." + EndMonth + "." + EndDay);
             //    int numOfTerminees = Utility.getNumberOfTermineeMembersInSheet(inputSheet);

             double [] reconCell_EmployeeBasicRefund = new double[numOfActives];
             double [] reconCell_EmployeeOptionalRefund = new double[numOfActives];
             String[] ReconcellEmployeeID = new String[numOfActives];

                 for(int readRow=11,g=0;g<numOfActives;readRow++,g++){
                     XSSFRow getRow = inputSheet.getRow(readRow);
                     //get employee id
                     XSSFCell reconCellA = getRow.getCell(0);  //employee number
                     if (reconCellA == null) {
                         reconCellA = getRow.createCell((short) 0);
                         reconCellA.setCellValue("");
                     }
                     String resultRecon = reconCellA.getStringCellValue();
                      ReconcellEmployeeID[g] = resultRecon.replaceAll("[-]", "");


                     XSSFCell reconCellK = getRow.getCell(10);
                     if (reconCellK == null) {
                         reconCellK = getRow.createCell((short) 10);
                         reconCellK.setCellValue("");
                     }
                    reconCell_EmployeeBasicRefund[g] = reconCellK.getNumericCellValue();

                     XSSFCell reconCellL = getRow.getCell(11);
                     if (reconCellL == null) {
                         reconCellL = getRow.createCell((short) 11);
                         reconCellL.setCellValue("");
                     }
                     reconCell_EmployeeOptionalRefund[g] = reconCellL.getNumericCellValue();

                 }


             int employeesBasicRefundIndex = readCol+3;
             int employeesOptionalRefundIndex = employeesBasicRefundIndex+1;
             int UnderOverIndex = readCol+5;

           //    readCol+=5;//move over by 3 columns to get to amount refunded
             ArrayList netFundYieldRates = new ArrayList();

                for (int I = 0, row=7; I < numOfActives; I++, row++) {

                    XSSFRow ActiveRow = TermineeSheet.getRow(row);

                    XSSFCell cellA = ActiveRow.getCell(0);
                    if (cellA == null) {
                        cellA = ActiveRow.createCell(0);
                        cellA.setCellValue("");
                    }

                    String name = cellA.getStringCellValue();
                    String name2 = name.replaceAll("[-]", "");
                    //end of year of termination
                    XSSFCell cellK = ActiveRow.getCell(10);
                    if (cellK == null) {
                        cellK = ActiveRow.createCell(10);
                        cellK.setCellValue(new Date("01-Jan-01"));
                    }
                    Date cellEndofPlanYearofTermination = cellK.getDateCellValue();//end of year of termination

                    //get the employee basic refund
                    XSSFCell cellEBAmtRefunded = ActiveRow.getCell(readCol + 3);
                    if (cellEBAmtRefunded == null) {
                        cellEBAmtRefunded = ActiveRow.createCell(readCol + 3);
                        //   cellEBAmtRefunded.setCellValue(0);
                    }
                    double CellEBAmtRefunded = cellEBAmtRefunded.getNumericCellValue();

                    //   AmtRefundedindex+=1;//move over to amount refunded for employee optional
                    //get the employee optional refund
                    XSSFCell cellEOAmtRefunded = ActiveRow.getCell(employeesBasicRefundIndex + 1);
                    if (cellEOAmtRefunded == null) {
                        cellEOAmtRefunded = ActiveRow.createCell(employeesBasicRefundIndex + 1);
                        //   cellEOAmtRefunded.setCellValue(0);
                    }
                    double CellEOAmtRefunded = cellEOAmtRefunded.getNumericCellValue();


                    //get non-Vested or Vested
                    XSSFCell cellType = ActiveRow.getCell(18);
                    if (cellType == null) {
                        cellType = ActiveRow.createCell(18);
                        //   cellEBAmtRefunded.setCellValue(0);
                    }
                    String CellType = cellType.getStringCellValue();

                    int f;
                    for (f = 0; f < numOfActives; f++) {
                        if (ReconcellEmployeeID[f].equals(name2)) {
                            cellR = ActiveRow.createCell(employeesBasicRefundIndex);
                            cellR.setCellValue(reconCell_EmployeeBasicRefund[f]);

                            cellR = ActiveRow.createCell(employeesOptionalRefundIndex);
                            cellR.setCellValue(reconCell_EmployeeOptionalRefund[f]);

                    //WRITE TO UNDER/OVER CELLS
                    // for(int col=0;col<2;col++) {
                    //write to employee basic of under/over column
                    cellR = ActiveRow.createCell(UnderOverIndex);
                    double resultEB = reconCell_EmployeeBasicRefund[f] - newAccEmployeeBalance[I];
                    cellR.setCellValue(resultEB);

                    //write to employee optional of under/over column
                    cellR = ActiveRow.createCell(UnderOverIndex + 1);
                    double resultEO = reconCell_EmployeeOptionalRefund[f] - newAccEmployeeOptional[I];
                    cellR.setCellValue(resultEO);
                    //    }
                            break;
                        }
                    }
                    //write to the vesting column
                    double vestSign = 0;
                    cellR = ActiveRow.createCell(UnderOverIndex + 2);

                    if (CellType.equals("Vested")) vestSign = 1;
                    cellR.setCellValue(vestSign);

                    //write to the Er Non-Vested Bal
                    cellR = ActiveRow.createCell(UnderOverIndex + 3);

                    double erVal = 1;
                    netFundYieldRates = getNetFundYieldRates(workingDir, years, cellEndofPlanYearofTermination);

                    Double[] values = new Double[netFundYieldRates.size()];

                    netFundYieldRates.toArray(values);

                    for (int a = 0; a < values.length; a++) {

                        // System.out.println("length" + netFundYieldRates.size());

                        erVal *= (1 + values[a]);
                        //   System.out.println("year: "+ (a+2004) + "="+netFundYieldRates.get(a));
                    }

//erVal*= newAccEmployerRequired[I];
                    erVal = newAccEmployerRequired[I] * erVal;
                    erVal = erVal * (1 - vestSign);
                    cellR.setCellValue(erVal);
                    // erVal=1;
                }

            }//at end of year

        }// END OF LOOP YEARS
        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Accumulated_Terminee_Sheet.xlsx"));
        workbook.write(outFile);
        fileInputStream.close();
        outFile.close();
    }//end of function Create Active Accumulated Balances

    public ArrayList getNetFundYieldRates(String workingDir, int numOfYears,Date dateofTermination) throws IOException {
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(workingDir + "\\nfy Rates.xlsx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet =workbook.getSheet("rates");
        XSSFCell [] cell = new XSSFCell[numOfYears];
        ArrayList values = new  ArrayList();

      //  SimpleDateFormat datetemp = new SimpleDateFormat("dd-MMM-yy");
    //    SimpleDateFormat datetemp2 = new SimpleDateFormat("yyyy.mm.dd");

      //  Date DateofTermination = null;
     //   Date EndDateofnfy = null;

      //      DateofTermination = dateofTermination;



        for(int row=1,I=0;I<numOfYears;row++,I++) {
            XSSFRow nfyRow = sheet.getRow(row);

            XSSFCell cellnfy = nfyRow.getCell(2);

       /*     if (cellnfy == null) {
                cellnfy = nfyRow.createCell(2);
              //  cellnfy.setCellValue(01-Jan-01);
            }*/
           Date CellDatenfy = cellnfy.getDateCellValue();//end of year of termination


            if(dateofTermination.before(CellDatenfy)){
                cell[I] = nfyRow.getCell(1);
                //    list.add(cell[I].getNumericCellValue(),I);
                values.add(cell[I].getNumericCellValue());
            }


        }

        return values;
    }

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

    public ArrayList View_Actives_Members(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IndexOutOfBoundsException, IOException {
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
        StringBuilder stringBuilder = new StringBuilder();

        stringBuilder.append("The following is a list of Active Members present as at " + EndYear + "." + EndMonth + "." + EndDay + " \n\n");

        SimpleDateFormat datetemp = new SimpleDateFormat("dd-MMM-yy");
        ArrayList list = new ArrayList<String>();

            //OPEN ACTIVE SHEET
            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Input Sheet.xlsx");
            XSSFWorkbook workbookInputSheet = new XSSFWorkbook(fileInputStream);

            XSSFSheet Initialsheet = workbookInputSheet.getSheet("Actives at End of Plan Yr "+StartYear);
            int InitialnumOfActives = Utility.getNumberOfMembersInSheet(Initialsheet);

            String [] InitialcellEmployeeID = new String[InitialnumOfActives];
            String[] InitialcellLastName = new String[InitialnumOfActives];
            String[] InitialcellFirstName = new String [InitialnumOfActives];

            //get the initial members in first year period
            for (int row = 0, readFromRow = 7; row < InitialnumOfActives; row++, readFromRow++) {

                //get the current row position based on value of row
                XSSFRow rowPosition = Initialsheet.getRow(readFromRow);
                //     XSSFRow rowPosition = Reconsheet.getRow(readFromRow);

                //get employee id
                XSSFCell cellA = rowPosition.getCell(0);  //employee number
                if (cellA == null) {
                    cellA = rowPosition.createCell((short) 0);
                    cellA.setCellValue("");
                }
                String result = cellA.getStringCellValue();
                InitialcellEmployeeID[row] = result.replaceAll("[-]", "");

                //get LAST NAME
                XSSFCell cellB = rowPosition.getCell((short) 1);  //last name
                if (cellB == null) {
                    cellB = rowPosition.createCell((short) 1);
                    cellB.setCellValue("");
                }
                InitialcellLastName[row] = cellB.getStringCellValue();

                //get FIRST NAME
                XSSFCell cellC = rowPosition.getCell((short) 2);  //first name
                if (cellC == null) {
                    cellC = rowPosition.createCell((short) 2);
                    cellC.setCellValue("");
                }
                InitialcellFirstName[row] = cellC.getStringCellValue();


                //get GENDER
                XSSFCell cellD = rowPosition.getCell((short) 3);  //gender
                if (cellD == null) {
                    cellD = rowPosition.createCell((short) 3);
                    cellD.setCellValue("");
                }
                String cellGender = cellD.getStringCellValue();

                //get Marital Status
                XSSFCell cellE = rowPosition.getCell((short) 4);  //Marital Status
                if (cellE == null) {
                    cellE = rowPosition.createCell((short) 4);
                    cellE.setCellValue("");
                }
                String cellMaritalStatus = cellE.getStringCellValue();

                //get Date of birth
                XSSFCell cellF = rowPosition.getCell((short) 5);  // Date of birth
                //   if (cellF == null) {
                //       cellF = rowPosition.createCell((short) 5);
                //      cellF.setCellValue(new Date("01-Jan-01"));
                //  }
                Date cellDateofBirth = cellF.getDateCellValue();

                //get Date of Employment
                XSSFCell cellG = rowPosition.getCell((short) 6);  //Date of Employment
                // if (cellG == null) {
                //     cellG = rowPosition.createCell((short) 6);
                //      cellG.setCellValue(new Date("01-Jan-01"));
                //   }
                Date cellDateofEmployment = cellG.getDateCellValue();


                //get Date of Enrollment
                XSSFCell cellH = rowPosition.getCell((short) 7);  //Date of Employment
                //  if (cellH == null) {
                //      cellH = rowPosition.createCell((short) 7);
                //   cellH.setCellValue(new Date("01-Jan-01"));
                //    }
                Date cellDateofEnrolment = cellH.getDateCellValue();

                stringBuilder.append("\nEmployee ID: " +    InitialcellEmployeeID[row] + "\n");
                stringBuilder.append("Last Name: " +    InitialcellLastName[row] + "\n");
                stringBuilder.append("First Name: " +InitialcellFirstName[row] + "\n");
                stringBuilder.append("Date of Birth: " + datetemp.format(cellDateofBirth)+ "\n");
                stringBuilder.append("Employment Date: " + datetemp.format(cellDateofEmployment) + "\n");
                stringBuilder.append("Plan Entry Date: " + datetemp.format(cellDateofEnrolment) + "\n");
                //        stringBuilder.append("Status Date: " + datetemp.format(c) + "\n");
                //    stringBuilder.append("Status: " + ReconCellStatus + "\n");
                //     stringBuilder.append("-------------------------------------------------------\n");
                list.add(   InitialcellEmployeeID[row]+ "," +    InitialcellLastName[row] + "," + InitialcellFirstName[row] + "," + datetemp.format(cellDateofBirth) + "," + datetemp.format(cellDateofEmployment) + "," + datetemp.format(cellDateofEnrolment));


            }
            boolean foundIt = false;
            //LOOP THROUGH THE OTHER YEARS
            int InitialMembers = InitialnumOfActives;
            int yearsRemaining = years-1;


            for (int x = 0; x < yearsRemaining; x++) {
                int newStartYear = StartYear+1;

                Calendar cal = Calendar.getInstance();
                cal.set((newStartYear+x), StartMonth, StartDay);
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy"); // Just the year, with 2 digits
                String formattedDate = sdf.format(cal.getTime());
                //  String formattedDate = "2014";
                String Recon = ("Actives at End of Plan Yr " + formattedDate);

//GET RECON SHEET
                XSSFSheet Reconsheet = workbookInputSheet.getSheet(Recon);

                int currentNumber = Utility.getNumberOfMembersInSheet(Reconsheet);//get curent members from recon row
                String [] cellEmployeeID = new String[currentNumber];
                String[] cellLastName = new String[currentNumber];
                String[] cellFirstName = new String [currentNumber];

                for (int readFromRow2 = 7, rowIterator2 = 0; rowIterator2 < currentNumber; readFromRow2++, rowIterator2++) {
                    XSSFRow getRow2 = Reconsheet.getRow(readFromRow2);

                    //get employee id
                    XSSFCell cellA = getRow2.getCell(0);  //employee number
                    if (cellA == null) {
                        cellA = getRow2.createCell((short) 0);
                        cellA.setCellValue("");
                    }
                    String result = cellA.getStringCellValue();
                    cellEmployeeID[rowIterator2]= result.replaceAll("[-]", "");

                    //get LAST NAME
                    XSSFCell cellB = getRow2.getCell((short) 1);  //last name
                    if (cellB == null) {
                        cellB = getRow2.createCell((short) 1);
                        cellB.setCellValue("");
                    }
                    cellLastName[rowIterator2]= cellB.getStringCellValue();

                    //get FIRST NAME
                    XSSFCell cellC = getRow2.getCell((short) 2);  //first name
                    if (cellC == null) {
                        cellC = getRow2.createCell((short) 2);
                        cellC.setCellValue("");
                    }
                    cellFirstName[rowIterator2] = cellC.getStringCellValue();

                    //get DOB
                    XSSFCell ReconCellF = getRow2.getCell((short) 5);
                    if (ReconCellF == null) {
                        ReconCellF = getRow2.createCell((short) 5);
                        // ReconCellF.setCellValue("");
                    }
                    Date ReconCellDOB = ReconCellF.getDateCellValue();

                    //get date of employment
                    XSSFCell ReconCellG = getRow2.getCell((short) 6);
                    if (ReconCellG == null) {
                        ReconCellG = getRow2.createCell((short) 6);
                        // ReconCellF.setCellValue("");
                    }
                    Date ReconCelldateofEmployment = ReconCellG.getDateCellValue();

                    //get DOB
                    XSSFCell ReconCellH = getRow2.getCell((short) 7);  //first name
                    if (ReconCellH == null) {
                        ReconCellH = getRow2.createCell((short) 7);
                        // ReconCellF.setCellValue("");
                    }
                    Date ReconCellDateofEnrolment = ReconCellH.getDateCellValue();

                    boolean findIt=false;

                    //test the value from recon sheet against initial
                    for (int initial = 0 ; initial < InitialMembers; initial++) {

                   //     System.out.print("initial " + initial);

                        if (cellEmployeeID[rowIterator2].equals(InitialcellEmployeeID[initial]) && cellLastName[rowIterator2].equals(InitialcellLastName[initial] )) {
                            findIt=true;
                            break;
                        }
                    }

                    if (!findIt){
                        stringBuilder.append("\nEmployee ID: " + cellEmployeeID[rowIterator2] + "\n");
                        stringBuilder.append("Last Name: " + cellLastName[rowIterator2] + "\n");
                        stringBuilder.append("First Name: " + cellFirstName[rowIterator2] + "\n");
                        stringBuilder.append("Date of Birth: " + datetemp.format(ReconCellDOB)+ "\n");
                        stringBuilder.append("Employment Date: " + datetemp.format(ReconCelldateofEmployment) + "\n");
                        stringBuilder.append("Plan Entry Date: " + datetemp.format(ReconCellDateofEnrolment) + "\n");

                        list.add(cellEmployeeID[rowIterator2]+ "," +  cellLastName[rowIterator2] + "," + cellFirstName[rowIterator2] + "," +datetemp.format(ReconCellDOB) + "," + datetemp.format(ReconCelldateofEmployment) + "," + datetemp.format(ReconCellDateofEnrolment));
                    }


                }//end loop years
                InitialcellEmployeeID = new String[currentNumber];
                InitialcellLastName =  new String[currentNumber];
                InitialcellFirstName = new String[currentNumber];
                for(int h=0;h<currentNumber;h++){
                    System.out.println("currentNumber"+currentNumber);
                    System.out.println("h"+h);
                    InitialcellEmployeeID[h]=cellEmployeeID[h];
                    InitialcellLastName[h] = cellLastName[h];
                    InitialcellFirstName[h]=cellFirstName[h];
                }

                InitialMembers = currentNumber;
            }//end of looping through each of members

                        this.setResult(String.valueOf(stringBuilder));
        return list;
    }// end of view active sheet

    public ArrayList View_Retired_Members(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IndexOutOfBoundsException {
        String ED[] = PensionPlanEndDate.split("/");
        int endMonth = Integer.parseInt(ED[0]);
        int endDay = Integer.parseInt(ED[1]);
        int endYear = Integer.parseInt(ED[2]);

        int EndYear = endYear;
        int EndMonth = endMonth;
        int EndDay = endDay;
        ArrayList<String> list= new ArrayList<>();
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("The following is a list of Retired Members present as at "+PensionPlanEndDate+" \n\n");
        SimpleDateFormat dF1 = new SimpleDateFormat();
        dF1.applyPattern("dd-MMM-yy");
        try {

            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Input Sheet.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet worksheet = workbook.getSheet("Terminated up to "+EndYear+"."+EndMonth+"."+EndDay);
            int numOfTerminees = Utility.getNumberOfTermineeMembersInSheet(worksheet);



            SimpleDateFormat dateF = new SimpleDateFormat("dd-MMM-yy");

            for (int readFromRow = 11,row=0; row < numOfTerminees; readFromRow++,row++) {

                XSSFRow getRow = worksheet.getRow(readFromRow);

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

                //get date of enrolment
                XSSFCell ReconCellH = getRow.getCell((short) 7);
                if (ReconCellH == null) {
                    ReconCellH = getRow.createCell((short) 7);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellDateofEnrolment = ReconCellH.getDateCellValue();


                XSSFCell ReconCellI = getRow.getCell((short) 8);  //status date
                if (ReconCellI == null) {
                    ReconCellI = getRow.createCell((short) 8);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellStatusDate = ReconCellI.getDateCellValue();

                XSSFCell ReconCellJ = getRow.getCell((short) 9);  //status date
                if (ReconCellJ == null) {
                    ReconCellJ = getRow.createCell((short) 9);
                    // ReconCellF.setCellValue("");
                }
                String ReconCellStatus = ReconCellJ.getStringCellValue();

                if (ReconCellStatus.equals("R")) {

                    stringBuilder.append("Employee ID: " + ReconcellEmployeeID + "\n");
                    stringBuilder.append("Last Name: " + ReconCellLastName + "\n");
                    stringBuilder.append("First Name: " + ReconCellFirstName + "\n");
                    stringBuilder.append("Date of Birth: " + dateF.format(ReconCellDOB)+ "\n");
                    stringBuilder.append("Employment Date: " + dateF.format(ReconCelldateofEmployment) + "\n");
                    stringBuilder.append("Plan Entry Date: " + dateF.format(ReconCellDateofEnrolment) + "\n");
                    stringBuilder.append("Status Date: " + dateF.format(ReconCellStatusDate) + "\n");
                    stringBuilder.append("Status: " + ReconCellStatus + "\n");
                    stringBuilder.append("-------------------------------------------------------\n");

                    this.setResult(String.valueOf(stringBuilder));
                    list.add(ReconcellEmployeeID + "," + ReconCellLastName + "," + ReconCellFirstName + "," + dateF.format(ReconCellDOB) + "," + dateF.format(ReconCelldateofEmployment) + "," + dateF.format(ReconCellDateofEnrolment));

                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (NullPointerException e) {
            e.printStackTrace();

        }
     //   return String.valueOf(stringBuilder);
        this.setResult(String.valueOf(stringBuilder));
        return list;
    }// end of view active sheet

    public ArrayList View_Terminee_Members(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IndexOutOfBoundsException {
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

        ArrayList<String> list= new ArrayList<>();
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("The following is a list of Terminee Members present as at "+PensionPlanEndDate+" \n\n");
        SimpleDateFormat dF1 = new SimpleDateFormat();
        dF1.applyPattern("dd-MMM-yy");
        try {
            // FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\Seperated Members.xlsx");
            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Input Sheet.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet worksheet = workbook.getSheet("Terminated up to "+EndYear+"."+EndMonth+"."+EndDay);
            //   XSSFSheet sheet = workbook.getSheetAt(0);

            int numOfTerminees = Utility.getNumberOfTermineeMembersInSheet(worksheet);


            SimpleDateFormat dateF = new SimpleDateFormat("dd-MMM-yy");

            for (int readFromRow = 11,row=0; row < numOfTerminees; readFromRow++,row++) {

                XSSFRow getRow = worksheet.getRow(readFromRow);

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

                //get date of enrolment
                XSSFCell ReconCellH = getRow.getCell((short) 7);
                if (ReconCellH == null) {
                    ReconCellH = getRow.createCell((short) 7);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellDateofEnrolment = ReconCellH.getDateCellValue();


                XSSFCell ReconCellI = getRow.getCell((short) 8);  //status date
                if (ReconCellI == null) {
                    ReconCellI = getRow.createCell((short) 8);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellStatusDate = ReconCellI.getDateCellValue();

                XSSFCell ReconCellJ = getRow.getCell((short) 9);  //status date
                if (ReconCellJ == null) {
                    ReconCellJ = getRow.createCell((short) 9);
                    // ReconCellF.setCellValue("");
                }
                String ReconCellStatus = ReconCellJ.getStringCellValue();

                if ((ReconCellStatus.equals("D")||ReconCellStatus.equals("R")|| ReconCellStatus.equals("T"))) {

                    stringBuilder.append("Employee ID: " + ReconcellEmployeeID + "\n");
                    stringBuilder.append("Last Name: " + ReconCellLastName + "\n");
                    stringBuilder.append("First Name: " + ReconCellFirstName + "\n");
                    stringBuilder.append("Date of Birth: " + dateF.format(ReconCellDOB)+ "\n");
                    stringBuilder.append("Employment Date: " + dateF.format(ReconCelldateofEmployment) + "\n");
                    stringBuilder.append("Plan Entry Date: " + dateF.format(ReconCellDateofEnrolment) + "\n");
                    stringBuilder.append("Status Date: " + dateF.format(ReconCellStatusDate) + "\n");
                    stringBuilder.append("Status: " + ReconCellStatus + "\n");
                    stringBuilder.append("-------------------------------------------------------\n");

                    this.setResult(String.valueOf(stringBuilder));
                    list.add(ReconcellEmployeeID + "," + ReconCellLastName + "," + ReconCellFirstName + "," + dateF.format(ReconCellDOB) + "," + dateF.format(ReconCelldateofEmployment) + "," + dateF.format(ReconCellDateofEnrolment));


                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (NullPointerException e) {
            e.printStackTrace();

        }
      //  return String.valueOf(stringBuilder);

        return  list;
    }// end of view active sheet

    public ArrayList View_Deceased_Members(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IndexOutOfBoundsException {
        String ED[] = PensionPlanEndDate.split("/");
        int endMonth = Integer.parseInt(ED[0]);
        int endDay = Integer.parseInt(ED[1]);
        int endYear = Integer.parseInt(ED[2]);

        int EndYear = endYear;
        int EndMonth = endMonth;
        int EndDay = endDay;
        ArrayList<String> list= new ArrayList<>();
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("The following is a list of Deceased Members present as at "+PensionPlanEndDate+" \n\n");
        SimpleDateFormat dF1 = new SimpleDateFormat();
        dF1.applyPattern("dd-MMM-yy");
        try {

            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Input Sheet.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet worksheet = workbook.getSheet("Terminated up to "+EndYear+"."+EndMonth+"."+EndDay);
            int numOfTerminees = Utility.getNumberOfTermineeMembersInSheet(worksheet);



            SimpleDateFormat dateF = new SimpleDateFormat("dd-MMM-yy");

            for (int readFromRow = 11,row=0; row < numOfTerminees; readFromRow++,row++) {

                XSSFRow getRow = worksheet.getRow(readFromRow);

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

                //get date of enrolment
                XSSFCell ReconCellH = getRow.getCell((short) 7);
                if (ReconCellH == null) {
                    ReconCellH = getRow.createCell((short) 7);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellDateofEnrolment = ReconCellH.getDateCellValue();


                XSSFCell ReconCellI = getRow.getCell((short) 8);  //status date
                if (ReconCellI == null) {
                    ReconCellI = getRow.createCell((short) 8);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellStatusDate = ReconCellI.getDateCellValue();

                XSSFCell ReconCellJ = getRow.getCell((short) 9);  //status date
                if (ReconCellJ == null) {
                    ReconCellJ = getRow.createCell((short) 9);
                    // ReconCellF.setCellValue("");
                }
                String ReconCellStatus = ReconCellJ.getStringCellValue();

                if (ReconCellStatus.equals("D")) {

                    stringBuilder.append("Employee ID: " + ReconcellEmployeeID + "\n");
                    stringBuilder.append("Last Name: " + ReconCellLastName + "\n");
                    stringBuilder.append("First Name: " + ReconCellFirstName + "\n");
                    stringBuilder.append("Date of Birth: " + dateF.format(ReconCellDOB)+ "\n");
                    stringBuilder.append("Employment Date: " + dateF.format(ReconCelldateofEmployment) + "\n");
                    stringBuilder.append("Plan Entry Date: " + dateF.format(ReconCellDateofEnrolment) + "\n");
                    stringBuilder.append("Status Date: " + dateF.format(ReconCellStatusDate) + "\n");
                    stringBuilder.append("Status: " + ReconCellStatus + "\n");
                    stringBuilder.append("-------------------------------------------------------\n");

                    this.setResult(String.valueOf(stringBuilder));
                    list.add(ReconcellEmployeeID + "," + ReconCellLastName + "," + ReconCellFirstName + "," + dateF.format(ReconCellDOB) + "," + dateF.format(ReconCelldateofEmployment) + "," + dateF.format(ReconCellDateofEnrolment));

                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (NullPointerException e) {
            e.printStackTrace();

        }
        //   return String.valueOf(stringBuilder);
        this.setResult(String.valueOf(stringBuilder));
        return list;
    }// end of view active sheet

    public ArrayList View_Terminated_Members(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IndexOutOfBoundsException {
        String ED[] = PensionPlanEndDate.split("/");
        int endMonth = Integer.parseInt(ED[0]);
        int endDay = Integer.parseInt(ED[1]);
        int endYear = Integer.parseInt(ED[2]);

        int EndYear = endYear;
        int EndMonth = endMonth;
        int EndDay = endDay;
        ArrayList<String> list= new ArrayList<>();
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("The following is a list of Terminated Members present as at"+PensionPlanEndDate+" \n\n");
        SimpleDateFormat dF1 = new SimpleDateFormat();
        dF1.applyPattern("dd-MMM-yy");
        try {

            FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Input Sheet.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet worksheet = workbook.getSheet("Terminated up to "+EndYear+"."+EndMonth+"."+EndDay);
            int numOfTerminees = Utility.getNumberOfTermineeMembersInSheet(worksheet);



            SimpleDateFormat dateF = new SimpleDateFormat("dd-MMM-yy");

            for (int readFromRow = 11,row=0; row < numOfTerminees; readFromRow++,row++) {

                XSSFRow getRow = worksheet.getRow(readFromRow);

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

                //get date of enrolment
                XSSFCell ReconCellH = getRow.getCell((short) 7);
                if (ReconCellH == null) {
                    ReconCellH = getRow.createCell((short) 7);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellDateofEnrolment = ReconCellH.getDateCellValue();


                XSSFCell ReconCellI = getRow.getCell((short) 8);  //status date
                if (ReconCellI == null) {
                    ReconCellI = getRow.createCell((short) 8);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellStatusDate = ReconCellI.getDateCellValue();

                XSSFCell ReconCellJ = getRow.getCell((short) 9);  //status date
                if (ReconCellJ == null) {
                    ReconCellJ = getRow.createCell((short) 9);
                    // ReconCellF.setCellValue("");
                }
                String ReconCellStatus = ReconCellJ.getStringCellValue();

                if (ReconCellStatus.equals("T")) {

                    stringBuilder.append("Employee ID: " + ReconcellEmployeeID + "\n");
                    stringBuilder.append("Last Name: " + ReconCellLastName + "\n");
                    stringBuilder.append("First Name: " + ReconCellFirstName + "\n");
                    stringBuilder.append("Date of Birth: " + dateF.format(ReconCellDOB)+ "\n");
                    stringBuilder.append("Employment Date: " + dateF.format(ReconCelldateofEmployment) + "\n");
                    stringBuilder.append("Plan Entry Date: " + dateF.format(ReconCellDateofEnrolment) + "\n");
                    stringBuilder.append("Status Date: " + dateF.format(ReconCellStatusDate) + "\n");
                    stringBuilder.append("Status: " + ReconCellStatus + "\n");
                    stringBuilder.append("-------------------------------------------------------\n");

                    this.setResult(String.valueOf(stringBuilder));
                    list.add(ReconcellEmployeeID + "," + ReconCellLastName + "," + ReconCellFirstName + "," + dateF.format(ReconCellDOB) + "," + dateF.format(ReconCelldateofEmployment) + "," + dateF.format(ReconCellDateofEnrolment));

                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (NullPointerException e) {
            e.printStackTrace();

        }
        //   return String.valueOf(stringBuilder);
        this.setResult(String.valueOf(stringBuilder));
        return list;
    }// end of view active sheet

    public double getExcessShortFall(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException{
        DecimalFormat DF = new DecimalFormat("#.##");//#.##
        FileInputStream file = new FileInputStream(workingDir + "\\Income_Expenditure_Table.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);
        DecimalFormat dF = new DecimalFormat("#");//#.##

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
            startDate = df.parse(StartYear+"."+StartMonth+"."+StartDay);
            endDate = df.parse(EndYear+"."+EndMonth+"."+EndDay);
        } catch (ParseException e) {
            e.printStackTrace();
        }

        int years= Utility.getDiffYears(startDate,endDate);
years+=1;
//get the Total investment income
       XSSFRow Row= sheet.getRow(37);
        XSSFCell cellTotalInvestmentIncome= Row.getCell(years+1);
        if(cellTotalInvestmentIncome==null){
            cellTotalInvestmentIncome=Row.createCell(37);
            cellTotalInvestmentIncome.setCellValue(0.00);
        }
     double   totalInvestmentIncome = cellTotalInvestmentIncome.getNumericCellValue();

        //get the Total expenses
     Row= sheet.getRow(36);
        XSSFCell cellTotalExpenses= Row.getCell(years+1);
        if(cellTotalExpenses==null){
            cellTotalExpenses=Row.createCell(36);
            cellTotalExpenses.setCellValue(0.00);
        }
        double   totalExpenses = cellTotalExpenses.getNumericCellValue();

        //calculate Net investment income
        double netInvestmentIncome = totalInvestmentIncome - totalExpenses;

        double interestCreditedActives =0;
        double interestCreditedTerminees=0;

SimpleDateFormat dateFormat = new SimpleDateFormat("dd-mm-yyyy");

        for(int x=0;x<years;x++) {
            Calendar cal = Calendar.getInstance();
            cal.set(StartYear, StartMonth, StartDay);
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy"); // Just the year, with 4 digits
            String formattedDate = sdf.format(cal.getTime());

            String startofCurrentYear = StartDay + "-" + StartMonth + "-" + StartYear; //"01-01-05";
            String endofCurrentYear = EndDay + "-" + EndMonth + "-" + StartYear;//"31-12-05";

            String Recon = ("Actives at End of Plan Yr "+formattedDate);

            interestCreditedTerminees+=Double.valueOf(dF.format(TermineeSumReader(PensionPlanEndDate,startofCurrentYear, endofCurrentYear, Recon,workingDir)));
double temp = Double.valueOf(dF.format(TermineeSumReader(PensionPlanEndDate,startofCurrentYear, endofCurrentYear, Recon,workingDir)));
            interestCreditedActives+=(Double.valueOf(dF.format(ActiveSumReader(PensionPlanEndDate,startofCurrentYear, endofCurrentYear, Recon,workingDir)))-temp);

/*            System.out.println("startofCurrentYear "+startofCurrentYear);
            System.out.println("endofCurrentYear "+endofCurrentYear);
            System.out.println("interestCreditedActives"+Double.valueOf(dF.format(ActiveSumReader(PensionPlanEndDate,startofCurrentYear, endofCurrentYear, Recon,workingDir)-temp)));
         System.out.println("interestCreditedTerminees"+  Double.valueOf(dF.format(TermineeSumReader(PensionPlanEndDate,startofCurrentYear, endofCurrentYear, Recon,workingDir))));*/
            StartYear++;
        }

        double sumActiveTermineeInterest = interestCreditedActives+interestCreditedTerminees;
     double result = netInvestmentIncome - sumActiveTermineeInterest;
        return Double.parseDouble(DF.format(result));
    }

    public double TermineeSumReader(String PensionPlanEndDate, String startofCurrentYear, String endofCurrentYear, String Recon, String workingDir) throws IndexOutOfBoundsException {

        String ED[] = PensionPlanEndDate.split("/");
        int endMonth = Integer.parseInt(ED[0]);
        int endDay = Integer.parseInt(ED[1]);
        int endYear = Integer.parseInt(ED[2]);

        int EndYear = endYear;
        int EndMonth = endMonth;
        int EndDay = endDay;


        double TActiveSum = 0;
        try {
            FileInputStream fileInputStream = new FileInputStream(workingDir+"\\Input Sheet.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

            XSSFSheet sheetTerminated = workbook.getSheet("Terminated up to "+EndYear+"."+EndMonth+"."+EndDay);
            XSSFSheet Reconsheet = workbook.getSheet(Recon);


            int numOfTerminees =Utility.getNumberOfTermineeMembersInSheet(sheetTerminated);
         int numOfActives = Utility.getNumberOfMembersInSheet(Reconsheet);


            double ActiveSum = 0;
            TActiveSum = 0;
            double TermineeSum = 0;

            int count = 0;
            for (int readFromTerminee = 11, row=0;row < numOfTerminees; readFromTerminee++,row++) {  //start looping through demo records

                XSSFRow getTermineeRow = sheetTerminated.getRow(readFromTerminee);

                //get employee id
                XSSFCell reconCellA = getTermineeRow.getCell(0);  //employee number
                if (reconCellA == null) {
                    reconCellA = getTermineeRow.createCell((short) 0);
                    reconCellA.setCellValue("");
                }
                String resultRecon = reconCellA.getStringCellValue();
                String ReconcellEmployeeID = resultRecon.replaceAll("[-]", "");

                //get LAST NAME
                XSSFCell ReconCellB = getTermineeRow.getCell((short) 1);  //last name
                if (ReconCellB == null) {
                    ReconCellB = getTermineeRow.createCell((short) 1);
                    ReconCellB.setCellValue("");
                }
                String ReconCellLastName = ReconCellB.getStringCellValue();

                //get FIRST NAME
                XSSFCell ReconCellC = getTermineeRow.getCell((short) 2);  //first name
                if (ReconCellC == null) {
                    ReconCellC = getTermineeRow.createCell((short) 2);
                    ReconCellC.setCellValue("");
                }
                String ReconCellFirstName = ReconCellC.getStringCellValue();

                //get status Date/ dTE OF TERMINATION
                XSSFCell ReconCellF = getTermineeRow.getCell((short) 8);
                if (ReconCellF == null) {
                    ReconCellF = getTermineeRow.createCell((short) 8);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellStatusdate= ReconCellF.getDateCellValue();

                //get date of employment
                XSSFCell ReconCellG = getTermineeRow.getCell((short) 9);
                if (ReconCellG == null) {
                    ReconCellG = getTermineeRow.createCell((short) 9);
                    // ReconCellF.setCellValue("");
                }
                String ReconCellStatus = ReconCellG.getStringCellValue();


           //loop through each member in active sheet and check i
                for (int readFromActive = 7, activeRow=0; activeRow < numOfActives; readFromActive++,activeRow++) {
                    XSSFRow Reconrow = Reconsheet.getRow(readFromActive);

                    XSSFCell reconCellA2 = Reconrow.getCell((short) 0);
                    //Update the value of cell
                    if (reconCellA2 == null) {
                        reconCellA2 = Reconrow.createCell(0);
                    }
                    String ReconA1Val = reconCellA2.getStringCellValue();

                    String ReconcellEmployeeID2 = ReconA1Val.replaceAll("[-]", "");



                    XSSFCell reconCellB = Reconrow.getCell((short) 1);
                    //Update the value of cell
                    if (reconCellB == null) {
                        reconCellB = Reconrow.createCell(1);
                    }
                    String ReconB1Val = reconCellB.getStringCellValue();


                    XSSFCell ReconCellC1 = Reconrow.getCell((short) 2);
                    //Update the value of cell
                    if (ReconCellC1 == null) {
                        ReconCellC1 = Reconrow.createCell(2);
                    }
                    String ReconC1Val = ReconCellC1.getStringCellValue();


                    XSSFCell ReconCellQ = Reconrow.getCell((short) 16);
                    if (ReconCellQ == null) {
                        ReconCellQ = Reconrow.createCell(16);
                    }
                    double ReconCell_EmployeeBasicInterest = ReconCellQ.getNumericCellValue();


                    XSSFCell ReconCellR= Reconrow.getCell((short) 17);
                    if (ReconCellR == null) {
                        ReconCellR = Reconrow.createCell(17);
                    }
                    double ReconCell_EmployeeOptionalInterest = ReconCellR.getNumericCellValue();


                    XSSFCell ReconCellS= Reconrow.getCell((short) 18);
                    //Update the value of cell
                    if (ReconCellS == null) {
                        ReconCellS = Reconrow.createCell(18);
                    }
                    double ReconCell_EmployerRequiredInterest = ReconCellS.getNumericCellValue();

                    SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yy");

                    Date StartofCurrentYear = null;
                    Date  EndofCurrentYear=null;

                    try {
                        StartofCurrentYear = sdf.parse(startofCurrentYear);
                        EndofCurrentYear = sdf.parse(endofCurrentYear);
                    } catch (ParseException e) {
                        e.printStackTrace();
                    }

                    if (ReconCellStatus.equals("R") || ReconCellStatus.equals("D")  && (ReconCellStatusdate.after(StartofCurrentYear) && ReconCellStatusdate.before(EndofCurrentYear)) && ReconcellEmployeeID.equals(ReconcellEmployeeID2)) {
                        ReconCellStatus = "T";
                    }

                    if (ReconCellStatus.equals("T") && !ReconcellEmployeeID.equals("ASSE88888") && ReconcellEmployeeID.equals(ReconcellEmployeeID2)&& (ReconCellStatusdate.after(StartofCurrentYear) && ReconCellStatusdate.before(EndofCurrentYear))) {
                        TActiveSum += ReconCell_EmployeeBasicInterest + ReconCell_EmployeeOptionalInterest + ReconCell_EmployerRequiredInterest;
                    }

                    count++;
                }


            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (NullPointerException e) {
            e.printStackTrace();

        }
        return TActiveSum;
    }

    public double ActiveSumReader(String PensionPlanEndDate, String startofCurrentYear, String endofCurrentYear, String Recon, String workingDir) throws IndexOutOfBoundsException {

        String ED[] = PensionPlanEndDate.split("/");
        int endMonth = Integer.parseInt(ED[0]);
        int endDay = Integer.parseInt(ED[1]);
        int endYear = Integer.parseInt(ED[2]);

        int EndYear = endYear;
        int EndMonth = endMonth;
        int EndDay = endDay;


        double TActiveSum = 0;
        try {
            FileInputStream fileInputStream = new FileInputStream(workingDir+"\\Input Sheet.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

            XSSFSheet sheetTerminated = workbook.getSheet("Terminated up to "+EndYear+"."+EndMonth+"."+EndDay);
            XSSFSheet Reconsheet = workbook.getSheet(Recon);


            int numOfTerminees =Utility.getNumberOfTermineeMembersInSheet(sheetTerminated);
            int numOfActives = Utility.getNumberOfMembersInSheet(Reconsheet);

System.out.println("numOfActives"+numOfActives);
            double ActiveSum = 0;

            double TermineeSum = 0;
            int count = 0;

            for (int readFromTerminee = 11, row=0;row < numOfTerminees; readFromTerminee++,row++) {  //start looping through demo records
                TActiveSum = 0;
                boolean isTerminated =false;
                XSSFRow getTermineeRow = sheetTerminated.getRow(readFromTerminee);

                //get employee id
                XSSFCell reconCellA = getTermineeRow.getCell(0);  //employee number
                if (reconCellA == null) {
                    reconCellA = getTermineeRow.createCell((short) 0);
                    reconCellA.setCellValue("");
                }
                String resultRecon = reconCellA.getStringCellValue();
                String ReconcellEmployeeID = resultRecon.replaceAll("[-]", "");

                //get LAST NAME
                XSSFCell ReconCellB = getTermineeRow.getCell((short) 1);  //last name
                if (ReconCellB == null) {
                    ReconCellB = getTermineeRow.createCell((short) 1);
                    ReconCellB.setCellValue("");
                }
                String ReconCellLastName = ReconCellB.getStringCellValue();

                //get FIRST NAME
                XSSFCell ReconCellC = getTermineeRow.getCell((short) 2);  //first name
                if (ReconCellC == null) {
                    ReconCellC = getTermineeRow.createCell((short) 2);
                    ReconCellC.setCellValue("");
                }
                String ReconCellFirstName = ReconCellC.getStringCellValue();

                //get status Date/ dTE OF TERMINATION
                XSSFCell ReconCellF = getTermineeRow.getCell((short) 8);
                if (ReconCellF == null) {
                    ReconCellF = getTermineeRow.createCell((short) 8);
                    // ReconCellF.setCellValue("");
                }
                Date ReconCellStatusdate= ReconCellF.getDateCellValue();

                //get date of employment
                XSSFCell ReconCellG = getTermineeRow.getCell((short) 9);
                if (ReconCellG == null) {
                    ReconCellG = getTermineeRow.createCell((short) 9);
                    // ReconCellF.setCellValue("");
                }
                String ReconCellStatus = ReconCellG.getStringCellValue();


                //loop through each member in active sheet and check i
                for (int readFromActive = 7, activeRow=0; activeRow < numOfActives; readFromActive++,activeRow++) {

                    XSSFRow Reconrow = Reconsheet.getRow(readFromActive);

                    XSSFCell reconCellA2 = Reconrow.getCell((short) 0);
                    //Update the value of cell
                    if (reconCellA2 == null) {
                        reconCellA2 = Reconrow.createCell(0);
                    }
                    String ReconA1Val = reconCellA2.getStringCellValue();
                    String ReconcellEmployeeID2 = ReconA1Val.replaceAll("[-]", "");


                    XSSFCell reconCellB = Reconrow.getCell((short) 1);
                    //Update the value of cell
                    if (reconCellB == null) {
                        reconCellB = Reconrow.createCell(1);
                    }
                    String ReconB1Val = reconCellB.getStringCellValue();


                    XSSFCell ReconCellC1 = Reconrow.getCell((short) 2);
                    //Update the value of cell
                    if (ReconCellC1 == null) {
                        ReconCellC1 = Reconrow.createCell(2);
                    }
                    String ReconC1Val = ReconCellC1.getStringCellValue();


                    XSSFCell ReconCellQ = Reconrow.getCell((short) 16);
                    if (ReconCellQ == null) {
                        ReconCellQ = Reconrow.createCell(16);
                    }
                    double ReconCell_EmployeeBasicInterest = ReconCellQ.getNumericCellValue();


                    XSSFCell ReconCellR= Reconrow.getCell((short) 17);
                    if (ReconCellR == null) {
                        ReconCellR = Reconrow.createCell(17);
                    }
                    double ReconCell_EmployeeOptionalInterest = ReconCellR.getNumericCellValue();


                    XSSFCell ReconCellS= Reconrow.getCell((short) 18);
                    //Update the value of cell
                    if (ReconCellS == null) {
                        ReconCellS = Reconrow.createCell(18);
                    }
                    double ReconCell_EmployerRequiredInterest = ReconCellS.getNumericCellValue();

                    SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yy");

                    Date StartofCurrentYear = null;
                    Date  EndofCurrentYear=null;

                    try {
                        StartofCurrentYear = sdf.parse(startofCurrentYear);
                        EndofCurrentYear = sdf.parse(endofCurrentYear);
                    } catch (ParseException e) {
                        e.printStackTrace();
                    }
                   // if (ReconCellStatusdate.after(StartofCurrentYear) || (ReconCellStatusdate.before(EndofCurrentYear))) {
                 //  isTerminated = true;
                //    }

                //    if (!isTerminated)
                  //      {
                            TActiveSum += ReconCell_EmployeeBasicInterest + ReconCell_EmployeeOptionalInterest + ReconCell_EmployerRequiredInterest;
                      //  }
                  // }

                    count++;
                }

            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (NullPointerException e) {
            e.printStackTrace();

        }
        return TActiveSum;
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
                        if(arraylist.get(Col) instanceof Date)
                            cellR.setCellValue((Date)arraylist.get(Col));
                        else if(arraylist.get(Col) instanceof Boolean)
                            cellR.setCellValue((Boolean)arraylist.get(Col));
                        else if(arraylist.get(Col) instanceof String)
                            cellR.setCellValue((String)arraylist.get(Col));
                        else if(arraylist.get(Col) instanceof Double)
                            cellR.setCellValue((Double)arraylist.get(Col));
                    //   cellR.setCellValue(String.valueOf(arraylist.get(Col)));
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

    public static void AddsheetintoExistingworkbook(String workDir, String sheetname) throws IOException, InvalidFormatException {

        //***************************Add a sheet into Existing workbook***********************************************


        String path=workDir+ "\\Tables\\All_Tables.xlsx";
       FileInputStream fileinp = new FileInputStream(path);
      XSSFWorkbook  workbook = new XSSFWorkbook(fileinp);
        workbook.createSheet(sheetname);

      FileOutputStream  fileOut = new FileOutputStream(path);
        workbook.write(fileOut);
        fileOut.close();
        System.out.println("File is written successfully");
    }

}

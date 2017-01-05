import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by akonowalchuk on 1/5/2017.
 */
public class Table {

    public void Create_Table_Summary_of_Active_Membership(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException{
        DecimalFormat dF = new DecimalFormat("#.#");//#.##
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

        FileInputStream fileR = new FileInputStream(workingDir + "\\Template_Summary_of_Active_Membership.xlsx");
        XSSFWorkbook workbookTemplate = new XSSFWorkbook(fileR);
        XSSFSheet sheetTemplate = workbookTemplate.getSheetAt(0);

        FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Accumulated_Actives_Sheet.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet ActiveSheet = workbook.getSheet("Actives");

    //    XSSFRow Row = ActiveSheet.getRow(row);
        int numOfActives = ActiveSheet.getLastRowNum();
        numOfActives+=1;

        //GET THE TOTAL MALE AND FEMALES
        int maleCount=0;
        int femaleCount=0;
        int total=0;

        for (int row = 7; row < numOfActives; row++) {

            XSSFRow Row = ActiveSheet.getRow(row);

            XSSFCell cellGender = Row.getCell(3);
            if (cellGender == null) {
                cellGender = Row.createCell(3);
                cellGender.setCellValue("");
            }
            String memberGender= cellGender.getStringCellValue();
            memberGender=memberGender.toLowerCase();

         //   System.out.println("F: "+femaleCount);
        //    System.out.println("m: "+maleCount);
            if (memberGender.equals("f"))femaleCount++;
            if (memberGender.equals("m"))maleCount++;
        }
        //END GET THE TOTAL MALE AND FEMALES


        //AVERAGE PENSIONABLE SALARY FOR MALE AND FEMALE
        double femalePensionableSalary=0;
        double malePensionableSalary=0;
        double pensionableSalaryAmount=0;
        int startofPensionableSalaryIndex=7;
        int endofPensionableSalaryIndex=7+years;
        int ageColumnIndex=endofPensionableSalaryIndex+1;
        double  memberAge=0;
        double sumMaleAge=0;
        double sumfemaleAge=0;

        int pensionableServiceColumnIndex=ageColumnIndex+1;
        double sumfemalePensionableService=0;
        double sumMalePensionableService=0;


      //  for(int k=0;k<years;k++) {
            //GET DATA FROM TEMPLATE
            for (int x = 0, row = 7; row < numOfActives; x++, row++) {

                XSSFRow Row = ActiveSheet.getRow(row);

                XSSFCell cellGender = Row.getCell(3);
                if (cellGender == null) {
                    cellGender = Row.createCell(3);
                    cellGender.setCellValue("");
                }
                String memberGender = cellGender.getStringCellValue();
                memberGender = memberGender.toLowerCase();

                //get age
                XSSFRow rowAge = ActiveSheet.getRow(row);
                XSSFCell cellAge = rowAge.getCell(ageColumnIndex);
                if (cellAge == null) {
                    cellAge = rowAge.createCell(ageColumnIndex);
                    cellAge.setCellValue(0.00);
                }
                 memberAge= cellAge.getNumericCellValue();

                if (memberGender.equals("f")) sumfemaleAge += memberAge;
                if (memberGender.equals("m")) sumMaleAge += memberAge;


                //get pensionable service
                XSSFRow rowPS = ActiveSheet.getRow(row);
                XSSFCell cellPS = rowPS.getCell(pensionableServiceColumnIndex);
                if (cellPS == null) {
                    cellPS = rowAge.createCell(pensionableServiceColumnIndex);
                    cellPS.setCellValue(0.00);
                }
                double pensionableService= cellPS.getNumericCellValue();

                if (memberGender.equals("f")) sumfemalePensionableService += pensionableService;
                if (memberGender.equals("m")) sumMalePensionableService += pensionableService;


//get pensionable salary
                XSSFRow Row1 = ActiveSheet.getRow(row);
                XSSFCell cellValues = Row1.getCell(endofPensionableSalaryIndex);
                if (cellValues == null) {
                    cellValues = Row1.createCell(endofPensionableSalaryIndex);
                    cellValues.setCellValue(0.00);
                }
                pensionableSalaryAmount = cellValues.getNumericCellValue();

                if (memberGender.equals("f")) femalePensionableSalary += pensionableSalaryAmount;
                if (memberGender.equals("m")) malePensionableSalary += pensionableSalaryAmount;
            }//end of looping through mwmbwea



        //WRITE female and male totals data to cells
        //male total
        XSSFRow row = sheetTemplate.getRow(3);
        XSSFCell writeCell= row.createCell(1);
        writeCell.setCellValue(maleCount);

        //female total
       row = sheetTemplate.getRow(3);
         writeCell= row.createCell(2);
        writeCell.setCellValue(femaleCount);

        //male and female total
        row = sheetTemplate.getRow(3);
        writeCell= row.createCell(3);
        total=femaleCount+maleCount;
        writeCell.setCellValue(Double.parseDouble(dF.format(total)));

        //WRITE AVERAGE AGE MALE AND FEMALE
        //male AGE
        row = sheetTemplate.getRow(4);
        writeCell= row.createCell(1);
        double avgMaleAge= sumMaleAge/maleCount;
        writeCell.setCellValue(Double.parseDouble(dF.format(avgMaleAge)));

        // female AGE
        row = sheetTemplate.getRow(4);
        writeCell= row.createCell(2);
        double avgFemaleAge= sumfemaleAge/femaleCount;
        writeCell.setCellValue(Double.parseDouble(dF.format(avgFemaleAge)));

        // averge for male and female AGE
        row = sheetTemplate.getRow(4);
        writeCell= row.createCell(3);
        double avgAge= (avgFemaleAge+avgMaleAge)/2;
        writeCell.setCellValue(Double.parseDouble(dF.format(avgAge)));


        //WRITE PENSIONABLE SERVICE FOR MALE AND FEMALE
        //male PS
        row = sheetTemplate.getRow(5);
        writeCell= row.createCell(1);
        double avgMalePS= sumMalePensionableService/maleCount;
        writeCell.setCellValue(Double.parseDouble(dF.format(avgMalePS)));

        // female AGE
        row = sheetTemplate.getRow(5);
        writeCell= row.createCell(2);
        double avgFemalePS= sumfemalePensionableService/femaleCount;
        writeCell.setCellValue(Double.parseDouble(dF.format(avgFemalePS)));

        // averge for male and female AGE
        row = sheetTemplate.getRow(5);
        writeCell= row.createCell(3);
        double avgPS= (avgMalePS+avgFemalePS)/2;
        writeCell.setCellValue(Double.parseDouble(dF.format(avgPS)));




        //WRITE AVERAGE PENSIONABLE SALARY TO THE CELLS
             row= sheetTemplate.getRow(6);
             double avgMalePensionableSalary=malePensionableSalary/maleCount;
            row.createCell(1).setCellValue(Double.parseDouble(dF.format(avgMalePensionableSalary)));

       // row = ActiveSheet.getRow(6);
        double avgFemalePensionableSalary=femalePensionableSalary/femaleCount;
        row.createCell(2).setCellValue(Double.parseDouble(dF.format(avgFemalePensionableSalary)));

      //  row = ActiveSheet.getRow(6);
        double avgtotalPensionableSalary=(avgFemalePensionableSalary+avgMalePensionableSalary)/2;
        row.createCell(3).setCellValue(Double.parseDouble(dF.format(avgtotalPensionableSalary)));

        for (int x = 0; x <4; x++) {
            //   sheet.autoSizeColumn(x);
            sheetTemplate.autoSizeColumn(x);

        }


        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Table_Summary_of_Active_Membership.xlsx"));
        workbookTemplate.write(outFile);
        fileInputStream.close();
        outFile.close();
    }

    public void Create_Table_Movement_in_Active_Membership(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException{
        DecimalFormat dF = new DecimalFormat("#.#");//#.##
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

        FileInputStream fileR = new FileInputStream(workingDir + "\\Template_Movement_in_Active_Membership.xlsx");
        XSSFWorkbook workbookTemplate = new XSSFWorkbook(fileR);
        XSSFSheet sheetTemplate = workbookTemplate.getSheetAt(0);

        FileInputStream fileInputStream = new FileInputStream(workingDir + "\\Accumulated_Actives_Sheet.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet ActiveSheet = workbook.getSheet("Actives");

        //get Active members as at start of plan year if any
        XSSFRow Row = ActiveSheet.getRow(3);
        //get male and female active members as at start date
        XSSFCell cellMaleActiveMember = Row.getCell(1);
        if (cellMaleActiveMember == null) {
            cellMaleActiveMember = Row.createCell(1);
            cellMaleActiveMember.setCellValue(0);
        }
       double maleActiveMembers = cellMaleActiveMember.getNumericCellValue();

        //get female active members as at start date
        XSSFCell cellFemaleActiveMember = Row.getCell(2);
        if (cellFemaleActiveMember == null) {
            cellFemaleActiveMember = Row.createCell(1);
            cellFemaleActiveMember.setCellValue(0);
        }
        double femaleActiveMembers = cellFemaleActiveMember.getNumericCellValue();

        //get both sex active members as at start date
        XSSFCell cellBothActiveMember = Row.getCell(3);
        if (cellBothActiveMember == null) {
            cellBothActiveMember = Row.createCell(1);
            cellBothActiveMember.setCellValue(0);
        }
        double bothActiveMembers = cellBothActiveMember.getNumericCellValue();


        //write the initial data

        //male initial active members
        XSSFRow row = sheetTemplate.createRow(3);
        XSSFCell writeCell= row.createCell(1);
        writeCell.setCellValue(maleActiveMembers);
        writeCell= row.createCell(2);
        writeCell.setCellValue(femaleActiveMembers);
        writeCell= row.createCell(3);
        writeCell.setCellValue(bothActiveMembers);





        for (int x = 0; x <4; x++) {
            sheetTemplate.autoSizeColumn(x);
        }

        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Table_Movement_in_Active_Membership.xlsx"));

        workbookTemplate.write(outFile);
        fileInputStream.close();
        outFile.close();
    }

    public void Create_Table_Analysis_of_Fund_Yield(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException{

    }

    public void Create_Table_Gains_and_Losses(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException{

    }

}

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.NoSuchFileException;
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

    public void Create_Table_Analysis_of_Fund_Yield(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {
        try {
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
            years += 1;

            FileInputStream fileTemplate = new FileInputStream(workingDir + "\\Template_Analysis_of_Fund_Yield.xlsx");
            XSSFWorkbook workbookTemplate= new XSSFWorkbook(fileTemplate);
            XSSFSheet sheetTemplate = workbookTemplate.getSheetAt(0);

            FileInputStream fileIncExp = new FileInputStream(workingDir + "\\Income_Expenditure_Table.xlsx");
            XSSFWorkbook workbookIncExp = new XSSFWorkbook(fileIncExp);
            XSSFSheet sheetIncExp = workbookIncExp.getSheetAt(0);

            FileInputStream fileInterest = new FileInputStream(workingDir + "\\Interest Rates.xlsx");
            XSSFWorkbook workbookInterest = new XSSFWorkbook(fileInterest);
            XSSFSheet sheetInterest = workbookInterest.getSheetAt(0);

            XSSFRow row;

            //XSSFCell cell;
            double[] valuesCIR = new double[years];
            //GET CIR
            for(int j = 0, cellIterator=1;j<years;j++,cellIterator++){
                row = sheetInterest.getRow(cellIterator);
                XSSFCell cellCIR = row.getCell(1);
                if(cellCIR==null){
                    cellCIR = row.createCell(cellIterator);
                    cellCIR.setCellValue(0.00);
                }
                valuesCIR[j]=cellCIR.getNumericCellValue();
            }


            //XSSFCell cell;
            double[] valuesGFY = new double[years];
            //GET GFY
            for(int j = 0, cellIterator=1;j<years;j++,cellIterator++){
                  row = sheetIncExp.getRow(39);
                XSSFCell cellGFY = row.getCell(cellIterator);
                if(cellGFY==null){
                    cellGFY = row.createCell(cellIterator);
                    cellGFY.setCellValue(0.00);
                }
                valuesGFY[j]=cellGFY.getNumericCellValue();
            }


            double[] valuesAFY = new double[years];
            //GET AFY
            for(int j = 0, cellIterator=1;j<years;j++,cellIterator++){
                row = sheetIncExp.getRow(40);
                XSSFCell cellAFY= row.getCell(cellIterator);
                if(cellAFY==null){
                    cellAFY = row.createCell(cellIterator);
                    cellAFY.setCellValue(0.00);
                }
                valuesAFY[j]=cellAFY.getNumericCellValue();
            }

            double[] valuesNFY = new double[years];
            //GET NFY
            for(int j = 0, cellIterator=1;j<years;j++,cellIterator++){
                row = sheetIncExp.getRow(41);
                XSSFCell cellNFY= row.getCell(cellIterator);
                if(cellNFY==null){
                    cellNFY = row.createCell(cellIterator);
                    cellNFY.setCellValue(0.00);
                }
                valuesNFY[j]=cellNFY.getNumericCellValue();
            }

            double[] valuesPYI= new double[years];
            //GET NFY
            for(int j = 0, cellIterator=1;j<years;j++,cellIterator++){
                row = sheetIncExp.getRow(43);
                XSSFCell cellPYI= row.getCell(cellIterator);
                if(cellPYI==null){
                    cellPYI = row.createCell(cellIterator);
                    cellPYI.setCellValue(0.00);
                }
                valuesPYI[j]=cellPYI.getNumericCellValue();
            }


            double[] valuesRAFY= new double[years];
            //GET NFY
            for(int j = 0, cellIterator=1;j<years;j++,cellIterator++){
                row = sheetIncExp.getRow(44);
                XSSFCell cellRAFY= row.getCell(cellIterator);
                if(cellRAFY==null){
                    cellRAFY = row.createCell(cellIterator);
                    cellRAFY.setCellValue(0.00);
                }
                valuesRAFY[j]=cellRAFY.getNumericCellValue();
            }

            //write data to table
            for(int y=0,rowIterator=3;y<years; y++,rowIterator++){
                row=sheetTemplate.getRow(rowIterator);

                XSSFCell cellGFY = row.createCell(1);
                        cellGFY.setCellValue(valuesGFY[y]);

                XSSFCell cellAFY = row.createCell(2);
                cellAFY.setCellValue(valuesAFY[y]);

                XSSFCell cellNFY = row.createCell(3);
                cellNFY.setCellValue(valuesNFY[y]);

                XSSFCell cellPYI = row.createCell(4);
                cellPYI.setCellValue(valuesPYI[y]);


                XSSFCell cellRAFY = row.createCell(5);
                cellRAFY.setCellValue(valuesRAFY[y]);


                XSSFCell cellCIR= row.createCell(6);
                cellCIR.setCellValue(100* valuesCIR[y]);
            }

//write average row
            DecimalFormat DF = new DecimalFormat("#.##");
            double sumGFY=0;
            double sumAFY=0;
            double sumNFY=0;
            double sumPYI=0;
            double sumCIR=0;
            double sumRAFY=0;

            for(int y=0,rowIterator=3;y<years; y++,rowIterator++) {
                sumGFY+=valuesGFY[y];
                sumAFY+=valuesAFY[y];
                sumNFY+=valuesNFY[y];
                sumPYI+=valuesPYI[y];
                sumCIR+=100*valuesCIR[y];
                sumRAFY+=valuesRAFY[y];
            }

            double avgGFY = Double.parseDouble(DF.format(sumGFY/years));
            double avgAFY = Double.parseDouble(DF.format(sumAFY/years));
            double avgNFY = Double.parseDouble(DF.format(sumNFY/years));
            double avgPYI = Double.parseDouble(DF.format(sumPYI/years));
            double avgCIR = Double.parseDouble(DF.format(sumCIR/years));
            double avgRAFY = Double.parseDouble(DF.format(sumRAFY/years));

            row=sheetTemplate.getRow(years+4);


           XSSFCell cellGFY= row.createCell(1);
            cellGFY.setCellValue(avgGFY);


            XSSFCell cellAFY= row.createCell(2);
            cellAFY.setCellValue(avgAFY);

            XSSFCell cellNFY= row.createCell(3);
            cellNFY.setCellValue(avgNFY);

            XSSFCell cellPYI= row.createCell(4);
            cellPYI.setCellValue(avgPYI);

            XSSFCell cellRAFY= row.createCell(5);
            cellRAFY.setCellValue(avgRAFY);

            XSSFCell cellCIR= row.createCell(6);
            cellCIR.setCellValue(avgCIR);



            for (int x = 0; x <6; x++) {
                //   sheet.autoSizeColumn(x);
                sheetTemplate.autoSizeColumn(x);
            }

            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File(workingDir + "\\Table_Analysis_of_Fund_Yield.xlsx"));
            workbookTemplate.write(out);
            out.close();
            workbookTemplate.close();
            System.out.println("Table_Analysis_of_Fund_Yield.xlsx written successfully");
        } catch (NoSuchFileException e1) {
            JOptionPane.showMessageDialog(null, "Please ensure the Plan Requirements are set, then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    public void Create_Table_Gains_and_Losses(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException{
        try{
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

            FileInputStream fileR = new FileInputStream(workingDir + "\\Template_Gains_Losses.xlsx");
            XSSFWorkbook workbookR = new XSSFWorkbook(fileR);
            XSSFSheet sheetTemplate = workbookR.getSheetAt(0);

            FileInputStream fileT = new FileInputStream(workingDir + "\\Completed_Terminee_Sheet.xlsx");
            XSSFWorkbook workbookTerminee = new XSSFWorkbook(fileT);
            XSSFSheet sheetTerminee = workbookTerminee.getSheetAt(0);

        XSSFWorkbook workbookCreate = new XSSFWorkbook();
        XSSFSheet sheet = workbookCreate.createSheet("Gains and Losses");

        //get initial data

            //get initial surplus
            XSSFRow row = sheetTemplate.getRow(3);
            XSSFCell cellSurplus = row.getCell(2);
            if(cellSurplus==null){
                cellSurplus = row.createCell(2);
                cellSurplus.setCellValue(0.00);
            }
            double surplusInitial = cellSurplus.getNumericCellValue();

            //get interest credited to previously
            row = sheetTemplate.getRow(5);
            XSSFCell cellInterestCredited = row.getCell(2);
            if(cellInterestCredited==null){
                cellInterestCredited = row.createCell(2);
                cellInterestCredited.setCellValue(0.00);
            }
            double surplusInterestCredited = cellInterestCredited.getNumericCellValue();


            //get Receivables
            row = sheetTemplate.getRow(7);
            XSSFCell cellReceivables = row.getCell(2);
            if(cellReceivables==null){
                cellReceivables = row.createCell(2);
                cellReceivables.setCellValue(0.00);
            }
            double receivables = cellReceivables.getNumericCellValue();


            //get Miscellaneous Sources
            row = sheetTemplate.getRow(8);
            XSSFCell cellMiscellaneousSources = row.getCell(2);
            if(cellMiscellaneousSources==null){
                cellMiscellaneousSources = row.createCell(2);
                cellMiscellaneousSources.setCellValue(0.00);
            }
            double MiscellaneousSources = cellMiscellaneousSources.getNumericCellValue();

            //get Excess(Shortfall) of Net Investment Income over Interested Credited
     double excessShortfall=  new ExcelReader().getExcessShortFall(PensionPlanStartDate,PensionPlanEndDate,workingDir);


     //get  Non-Vested Employer's Balances
            int cellNumbers=18+ (years*9)+4+9; //fees
         //   int cellNumbers=18+ (years*8)+4+9; //no fees
            int lastcol = sheetTerminee.getLastRowNum();
            lastcol-=3;

       //     System.out.println("lastcol"+lastcol);
         //   System.out.println("cellnumber"+cellNumbers);

            row = sheetTerminee.getRow(lastcol);
            XSSFCell cellNonVestedEmployer = row.getCell(cellNumbers);
            if(cellNonVestedEmployer==null){
                cellNonVestedEmployer = row.createCell(cellNumbers);
                cellNonVestedEmployer.setCellValue(0.00);
            }
            double NonVestedEmployer = cellNonVestedEmployer.getNumericCellValue();

            row = sheetTemplate.getRow(3);
            row.createCell(2).setCellValue(surplusInitial);

            row = sheetTemplate.getRow(4);
            row.createCell(2).setCellValue(excessShortfall);

            row = sheetTemplate.getRow(5);
            row.createCell(2).setCellValue(surplusInterestCredited);


            row = sheetTemplate.getRow(6);
            row.createCell(2).setCellValue(NonVestedEmployer);


            row = sheetTemplate.getRow(7);
            row.createCell(2).setCellValue(receivables);

            row = sheetTemplate.getRow(8);
            row.createCell(2).setCellValue(MiscellaneousSources);


            double result = surplusInitial - excessShortfall - surplusInterestCredited + NonVestedEmployer  + receivables + MiscellaneousSources;
            row = sheetTemplate.getRow(10);
            row.createCell(2).setCellValue(result);


            System.out.println("excessShortfall"+excessShortfall);
            System.out.println("NonVestedEmployer"+NonVestedEmployer);
            System.out.println("surplusInterestCredited"+surplusInterestCredited);
            System.out.println("MiscellaneousSources"+MiscellaneousSources);
            System.out.println("result"+result);
            System.out.println("receivables"+receivables);

            for (int x = 0; x <4; x++) {
            //   sheet.autoSizeColumn(x);
            sheet.autoSizeColumn(x);
        }
        //Write the workbook in file system
        FileOutputStream out = new FileOutputStream(new File(workingDir + "\\Table_Gains_Losses.xlsx"));
            workbookR.write(out);
        out.close();
        workbookR.close();
        System.out.println("Table_Gains_Losses.xlsx written successfully");
    } catch (NoSuchFileException e1) {
        JOptionPane.showMessageDialog(null, "Please ensure the Plan Requirements are set, then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
    } catch (Exception e) {
        e.printStackTrace();
    }
    }

}

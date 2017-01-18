import javafx.scene.control.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.*;
import java.nio.file.NoSuchFileException;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
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

        FileInputStream fileR = new FileInputStream(workingDir + "\\Templates\\Template_Summary_of_Active_Membership.xlsx");
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

        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Tables\\Table_Summary_of_Active_Membership.xlsx"));
        workbookTemplate.write(outFile);
        fileInputStream.close();
        outFile.close();
    }

    public void Create_Table_Movement_in_Active_Membership(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException{

        FileInputStream fileR = new FileInputStream(workingDir + "\\Templates\\Template_Movements_in_Active_Membership.xlsx");
        XSSFWorkbook workbookTemplate = new XSSFWorkbook(fileR);
        XSSFSheet sheetTemplate = workbookTemplate.getSheetAt(0);


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
      //  try {

            //get access to the DATA of the Active Members
            FileInputStream fileInputStreamWorkBookData = new FileInputStream(workingDir + "\\Input Sheet.xlsx");
            XSSFWorkbook workbookData = new XSSFWorkbook(fileInputStreamWorkBookData);

            XSSFSheet Initialsheet = workbookData.getSheet("Actives at End of Plan Yr " + StartYear);
            int InitialnumOfActives = Utility.getNumberOfMembersInSheet(Initialsheet);

            String[] InitialcellEmployeeID = new String[InitialnumOfActives];
            String[] InitialcellLastName = new String[InitialnumOfActives];
            String[] InitialcellFirstName = new String[InitialnumOfActives];

//get the initial members in first year period
            int countMale = 0;
            int countFemale = 0;
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
                cellGender = cellGender.toLowerCase();

                if (cellGender.equals("m")) countMale++;
                if (cellGender.equals("f")) countFemale++;

            }
//write the active members at start of year
            XSSFRow writeRow=sheetTemplate.getRow(3);

            XSSFCell writeCell=writeRow.createCell(1);
            writeCell.setCellValue(countMale);

          writeCell=writeRow.createCell(2);
        writeCell.setCellValue(countFemale);

        writeCell=writeRow.createCell(3);
        writeCell.setCellValue(countFemale+countMale);

        //not its time to get new entrants over the years
        boolean foundIt = false;
        //LOOP THROUGH THE OTHER YEARS
        int InitialMembers = InitialnumOfActives;
        int yearsRemaining = years-1;

        int entrantCountMale=0;
        int entrantCountfemale=0;
        //GET OTHER ACTIVE MEMBERS FROM OTHER YEARS
        //   else{
        for (int x = 0; x < yearsRemaining; x++) {

            int newStartYear = StartYear + 1;

            Calendar cal = Calendar.getInstance();
            cal.set((newStartYear + x), StartMonth, StartDay);
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy"); // Just the year, with 2 digits
            String formattedDate = sdf.format(cal.getTime());
            //  String formattedDate = "2014";
            String Recon = ("Actives at End of Plan Yr " + formattedDate);

//GET RECON SHEET
            XSSFSheet Reconsheet = workbookData.getSheet(Recon);

            int currentNumber = Utility.getNumberOfMembersInSheet(Reconsheet);//get curent members from recon row
            String[] cellEmployeeID = new String[currentNumber];
            String[] cellLastName = new String[currentNumber];
            String[] cellFirstName = new String[currentNumber];

            for (int row = 0, readFromRecon = 7; row < currentNumber; row++, readFromRecon++) {

                XSSFRow rowPosition = Reconsheet.getRow(readFromRecon);

                //get employee id
                XSSFCell cellA = rowPosition.getCell(0);  //employee number
                if (cellA == null) {
                    cellA = rowPosition.createCell((short) 0);
                    cellA.setCellValue("");
                }
                String result = cellA.getStringCellValue();
                cellEmployeeID[row] = result.replaceAll("[-]", "");

                //get LAST NAME
                XSSFCell cellB = rowPosition.getCell((short) 1);  //last name
                if (cellB == null) {
                    cellB = rowPosition.createCell((short) 1);
                    cellB.setCellValue("");
                }
                cellLastName[row] = cellB.getStringCellValue();

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
                String CellGender = cellD.getStringCellValue();
                CellGender = CellGender.toLowerCase();
                boolean findIt = false;

                //test the value from recon sheet against initial
                for (int initial = 0; initial < InitialMembers; initial++) {

                  //  System.out.print("initial " + initial);

                    if (cellEmployeeID[row].equals(InitialcellEmployeeID[initial]) && cellLastName[row].equals(InitialcellLastName[initial])) {
                        findIt = true;
                        break;
                    }
                }

                if (!findIt) {
if(CellGender.equals("m"))  entrantCountMale++;
if(CellGender.equals("f"))  entrantCountfemale++;
                }

            }
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
        }//end loop years

//write the active members at start of year
       writeRow=sheetTemplate.getRow(4);

 writeCell=writeRow.createCell(1);
        writeCell.setCellValue(entrantCountMale);

        writeCell=writeRow.createCell(2);
        writeCell.setCellValue(entrantCountfemale);

        writeCell=writeRow.createCell(3);
        writeCell.setCellValue(entrantCountMale+entrantCountfemale);


        //loop through eqach member in Termination sheet
        XSSFSheet termineeSheet = workbookData.getSheet("Terminated up to "+EndYear+"."+EndMonth+"."+EndDay);
        int numofTerminees =Utility.getNumberOfTermineeMembersInSheet(termineeSheet);

        int [] countTermineeMale=  new int[3];
        int[] countTermineeFemale=new int[3];

        countTermineeMale[0]=0;
        countTermineeMale[1]=0;
        countTermineeMale[2]=0;

        countTermineeFemale[0]=0;
        countTermineeFemale[1]=0;
        countTermineeFemale[2]=0;

        for(int readFrom = 11, r=0;r<numofTerminees;r++,readFrom++){
            XSSFRow rowPosition = termineeSheet.getRow(readFrom);

            //get employee id
            XSSFCell cellA = rowPosition.getCell(0);  //employee number
            if (cellA == null) {
                cellA = rowPosition.createCell((short) 0);
                cellA.setCellValue("");
            }
            String result = cellA.getStringCellValue();
          String  cellEmployeeID= result.replaceAll("[-]", "");

            //get LAST NAME
            XSSFCell cellB = rowPosition.getCell((short) 1);  //last name
            if (cellB == null) {
                cellB = rowPosition.createCell((short) 1);
                cellB.setCellValue("");
            }
            String  cellLastName= cellB.getStringCellValue();

            //get FIRST NAME
            XSSFCell cellC = rowPosition.getCell((short) 2);  //first name
            if (cellC == null) {
                cellC = rowPosition.createCell((short) 2);
                cellC.setCellValue("");
            }
            String cellFirstName = cellC.getStringCellValue();

            //get GENDER
            XSSFCell cellD = rowPosition.getCell((short) 3);  //gender
            if (cellD == null) {
                cellD = rowPosition.createCell((short) 3);
                cellD.setCellValue("");
            }
            String cellGender = cellD.getStringCellValue();
            cellGender= cellGender.toLowerCase();

            //get GENDER
            XSSFCell cellJ = rowPosition.getCell((short) 9);  //gender
            if (cellJ == null) {
                cellJ = rowPosition.createCell((short) 9);
                cellJ.setCellValue("");
            }
            String cellStatus = cellJ.getStringCellValue();

            if (cellStatus.equals("T")){
                if (cellGender.equals("m")) countTermineeMale[0]++;
                if(cellGender.equals("f")) countTermineeFemale[0]++;
            }

            if (cellStatus.equals("R")){
                if (cellGender.equals("m")) countTermineeMale[1]++;
                if(cellGender.equals("f")) countTermineeFemale[1]++;
            }

            if (cellStatus.equals("D")){
                if (cellGender.equals("m")) countTermineeMale[2]++;
                if(cellGender.equals("f")) countTermineeFemale[2]++;
            }

        }

        writeRow=sheetTemplate.getRow(7);

        writeCell=writeRow.createCell(1);
        writeCell.setCellValue(countTermineeMale[0]);

        writeCell=writeRow.createCell(2);
        writeCell.setCellValue(countTermineeFemale[0]);

        writeCell=writeRow.createCell(3);
        writeCell.setCellValue(countTermineeMale[0]+countTermineeFemale[0]);

//RETIREMENT
        writeRow=sheetTemplate.getRow(8);

        writeCell=writeRow.createCell(1);
        writeCell.setCellValue(countTermineeMale[1]);

        writeCell=writeRow.createCell(2);
        writeCell.setCellValue(countTermineeFemale[1]);

        writeCell=writeRow.createCell(3);
        writeCell.setCellValue(countTermineeMale[1]+countTermineeFemale[1]);

        //DEATHS
        writeRow=sheetTemplate.getRow(9);

        writeCell=writeRow.createCell(1);
        writeCell.setCellValue(countTermineeMale[2]);

        writeCell=writeRow.createCell(2);
        writeCell.setCellValue(countTermineeFemale[2]);

        writeCell=writeRow.createCell(3);
        writeCell.setCellValue(countTermineeMale[2]+countTermineeFemale[2]);

        //last row
        writeRow=sheetTemplate.getRow(12);

        writeCell=writeRow.createCell(1);
        writeCell.setCellValue((countMale+entrantCountMale)-(countTermineeMale[0]+countTermineeMale[1]+countTermineeMale[2]));

        writeCell=writeRow.createCell(2);
        writeCell.setCellValue((countMale+entrantCountMale)-(countTermineeMale[0]+countTermineeMale[1]+countTermineeMale[2]));
        writeCell.setCellValue((countFemale+entrantCountfemale) - (countTermineeFemale[0]+countTermineeFemale[1]+countTermineeFemale[2]));

        writeCell=writeRow.createCell(3);
        writeCell.setCellValue((countMale+entrantCountMale)-(countTermineeMale[0]+countTermineeMale[1]+countTermineeMale[2])  +  (countFemale+entrantCountfemale) - (countTermineeFemale[0]+countTermineeFemale[1]+countTermineeFemale[2])  );


        for (int x = 0; x < 4; x++) {
                sheetTemplate.autoSizeColumn(x);
            }

        FileOutputStream outFile = new FileOutputStream(new File(workingDir+"\\Tables\\Table_Movement_in_Active_Membership.xlsx"));
        workbookTemplate.write(outFile);
        fileInputStreamWorkBookData.close();
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

            FileInputStream fileTemplate = new FileInputStream(workingDir + "\\Templates\\Template_Analysis_of_Fund_Yield.xlsx");
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
            FileOutputStream out = new FileOutputStream(new File(workingDir + "\\Tables\\Table_Analysis_of_Fund_Yield.xlsx"));
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

            FileInputStream fileR = new FileInputStream(workingDir + "\\Templates\\Template_Gains_Losses.xlsx");
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
      int cellNumbers=0;
            if (new File(workingDir + "\\f.txt").exists()) {
                cellNumbers=18+ (years*9)+4+9; //fees
            }else if(!new File(workingDir + "\\f.txt").exists()) {
                cellNumbers=18+ (years*8)+4+9; //fees
            }

     //get  Non-Vested Employer's Balances

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


/*            System.out.println("excessShortfall"+excessShortfall);
            System.out.println("NonVestedEmployer"+NonVestedEmployer);
            System.out.println("surplusInterestCredited"+surplusInterestCredited);
            System.out.println("MiscellaneousSources"+MiscellaneousSources);
            System.out.println("result"+result);
            System.out.println("receivables"+receivables);*/

            for (int x = 0; x <4; x++) {
            //   sheet.autoSizeColumn(x);
            sheet.autoSizeColumn(x);
        }
        //Write the workbook in file system
        FileOutputStream out = new FileOutputStream(new File(workingDir + "\\Tables\\Table_Gains_Losses.xlsx"));
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

    public void Create_Balance_Sheet_Table(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException{

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
        double employeeBasicContribution=0.00;
        double employeeOptionalContribution=0.00;
        double employerContribution=0.00;
        double subTotalAccruedBenefits=0.00;
        double refundsOutstandingToUnclaimedMembers=0.00;
        double refundsOutstandingToTerminatedMembers=0.00;
        double accruedBenefitsToRetiredMembers=0.00;
        double accruedBenefitsToDeferredVestedPensioners;
        double totalActuarialLiability=0.00;
        double marketValueOfNetAssets=0.00;
        double actuarialSurplusDefecit=0.00;
        double solvencyLevel=0.00;

        try{
            FileInputStream fileR = new FileInputStream(workingDir + "\\Templates\\Template_Balance_Sheet.xlsx");
            XSSFWorkbook workbookR = new XSSFWorkbook(fileR);
            XSSFSheet sheetVal_Bal = workbookR.getSheetAt(0);

            FileInputStream file = new FileInputStream(workingDir + "\\Income_Expenditure_Table.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0);

//GET DATA FROM TEMPLATE

            XSSFRow Row= sheetVal_Bal.getRow(5);
            XSSFCell cellemployeeBasicContribution = Row.getCell(3);
            if(cellemployeeBasicContribution==null){
                cellemployeeBasicContribution=Row.createCell(3);
                cellemployeeBasicContribution.setCellValue(0.00);
            }
            employeeBasicContribution = cellemployeeBasicContribution.getNumericCellValue();

            Row= sheetVal_Bal.getRow(6);
            XSSFCell cellemployeeOptionalContribution = Row.getCell(3);
            if(cellemployeeOptionalContribution==null){
                cellemployeeOptionalContribution=Row.createCell(3);
                cellemployeeOptionalContribution.setCellValue(0.00);
            }
            employeeOptionalContribution = cellemployeeOptionalContribution.getNumericCellValue();


            Row= sheetVal_Bal.getRow(7);
            XSSFCell cellemployerContribution = Row.getCell(3);
            if(cellemployerContribution==null){
                cellemployerContribution=Row.createCell(3);
                cellemployerContribution.setCellValue(0.00);
            }
            employerContribution = cellemployerContribution.getNumericCellValue();


            Row= sheetVal_Bal.getRow(10);
            XSSFCell cellrefundsOutstandingToUnclaimedMembers = Row.getCell(3);
            if(cellrefundsOutstandingToUnclaimedMembers==null){
                cellrefundsOutstandingToUnclaimedMembers=Row.createCell(3);
                cellrefundsOutstandingToUnclaimedMembers.setCellValue(0.00);
            }
            refundsOutstandingToUnclaimedMembers = cellrefundsOutstandingToUnclaimedMembers.getNumericCellValue();


            Row= sheetVal_Bal.getRow(11);
            XSSFCell cellrefundsOutstandingToTerminatedMembers= Row.getCell(3);
            if(cellrefundsOutstandingToTerminatedMembers==null){
                cellrefundsOutstandingToTerminatedMembers=Row.createCell(3);
                cellrefundsOutstandingToTerminatedMembers.setCellValue(0.00);
            }
            refundsOutstandingToTerminatedMembers = cellrefundsOutstandingToTerminatedMembers.getNumericCellValue();


            Row= sheetVal_Bal.getRow(12);
            XSSFCell cellaccruedBenefitsToRetiredMembers= Row.getCell(3);
            if(cellaccruedBenefitsToRetiredMembers==null){
                cellaccruedBenefitsToRetiredMembers=Row.createCell(3);
                cellaccruedBenefitsToRetiredMembers.setCellValue(0.00);
            }
            accruedBenefitsToRetiredMembers = cellaccruedBenefitsToRetiredMembers.getNumericCellValue();


            Row= sheetVal_Bal.getRow(13);
            XSSFCell cellaccruedBenefitsToDeferredVestedPensioners= Row.getCell(3);
            if(cellaccruedBenefitsToDeferredVestedPensioners==null){
                cellaccruedBenefitsToDeferredVestedPensioners=Row.createCell(3);
                cellaccruedBenefitsToDeferredVestedPensioners.setCellValue(0.00);
            }
            accruedBenefitsToDeferredVestedPensioners = cellaccruedBenefitsToDeferredVestedPensioners.getNumericCellValue();

            Row= sheet.getRow(32);
            XSSFCell cellNetIncome= Row.getCell(years+1);
            if(cellNetIncome==null){
                cellNetIncome=Row.createCell(3);
                cellNetIncome.setCellValue(0.00);
            }
            marketValueOfNetAssets = cellNetIncome.getNumericCellValue();

            //CALCULATIONS
            subTotalAccruedBenefits = employeeBasicContribution+employeeOptionalContribution+employerContribution;
            totalActuarialLiability=subTotalAccruedBenefits+refundsOutstandingToUnclaimedMembers+refundsOutstandingToTerminatedMembers+accruedBenefitsToRetiredMembers+accruedBenefitsToDeferredVestedPensioners;
            actuarialSurplusDefecit=marketValueOfNetAssets-totalActuarialLiability;
            if(totalActuarialLiability==0)  totalActuarialLiability=1;
            solvencyLevel=(marketValueOfNetAssets/totalActuarialLiability)*100;


//WRITE CALCULATIONS TO WORKBOOK
            XSSFRow row = sheetVal_Bal.getRow(8);
            row.createCell(3).setCellValue(Double.parseDouble(dF.format(subTotalAccruedBenefits)));

            row = sheetVal_Bal.getRow(15);
            row.createCell(3).setCellValue(Double.parseDouble(dF.format(totalActuarialLiability)));

            row = sheetVal_Bal.getRow(17);
            row.createCell(3).setCellValue(Double.parseDouble(dF.format(marketValueOfNetAssets)));

            row = sheetVal_Bal.getRow(19);
            row.createCell(3).setCellValue(Double.parseDouble(dF.format(actuarialSurplusDefecit)));

            row = sheetVal_Bal.getRow(21);
            row.createCell(3).setCellValue(Double.parseDouble(dF.format(solvencyLevel)));



            FileOutputStream outFile = new FileOutputStream(new File(workingDir +"\\Tables\\Balance_Sheet_Table.xlsx"));
            workbookR.write(outFile);
            fileR.close();
            file.close();
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

    public double[] getInflationRates(String workingDir, int numOfYears) throws IOException {

        FileInputStream fis = null;
        try {
            fis = new FileInputStream(workingDir + "\\Inflation Rates.xlsx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet =workbook.getSheet("rates");
        XSSFCell [] cell = new XSSFCell[numOfYears];
        double [] values = new double[numOfYears];

        for(int row=2,I=0;I<numOfYears;row++,I++) {
            XSSFRow interestRow = sheet.getRow(row);

            cell[I] = interestRow.getCell(1);
            //    list.add(cell[I].getNumericCellValue(),I);
            values[I]=cell[I].getNumericCellValue();

        }

        return values;
    }

    public void Create_Income_Expenditure_Table(String PensionPlanStartDate, String PensionPlanEndDate, String workingDir) throws IOException {
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


        try{
            FileInputStream fileR = new FileInputStream(workingDir + "\\Templates\\Template_Inc_Exp_Sheet.xlsx");
            XSSFWorkbook workbookR = new XSSFWorkbook(fileR);
            XSSFSheet sheetInc_Exp = workbookR.getSheetAt(0);

//int counter=6;
            double consolidatedEmployeeContribution=0;
            double consolidatedEmployersContributions=0;
            double consolidatedInterest=0;
            double consolidatedTotalIncome=0;
            double consolidatedEmployeeRequiredExpenditure=0;
            double consolidatedEmployeeOptionalExpenditure=0;
            double consolidatedEmployerRequiredExpenditure=0;
            double consolidatedAdministrativeFees=0;
            double consolidatedTotalExpenditure=0;
            double consolidatedNetIncome=0;
            double consolidatedFundAtEndofPeriod=0;
            double consolidatedFundAtBeginningofPeriod =0;
            double consolidatedAdministrativeOtherExpenses=0;
            double consolidatedInvestmentExpenses=0;
            double consolidatedTotalExpenses=0;
            double consolidatedInvestmentIncome=0;

            double consolidatedGrossFundYield=1;
            double consolidatedAdjustedFundYield=1;
            double consolidatedNetFundYield=1;
            double consolidatedPlanYearInflation=1;
            double consolidatedRealAdjustedFundYield=1;

            double employeeBasic_Optional[] = new double[years];
            double EmployerRequired[] = new double[years];
            double empInterest[] = new double[years];
            double expenditureEmployeeBasic[] = new double[years];
            double expenditureEmployeeOptional[]=new double[years];
            double expenditureEmployerRequired[]= new double[years];
            double administrativeFees[]= new double[years];

            ArrayList listEBO = new ArrayList();
            ArrayList listER =new ArrayList();
            ArrayList listInterest =new ArrayList();
            //     ArrayList list
            double[]totalIncomeSum = new double[years];
            double[]totalExpenditure = new double[years];
            double[] netIncome= new double[years];
            double[] fundAtEndofPeriod = new double[years];
            double[] fundAtBeginning =new double[years];
            double[] priorYearAdjustment = new double[years];
            double[] netRealizedGainLoss = new double[years];
            double[] netUnrealizedGainLoss = new double[years];
            double[] purchasOfImmediatePensions = new double[years];
            double[] purchasOfDeferredPensions = new double[years];
            double[] lumpSumtoRetirees = new double[years];
            double[] monthlyPensionsPaidtoPesioners = new double[years];
            double[] amountsPurchaseImmediateDeferredAnnuities = new double[years];
            double[] transfers = new double[years];
            double[] investmentManagementFees = new double[years];
            double[] feesForProfessionalServices = new double[years];
            double[] otherExpenses = new double[years];

            double[] administrativeAndOtherExpenses = new double[years];
            double[] investmentExpenses = new double[years];
            double[] totalExpenses = new double[years];
            double[] totalInvestmentIncome = new double[years];

            double[] grossFundYield = new double[years];
            double[] adjuestedFundYield = new double[years];
            double[] netFundYield = new double[years];

            double[] planYearInflation = new double[years];
            double[] realAdjustedFundYield = new double[years];

            //get the interest rates values
            double[] inflationRates = new double [years];
            inflationRates= getInflationRates(workingDir,years);


//GET THE FUND AT BEGINNING OF PERIOD IF ANY
            XSSFRow Row= sheetInc_Exp.getRow(6);
            XSSFCell cellFundAtBeginning = Row.getCell(1);
            if(cellFundAtBeginning==null){
                cellFundAtBeginning=Row.createCell(1);
                cellFundAtBeginning.setCellValue(0.00);
            }
            fundAtBeginning[0] = cellFundAtBeginning.getNumericCellValue();


//GET DATA FROM TEMPLATE
            for (int x = 0; x <years; x++) {

                //get the prior year adjustment if any
                Row= sheetInc_Exp.getRow(7);
                XSSFCell cellPriorYearAdjustment = Row.getCell(1+x);
                if(cellPriorYearAdjustment==null){
                    cellPriorYearAdjustment=Row.createCell(1+x);
                    cellPriorYearAdjustment.setCellValue(0.00);
                }
                priorYearAdjustment[x] = cellPriorYearAdjustment.getNumericCellValue();
                //get the prior year adjustment if any


                //get  Employees' Contributions (Basic and Optional)
                Row= sheetInc_Exp.getRow(9);
                XSSFCell cellemployeeBasic_Optional = Row.getCell(1+x);
                if(cellemployeeBasic_Optional==null){
                    cellemployeeBasic_Optional=Row.createCell(1+x);
                    cellemployeeBasic_Optional.setCellValue(0.00);
                }
                employeeBasic_Optional[x] = cellemployeeBasic_Optional.getNumericCellValue();


                //get the employers contribution
                Row= sheetInc_Exp.getRow(10);
                XSSFCell cellEmployerRequired = Row.getCell(1+x);
                if(cellEmployerRequired==null){
                    cellEmployerRequired=Row.createCell(1+x);
                    cellEmployerRequired.setCellValue(0.00);
                }
                EmployerRequired[x] = cellEmployerRequired.getNumericCellValue();


                //get the interest/dividend
                Row= sheetInc_Exp.getRow(11);
                XSSFCell cellInterestDividend = Row.getCell(1+x);
                if(cellInterestDividend==null){
                    cellInterestDividend=Row.createCell(1+x);
                    cellInterestDividend.setCellValue(0.00);
                }
                empInterest[x] = cellInterestDividend.getNumericCellValue();


                //get the net realized gain if any
                Row= sheetInc_Exp.getRow(12);
                XSSFCell cellnetRealizedGainLoss = Row.getCell(1+x);
                if(cellnetRealizedGainLoss==null){
                    cellnetRealizedGainLoss=Row.createCell(1+x);
                    cellnetRealizedGainLoss.setCellValue(0.00);
                }
                netRealizedGainLoss[x] = cellnetRealizedGainLoss.getNumericCellValue();


                //get the net unrealized gain if any
                Row= sheetInc_Exp.getRow(13);
                XSSFCell cellnetUnrealizedGainLoss= Row.getCell(1+x);
                if(cellnetUnrealizedGainLoss==null){
                    cellnetUnrealizedGainLoss=Row.createCell(1+x);
                    cellnetUnrealizedGainLoss.setCellValue(0.00);
                }
                netUnrealizedGainLoss[x] = cellnetUnrealizedGainLoss.getNumericCellValue();


                //get Expenditure Employee Basic
                Row= sheetInc_Exp.getRow(18);
                XSSFCell cellexpenditureEmployeeBasic= Row.getCell(1+x);
                if(cellexpenditureEmployeeBasic==null){
                    cellexpenditureEmployeeBasic=Row.createCell(1+x);
                    cellexpenditureEmployeeBasic.setCellValue(0.00);
                }
                expenditureEmployeeBasic[x] = cellexpenditureEmployeeBasic.getNumericCellValue();

                //get Expenditure Employee Optional
                Row= sheetInc_Exp.getRow(19);
                XSSFCell cellexpenditureEmployeeOptional= Row.getCell(1+x);
                if(cellexpenditureEmployeeOptional==null){
                    cellexpenditureEmployeeOptional=Row.createCell(1+x);
                    cellexpenditureEmployeeOptional.setCellValue(0.00);
                }
                expenditureEmployeeOptional[x] = cellexpenditureEmployeeOptional.getNumericCellValue();


                //get Expenditure Employer Required
                Row= sheetInc_Exp.getRow(20);
                XSSFCell cellexpenditureEmployerRequired= Row.getCell(1+x);
                if(cellexpenditureEmployerRequired==null){
                    cellexpenditureEmployerRequired=Row.createCell(1+x);
                    cellexpenditureEmployerRequired.setCellValue(0.00);
                }
                expenditureEmployerRequired[x] = cellexpenditureEmployerRequired.getNumericCellValue();


                //get the purchase of immediate pensionsi f any
                Row= sheetInc_Exp.getRow(21);
                XSSFCell cellpurchasOfImmediatePensions= Row.getCell(1+x);
                if(cellpurchasOfImmediatePensions==null){
                    cellpurchasOfImmediatePensions=Row.createCell(1+x);
                    cellpurchasOfImmediatePensions.setCellValue(0.00);
                }
                purchasOfImmediatePensions[x] = cellpurchasOfImmediatePensions.getNumericCellValue();


                //get the purchase of Deferred pensionsi f any
                Row= sheetInc_Exp.getRow(22);
                XSSFCell cellpurchasOfDeferredPensions= Row.getCell(1+x);
                if(cellpurchasOfDeferredPensions==null){
                    cellpurchasOfDeferredPensions=Row.createCell(1+x);
                    cellpurchasOfDeferredPensions.setCellValue(0.00);
                }
                purchasOfDeferredPensions[x] = cellpurchasOfDeferredPensions.getNumericCellValue();


                //get the lump sum to retirees if any
                Row= sheetInc_Exp.getRow(23);
                XSSFCell celllumpSumtoRetirees= Row.getCell(1+x);
                if(celllumpSumtoRetirees==null){
                    celllumpSumtoRetirees=Row.createCell(1+x);
                    celllumpSumtoRetirees.setCellValue(0.00);
                }
                lumpSumtoRetirees[x] = celllumpSumtoRetirees.getNumericCellValue();


                Row= sheetInc_Exp.getRow(24);
                XSSFCell cellmonthlyPensionsPaidtoPesioners= Row.getCell(1+x);
                if(cellmonthlyPensionsPaidtoPesioners==null){
                    cellmonthlyPensionsPaidtoPesioners=Row.createCell(1+x);
                    cellmonthlyPensionsPaidtoPesioners.setCellValue(0.00);
                }
                monthlyPensionsPaidtoPesioners[x] = cellmonthlyPensionsPaidtoPesioners.getNumericCellValue();


                Row= sheetInc_Exp.getRow(25);
                XSSFCell cellamountsPurchaseImmediateDeferredAnnuities= Row.getCell(1+x);
                if(cellamountsPurchaseImmediateDeferredAnnuities==null){
                    cellamountsPurchaseImmediateDeferredAnnuities=Row.createCell(1+x);
                    cellamountsPurchaseImmediateDeferredAnnuities.setCellValue(0.00);
                }
                amountsPurchaseImmediateDeferredAnnuities[x] = cellamountsPurchaseImmediateDeferredAnnuities.getNumericCellValue();


                Row= sheetInc_Exp.getRow(26);
                XSSFCell celltransfers= Row.getCell(1+x);
                if(celltransfers==null){
                    celltransfers=Row.createCell(1+x);
                    celltransfers.setCellValue(0.00);
                }
                transfers[x] = celltransfers.getNumericCellValue();

                //get Administrative Fees
                Row= sheetInc_Exp.getRow(27);
                XSSFCell celladministrativeFees= Row.getCell(1+x);
                if(celladministrativeFees==null){
                    celladministrativeFees=Row.createCell(1+x);
                    celladministrativeFees.setCellValue(0.00);
                }
                administrativeFees[x] = celladministrativeFees.getNumericCellValue();



                Row= sheetInc_Exp.getRow(28);
                XSSFCell cellinvestmentManagementFees = Row.getCell(1+x);
                if(cellinvestmentManagementFees==null){
                    cellinvestmentManagementFees=Row.createCell(1+x);
                    cellinvestmentManagementFees.setCellValue(0.00);
                }
                investmentManagementFees[x] = cellinvestmentManagementFees.getNumericCellValue();


                Row= sheetInc_Exp.getRow(29);
                XSSFCell cellfeesForProfessionalServices = Row.getCell(1+x);
                if(cellfeesForProfessionalServices==null){
                    cellfeesForProfessionalServices=Row.createCell(1+x);
                    cellfeesForProfessionalServices.setCellValue(0.00);
                }
                feesForProfessionalServices[x] = cellfeesForProfessionalServices.getNumericCellValue();


                Row= sheetInc_Exp.getRow(30);
                XSSFCell cellotherExpenses = Row.getCell(1+x);
                if(cellotherExpenses==null){
                    cellotherExpenses=Row.createCell(1+x);
                    cellotherExpenses.setCellValue(0.00);
                }
                otherExpenses[x] = cellotherExpenses.getNumericCellValue();


            }//end of loop to get INTITIAL DATA

//MAIN LOOP
            for (int x = 0; x <years; x++) {

                administrativeAndOtherExpenses[x]=administrativeFees[x]+feesForProfessionalServices[x]+otherExpenses[x];
                investmentExpenses[x]=investmentManagementFees[x];
                totalExpenses[x]=administrativeAndOtherExpenses[x]+investmentExpenses[x];
                totalInvestmentIncome[x]=empInterest[x]+netRealizedGainLoss[x]+netUnrealizedGainLoss[x];

                listEBO.add(employeeBasic_Optional[x]);
                listER.add(EmployerRequired[x]);
                listInterest.add(empInterest[x]);



                for (int b = 0; b < listEBO.size(); b++) {
                    XSSFRow row = sheetInc_Exp.getRow(9);
                    row.createCell(1+x).setCellValue((Double)listEBO.get(b));
                }


                //write the Employers' Required Contributions
                for (int b = 0; b < listER.size(); b++) {
                    XSSFRow row = sheetInc_Exp.getRow(10);
                    row.createCell(1+x).setCellValue((Double)listER.get(b));
                }

                //WRITE INTEREST SUMS
                for (int b = 0; b < listInterest.size(); b++) {
                    XSSFRow row = sheetInc_Exp.getRow(11);
                    row.createCell(1+x).setCellValue((Double)listInterest.get(b));
                }


                //WRITE TOTAL INCOME and Total Expenditure
                for (int b = 0; b <years; b++) {
                    XSSFRow row = sheetInc_Exp.getRow(14);
                    totalIncomeSum[x]=employeeBasic_Optional[x]+EmployerRequired[x]+empInterest[x]+netRealizedGainLoss[x]+netUnrealizedGainLoss[x];
                    row.createCell(1+x).setCellValue((Double)totalIncomeSum[x]);

                    //Total Expenditure
                    row = sheetInc_Exp.getRow(31);
                    totalExpenditure[x]=expenditureEmployeeBasic[x]+expenditureEmployeeOptional[x]+expenditureEmployerRequired[x]+purchasOfImmediatePensions[x]+purchasOfDeferredPensions[x]+lumpSumtoRetirees[x]+monthlyPensionsPaidtoPesioners[x]+amountsPurchaseImmediateDeferredAnnuities[x]+transfers[x]+investmentManagementFees[x]+feesForProfessionalServices[x]+otherExpenses[x]+administrativeFees[x];
                    row.createCell(1+x).setCellValue((Double)totalExpenditure[x]);

                    //Total Expenditure
                    row = sheetInc_Exp.getRow(32);
                    netIncome[x]= totalIncomeSum[x]-totalExpenditure[x];
                    row.createCell(1+x).setCellValue((Double)netIncome[x]);

                    //Total Expenditure
                    row = sheetInc_Exp.getRow(33);
                    fundAtEndofPeriod[x]= priorYearAdjustment[x]+fundAtBeginning[x]+netIncome[x];
                    row.createCell(1+x).setCellValue((Double)fundAtEndofPeriod[x]);
                }

                System.out.print("totalInvestmentIncome[x]"+totalInvestmentIncome[x]);
                System.out.print("fundAtBeginning[x]"+fundAtBeginning[x]);
                System.out.print("fundAtEndofPeriod[x]"+fundAtEndofPeriod[x]);
                grossFundYield[x]= ((2*totalInvestmentIncome[x]) / (fundAtBeginning[x]+fundAtEndofPeriod[x]-totalInvestmentIncome[x]))* 100;
                adjuestedFundYield[x]= ((2* ((totalInvestmentIncome[x]-investmentExpenses[x]))) / ((fundAtBeginning[x]+fundAtEndofPeriod[x]-totalInvestmentIncome[x]+investmentExpenses[x]))) * 100;
                netFundYield[x]= ((2* (totalInvestmentIncome[x]-totalExpenses[x])) / ((fundAtBeginning[x]+fundAtEndofPeriod[x]-totalInvestmentIncome[x]-totalExpenses[x]))) *100;

                planYearInflation[x]=inflationRates[x]*100;
                realAdjustedFundYield[x]=adjuestedFundYield[x]-planYearInflation[x];

                //put fund at end of period to beginning of next year period
                for(int h=1,l=0;h<years;h++,l++){
                    fundAtBeginning[h]=fundAtEndofPeriod[l];
                }

//WRITE EXPENDITURE ROWS
                for(int y=0;y<years;y++){
                    XSSFRow row = sheetInc_Exp.getRow(18);
                    row.createCell(1+x).setCellValue((Double)expenditureEmployeeBasic[x]);

                    row = sheetInc_Exp.getRow(19);
                    row.createCell(1+x).setCellValue((Double)expenditureEmployeeOptional[x]);

                    row = sheetInc_Exp.getRow(20);
                    row.createCell(1+x).setCellValue((Double)expenditureEmployerRequired[x]);

                    row = sheetInc_Exp.getRow(27);
                    row.createCell(1+x).setCellValue((Double)administrativeFees[x]);

                    row = sheetInc_Exp.getRow(34);
                    row.createCell(1+x).setCellValue((Double)administrativeAndOtherExpenses[x]);

                    row = sheetInc_Exp.getRow(35);
                    row.createCell(1+x).setCellValue((Double)investmentExpenses[x]);

                    row = sheetInc_Exp.getRow(36);
                    row.createCell(1+x).setCellValue(Double.parseDouble(dF.format(totalExpenses[x])));

                    row = sheetInc_Exp.getRow(37);
                    row.createCell(1+x).setCellValue(Double.parseDouble(dF.format(totalInvestmentIncome[x])));

                    //FUND YIELD
                    System.out.println("grossFundYield[x]"+grossFundYield[x]);
                    row = sheetInc_Exp.getRow(39);
                    row.createCell(1+x).setCellValue(Double.parseDouble(dF.format(grossFundYield[x])));

                    row = sheetInc_Exp.getRow(40);
                    row.createCell(1+x).setCellValue(Double.parseDouble(dF.format(adjuestedFundYield[x])));

                    row = sheetInc_Exp.getRow(41);
                    row.createCell(1+x).setCellValue(Double.parseDouble(dF.format(netFundYield[x])));

                    row = sheetInc_Exp.getRow(43);
                    row.createCell(1+x).setCellValue(Double.parseDouble(dF.format(planYearInflation[x])));

                    row = sheetInc_Exp.getRow(44);
                    row.createCell(1+x).setCellValue(Double.parseDouble(dF.format(realAdjustedFundYield[x])));
                }

                //Consolidated Totals
                //   System.out.println(totalIncomeSum[x]);
                consolidatedEmployeeContribution+=employeeBasic_Optional[x];
                consolidatedEmployersContributions+=EmployerRequired[x];
                consolidatedInterest+=empInterest[x];
                consolidatedTotalIncome+=totalIncomeSum[x];//consolidated Total income
                consolidatedEmployeeRequiredExpenditure+=expenditureEmployeeBasic[x];
                consolidatedEmployeeOptionalExpenditure+=expenditureEmployeeOptional[x];
                consolidatedEmployerRequiredExpenditure+=expenditureEmployerRequired[x];
                consolidatedAdministrativeFees+=administrativeFees[x];
                consolidatedTotalExpenditure+=totalExpenditure[x];
                consolidatedNetIncome+=netIncome[x];
                consolidatedFundAtEndofPeriod+=fundAtEndofPeriod[x];
                consolidatedFundAtBeginningofPeriod+=fundAtBeginning[x];
                consolidatedAdministrativeOtherExpenses+=administrativeAndOtherExpenses[x];
                consolidatedInvestmentExpenses+=investmentExpenses[x];
                consolidatedTotalExpenses+=totalExpenses[x];
                consolidatedInvestmentIncome+=totalInvestmentIncome[x];

                consolidatedGrossFundYield*=(1+grossFundYield[x]/100);
                consolidatedAdjustedFundYield*=(1+adjuestedFundYield[x]/100);
                consolidatedNetFundYield*=(1+netFundYield[x]/100);

                consolidatedPlanYearInflation*=(1+planYearInflation[x]/100);
                consolidatedRealAdjustedFundYield*=(1+realAdjustedFundYield[x]/100);

                for (int b = 0; b <years; b++) {
                    //WRITE Consolidated Totals
                    if(b==years-1){
                        XSSFRow row = sheetInc_Exp.getRow(9);
                        row.createCell(years+1).setCellValue((Double)consolidatedEmployeeContribution);

                        row = sheetInc_Exp.getRow(10);
                        row.createCell(years+1).setCellValue((Double)consolidatedEmployersContributions);

                        row = sheetInc_Exp.getRow(11);
                        row.createCell(years+1).setCellValue((Double)consolidatedInterest);

                        row = sheetInc_Exp.getRow(14);
                        row.createCell(years+1).setCellValue((Double)consolidatedTotalIncome);


                        row = sheetInc_Exp.getRow(18);
                        row.createCell(years+1).setCellValue((Double)consolidatedEmployeeRequiredExpenditure);

                        row = sheetInc_Exp.getRow(19);
                        row.createCell(years+1).setCellValue((Double)consolidatedEmployeeOptionalExpenditure);

                        row = sheetInc_Exp.getRow(20);
                        row.createCell(years+1).setCellValue((Double)consolidatedEmployerRequiredExpenditure);


                        row = sheetInc_Exp.getRow(27);
                        row.createCell(years+1).setCellValue((Double)consolidatedAdministrativeFees);

                        row = sheetInc_Exp.getRow(31);
                        row.createCell(years+1).setCellValue((Double)consolidatedTotalExpenditure);

                        row = sheetInc_Exp.getRow(32);
                        row.createCell(years+1).setCellValue((Double)consolidatedNetIncome);

                        row = sheetInc_Exp.getRow(33);
                        row.createCell(years+1).setCellValue((Double)consolidatedFundAtEndofPeriod);

                        row = sheetInc_Exp.getRow(6);
                        row.createCell(years+1).setCellValue((Double)consolidatedFundAtBeginningofPeriod);


                        row = sheetInc_Exp.getRow(34);
                        row.createCell(years+1).setCellValue((Double)consolidatedAdministrativeOtherExpenses);


                        row = sheetInc_Exp.getRow(35);
                        row.createCell(years+1).setCellValue((Double)consolidatedInvestmentExpenses);

                        row = sheetInc_Exp.getRow(36);
                        row.createCell(years+1).setCellValue((Double)consolidatedTotalExpenses);

                        row = sheetInc_Exp.getRow(37);
                        row.createCell(years+1).setCellValue((Double)consolidatedInvestmentIncome);

                        row = sheetInc_Exp.getRow(39);

                        double yearFloat = years;
                        double indices = 1/yearFloat;

                        double gfyVal = ((Math.pow(consolidatedGrossFundYield,indices)-1)*100);
                        row.createCell(years+1).setCellValue(Double.parseDouble(dF.format(gfyVal)));

                        row = sheetInc_Exp.getRow(40);
                        double afyVal = ((Math.pow(consolidatedAdjustedFundYield,indices)-1)*100);
                        row.createCell(years+1).setCellValue(Double.parseDouble(dF.format(afyVal)));

                        row = sheetInc_Exp.getRow(41);
                        double nfyVal = ((Math.pow(consolidatedNetFundYield,indices)-1)*100);
                        row.createCell(years+1).setCellValue(Double.parseDouble(dF.format(nfyVal)));

                        row = sheetInc_Exp.getRow(43);
                        double PlanYearInflationVal = ((Math.pow(consolidatedPlanYearInflation,indices)-1)*100);
                        row.createCell(years+1).setCellValue(Double.parseDouble(dF.format(PlanYearInflationVal)));

                        row = sheetInc_Exp.getRow(44);
                        double RealAdjustedFundYieldVal = ((Math.pow(consolidatedRealAdjustedFundYield,indices)-1)*100);
                        row.createCell(years+1).setCellValue(Double.parseDouble(dF.format(RealAdjustedFundYieldVal)));
                    }

                }

                //WRITE FUND AT BEGINING OF PERIOD
                for(int t=0;t<years-1;t++){
                    XSSFRow row = sheetInc_Exp.getRow(6);
                    row.createCell(2+t).setCellValue((Double)fundAtEndofPeriod[t]);
                }




                StartYear++;//incrememnt year at end of loping year
            }//end of looping through each year

            FileOutputStream outFile = new FileOutputStream(new File(workingDir +"\\Tables\\Income_Expenditure_Table.xlsx"));
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

}

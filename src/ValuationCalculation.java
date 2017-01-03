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

/**
 * Created by Ayshahvez on 1/1/2017.
 */
public class ValuationCalculation {

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
    FileInputStream fileR = new FileInputStream(workingDir + "\\Template_Inc_Exp_Sheet.xlsx");
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
        double consolidatedGrossFundYield=0;
        double consolidatedAdjustedFundYield=0;
        double consolidatedNetFundYield=0;
        double consolidatedPlanYearInflation=0;
        double consolidatedRealAdjustedFundYield=0;

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

            Calendar cal = Calendar.getInstance();
            cal.set(StartYear, StartMonth, StartDay);
            SimpleDateFormat sdf = new SimpleDateFormat("yy"); // Just the year, with 2 digits
            String formattedDate = sdf.format(cal.getTime());


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
         //   System.out.println(formattedDate);

            int CountReconRow = Reconsheet.getPhysicalNumberOfRows();

            for (int y = 6; y < CountReconRow; y++) {
                XSSFRow reconRow = Reconsheet.getRow(y);

                XSSFCell cellH = reconRow.getCell((short) 7);
                //Update the value of cell
                if (cellH == null) {
                    cellH = reconRow.createCell(7);
                    cellH.setCellValue(0.00);
                }
                double employeeBasic = cellH.getNumericCellValue();

                XSSFCell cellI = reconRow.getCell((short) 8);
                //Update the value of cell
                if (cellI == null) {
                    cellI = reconRow.createCell(8);
                    cellI.setCellValue(0.00);
                }
                double employeeVoluntary = cellI.getNumericCellValue();

                XSSFCell cellJ = reconRow.getCell((short) 9);
                //Update the value of cell
                if (cellJ == null) {
                    cellJ = reconRow.createCell(9);
                    cellJ.setCellValue(0.00);
                }
                double employerRequired = cellJ.getNumericCellValue();

//GET THE INTEREST
                XSSFCell cellT = reconRow.getCell((short) 19);
                //Update the value of cell
                if (cellT== null) {
                    cellT = reconRow.createCell(19);
                    cellT.setCellValue(0.00);
                }
                double employeeRequiredInterest = cellT.getNumericCellValue();

                XSSFCell cellU = reconRow.getCell((short) 20);
                //Update the value of cell
                if (cellU== null) {
                    cellU = reconRow.createCell(20);
                    cellU.setCellValue(0.00);
                }
                double employeeVoluntaryInterest = cellU.getNumericCellValue();

                XSSFCell cellV = reconRow.getCell((short) 21);
                //Update the value of cell
                if (cellV== null) {
                    cellV = reconRow.createCell(21);
                    cellV.setCellValue(0.00);
                }
                double employerRequiredInterest = cellV.getNumericCellValue();

                //GET THE employee Required WITHDRAWAL
                XSSFCell cellL = reconRow.getCell((short) 11);
                //Update the value of cell
                if (cellL== null) {
                    cellL = reconRow.createCell(11);
                    cellL.setCellValue(0.00);
                }
                double employeeRequiredWithdrawal = cellL.getNumericCellValue();

                //GET THE employee OPTIONAL WITHDRAWAL
                XSSFCell cellM = reconRow.getCell((short) 12);
                //Update the value of cell
                if (cellM== null) {
                    cellM = reconRow.createCell(12);
                    cellM.setCellValue(0.00);
                }
                double employeeOptionalWithdrawal = cellM.getNumericCellValue();

                //GET THE employer Required WITHDRAWAL
                XSSFCell cellN = reconRow.getCell((short) 13);
                //Update the value of cell
                if (cellN== null) {
                    cellN = reconRow.createCell(13);
                    cellN.setCellValue(0.00);
                }
                double employerRequiredWithdrawal = cellN.getNumericCellValue();

                //get fees amount
                //GET THE employer Required WITHDRAWAL
                XSSFCell cellP= reconRow.getCell((short) 15);
                //Update the value of cell
                if (cellP== null) {
                    cellP = reconRow.createCell(15);
                    cellP.setCellValue(0.00);
                }
                double employeeRequiredFee = cellP.getNumericCellValue();

                //GET THE employer Required WITHDRAWAL
                XSSFCell cellQ = reconRow.getCell((short) 16);
                //Update the value of cell
                if (cellQ== null) {
                    cellQ = reconRow.createCell(16);
                    cellQ.setCellValue(0.00);
                }
                double employeeOptionalFee = cellQ.getNumericCellValue();


                //GET THE employer Required WITHDRAWAL
                XSSFCell cellR = reconRow.getCell((short) 17);
                //Update the value of cell
                if (cellR== null) {
                    cellR = reconRow.createCell(17);
                    cellR.setCellValue(0.00);
                }
                double employerFee = cellR.getNumericCellValue();


                employeeBasic_Optional[x]+=(employeeBasic+employeeVoluntary);
                EmployerRequired[x]+=employerRequired;
                empInterest[x]+=(employeeRequiredInterest+employeeVoluntaryInterest+employerRequiredInterest);

                expenditureEmployeeBasic[x]+=employeeRequiredWithdrawal;
                expenditureEmployeeOptional[x]+=employeeOptionalWithdrawal;
                expenditureEmployerRequired[x]+=employerRequiredWithdrawal;

                administrativeFees[x]+=(employeeRequiredFee-employeeOptionalFee-employerFee);

                administrativeAndOtherExpenses[x]=administrativeFees[x]+feesForProfessionalServices[x]+otherExpenses[x];
           investmentExpenses[x]=investmentManagementFees[x];
           totalExpenses[x]=administrativeAndOtherExpenses[x]+investmentExpenses[x];
           totalInvestmentIncome[x]=empInterest[x]+netRealizedGainLoss[x]+netUnrealizedGainLoss[x];

            }//end of looping through recon sheet


            listEBO.add(employeeBasic_Optional[x]);
            listER.add(EmployerRequired[x]);
            listInterest.add(empInterest[x]);


//write the Employees' Contributions (Basic and Optional)
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


            grossFundYield[x]= ((2*totalInvestmentIncome[x]) / (fundAtBeginning[x]+fundAtEndofPeriod[x]-totalInvestmentIncome[x]))* 100;
            adjuestedFundYield[x]= ((2* ((totalInvestmentIncome[x]-investmentExpenses[x]))) / ((fundAtBeginning[x]+fundAtEndofPeriod[x]-totalInvestmentIncome[x]+investmentExpenses[x]))) * 100;
            netFundYield[x]= ((2* (totalInvestmentIncome[x]-investmentExpenses[x])) / ((fundAtBeginning[x]+fundAtEndofPeriod[x]-totalInvestmentIncome[x]-totalExpenses[x]))) *100;
      /*      int g=2004;
            System.out.println("years:"+(g+x));
            System.out.println("totalInvestmentIncome:"+totalInvestmentIncome[x]);
            System.out.println("fundAtBeginning:"+fundAtBeginning[x]);
            System.out.println("fundAtEnd:"+fundAtEndofPeriod[x]);
            System.out.println("totalInvestmentIncome:"+totalInvestmentIncome[x]);
            System.out.println("investmentExpenses:"+investmentExpenses[x]);
            System.out.println(grossFundYield[x]);
            System.out.println(adjuestedFundYield[x]);
            System.out.println(netFundYield[x]);
            System.out.println();
*/

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
            consolidatedGrossFundYield+=grossFundYield[x];
            consolidatedAdjustedFundYield+=adjuestedFundYield[x];
            consolidatedNetFundYield+=netFundYield[x];
            consolidatedPlanYearInflation+=planYearInflation[x];
            consolidatedRealAdjustedFundYield+=realAdjustedFundYield[x];

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
                    row.createCell(years+1).setCellValue((Double)consolidatedGrossFundYield);

                    row = sheetInc_Exp.getRow(40);
                    row.createCell(years+1).setCellValue((Double)consolidatedAdjustedFundYield);

                    row = sheetInc_Exp.getRow(41);
                    row.createCell(years+1).setCellValue((Double)consolidatedNetFundYield);

                    row = sheetInc_Exp.getRow(43);
                    row.createCell(years+1).setCellValue((Double)consolidatedPlanYearInflation);

                    row = sheetInc_Exp.getRow(44);
                    row.createCell(years+1).setCellValue((Double)consolidatedRealAdjustedFundYield);
                }

            }

            //WRITE FUND AT BEGINING OF PERIOD
            for(int t=0;t<years-1;t++){
               XSSFRow row = sheetInc_Exp.getRow(6);
                row.createCell(2+t).setCellValue((Double)fundAtEndofPeriod[t]);
            }




            StartYear++;//incrememnt year at end of loping year
        }//end of looping through each year



    FileOutputStream outFile = new FileOutputStream(new File(workingDir +"\\Income_Expenditure_Table.xlsx"));
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

    public double[] getInflationRates(String workingDir, int numOfYears) throws IOException {
        //    ArrayList list = new ArrayList();
//System.out.print(numOfYears);
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


}

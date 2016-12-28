import java.io.IOException;
import java.util.Date;

public class CreateWorkBook
{
    public static void main(String[] args) {
   //     double[] interestValues = new double[12];
     //   getInterestRates(workingDir,years);
      try {
          new MainWindow();
    //       ExcelReader excelReader = new ExcelReader();
    //      interestValues= excelReader.getInterestRates("C:\\Users\\akonowalchuk\\GFRAM",12);
       } catch (IOException e) {
            e.printStackTrace();
        }

     //   for(int x=0;x<12;x++){
     //     System.out.println(interestValues[x]);
     //   }
//Utility utility= new Utility();
//System.out.println(utility.getAge(new Date("01/01/2004"), new Date("12/31/2015")));

    }
}
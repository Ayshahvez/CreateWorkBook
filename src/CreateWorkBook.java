import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;

//import com.sun.glass.ui.Pen;

public class CreateWorkBook
{
    public static void main(String[] args) throws IOException {
/*     //   String m="JANUARY";
        Date StartDate = new GregorianCalendar(2004,Calendar.JANUARY,01).getTime();
        Date EndDate = new GregorianCalendar(2015,Calendar.DECEMBER,31).getTime();
        String PensionPlanName = "Testing PENSION FUND";*/

        try {
         //   new Utility().read();
           new MainWindow();
    /*        String fn = JOptionPane.showInputDialog("Enter first number: ");
            String sn = JOptionPane.showInputDialog("Enter second number: ");

            int num1 = Integer.parseInt(fn);
            int num2 = Integer.parseInt(sn);
            int sum = num1 + num2;
            JOptionPane.showMessageDialog(null,"The answer is"+ sum,"Title",JOptionPane.PLAIN_MESSAGE);
            new MainWindow();*/
            //create objects
           /* TemplateSheets templateSheets = new TemplateSheets();
            templateSheets.Create_Template_Active_Terminee_Sheet("C:\\Users\\akonowalchuk\\GFRAM");*/
        //    ExcelReader reader = new ExcelReader();

            //create templates
     //   templateSheets.Create_Template_Terminee_Sheet(StartDate,EndDate,PensionPlanName);
    //         templateSheets.Create_Template_Active_Sheet(StartDate,EndDate,PensionPlanName);


//OPERATIONS
         //  reader.Separate_Actives_Terminees();
        //    reader.Create_Actives_Sheet();
       //    reader.Create_Activee_Contribution();
        //   reader.Create_Terminee_Sheet();
       //    reader.Create_Terminee_Contribution();
         //  reader.Results();
       //  ValidationChecks Check = new ValidationChecks();
         // Check.Check_FivePercent_PS();
         //  Check.Check_For_Duplicates();
        //   Check.Check_Plan_EntryDate_empDATE();
         //  Check.Check_BirthDate_EmpDate();
        //   Check.Check_BirthDate();
         //  Check.Check_DateofBirth();
         //  Check.Check_Age();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    }
import java.io.IOException;
import java.util.Date;

public class CreateWorkBook
{
    public static void main(String[] args) {

       try {
           new MainWindow();
       } catch (IOException e) {
            e.printStackTrace();
        }
//Utility utility= new Utility();
//System.out.println(utility.getAge(new Date("01/01/2004"), new Date("12/31/2015")));

    }
}
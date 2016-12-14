import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.IOException;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.Period;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.Calendar;
import java.util.Date;
import java.util.Locale;

import static java.util.Calendar.DATE;
import static java.util.Calendar.MONTH;
import static java.util.Calendar.YEAR;

/**
 * Created by Ayshahvez on 12/7/2016.
 */
public class Utility extends Component {

    public String getFilePath(){
        String filePath=null;
        JFileChooser chooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter(
                "Excel Files", "xlsx","xls");
        chooser.setFileFilter(filter);
        int returnVal = chooser.showOpenDialog(this);
        if(returnVal == JFileChooser.APPROVE_OPTION) {
            System.out.println("You chose to open this file: " +
                    //   chooser.getSelectedFile().getName());
                    chooser.getSelectedFile().getAbsolutePath());
            filePath= chooser.getSelectedFile().getAbsolutePath();
        }
        return filePath;
    }


    public String getWorkingDir(){
        String filePathWorkingDir=null;
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY); //to get only folders
        FileNameExtensionFilter filter = new FileNameExtensionFilter(
                "Excel Files, PDF", "xlsx","xls","pdf","doc","docx");
        chooser.setFileFilter(filter);
        int returnVal = chooser.showSaveDialog(this);
        if(returnVal == JFileChooser.APPROVE_OPTION) {
            System.out.println("You chose to open this file: " + chooser.getSelectedFile().getPath());
           filePathWorkingDir= chooser.getSelectedFile().getPath();
        }
        return filePathWorkingDir;
    }

    public static String getStartDate(int year, int month, int day){

        return LocalDate.of(year, month, day).minusYears(1).plusDays(1).format(DateTimeFormatter.ofPattern("dd-MMM-yy"));
    }

    public static long betweenDates(Date firstDate, Date secondDate) throws IOException
    {
        return ChronoUnit.DAYS.between(firstDate.toInstant(), secondDate.toInstant());
    }


    public static int calculateAge(LocalDate birthDate, LocalDate currentDate) {
        if ((birthDate != null) && (currentDate != null)) {
            return Period.between(birthDate, currentDate).getYears();
        } else {
            return 0;
        }
    }

    public static int getDiffYears(Date first, Date last) {
        Calendar a = getCalendar(first);
        Calendar b = getCalendar(last);
        int diff = b.get(YEAR) - a.get(YEAR);
        if (a.get(MONTH) > b.get(MONTH) ||
                (a.get(MONTH) == b.get(MONTH) && a.get(DATE) > b.get(DATE))) {
            diff--;
        }
        return diff;
    }

    public static Calendar getCalendar(Date date) {
        Calendar cal = Calendar.getInstance(Locale.US);
        cal.setTime(date);
        return cal;
    }


    public static String getDate(int year, int month, int day){

        return LocalDate.of(year, month, day).plusYears(1).minusDays(1).format(DateTimeFormatter.ofPattern("yyyy.MM.dd"));
    }


    public static String get1YearDateFromNow(int year, int month, int day){

        return LocalDate.of(year, month, day).plusYears(1).format(DateTimeFormatter.ofPattern("yyyy.MM.dd"));
    }

    public static String getEndDate(int year, int month, int day){

        return LocalDate.of(year, month, day).plusYears(1).minusDays(1).format(DateTimeFormatter.ofPattern("dd-MMM-yy"));
    }

    public static int getAge(Date dateofBirth, Date PlanEntry){
      //  Calendar today = Calendar.getInstance();

        Calendar PlanEntryDate = Calendar.getInstance();

        Calendar birthDate = Calendar.getInstance();

        int age =0;

        birthDate.setTime(dateofBirth);
        PlanEntryDate.setTime(PlanEntry);
      //  if(birthDate.after(today)){
      //      throw new IllegalArgumentException("You cannot be born in the future");
     //   }

        age =  PlanEntryDate.get(Calendar.YEAR) - birthDate.get(Calendar.YEAR);

        if((birthDate.get(Calendar.DAY_OF_YEAR) -  PlanEntryDate.get(Calendar.DAY_OF_YEAR)>3 ||
                (birthDate.get(Calendar.MONTH) > PlanEntryDate.get(Calendar.MONTH)))){
            age--;

        }else if((birthDate.get(Calendar.MONTH)== PlanEntryDate.get(Calendar.MONTH)) &&
        (birthDate.get(Calendar.DAY_OF_MONTH) >  PlanEntryDate.get(Calendar.DAY_OF_MONTH))){
    age--;
        }
        return age;
    }
}

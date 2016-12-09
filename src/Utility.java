import java.io.IOException;
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
public class Utility {

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

    public static int getAge(Date dateofBirth){
        Calendar today = Calendar.getInstance();
        Calendar birthDate = Calendar.getInstance();

        int age =0;

        birthDate.setTime(dateofBirth);
        if(birthDate.after(today)){
            throw new IllegalArgumentException("You cannot be born in the future");
        }

        age = today.get(Calendar.YEAR) - birthDate.get(Calendar.YEAR);

        if((birthDate.get(Calendar.DAY_OF_YEAR) - today.get(Calendar.DAY_OF_YEAR)>3 ||
                (birthDate.get(Calendar.MONTH) > today.get(Calendar.MONTH)))){
            age--;

        }else if((birthDate.get(Calendar.MONTH)==today.get(Calendar.MONTH)) &&
        (birthDate.get(Calendar.DAY_OF_MONTH) > today.get(Calendar.DAY_OF_MONTH))){
    age--;
        }
        return age;
    }
}

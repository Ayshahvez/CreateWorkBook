import java.io.*;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
import java.util.*;

//import com.sun.glass.ui.Pen;
import org.apache.poi.hssf.record.pivottable.StreamIDRecord;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.commons.collections4.*;

import java.time.LocalDate;
import   java.time.temporal.ChronoUnit;

import static java.util.Calendar.DATE;
import static java.util.Calendar.MONTH;
import static java.util.Calendar.YEAR;

public class CreateWorkBook
{
    public static void main(String[] args) throws IOException {
        Date StartDate = new GregorianCalendar(2004,Calendar.JANUARY,01).getTime();
        Date EndDate = new GregorianCalendar(2015,Calendar.DECEMBER,31).getTime();
        String PensionPlanName = "HAS PENSION FUND";

        try {

       TemplateSheets templateSheets = new TemplateSheets();
   //    templateSheets.Create_Template_Terminee_Sheet(StartDate,EndDate,PensionPlanName);
           // templateSheets.Create_Template_Terminee_Con();

       ExcelReader excelReader = new ExcelReader();
       excelReader.Create_Terminee_Contribution();


        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    }
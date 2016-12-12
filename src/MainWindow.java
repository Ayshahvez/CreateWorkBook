//import javax.swing.*;
//import static javax.swing.*;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;

/**
 * Created by Ayshahvez on 12/11/2016.
 */
public class MainWindow extends JFrame implements ActionListener{
    ExcelReader excelReader = new ExcelReader();
    TemplateSheets templateSheets = new TemplateSheets();
    ValidationChecks validationChecks = new ValidationChecks();
    Utility utility = new Utility();

//GLOBAL VARIABLES NEEDED FOR NOW
    String filePathValData=null;
    String filePathOutputTemplate=null;
    String filePathWorkingDir= null;
    String PensionPlanName=null;
    String PensionPlanStartDate=null;
    String PensionPlanEndDate=null;
    Date startDate = null;
    Date endDate=null;

//main MenuBar
JMenuBar menuBar;

    JPanel codePanel = new JPanel(new BorderLayout());
  ResultsWindow resultsWindow = new ResultsWindow();
    JScrollPane scrollPane;

    //Main menu items
    JMenu MenuLoadWorkbook;
    JMenu menuSingleCheck;

    JMenu MenuTemplateSheet;
    private  JMenu MenuValidationChecks;
    private JMenu MenuCreateWorkBook;

    private JMenu MenuSetPlanRequirements;
    private JMenu MenuStart_End_Dates;

    private JMenuItem MenuItemPensionPlanName;
    private JMenuItem MenuItemStartDate;
    private JMenuItem MenuItemEndDate;


    private JMenuItem CheckDuplicate;
    private JMenuItem CheckAge;
    private JMenuItem CheckEmployeePS;
    private JMenuItem CheckPlanEntry;
    private JMenuItem CheckDateofBirth;
    private JMenuItem CheckAll;

    private JMenuItem LoadValDataWorkBook;
    private  JMenuItem LoadPensionableSalaryWorkBook;

    private JMenuItem MenuItemLoadTemplateActiveSheet;
    private JMenuItem MenuItemLoadTemplateTermineeSheet;

    private JMenuItem MenuItemCreateActiveSheetTemplate;
    private JMenuItem MenuItemCreateTermineeSheetTemplate;

    private JMenuItem MenuItemCreateActiveSheet;
    private JMenuItem MenuItemCreateTermineeSheet;

    private JMenu MenuMembers;
    private JMenuItem MenuItemSeperateMembers;
    private JMenuItem MenuItemViewActiveMember;
    private JMenuItem MenuItemViewTermineeMember;
    private JMenuItem MenuItemViewAllMembers;

    //load template sheet menu
    private JMenu MenuCreateTemplateSheet;
    private JMenu MenuLoadTemplateSheet;
    private JMenuItem MenuItemLoadOutputTemplate;

    private JMenuItem MenuItemWorkingDir;



private JTextField txtVal,txtWelcome;
private JButton btnBrowse;
private JPanel jPanel,panWelcome;

    public MainWindow() {
        super("GFRAM Pension Automation Process Beta");
     //   this.curcon = curcon;
        this.setLayout(new BorderLayout());
        this.  initiliazeComponents();
        this.addComponentsTopanels();
        this.addPanelsToWindow();
        this.setWindowProperties();
        this.registerListener();
     //  String dir = utility.getFilePath();
    }

    private void setWindowProperties() {
        // TODO Auto-generated method stub
        this.setSize(700, 700);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setVisible(true);
     //   this.pack();
        this.setResizable(true);
    }


    public void initiliazeComponents () {
        MenuItemWorkingDir = new JMenuItem("Set default working directory");


        MenuSetPlanRequirements = new JMenu("Plan Requirements");
        MenuStart_End_Dates = new JMenu("Set Start and End Date");
        MenuItemStartDate = new JMenuItem("Set Start Date");
        MenuItemEndDate = new JMenuItem("Set End Date");
        MenuItemPensionPlanName = new JMenuItem("Set Pension Plan Name");


        MenuTemplateSheet = new JMenu("Template Sheet");
        MenuCreateTemplateSheet = new JMenu("Create Template Sheet");
        MenuItemCreateActiveSheetTemplate = new JMenuItem("Create Active Sheet Template");
        MenuItemCreateTermineeSheetTemplate = new JMenuItem("Create Terminee Sheet Template");
        MenuLoadTemplateSheet = new JMenu("Load Template Sheet");
        MenuItemLoadOutputTemplate = new JMenuItem("Load Template for Active and Terminated Members");



        MenuValidationChecks = new JMenu("Validation Checks");
        menuSingleCheck = new JMenu("Perform Single Check");
        CheckAll= new JMenuItem("Perform All Checks");
        MenuLoadWorkbook = new JMenu("Load WorkBook");
        MenuCreateWorkBook = new JMenu("Create WorkBook");


        MenuItemLoadTemplateActiveSheet = new JMenuItem("Load Template Sheet for Active Members");
        MenuItemLoadTemplateTermineeSheet = new JMenuItem("Load Template Sheet for Terminee Members");


        MenuItemCreateActiveSheet = new JMenuItem("Create Active Sheet");
        MenuItemCreateTermineeSheet = new JMenuItem("Create Terminee Sheet");


        scrollPane = new JScrollPane(resultsWindow);
        menuBar = new JMenuBar();

        CheckDuplicate = new JMenuItem("Duplicate Check");
        CheckAge = new JMenuItem("Age Check");
        CheckDateofBirth= new JMenuItem("Date of Birth Check");
        CheckEmployeePS= new JMenuItem("Employee Pensionable Salary Check");
        CheckPlanEntry= new JMenuItem("Plan Entry Check");



        txtVal = new JTextField("Valuation Data Workbook");
        txtWelcome = new JTextField("GFRAM Pension Automation Process Beta");
        txtWelcome.setEditable(false);
        btnBrowse = new JButton("Browse");
        jPanel = new JPanel(new FlowLayout());
        panWelcome = new JPanel(new FlowLayout());

        LoadValDataWorkBook = new JMenuItem("Browse for Valuation Data Workbook");
        LoadPensionableSalaryWorkBook = new JMenuItem("Browse for Pensionable Salary Workbook");

        MenuMembers = new JMenu("Members");
        MenuItemSeperateMembers = new JMenuItem("Seperate Active & Terminee Members");
        MenuItemViewActiveMember = new JMenuItem("View Active Members");
        MenuItemViewTermineeMember = new JMenuItem("View Terminated Members");
        MenuItemViewAllMembers = new JMenuItem("View All Members");

    }

    private void addComponentsTopanels() {
        MenuTemplateSheet.add(MenuCreateTemplateSheet);
        MenuCreateTemplateSheet.add(MenuItemCreateActiveSheetTemplate);
        MenuCreateTemplateSheet.add(MenuItemCreateTermineeSheetTemplate);
        MenuTemplateSheet.add(MenuLoadTemplateSheet);
        MenuLoadTemplateSheet.add(MenuItemLoadOutputTemplate);
        MenuLoadTemplateSheet.add(MenuItemLoadTemplateActiveSheet);
        MenuLoadTemplateSheet.add(MenuItemLoadTemplateTermineeSheet);

        MenuMembers.add(MenuItemSeperateMembers);
        MenuMembers.add(MenuItemViewActiveMember);
        MenuMembers.add(MenuItemViewTermineeMember);
        MenuMembers.add(MenuItemViewAllMembers);

        MenuSetPlanRequirements.add(MenuItemPensionPlanName);
        MenuSetPlanRequirements.add(MenuStart_End_Dates);
        MenuSetPlanRequirements.add(MenuItemWorkingDir);
        MenuStart_End_Dates.add(MenuItemStartDate);
        MenuStart_End_Dates.add(MenuItemEndDate);

        panWelcome.add(txtWelcome);

        menuSingleCheck.add(CheckDuplicate);
        menuSingleCheck.add(CheckAge);
        menuSingleCheck.add(CheckDateofBirth);
        menuSingleCheck.add(CheckPlanEntry);
        menuSingleCheck.add(CheckEmployeePS);
        MenuValidationChecks.add(menuSingleCheck);
       MenuValidationChecks.add(CheckAll);



        menuBar.add(MenuLoadWorkbook);
        menuBar.add(MenuSetPlanRequirements);
        menuBar.add(MenuMembers);
        menuBar.add(MenuTemplateSheet);
        menuBar.add(MenuValidationChecks);
        menuBar.add(MenuCreateWorkBook);

        MenuCreateWorkBook.add(MenuItemCreateActiveSheet);
        MenuCreateWorkBook.add(MenuItemCreateTermineeSheet);

        jPanel.add(menuBar);

       MenuLoadWorkbook.add(LoadValDataWorkBook);
       MenuLoadWorkbook.add(LoadPensionableSalaryWorkBook);
    }

    private void addPanelsToWindow() {
        this.add(panWelcome,BorderLayout.NORTH);
        this.add(jPanel,BorderLayout.SOUTH);
        this.add(scrollPane,BorderLayout.CENTER);
    }


    public void actionPerformed(ActionEvent e) {
        Color LINES = new Color(130, 125, 127);
        if (e.getSource().equals(CheckAge)) {
            JOptionPane.showMessageDialog(null, "Now Performing Age Check, Press Ok to Continue", "Age Check", JOptionPane.PLAIN_MESSAGE);
            String result = null;
            try {
                result = validationChecks.Check_Age(filePathValData);
            } catch (IOException e1) {
                e1.printStackTrace();
            }
            resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);

        }

        if (e.getSource().equals(CheckDateofBirth)) {
            JOptionPane.showMessageDialog(null, "Now Performing Date of Birth Check, Press Ok to Continue", "Date of Birth Check", JOptionPane.PLAIN_MESSAGE);
            String result = null;
            try {
                result = validationChecks.Check_DateofBirth(filePathWorkingDir);
            } catch (IOException e1) {
                e1.printStackTrace();
            }
            resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);

        }

        if (e.getSource().equals(CheckDuplicate)) {

            JOptionPane.showMessageDialog(null, "Now Performing Duplicate Check, Press Ok to Continue", "Duplicate Check", JOptionPane.PLAIN_MESSAGE);
            String s = null;
            try {
                s = validationChecks.Check_For_Duplicates(filePathValData);
            } catch (IOException e1) {
                e1.printStackTrace();
            }
            resultsWindow.appendToPane(resultsWindow, s + "\n", LINES, true);
        }

        if(e.getSource().equals(LoadValDataWorkBook)){
            JFileChooser chooser = new JFileChooser();
            FileNameExtensionFilter filter = new FileNameExtensionFilter(
                    "Excel Files", "xlsx","xls");
            chooser.setFileFilter(filter);
            int returnVal = chooser.showOpenDialog(this);
            if(returnVal == JFileChooser.APPROVE_OPTION) {
                System.out.println("You chose to open this file: " +
                     //   chooser.getSelectedFile().getName());
              chooser.getSelectedFile().getAbsolutePath());
             filePathValData= chooser.getSelectedFile().getAbsolutePath();
            }
        }

        if(e.getSource().equals(MenuItemSeperateMembers)){
         //   excelReader.Separate_Actives_Terminees(filePathValData);

            String s = null;
            s = excelReader.Separate_Actives_Terminees(filePathValData,filePathOutputTemplate,filePathWorkingDir);
            resultsWindow.appendToPane(resultsWindow, s + "\n", LINES, true);
        }

        if(e.getSource().equals(MenuItemPensionPlanName)){
            PensionPlanName = JOptionPane.showInputDialog("Please enter the name of the Pension Plan: ");
            JOptionPane.showMessageDialog(null,"You have entered, the pension plan name as: "+PensionPlanName,"Plan Name",JOptionPane.PLAIN_MESSAGE);
        }

        if(e.getSource().equals(MenuItemStartDate)){
          PensionPlanStartDate = JOptionPane.showInputDialog("Please enter the Start Date of the Pension Plan [mm/dd/yyyy]: ");


            try {
               startDate = new SimpleDateFormat("mm/dd/yyyy").parse(PensionPlanStartDate);
            } catch (ParseException ee) {
                ee.printStackTrace();
            }

            JOptionPane.showMessageDialog(null,"You have entered the Start Date as: "+ PensionPlanStartDate,"End Date",JOptionPane.PLAIN_MESSAGE);
        }

        if(e.getSource().equals(MenuItemEndDate)){
             PensionPlanEndDate = JOptionPane.showInputDialog("Please enter the End Date of the Pension Plan [mm/dd/yyyy]: ");

            try {
                endDate = new SimpleDateFormat("mm/dd/yyyy").parse(PensionPlanEndDate);
            } catch (ParseException ee) {
                ee.printStackTrace();
            }
            JOptionPane.showMessageDialog(null,"You have entered the End Date as: "+ PensionPlanEndDate,"End Date",JOptionPane.PLAIN_MESSAGE);
        }

        if(e.getSource().equals(MenuItemLoadOutputTemplate)){
            filePathOutputTemplate = utility.getFilePath();
        }

        if(e.getSource().equals(MenuItemWorkingDir)){
            filePathWorkingDir = utility.getWorkingDir();
        }

       if(e.getSource().equals(MenuItemCreateActiveSheetTemplate)){
           try {
               templateSheets.Create_Template_Active_Sheet(startDate,endDate,PensionPlanName,filePathWorkingDir);
           } catch (IOException e1) {
               e1.printStackTrace();
           }
       }

        if(e.getSource().equals(MenuItemCreateTermineeSheetTemplate)){
            try {
                templateSheets.Create_Template_Terminee_Sheet(startDate,endDate,PensionPlanName,filePathWorkingDir);
            } catch (IOException e1) {
                e1.printStackTrace();
            }
            JOptionPane.showMessageDialog(null,"The Terminee Template Sheet was created Successfully","Success",JOptionPane.PLAIN_MESSAGE);
        }

        if(e.getSource().equals(MenuItemCreateActiveSheet)){
            String result = excelReader.Create_Actives_Sheet(filePathWorkingDir);
            try {
                excelReader.Create_Activee_Contribution(PensionPlanStartDate,PensionPlanEndDate,filePathWorkingDir);
            } catch (IOException e1) {
                e1.printStackTrace();
            }
            resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
            JOptionPane.showMessageDialog(null,"The Active Sheet was created Successfully","Success",JOptionPane.PLAIN_MESSAGE);
        }


        if(e.getSource().equals(MenuItemCreateTermineeSheet)){
            String result = excelReader.Create_Terminee_Sheet(PensionPlanStartDate,PensionPlanEndDate,filePathWorkingDir);
            try {
               excelReader.Create_Terminee_Contribution(PensionPlanStartDate,PensionPlanEndDate,filePathWorkingDir);
            } catch (IOException e1) {
                e1.printStackTrace();
            }
            resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
            JOptionPane.showMessageDialog(null,"The Terminee Sheet was created Successfully","Success",JOptionPane.PLAIN_MESSAGE);
        }

    }


    private void registerListener() {
        CheckDuplicate.addActionListener(this);
        CheckAge.addActionListener(this);
        CheckDateofBirth.addActionListener(this);
        CheckEmployeePS.addActionListener(this);
        CheckPlanEntry.addActionListener(this);


       LoadValDataWorkBook.addActionListener(this);
        MenuItemSeperateMembers.addActionListener(this);
        MenuItemPensionPlanName.addActionListener(this);
MenuItemStartDate.addActionListener(this);
MenuItemEndDate.addActionListener(this);
MenuItemLoadOutputTemplate.addActionListener(this);
MenuItemWorkingDir.addActionListener(this);

        MenuItemCreateActiveSheetTemplate.addActionListener(this);
        MenuItemCreateTermineeSheetTemplate.addActionListener(this);

        MenuItemCreateActiveSheet.addActionListener(this);
        MenuItemCreateTermineeSheet.addActionListener(this);
    }


}


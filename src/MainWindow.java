//import javax.swing.*;
//import static javax.swing.*;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;

/**
 * Created by Ayshahvez konowalchuk xbox one XL NARUTO on 12/11/2016.
 */
public class MainWindow extends JFrame implements ActionListener {
    ExcelReader excelReader = new ExcelReader();
    TemplateSheets templateSheets = new TemplateSheets();
    ValidationChecks validationChecks = new ValidationChecks();
    Utility utility = new Utility();

    //GLOBAL VARIABLES NEEDED FOR NOW
    String filePathValData = null;
    String filePathOutputTemplate = null;
    Date startDate = null;
    Date endDate = null;
    //  WorkingDir

    private String WorkingDir = utility.read();
    String filePathWorkingDir = WorkingDir;
    String PensionPlanName = null;// utility.readFile("PN",filePathWorkingDir);
    String PensionPlanStartDate = null;//utility.readFile("SD",filePathWorkingDir);
    String PensionPlanEndDate = null; //utility.readFile("ED",filePathWorkingDir);


    //main MenuBar
    JMenuBar menuBar;


    JMenu MenuEditor;
    JMenuItem MenuItemViewPlanDeatils;
    JMenuItem MenuItemClearScreen;
    JMenuItem MenuItemRefreshData;
    JMenuItem MenuItemSaveData;

    JPanel codePanel = new JPanel(new BorderLayout());
    ResultsWindow resultsWindow = new ResultsWindow();
    JScrollPane scrollPane;

    //Main menu items
    JMenu MenuLoadWorkbook;
    JMenu menuSingleCheck;

    JMenu MenuTemplateSheet;
    private JMenu MenuValidationChecks;
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
    private JMenuItem LoadPensionableSalaryWorkBook;

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
    private JMenuItem MenuItemViewRetiredMember;
    private JMenuItem MenuItemViewAllMembers;
    private JMenuItem MenuItemViewDeferredMember;
    private JMenuItem MenuItemViewDeceasedMember;
    private JMenuItem MenuItemViewTerminatedMember;
    //load template sheet menu
    private JMenu MenuCreateTemplateSheet;
    private JMenu MenuLoadTemplateSheet;
    private JMenuItem MenuItemLoadOutputTemplate;

    private JMenuItem MenuItemWorkingDir;

    JLabel imgLabel;

    private JTextField txtVal, txtWelcome;
    private JButton btnBrowse;
    private JPanel jPanel, panWelcome, eastPanel, westPanel;

    public MainWindow() throws IOException {
        super("GFRAM Pension Automation Process Beta");
        //   this.curcon = curcon;
        this.setLayout(new BorderLayout());
        this.initiliazeComponents();
        this.addComponentsTopanels();
        this.addPanelsToWindow();
        this.setWindowProperties();
        this.registerListener();
    }


    private void setWindowProperties() {
        // TODO Auto-generated method stub
        this.setSize(700, 800);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setVisible(true);
        //   this.pack();
        this.setResizable(true);
    }


    public void initiliazeComponents() {
        eastPanel = new JPanel(new FlowLayout());
        westPanel = new JPanel(new FlowLayout());

        //  imgLabel = new JLabel(new ImageIcon(filePathWorkingDir+"\\dp.png"));
        imgLabel = new JLabel(new ImageIcon("C:\\Users\\akonowalchuk\\GFRAM\\dp.png"));
        //    imgLabel = new JLabel(new ImageIcon("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\dp.png"));


        MenuEditor = new JMenu("Editor");
        MenuItemClearScreen = new JMenuItem("Clear Screen");
        MenuItemViewPlanDeatils = new JMenuItem("View Current Pension Plan Details");
        MenuItemRefreshData = new JMenuItem("Refresh Plan Requirements Data");
        MenuItemSaveData = new JMenuItem("Save Plan Requirement Data");


        MenuItemWorkingDir = new JMenuItem("Set Default Working Directory");

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
        CheckAll = new JMenuItem("Perform All Checks");
        MenuLoadWorkbook = new JMenu("Load WorkBook");
        MenuCreateWorkBook = new JMenu("Create WorkBook");

        MenuItemLoadTemplateActiveSheet = new JMenuItem("Load Template Sheet for Active Members");
        MenuItemLoadTemplateTermineeSheet = new JMenuItem("Load Template Sheet for Terminee Members");

        MenuItemSeperateMembers = new JMenuItem("Separate Active & Terminee Members into a new WorkBook");
        MenuItemCreateActiveSheet = new JMenuItem("Create Active Members Sheet");
        MenuItemCreateTermineeSheet = new JMenuItem("Create Terminee Members Sheet");

        scrollPane = new JScrollPane(resultsWindow);
        menuBar = new JMenuBar();

        CheckDuplicate = new JMenuItem("Duplicate Check");
        CheckAge = new JMenuItem("Age Check");
        CheckDateofBirth = new JMenuItem("Date of Birth Check");
        CheckEmployeePS = new JMenuItem("Employee Pensionable Salary Check");
        CheckPlanEntry = new JMenuItem("Plan Entry Check");


        txtVal = new JTextField("Valuation Data Workbook");
        txtWelcome = new JTextField("\t\tGFRAM Pension Automation Process Beta");
        txtWelcome.setEditable(false);
        btnBrowse = new JButton("Browse");
        jPanel = new JPanel(new FlowLayout());
        panWelcome = new JPanel(new BorderLayout());

        LoadValDataWorkBook = new JMenuItem("Browse for Valuation Data Workbook");
        LoadPensionableSalaryWorkBook = new JMenuItem("Browse for Pensionable Salary Workbook");

        MenuMembers = new JMenu("Members");

        MenuItemViewActiveMember = new JMenuItem("View all Active Members");
        MenuItemViewTermineeMember = new JMenuItem("View all Terminee Members");
        MenuItemViewRetiredMember = new JMenuItem("View Retired Members");
        MenuItemViewDeferredMember = new JMenuItem("View Deferred Members");
        MenuItemViewDeceasedMember = new JMenuItem("View Deceased Members");
        MenuItemViewAllMembers = new JMenuItem("View All Members");
        MenuItemViewTerminatedMember = new JMenuItem("View Terminated Members");

    }

    private void addComponentsTopanels() {
        MenuTemplateSheet.add(MenuCreateTemplateSheet);
        //    MenuTemplateSheet.add(MenuLoadTemplateSheet);

        MenuCreateTemplateSheet.add(MenuItemCreateActiveSheetTemplate);
        MenuCreateTemplateSheet.add(MenuItemCreateTermineeSheetTemplate);

        MenuLoadTemplateSheet.add(MenuItemLoadOutputTemplate);
        MenuLoadTemplateSheet.add(MenuItemLoadTemplateActiveSheet);
        MenuLoadTemplateSheet.add(MenuItemLoadTemplateTermineeSheet);


        MenuMembers.add(MenuItemViewActiveMember);
        MenuMembers.add(MenuItemViewTerminatedMember);
        MenuMembers.add(MenuItemViewTermineeMember);
        MenuMembers.add(MenuItemViewRetiredMember);
        MenuMembers.add(MenuItemViewDeferredMember);
        MenuMembers.add(MenuItemViewDeceasedMember);
        MenuMembers.add(MenuItemViewAllMembers);

        MenuSetPlanRequirements.add(MenuItemWorkingDir);
        MenuSetPlanRequirements.add(MenuItemPensionPlanName);
        MenuSetPlanRequirements.add(MenuStart_End_Dates);

        MenuStart_End_Dates.add(MenuItemStartDate);
        MenuStart_End_Dates.add(MenuItemEndDate);

        //   panWelcome.add(txtWelcome,BorderLayout.NORTH);
        panWelcome.add(imgLabel, BorderLayout.SOUTH);
        //  panWelcome.a

        menuSingleCheck.add(CheckDuplicate);
        menuSingleCheck.add(CheckAge);
        menuSingleCheck.add(CheckDateofBirth);
        menuSingleCheck.add(CheckPlanEntry);
        menuSingleCheck.add(CheckEmployeePS);
        MenuValidationChecks.add(menuSingleCheck);
        MenuValidationChecks.add(CheckAll);

        //    menuBar.add(MenuLoadWorkbook);
        menuBar.add(MenuSetPlanRequirements);
        menuBar.add(MenuTemplateSheet);
        menuBar.add(MenuMembers);
        menuBar.add(MenuValidationChecks);
        menuBar.add(MenuCreateWorkBook);
        menuBar.add(MenuEditor);

        MenuEditor.add(MenuItemClearScreen);
        MenuEditor.add(MenuItemViewPlanDeatils);
        MenuEditor.add(MenuItemSaveData);
        MenuEditor.add(MenuItemRefreshData);

        MenuCreateWorkBook.add(MenuItemSeperateMembers);
        MenuCreateWorkBook.add(MenuItemCreateActiveSheet);
        MenuCreateWorkBook.add(MenuItemCreateTermineeSheet);

        jPanel.add(menuBar);

        MenuLoadWorkbook.add(LoadValDataWorkBook);
        MenuLoadWorkbook.add(LoadPensionableSalaryWorkBook);
    }

    private void addPanelsToWindow() {
        this.add(panWelcome, BorderLayout.NORTH);
        this.add(jPanel, BorderLayout.SOUTH);
        this.add(scrollPane, BorderLayout.CENTER);
        this.add(eastPanel, BorderLayout.EAST);
        this.add(westPanel, BorderLayout.WEST);
    }


    public void actionPerformed(ActionEvent e) {
        Color LINES = new Color(130, 125, 127);

        if (e.getSource().equals(MenuItemClearScreen)) {
            resultsWindow.ClearScreen(resultsWindow);
        }


        if (e.getSource().equals(MenuItemViewPlanDeatils)) {
            // try {
            //     JOptionPane.showMessageDialog(null, "Plan Name: " + utility.readFile("PN",filePathWorkingDir) + "\nStart Date: "+utility.readFile("SD",filePathWorkingDir)+" \nEnd Date: "+utility.readFile("ED",filePathWorkingDir)+"\nCurrent Working Directory: "+utility.readFile("WD",filePathWorkingDir),"Info", JOptionPane.PLAIN_MESSAGE);
            JOptionPane.showMessageDialog(null, "Plan Name: " + PensionPlanName + "\nStart Date: " + PensionPlanStartDate + " \nEnd Date: " + PensionPlanEndDate + "\nCurrent Working Directory: " + filePathWorkingDir, "Info", JOptionPane.PLAIN_MESSAGE);
            //  } catch (IOException e1) {
            //      e1.printStackTrace();
            //  }
        }

        if (e.getSource().equals(MenuItemSaveData)) {

            if (PensionPlanEndDate != null)
                utility.writeDefaultsToFile("ED", filePathWorkingDir, PensionPlanEndDate);

            if (PensionPlanStartDate != null)
                utility.writeDefaultsToFile("SD", filePathWorkingDir, PensionPlanStartDate);

            if (PensionPlanName != null)
                utility.writeDefaultsToFile("PN", filePathWorkingDir, PensionPlanName);


            if (PensionPlanName != null && PensionPlanEndDate != null && PensionPlanStartDate != null) {
                JOptionPane.showMessageDialog(null, "You have successfully stored the Plan Name, Start Date and End Date for Future use", "Success", JOptionPane.PLAIN_MESSAGE);
            } else if (PensionPlanName != null && PensionPlanEndDate != null) {
                JOptionPane.showMessageDialog(null, "You have successfully stored the Plan Name and Plan End Date for Future use", "Success", JOptionPane.PLAIN_MESSAGE);
            } else if (PensionPlanStartDate != null && PensionPlanEndDate != null) {
                JOptionPane.showMessageDialog(null, "You have successfully stored the Plan Start Date and Plan End Date for Future use", "Success", JOptionPane.PLAIN_MESSAGE);
            } else if (PensionPlanStartDate != null && PensionPlanName != null) {
                JOptionPane.showMessageDialog(null, "You have successfully stored the Plan Start Date and Plan Name for Future use", "Success", JOptionPane.PLAIN_MESSAGE);
            } else if (PensionPlanName != null) {
                JOptionPane.showMessageDialog(null, "You have successfully stored the Plan Name for Future use", "Success", JOptionPane.PLAIN_MESSAGE);
            } else if (PensionPlanStartDate != null) {
                JOptionPane.showMessageDialog(null, "You have successfully stored the Plan Start Date for Future use", "Success", JOptionPane.PLAIN_MESSAGE);
            } else if (PensionPlanEndDate != null) {
                JOptionPane.showMessageDialog(null, "You have successfully stored the Plan End Date for Future use", "Success", JOptionPane.PLAIN_MESSAGE);
            }

        }


        if (e.getSource().equals(MenuItemRefreshData)) {
            try {
                if (new File(filePathWorkingDir + "\\WD.txt").exists()) {
                    this.filePathWorkingDir = utility.readFile("WD", filePathWorkingDir);
                }

                if (new File(filePathWorkingDir + "\\PN.txt").exists()) {
                    PensionPlanName = utility.readFile("PN", filePathWorkingDir);
                }
                if (new File(filePathWorkingDir + "\\SD.txt").exists()) {
                    this.PensionPlanStartDate = utility.readFile("SD", filePathWorkingDir);
                }

                if (new File(filePathWorkingDir + "\\ED.txt").exists()) {
                    this.PensionPlanEndDate = utility.readFile("ED", filePathWorkingDir);
                }

            } catch (IOException e1) {
                e1.printStackTrace();
            }
            JOptionPane.showMessageDialog(null, "Refreshed Successfully ", "Success", JOptionPane.PLAIN_MESSAGE);
        }


        if (e.getSource().equals(CheckAll)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                JOptionPane.showMessageDialog(null, "Now Performing Employee Plan Entry Check, Press Ok to Continue", "Plan Entry Check", JOptionPane.PLAIN_MESSAGE);
                String result = null;
                try {
                    result = validationChecks.Check_For_Duplicates(filePathWorkingDir);
                    result += validationChecks.Check_Plan_EntryDate_empDATE(filePathWorkingDir);
                    result += validationChecks.Check_Age(filePathWorkingDir);
                    result += validationChecks.Check_DateofBirth(filePathWorkingDir);
                    result += validationChecks.Check_FivePercent_PS(filePathWorkingDir);
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
                resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
            }
        }

        if (e.getSource().equals(CheckPlanEntry)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                JOptionPane.showMessageDialog(null, "Now Performing Employee Plan Entry Check, Press Ok to Continue", "Plan Entry Check", JOptionPane.PLAIN_MESSAGE);
                String result = null;
                try {
                    result = validationChecks.Check_Plan_EntryDate_empDATE(filePathWorkingDir);
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
                resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
            }
        }

        if (e.getSource().equals(CheckEmployeePS)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                JOptionPane.showMessageDialog(null, "Now Performing Pensionable Salary and Contributing Check, Press Ok to Continue", "Pensionable Check", JOptionPane.PLAIN_MESSAGE);
                String result = null;
                try {
                    result = validationChecks.Check_FivePercent_PS(filePathWorkingDir);
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
                resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
            }
        }

        if (e.getSource().equals(CheckAge)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                JOptionPane.showMessageDialog(null, "Now Performing Age Check, Press Ok to Continue", "Age Check", JOptionPane.PLAIN_MESSAGE);
                String result = null;
                try {
                    result = validationChecks.Check_Age(filePathWorkingDir);
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
                resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
            }
        }

        if (e.getSource().equals(CheckDateofBirth)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                JOptionPane.showMessageDialog(null, "Now Performing Date of Birth Check, Press Ok to Continue", "Date of Birth Check", JOptionPane.PLAIN_MESSAGE);
                String result = null;
                try {
                    result = validationChecks.Check_DateofBirth(filePathWorkingDir);
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
                resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
            }
        }

        if (e.getSource().equals(CheckDuplicate)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                JOptionPane.showMessageDialog(null, "Now Performing Duplicate Check, Press Ok to Continue", "Duplicate Check", JOptionPane.PLAIN_MESSAGE);
                String s = null;
                try {
                    s = validationChecks.Check_For_Duplicates(filePathWorkingDir);
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
                resultsWindow.appendToPane(resultsWindow, s + "\n", LINES, true);
            }
        }

        if (e.getSource().equals(LoadValDataWorkBook)) {
            JFileChooser chooser = new JFileChooser();
            FileNameExtensionFilter filter = new FileNameExtensionFilter(
                    "Excel Files", "xlsx", "xls");
            chooser.setFileFilter(filter);
            int returnVal = chooser.showOpenDialog(this);
            if (returnVal == JFileChooser.APPROVE_OPTION) {
                System.out.println("You chose to open this file: " +
                        //   chooser.getSelectedFile().getName());
                        chooser.getSelectedFile().getAbsolutePath());
                filePathValData = chooser.getSelectedFile().getAbsolutePath();
            }
        }

        if (e.getSource().equals(MenuItemSeperateMembers)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                       if(new File(filePathWorkingDir+"\\Hose Valuation Data (Actuary's copy).xlsx ").exists()) {
                           String result = null;
                           //     result = excelReader.Separate_Actives_Terminees(filePathValData,filePathOutputTemplate,filePathWorkingDir);
                           result = excelReader.Separate_Actives_Terminees(filePathWorkingDir);
                           resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                           JOptionPane.showMessageDialog(null, "Active and Terminated Members now separated, Please remember to input date of refunds for Terminee members", "Success", JOptionPane.PLAIN_MESSAGE);
                       }
                       else{
                           JOptionPane.showMessageDialog(null, "Please ensure the Valuation Data Workbook is present in the Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
                       }
            }
        }
        if (e.getSource().equals(MenuItemPensionPlanName)) {
            if (filePathWorkingDir != null) {
                PensionPlanName = JOptionPane.showInputDialog("Please enter the name of the Pension Plan: ");
                if (PensionPlanName != null) {
                    JOptionPane.showMessageDialog(null, "You have entered, the pension plan name as: " + PensionPlanName, "Plan Name", JOptionPane.PLAIN_MESSAGE);
                    //   utility.writeDefaultsToFile("PN",filePathWorkingDir,PensionPlanName);
                }
            } else {
                JOptionPane.showMessageDialog(null, "Please Ensure you Set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            }

        }

        if (e.getSource().equals(MenuItemStartDate)) {

            if (filePathWorkingDir != null) {
                PensionPlanStartDate = JOptionPane.showInputDialog("Please enter the Start Date of the Pension Plan [mm/dd/yyyy]: ");

                if (PensionPlanStartDate != null) {
                    try {
                        startDate = new SimpleDateFormat("MM/dd/yyyy").parse(PensionPlanStartDate);
                    } catch (ParseException ee) {
                        ee.printStackTrace();
                    }

                    JOptionPane.showMessageDialog(null, "You have entered the Start Date as: " + PensionPlanStartDate, "End Date", JOptionPane.PLAIN_MESSAGE);
                    //    utility.writeDefaultsToFile("SD",filePathWorkingDir,PensionPlanStartDate);
                }
            } else {
                JOptionPane.showMessageDialog(null, "Please Ensure you Set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            }

        }

        if (e.getSource().equals(MenuItemEndDate)) {

            if (filePathWorkingDir != null) {
                PensionPlanEndDate = JOptionPane.showInputDialog("Please enter the End Date of the Pension Plan [mm/dd/yyyy]: ");
                if (PensionPlanEndDate != null) {
                    try {
                        endDate = new SimpleDateFormat("MM/dd/yyyy").parse(PensionPlanEndDate);
                    } catch (ParseException ee) {
                        ee.printStackTrace();
                    }
                    JOptionPane.showMessageDialog(null, "You have entered the End Date as: " + PensionPlanEndDate, "End Date", JOptionPane.PLAIN_MESSAGE);
                    //     utility.writeDefaultsToFile("ED",filePathWorkingDir,PensionPlanEndDate);
                }
            } else {
                JOptionPane.showMessageDialog(null, "Please Ensure you Set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            }

        }

        if (e.getSource().equals(MenuItemLoadOutputTemplate)) {
            filePathOutputTemplate = utility.getFilePath();
        }


        if (e.getSource().equals(MenuItemWorkingDir)) {
            filePathWorkingDir = utility.getWorkingDir();
            if (filePathWorkingDir != null) {
                JOptionPane.showMessageDialog(null, "You have set the working directory to: " + filePathWorkingDir, "Success", JOptionPane.PLAIN_MESSAGE);

                utility.writeDefaultsToFile("WD", filePathWorkingDir, filePathWorkingDir);

            }
        }


        if (e.getSource().equals(MenuItemCreateActiveSheetTemplate)) {
            //   File f = new File(WorkingDir + "//Updated_Actives_Sheet.xlsx");
            if (PensionPlanStartDate == null && PensionPlanEndDate == null && PensionPlanName == null) {
                JOptionPane.showMessageDialog(null, "Please Ensure you Input the Plan Name, Plan Start Date and Plan End Date", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else if (PensionPlanStartDate == null && PensionPlanEndDate == null) {
                JOptionPane.showMessageDialog(null, "Please Ensure you Input the Plan Start Date and Plan End Date", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else if (PensionPlanName == null && PensionPlanEndDate == null) {
                JOptionPane.showMessageDialog(null, "Please Ensure you Input the Plan Name and Plan End Date", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else if (PensionPlanName == null && PensionPlanStartDate == null) {
                JOptionPane.showMessageDialog(null, "Please Ensure you Input the Plan Name and Plan Start Date", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else if (PensionPlanStartDate == null) {
                JOptionPane.showMessageDialog(null, "Please Ensure you Input the Start Date", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else if (PensionPlanEndDate == null) {
                JOptionPane.showMessageDialog(null, "Please Ensure you Input the End Date", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else if (PensionPlanName == null) {
                JOptionPane.showMessageDialog(null, "Please Ensure you Input the Plan Name", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please Ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            }


            if (PensionPlanStartDate != null && PensionPlanEndDate != null && PensionPlanName != null) {
                try {

                    TemplateSheets.Create_Template_Active_Sheet(PensionPlanStartDate, PensionPlanEndDate, PensionPlanName, filePathWorkingDir);
                    //    TemplateSheets.Create_Template_Active_Sheet(utility.readFile("SD",filePathWorkingDir), utility.readFile("ED",filePathWorkingDir), utility.readFile("PN",filePathWorkingDir), filePathWorkingDir);
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
                if (new File(filePathWorkingDir + "//Template_Active_Sheet.xlsx").exists()) {
                    JOptionPane.showMessageDialog(null, "The Active Sheet Template was created Successfully", "Success", JOptionPane.PLAIN_MESSAGE);
                }
            }
        }

        //create template sheet for active members
        if (e.getSource().equals(MenuItemCreateTermineeSheetTemplate)) {
            if (PensionPlanStartDate == null && PensionPlanEndDate == null && PensionPlanName == null) {
                JOptionPane.showMessageDialog(null, "Please Ensure you Input the Plan Name, Plan Start Date and Plan End Date", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else if (PensionPlanStartDate == null && PensionPlanEndDate == null) {
                JOptionPane.showMessageDialog(null, "Please Ensure you Input the Plan Start Date and Plan End Date", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else if (PensionPlanName == null && PensionPlanEndDate == null) {
                JOptionPane.showMessageDialog(null, "Please Ensure you Input the Plan Name and Plan End Date", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else if (PensionPlanName == null && PensionPlanStartDate == null) {
                JOptionPane.showMessageDialog(null, "Please Ensure you Input the Plan Name and Plan Start Date", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else if (PensionPlanStartDate == null) {
                JOptionPane.showMessageDialog(null, "Please Ensure you Input the Start Date", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else if (PensionPlanEndDate == null) {
                JOptionPane.showMessageDialog(null, "Please Ensure you Input the End Date", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else if (PensionPlanName == null) {
                JOptionPane.showMessageDialog(null, "Please Ensure you Input the Plan Name", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please Ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            }


            if (PensionPlanStartDate != null && PensionPlanEndDate != null && PensionPlanName != null) {
                try {
                    TemplateSheets.Create_Template_Terminee_Sheet(PensionPlanStartDate, PensionPlanEndDate, PensionPlanName, filePathWorkingDir);;
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
                JOptionPane.showMessageDialog(null, "The Terminee Sheet  Template was created Successfully", "Success", JOptionPane.PLAIN_MESSAGE);
            }
        }//end
        if (e.getSource().equals(MenuItemCreateActiveSheet)) {
            String result=null;
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                if (new File(filePathWorkingDir + "\\Template_Active_Sheet.xlsx ").exists()) {

                    if(PensionPlanEndDate!=null && PensionPlanStartDate!=null) {
                        result = excelReader.Create_Actives_Sheet(filePathWorkingDir);
                    //    result = excelReader.Create_Terminee_Sheet(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                        try {
                            excelReader.Create_Activee_Contribution(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                        } catch (IOException e1) {
                            e1.printStackTrace();
                        }
                        resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                        JOptionPane.showMessageDialog(null, "The Active Sheet was created Successfully", "Success", JOptionPane.PLAIN_MESSAGE);
                    }  else{
                        JOptionPane.showMessageDialog(null, "Please Refresh The Plan Requirement Data, then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                } else {

                    JOptionPane.showMessageDialog(null, "Please Ensure you Create the Template Sheet for the Active Sheet", "Notice", JOptionPane.PLAIN_MESSAGE);
                }

            }
                }


        if (e.getSource().equals(MenuItemCreateTermineeSheet)) {
            String result = null;
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                if (new File(filePathWorkingDir + "\\Template_Terminee_Sheet.xlsx ").exists()) {
                        if(PensionPlanEndDate!=null && PensionPlanStartDate!=null){
                            result = excelReader.Create_Terminee_Sheet(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);


                    if (new File(filePathWorkingDir + "\\Terminees_Sheet.xlsx ").exists()) {
                        try {
                            excelReader.Create_Terminee_Contribution(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                        } catch (IOException e1) {
                            e1.printStackTrace();
                        }
                        resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                        JOptionPane.showMessageDialog(null, "The Terminee Sheet was created Successfully", "Success", JOptionPane.PLAIN_MESSAGE);
                    }
                    else{
                        JOptionPane.showMessageDialog(null, "Please Ensure the Date of Refunds were Inputted for the Terminated Members", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                        }
                        else{
                            JOptionPane.showMessageDialog(null, "Please Refresh The Plan Requirement Data, then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                        }
                } else {
                    JOptionPane.showMessageDialog(null, "Please Ensure you Create the Template Sheet for the Terminee Sheet", "Notice", JOptionPane.PLAIN_MESSAGE);
                }

            }
        }

        if (e.getSource().equals(MenuItemViewActiveMember)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {


                if (new File(filePathWorkingDir + "\\Output.xlsx").exists()) {


                    String result = excelReader.View_Actives_Members(filePathWorkingDir, PensionPlanEndDate);
                    JOptionPane.showMessageDialog(null, "Please wait for the list of the Active Members as at " + PensionPlanEndDate, "Success", JOptionPane.PLAIN_MESSAGE);
                    resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                } else {
                    JOptionPane.showMessageDialog(null, "Please ensure you Create the Workbook with Active and Terminee Members separated", "Notice", JOptionPane.PLAIN_MESSAGE);
                }
            }
        }

        if (e.getSource().equals(MenuItemViewTermineeMember)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {


                if (new File(filePathWorkingDir + "\\Output.xlsx").exists()) {

                    String result = excelReader.View_Terminee_Members(filePathWorkingDir, PensionPlanEndDate);
                    JOptionPane.showMessageDialog(null, "Please wait for the list of the Terminee Members as at " + PensionPlanEndDate, "Success", JOptionPane.PLAIN_MESSAGE);
                    resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);

                } else {
                    JOptionPane.showMessageDialog(null, "Please ensure you Create the Workbook with Active and Terminee Members separated", "Notice", JOptionPane.PLAIN_MESSAGE);
                }
            }
        }


        if (e.getSource().equals(MenuItemViewTerminatedMember)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {

                if (new File(filePathWorkingDir + "\\Output.xlsx").exists()) {

                    String result = excelReader.View_Terminated_Members(filePathWorkingDir, PensionPlanEndDate);
                    JOptionPane.showMessageDialog(null, "Please wait for the list of the Terminated Members as at " + PensionPlanEndDate, "Success", JOptionPane.PLAIN_MESSAGE);
                    resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                } else {
                    JOptionPane.showMessageDialog(null, "Please ensure you Create the Workbook with Active and Terminee Members separated", "Notice", JOptionPane.PLAIN_MESSAGE);
                }
            }
        }


        if (e.getSource().equals(MenuItemViewAllMembers)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {

                if (new File(filePathWorkingDir + "\\Output.xlsx").exists()) {

                    String result = excelReader.View_Actives_Members(filePathWorkingDir, PensionPlanEndDate);
                    result += excelReader.View_Terminee_Members(filePathWorkingDir, PensionPlanEndDate);
                    JOptionPane.showMessageDialog(null, "Please wait for the list of all the Members as at " + PensionPlanEndDate, "Success", JOptionPane.PLAIN_MESSAGE);
                    resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                } else {
                    JOptionPane.showMessageDialog(null, "Please ensure you Create the Workbook with Active and Terminee Members separated", "Notice", JOptionPane.PLAIN_MESSAGE);
                }
            }
        }


        if (e.getSource().equals(MenuItemViewRetiredMember)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {

                if (new File(filePathWorkingDir + "\\Output.xlsx").exists()) {

                    String result = excelReader.View_Retired_Members(filePathWorkingDir, PensionPlanEndDate);
                    JOptionPane.showMessageDialog(null, "Please wait for the list of the Retired Members as at " + PensionPlanEndDate, "Success", JOptionPane.PLAIN_MESSAGE);
                    resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                } else {
                    JOptionPane.showMessageDialog(null, "Please ensure you Create the Workbook with Active and Terminee Members separated", "Notice", JOptionPane.PLAIN_MESSAGE);
                }
            }
        }
        if (e.getSource().equals(MenuItemViewDeceasedMember)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {

                if (new File(filePathWorkingDir + "\\Output.xlsx").exists()) {

                    String result = excelReader.View_Deceased_Members(filePathWorkingDir, PensionPlanEndDate);
                    JOptionPane.showMessageDialog(null, "Please wait for the list of the Deceased Members as at " + PensionPlanEndDate, "Success", JOptionPane.PLAIN_MESSAGE);
                    resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                } else {
                    JOptionPane.showMessageDialog(null, "Please ensure you Create the Workbook with Active and Terminee Members separated", "Notice", JOptionPane.PLAIN_MESSAGE);
                }
            }
        }

        if (e.getSource().equals(MenuItemViewDeferredMember)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {

                if (new File(filePathWorkingDir + "\\Output.xlsx").exists()) {

                    String result = excelReader.View_Deferred_Members(filePathWorkingDir, PensionPlanEndDate);
                    JOptionPane.showMessageDialog(null, "Please wait for the list of the Deferred Members as at " + PensionPlanEndDate, "Success", JOptionPane.PLAIN_MESSAGE);
                    resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                } else {
                    JOptionPane.showMessageDialog(null, "Please ensure you Create the Workbook with Active and Terminee Members separated", "Notice", JOptionPane.PLAIN_MESSAGE);
                }
            }
        }
    }

    private void registerListener() {
        //MEMBERS LISTENERS
        MenuItemViewActiveMember.addActionListener(this);
        MenuItemViewTermineeMember.addActionListener(this);
        MenuItemViewAllMembers.addActionListener(this);
        MenuItemViewRetiredMember.addActionListener(this);
        MenuItemViewDeceasedMember.addActionListener(this);
        MenuItemViewDeferredMember.addActionListener(this);
        MenuItemViewTerminatedMember.addActionListener(this);

        MenuItemClearScreen.addActionListener(this);
        MenuItemRefreshData.addActionListener(this);
        MenuItemViewPlanDeatils.addActionListener(this);
        MenuItemSaveData.addActionListener(this);

        CheckDuplicate.addActionListener(this);
        CheckAge.addActionListener(this);
        CheckDateofBirth.addActionListener(this);
        CheckEmployeePS.addActionListener(this);
        CheckPlanEntry.addActionListener(this);
        CheckAll.addActionListener(this);

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


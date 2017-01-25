import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;


public class MainWindow extends JFrame implements ActionListener {
    Desktop dt = Desktop.getDesktop();
    ExcelReader excelReader = new ExcelReader();
    ValidationChecks validationChecks = new ValidationChecks();
    Table table = new Table();
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
    String planEntryAge = null;

    JMenu MenuCreateTableTemplate;
    JMenuItem MenuItemCreateTemplateSummaryofActiveMembership;
    JMenuItem MenuItemCreateTemplateMovementsinActiveMemberships;
    JMenuItem MenuItemCreateTemplateAnalysisofFundYield;
    JMenuItem MenuItemCreateTemplateGainsLosses;

    JMenuItem MenuItemCreateTableSummaryofActiveMembership;
    JMenuItem MenuItemCreateTableMovementsinActiveMemberships;
    JMenuItem MenuItemCreateTableAnalysisofFundYield;
    JMenuItem MenuItemCreateTableGainsLosses;
    //main MenuBar
    JMenuBar menuBar;
    JMenuBar MenuBarCreateTemplateSheets;

    JMenu MenuEditor;
    JMenuItem MenuItemViewPlanDeatils;
    JMenuItem MenuItemClearScreen;
    JMenuItem MenuItemRefreshData;
    JMenuItem MenuItemSaveData;
    JMenuItem MenuItemDeleteData;


    JPanel codePanel = new JPanel(new BorderLayout());
    ResultsWindow resultsWindow = new ResultsWindow();
    //   tableWindow tableWindow = new t
    JScrollPane scrollPane;

    // JTable table;

    //Main menu items
    JMenu MenuLoadWorkbook;
    JMenu menuSingleCheck;

    JMenu MenuTemplateSheet;
    //  private JMenuItem MenuItemCreateSeperatedTemplate;
    private JMenu MenuValidationChecks;
    private JMenu MenuCreateWorkBook;

    private JMenu MenuSetPlanRequirements;
    private JMenu MenuStart_End_Dates;

    private JMenu MenuFeesTemplate;
    private JMenuItem MenuItemCreateFeesActiveSheetTemplate;
    private JMenuItem MenuItemCreateFeesTermineeSheetTemplate;

    private JMenu MenuNoFeesTemplate;
    private JMenuItem MenuItemCreateNoFeesActiveSheetTemplate;
    private JMenuItem MenuItemCreateNoFeesTermineeSheetTemplate;

    private JMenuItem MenuItemPensionPlanName;
    private JMenuItem MenuItemStartDate;
    private JMenuItem MenuItemEndDate;
    JMenuItem MenuItemInterestRates;

    private JMenuItem CheckDuplicate;
    private JMenuItem CheckAge;
    private JMenuItem CheckEmployeePS;
    private JMenuItem CheckPlanEntry;
    private JMenuItem CheckDateofBirth;
    private JMenuItem CheckAll;

    //   private JMenuItem LoadValDataWorkBook;
    //  private JMenuItem LoadPensionableSalaryWorkBook;

    private JMenuItem MenuItemLoadTemplateActiveSheet;
    private JMenuItem MenuItemLoadTemplateTermineeSheet;


    private JMenu MenuCreateWorkbookNoFees;
    private JMenuItem MenuItemCreateActiveSheet;
    private JMenuItem MenuItemCreateTermineeSheet;

    private JMenu MenuCreateWorkbookFees;
    private JMenuItem MenuItemCreateFeesActiveSheet;
    private JMenuItem MenuItemCreateFeesTermineeSheet;

    private JMenu MenuMembers;
    //  private JMenuItem MenuItemSeperateMembers;
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

    //CALCULATION MENU
    JMenu MenuCreateTable;
    JMenuItem MenuItemCreateIncExpTemplate;
    JMenuItem MenuItemCreateBalSheetTemplate;

    JMenuItem MenuItemCreateIncExpTable;
    JMenuItem MenuItemCreateBalSheetTable;

    public MainWindow() throws IOException {
        super("GFRAM Direct Contribution Pension Automation Process Beta");
        try {
            //   UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            Font f = new Font("Default", Font.PLAIN, 14);
            UIManager.put("Menu.font", f);
            UIManager.put("MenuItem.font", f);

        } catch (Exception e) {
            e.printStackTrace();
        }
        //   this.curcon = curcon;
        this.setLayout(new BorderLayout());
        this.initiliazeComponents();
        this.addComponentsTopanels();
        this.addPanelsToWindow();
        this.setWindowProperties();
        this.registerListener();

    }

    private void setWindowProperties() {
        //TODO Auto-generated method stub
        this.setSize(800, 800);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setVisible(true);
        //  this.pack();
        this.setResizable(true);
        resultsWindow.setBorder(BorderFactory.createLineBorder(Color.gray, 3));
        resultsWindow.setEditable(false);
    }

    public void initiliazeComponents() {
        //TABLE TEMPLATES
        MenuCreateTableTemplate = new JMenu("Create Table Template");
        MenuItemCreateTemplateSummaryofActiveMembership = new JMenuItem("Create Template Summary of Active Membership");
        MenuItemCreateTemplateMovementsinActiveMemberships = new JMenuItem("Create Template Movements in Active Memberships");
        MenuItemCreateTemplateAnalysisofFundYield = new JMenuItem("Create Template Analysis of Fund Yield");
        MenuItemCreateTemplateGainsLosses = new JMenuItem("Create Template Gains and Losses");

        MenuItemCreateTableSummaryofActiveMembership = new JMenuItem("Create Table Summary of Active Membership");
        MenuItemCreateTableMovementsinActiveMemberships = new JMenuItem("Create Table Movements in Active Memberships");
        MenuItemCreateTableAnalysisofFundYield = new JMenuItem("Create Table Analysis of Fund Yield");
        MenuItemCreateTableGainsLosses = new JMenuItem("Create Table Gains and Losses");


        //CALCULATIONS
        MenuCreateTable = new JMenu("Create Table");
        MenuItemCreateIncExpTemplate = new JMenuItem("Create Income and Expenditure Template");
        MenuItemCreateBalSheetTemplate = new JMenuItem("Create Balance Sheet Template");

        MenuItemCreateIncExpTable = new JMenuItem("Create Income and Expenditure Table");
        MenuItemCreateBalSheetTable = new JMenuItem("Create Balance Sheet Table");

        MenuBarCreateTemplateSheets = new JMenuBar();
/*        String[] columnHeaders = {"Popularity","Position","Team","Manager","Points"};
        Object[][] data = {
                {"1","Famous","Man Utd","Luis Van Gaal","86"},

        };
        table =  new JTable(data,columnHeaders);
        table.setPreferredScrollableViewportSize(new Dimension(500,80));*/

        MenuCreateWorkBook = new JMenu("Create WorkBook");


        MenuFeesTemplate = new JMenu("Create Template With Fees");
        MenuNoFeesTemplate = new JMenu("Create Template With No Fees");

        MenuCreateWorkbookFees = new JMenu("Create Workbook with Fees");
        MenuCreateWorkbookNoFees = new JMenu("Create Workbok with No Fees");


        MenuItemCreateFeesActiveSheet = new JMenuItem("Create Active Member Sheet with Fees");
        MenuItemCreateFeesTermineeSheet = new JMenuItem("Create Terminee Member Sheet with Fees");

        MenuItemCreateActiveSheet = new JMenuItem("Create Active Members Sheet with no Fees");
        MenuItemCreateTermineeSheet = new JMenuItem("Create Terminee Members Sheet with no Fees");

        jPanel = new JPanel(new FlowLayout());
        panWelcome = new JPanel(new BorderLayout());
        menuBar = new JMenuBar();

        scrollPane = new JScrollPane(resultsWindow);
        //   scrollPane = new JScrollPane(table);

        eastPanel = new JPanel(new FlowLayout());
        westPanel = new JPanel(new FlowLayout());

        //  imgLabel = new JLabel(new ImageIcon(filePathWorkingDir+"\\dp.png"));
        imgLabel = new JLabel(new ImageIcon("C:\\Users\\akonowalchuk\\OneDrive\\GFRAM\\dp.png"));
       //   imgLabel = new JLabel(new ImageIcon("C:\\Users\\Ayshahvez\\OneDrive\\GFRAM\\dp.png"));

        //PLAN REQUIREMENTS
        MenuSetPlanRequirements = new JMenu("Plan Requirements");
        MenuItemWorkingDir = new JMenuItem("Set Default Working Directory");
        MenuItemPensionPlanName = new JMenuItem("Set Pension Plan Name");
        MenuStart_End_Dates = new JMenu("Set Start and End Date");
        MenuItemStartDate = new JMenuItem("Set Start Date");
        MenuItemEndDate = new JMenuItem("Set End Date");
        MenuItemInterestRates = new JMenuItem("Input Interest Rates");

        //TEMPLATE SHEETS
        MenuTemplateSheet = new JMenu("Template Sheet");
        MenuCreateTemplateSheet = new JMenu("Create Template Sheet");
        //   MenuItemCreateSeperatedTemplate = new JMenuItem("Create Template for Seperated Workbook");
        MenuLoadTemplateSheet = new JMenu("Load Template Sheet");
        MenuItemLoadOutputTemplate = new JMenuItem("Load Template for Active and Terminated Members");

        //FEES
        MenuItemCreateFeesActiveSheetTemplate = new JMenuItem("Create Active Sheet Template with Fees");
        MenuItemCreateFeesTermineeSheetTemplate = new JMenuItem("Create Terminee Sheet Template with Fees");

        //NO FEES
        MenuItemCreateNoFeesActiveSheetTemplate = new JMenuItem("Create Active Sheet Template with No Fees");
        MenuItemCreateNoFeesTermineeSheetTemplate = new JMenuItem("Create Terminee Sheet Template with No Fees");

        //MEMBERS
        MenuMembers = new JMenu("Members");
        MenuItemViewActiveMember = new JMenuItem("View all Active Members");
        MenuItemViewTermineeMember = new JMenuItem("View all Terminee Members");
        MenuItemViewRetiredMember = new JMenuItem("View Retired Members");
        MenuItemViewDeferredMember = new JMenuItem("View Deferred Members");
        MenuItemViewDeceasedMember = new JMenuItem("View Deceased Members");
        MenuItemViewAllMembers = new JMenuItem("View All Members");
        MenuItemViewTerminatedMember = new JMenuItem("View Terminated Members");


        MenuValidationChecks = new JMenu("Data Quality Checks");
        menuSingleCheck = new JMenu("Perform Validation Check");
        CheckDuplicate = new JMenuItem("Duplicate Check");
        CheckAge = new JMenuItem("Age Check");
        CheckDateofBirth = new JMenuItem("Date of Birth Check");
        CheckEmployeePS = new JMenuItem("Employee Pensionable Salary Check");
        CheckPlanEntry = new JMenuItem("Plan Entry Check");
        CheckAll = new JMenuItem("Perform All Checks");


        MenuEditor = new JMenu("Editor");
        MenuItemClearScreen = new JMenuItem("Clear Screen");
        MenuItemViewPlanDeatils = new JMenuItem("View Current Pension Plan Details");
        MenuItemRefreshData = new JMenuItem("Refresh Plan Requirements Data");
        MenuItemSaveData = new JMenuItem("Save Plan Requirement Data");
        MenuItemDeleteData = new JMenuItem("Delete Plan Requirement Data");

        MenuLoadWorkbook = new JMenu("Load WorkBook");
        MenuItemLoadTemplateActiveSheet = new JMenuItem("Load Template Sheet for Active Members");
        MenuItemLoadTemplateTermineeSheet = new JMenuItem("Load Template Sheet for Terminee Members");
        //    MenuItemSeperateMembers = new JMenuItem("Separate Active & Terminee Members into a new WorkBook");

    }

    private void addComponentsTopanels() {


        //    MenuCreateWorkBook.add(MenuItemSeperateMembers);
        MenuCreateWorkBook.add(MenuCreateWorkbookFees);
        MenuCreateWorkBook.add(MenuCreateWorkbookNoFees);

        MenuCreateWorkbookFees.add(MenuItemCreateFeesActiveSheet);
        MenuCreateWorkbookFees.add(MenuItemCreateFeesTermineeSheet);

        MenuCreateWorkbookNoFees.add(MenuItemCreateActiveSheet);
        MenuCreateWorkbookNoFees.add(MenuItemCreateTermineeSheet);

        MenuTemplateSheet.add(MenuCreateTemplateSheet);
        //    MenuTemplateSheet.add(MenuLoadTemplateSheet);

        //   MenuCreateTemplateSheet.add(MenuItemCreateSeperatedTemplate);

        MenuCreateTemplateSheet.add(MenuFeesTemplate);
        MenuCreateTemplateSheet.add(MenuNoFeesTemplate);


        MenuFeesTemplate.add(MenuItemCreateFeesActiveSheetTemplate);
        MenuFeesTemplate.add(MenuItemCreateFeesTermineeSheetTemplate);

        MenuNoFeesTemplate.add(MenuItemCreateNoFeesActiveSheetTemplate);
        MenuNoFeesTemplate.add(MenuItemCreateNoFeesTermineeSheetTemplate);
        // MenuCreateTemplateSheet.add(MenuItemCreateIncExpTemplate);
        //   MenuCreateTemplateSheet.add(MenuItemCreateBalSheetTemplate);

        //   MenuLoadTemplateSheet.add(MenuItemLoadOutputTemplate);
        //   MenuLoadTemplateSheet.add(MenuItemLoadTemplateActiveSheet);
        //   MenuLoadTemplateSheet.add(MenuItemLoadTemplateTermineeSheet);

        MenuMembers.add(MenuItemViewActiveMember);
        MenuMembers.add(MenuItemViewTerminatedMember);
        MenuMembers.add(MenuItemViewTermineeMember);
        MenuMembers.add(MenuItemViewRetiredMember);
        //      MenuMembers.add(MenuItemViewDeferredMember);
        MenuMembers.add(MenuItemViewDeceasedMember);
        //     MenuMembers.add(MenuItemViewAllMembers);

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
        //     MenuValidationChecks.add(CheckAll);

        //    menuBar.add(MenuLoadWorkbook);
        menuBar.add(MenuSetPlanRequirements);

        menuBar.add(MenuTemplateSheet);

        menuBar.add(MenuMembers);
        menuBar.add(MenuValidationChecks);
        menuBar.add(MenuCreateWorkBook);
        menuBar.add(MenuCreateTable);
        menuBar.add(MenuEditor);

        MenuEditor.add(MenuItemClearScreen);
        MenuEditor.add(MenuItemViewPlanDeatils);
        MenuEditor.add(MenuItemRefreshData);
        MenuEditor.add(MenuItemSaveData);
        MenuEditor.add(MenuItemDeleteData);


        jPanel.add(menuBar);
        //  westPanel.add(MenuBarCreateTemplateSheets);
        //  eastPanel.add(MenuCreateTable);
        //     MenuBarCreateTemplateSheets.add(MenuItemCreateIncExpTemplate);

//        MenuLoadWorkbook.add(LoadValDataWorkBook);
        //   MenuLoadWorkbook.add(LoadPensionableSalaryWorkBook);
        //TABLES
        MenuCreateTemplateSheet.add(MenuCreateTableTemplate);

        //add tables template to create table menu
        MenuCreateTableTemplate.add(MenuItemCreateIncExpTemplate);
        MenuCreateTableTemplate.add(MenuItemCreateBalSheetTemplate);
        MenuCreateTableTemplate.add(MenuItemCreateTemplateSummaryofActiveMembership);
        MenuCreateTableTemplate.add(MenuItemCreateTemplateMovementsinActiveMemberships);
        MenuCreateTableTemplate.add(MenuItemCreateTemplateAnalysisofFundYield);
        MenuCreateTableTemplate.add(MenuItemCreateTemplateGainsLosses);

        //CALCULATIONS
        MenuCreateTable.add(MenuItemCreateIncExpTable);
        MenuCreateTable.add(MenuItemCreateBalSheetTable);
        MenuCreateTable.add(MenuItemCreateTableSummaryofActiveMembership);
        MenuCreateTable.add(MenuItemCreateTableMovementsinActiveMemberships);
        MenuCreateTable.add(MenuItemCreateTableAnalysisofFundYield);
        MenuCreateTable.add(MenuItemCreateTableGainsLosses);
    }

    private void addPanelsToWindow() {
        this.add(panWelcome, BorderLayout.NORTH);
        this.add(jPanel, BorderLayout.SOUTH);
        this.add(scrollPane, BorderLayout.CENTER);
        this.add(eastPanel, BorderLayout.EAST);
        this.add(westPanel, BorderLayout.WEST);
    }

    public void actionPerformed(ActionEvent e) {
        //  Color LINES = new Color(130, 125, 127);
        Color LINES = new Color(105, 105, 107);

        if (e.getSource().equals(MenuItemCreateTableSummaryofActiveMembership)) {
            try {
                if (PensionPlanName == null && PensionPlanStartDate == null && PensionPlanEndDate == null) {
                    if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please refresh the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure you Input the Plan Requiremenets Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                } else {
                    if(new File(filePathWorkingDir+"\\Templates\\Template_Summary_of_Active_Membership.xlsx").exists()) {

                        if(new File(filePathWorkingDir+"\\Accumulated_Actives_Sheet.xlsx").exists()) {
                            table.Create_Table_Summary_of_Active_Membership(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                            JOptionPane.showMessageDialog(null, "Summary of Active Membership Table Was Successfully Created", "Notice", JOptionPane.PLAIN_MESSAGE);
                        }else{
                            JOptionPane.showMessageDialog(null, "The Active Sheet was not found, Please Create it before you proceed!", "Notice", JOptionPane.PLAIN_MESSAGE);
                        }

                }
                    else{
                    JOptionPane.showMessageDialog(null, "Please ensure the Template for the Summary of Active Membership is Created", "Notice", JOptionPane.PLAIN_MESSAGE);
                }
                }
            } catch (IOException e1) {
                e1.printStackTrace();
            }

        }

        if (e.getSource().equals(MenuItemCreateTableAnalysisofFundYield)) {
            try {
                if (PensionPlanName == null && PensionPlanStartDate == null && PensionPlanEndDate == null) {
                    if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please refresh the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure you Input the Plan Requiremenets Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                } else {
                    if(new File(filePathWorkingDir+"\\Templates\\Template_Analysis_of_Fund_Yield.xlsx").exists()) {
                        table.Create_Table_Analysis_of_Fund_Yield(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                        JOptionPane.showMessageDialog(null, "Analysis of Fund Yield Table Was Successfully Created", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                    else{
                        JOptionPane.showMessageDialog(null, "Please ensure the Template for the Analysis of Fund Yield Table is Created", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                }
            } catch (IOException e1) {
                e1.printStackTrace();
            }

        }

        if (e.getSource().equals(MenuItemCreateTableMovementsinActiveMemberships)) {
            try {
                if (PensionPlanName == null && PensionPlanStartDate == null && PensionPlanEndDate == null) {
                    if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please refresh the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure you Input the Plan Requiremenets Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                } else {
                    if(new File(filePathWorkingDir+"\\Templates\\Template_Movements_in_Active_Membership.xlsx").exists()) {
                    table.Create_Table_Movement_in_Active_Membership(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                    JOptionPane.showMessageDialog(null, "Movement in Active Membership Table Was Successfully Created", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                    else{
                        JOptionPane.showMessageDialog(null, "Please ensure the Template for the Movements in Active Membership is Created", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                }
            } catch (IOException e1) {
                e1.printStackTrace();
            }

        }


        if (e.getSource().equals(MenuItemCreateTemplateSummaryofActiveMembership)) {
            try {
                if (PensionPlanName == null && PensionPlanStartDate == null && PensionPlanEndDate == null) {
                    if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please refresh the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure you Input the Plan Requiremenets Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                } else {
                    TemplateSheets.Create_Template_Summary_of_Active_Memberships(PensionPlanStartDate, PensionPlanEndDate, PensionPlanName, filePathWorkingDir);
                    JOptionPane.showMessageDialog(null, "The Template was created Successfully", "Notice", JOptionPane.PLAIN_MESSAGE);
                    JOptionPane.showMessageDialog(null, "Now Please Input any other Input Data in the Summary of Active Membership Template and SAVE the WorkBook", "Notice", JOptionPane.PLAIN_MESSAGE);
                    if (new File(filePathWorkingDir + "\\Templates\\Template_Summary_of_Active_Membership.xlsx").exists()) {
                        dt.open(new File(filePathWorkingDir + "\\Templates\\Template_Summary_of_Active_Membership.xlsx"));
                    }
                }
            } catch (IOException e1) {
                e1.printStackTrace();
            }

        }

        if (e.getSource().equals(MenuItemCreateTemplateMovementsinActiveMemberships)) {
            try {
                if (PensionPlanName == null && PensionPlanStartDate == null && PensionPlanEndDate == null) {
                    if (new File(filePathWorkingDir + "\\Program Files\\Program Files\\PN.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please refresh the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure you Input the Plan Requiremenets Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                } else {
                    TemplateSheets.Create_Template_Movement_in_Active_Memberships(PensionPlanStartDate, PensionPlanEndDate, PensionPlanName, filePathWorkingDir);
                    JOptionPane.showMessageDialog(null, "The Template was created Successfully", "Notice", JOptionPane.PLAIN_MESSAGE);
                    JOptionPane.showMessageDialog(null, "Now Please Input any other Input Data in the Movement in Active Membership Template", "Notice", JOptionPane.PLAIN_MESSAGE);
                    if (new File(filePathWorkingDir + "\\Templates\\Template_Movements_in_Active_Membership.xlsx").exists()) {
                        dt.open(new File(filePathWorkingDir + "\\Templates\\Template_Movements_in_Active_Membership.xlsx"));
                    }
                }
            } catch (IOException e1) {
                e1.printStackTrace();
            }
        }

        if (e.getSource().equals(MenuItemCreateTemplateAnalysisofFundYield)) {
            try {
                if (PensionPlanName == null && PensionPlanStartDate == null && PensionPlanEndDate == null) {
                    if (new File(filePathWorkingDir + "\\Program Files\\Program Files\\PN.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please refresh the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure you Input the Plan Requiremenets Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                } else {
                    TemplateSheets.Create_Template_Analysis_of_Fund_Yield(PensionPlanStartDate, PensionPlanEndDate, PensionPlanName, filePathWorkingDir);

                    JOptionPane.showMessageDialog(null, "The Template was created Successfully", "Notice", JOptionPane.PLAIN_MESSAGE);
                    JOptionPane.showMessageDialog(null, "Now Please Input any other Input Data in the Analysis of Fund Yield Template and SAVE the WorkBook", "Notice", JOptionPane.PLAIN_MESSAGE);
                    if (new File(filePathWorkingDir + "\\Templates\\Template_Analysis_of_Fund_Yield.xlsx").exists()) {

                        dt.open(new File(filePathWorkingDir + "\\Templates\\Template_Analysis_of_Fund_Yield.xlsx"));
                    }
                }
            } catch (IOException e1) {
                e1.printStackTrace();
            } catch (ParseException e1) {
                e1.printStackTrace();
            }
    }


        if (e.getSource().equals(MenuItemCreateTemplateGainsLosses)) {
            try {
                if (PensionPlanName == null && PensionPlanStartDate == null && PensionPlanEndDate == null) {
                    if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please refresh the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure you Input the Plan Requiremenets Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                } else {
                    TemplateSheets.Create_Template_Gains_Losses(PensionPlanStartDate, PensionPlanEndDate, PensionPlanName, filePathWorkingDir);
                    JOptionPane.showMessageDialog(null, "The Template was created Successfully", "Notice", JOptionPane.PLAIN_MESSAGE);
                    JOptionPane.showMessageDialog(null, "Now Please Input any other Input Data in the Gains and Losses Template and SAVE the WorkBook", "Notice", JOptionPane.PLAIN_MESSAGE);
                    if (new File(filePathWorkingDir + "\\Templates\\Template_Gains_Losses.xlsx").exists()) {
                        dt.open(new File(filePathWorkingDir + "\\Templates\\Template_Gains_Losses.xlsx"));
                    }
                }
                } catch(IOException e1){
                    e1.printStackTrace();
                }
            }

        if (e.getSource().equals(MenuItemCreateBalSheetTemplate)) {
            try {
                if (PensionPlanName == null && PensionPlanStartDate == null && PensionPlanEndDate == null) {
                    if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please refresh the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure you Input the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                } else {
                    TemplateSheets.Create_Template_Balance_Sheet(PensionPlanStartDate, PensionPlanEndDate, PensionPlanName, filePathWorkingDir);
                    JOptionPane.showMessageDialog(null, "The Template was created Successfully", "Notice", JOptionPane.PLAIN_MESSAGE);
                    JOptionPane.showMessageDialog(null, "Now Please Input any other Input Data in the Balance Sheet Template and SAVE the WorkBook", "Notice", JOptionPane.PLAIN_MESSAGE);
                    if (new File(filePathWorkingDir + "\\Templates\\Template_Balance_Sheet.xlsx").exists()) {
                        dt.open(new File(filePathWorkingDir + "\\Templates\\Template_Balance_Sheet.xlsx"));
                    }
                }
                } catch(IOException e1){
                    e1.printStackTrace();
                } catch(ParseException e1){
                    e1.printStackTrace();
                }
            }

        if (e.getSource().equals(MenuItemCreateIncExpTemplate)) {
            try {
                if (PensionPlanName == null && PensionPlanStartDate == null && PensionPlanEndDate == null) {
                    if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please refresh the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure you Input the Plan Requiremenets Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                } else {
                    TemplateSheets.Create_Template_Inc_Exp_Sheet(PensionPlanStartDate, PensionPlanEndDate, PensionPlanName, filePathWorkingDir);

                    JOptionPane.showMessageDialog(null, "The Template was created Successfully", "Notice", JOptionPane.PLAIN_MESSAGE);
                    JOptionPane.showMessageDialog(null, "Now Please Input any other Input Data in the Income and Expenditure Template and SAVE the WorkBook", "Notice", JOptionPane.PLAIN_MESSAGE);
                    if (new File(filePathWorkingDir + "\\Templates\\Template_Inc_Exp_Sheet.xlsx").exists()) {
                        dt.open(new File(filePathWorkingDir + "\\Templates\\Template_Inc_Exp_Sheet.xlsx"));
                    }
                }
                } catch(IOException e1){
                    e1.printStackTrace();
                }
        }


        if (e.getSource().equals(MenuItemCreateIncExpTable)) {
            try {
                if (new File(filePathWorkingDir + "\\Templates\\Template_Inc_Exp_Sheet.xlsx").exists()) {
                    if (PensionPlanName == null && PensionPlanStartDate == null && PensionPlanEndDate == null) {
                        if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                            JOptionPane.showMessageDialog(null, "Please refresh the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                        } else {
                            JOptionPane.showMessageDialog(null, "Please ensure you Input the Plan Requiremenets Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                        }
                    } else {
                        if(new File(filePathWorkingDir+"\\Templates\\Template_Inc_Exp_Sheet.xlsx").exists()) {

                            table.Create_Income_Expenditure_Table(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                        JOptionPane.showMessageDialog(null, "Income and Expenditure Table Was Successfully Created", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                    else{
                        JOptionPane.showMessageDialog(null, "Please ensure the Template for the Income and Expenditure Table is Created", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Please Ensure you Create the Income and Expenditure Template before Proceeding", "Notice", JOptionPane.PLAIN_MESSAGE);
                }
            } catch (IOException e1) {
                e1.printStackTrace();
            }
        }

        if (e.getSource().equals(MenuItemCreateBalSheetTable)) {
            try {
                if (PensionPlanName == null && PensionPlanStartDate == null && PensionPlanEndDate == null) {
                    if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please refresh the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure you Input the Plan Requiremenets Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                } else {

                    if(new File(filePathWorkingDir+"\\Templates\\Template_Balance_Sheet.xlsx").exists()) {
                        table.Create_Balance_Sheet_Table(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                    JOptionPane.showMessageDialog(null, "The Valuation Balance Sheet Table Was Successfully Created", "Notice", JOptionPane.PLAIN_MESSAGE);
                }
                    else{
                    JOptionPane.showMessageDialog(null, "Please ensure the Template for the Balance Sheet Table is Created", "Notice", JOptionPane.PLAIN_MESSAGE);
                }
                }
            } catch (IOException e1) {
                e1.printStackTrace();
            }
        }

        if (e.getSource().equals(MenuItemCreateTableGainsLosses)) {
            try {
                if (PensionPlanName == null && PensionPlanStartDate == null && PensionPlanEndDate == null) {
                    if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please refresh the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure you Input the Plan Requiremenets Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                } else {
                    table.Create_Table_Gains_and_Losses(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                    JOptionPane.showMessageDialog(null, "The Gains and Losses Table Was Successfully Created", "Notice", JOptionPane.PLAIN_MESSAGE);
                }
            } catch (IOException e1) {
                e1.printStackTrace();
            }
        }

        if (e.getSource().equals(MenuItemClearScreen)) {
            resultsWindow.ClearScreen(resultsWindow);
        }

        if (e.getSource().equals(MenuItemDeleteData)) {
            if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() || new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() || new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                int option = JOptionPane.showConfirmDialog(null, "Are you Sure you want to Delete Saved Data?", "Notice", JOptionPane.OK_CANCEL_OPTION);
                if (option == JOptionPane.OK_OPTION) { // Afirmative
                    //....
                    if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists()) {
                        new File(filePathWorkingDir + "\\Program Files\\PN.txt").delete();
                    }

                    if (new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists()) {
                        new File(filePathWorkingDir + "\\Program Files\\SD.txt").delete();
                    }

                    if (new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        new File(filePathWorkingDir + "\\Program Files\\ED.txt").delete();
                    }
                    JOptionPane.showMessageDialog(null, "Saved Data Was Successfully Deleted!", "Notice", JOptionPane.PLAIN_MESSAGE);
                }

            } else {
                JOptionPane.showMessageDialog(null, "There was no Saved Data to be Deleted!", "Notice", JOptionPane.PLAIN_MESSAGE);
            }
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
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {

                if (PensionPlanEndDate != null)
                    utility.writeDefaultsToFile("ED", filePathWorkingDir, PensionPlanEndDate);

                if (PensionPlanStartDate != null)
                    utility.writeDefaultsToFile("SD", filePathWorkingDir, PensionPlanStartDate);

                if (PensionPlanName != null)
                    utility.writeDefaultsToFile("PN", filePathWorkingDir, PensionPlanName);

                if (filePathWorkingDir != null)
                    utility.writeDefaultsToFile("WD", filePathWorkingDir, filePathWorkingDir);

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
                } else {
                    JOptionPane.showMessageDialog(null, "There was no Data to be Saved!", "Notice", JOptionPane.PLAIN_MESSAGE);
                }
            }
        }


        if (e.getSource().equals(MenuItemRefreshData)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                if (!new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() && !new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && !new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                    JOptionPane.showMessageDialog(null, "Sorry, There was no Data to Load, Please Input all the Plan Requirements Data ", "Notice", JOptionPane.PLAIN_MESSAGE);
                } else {
                    try {
                        if (new File(filePathWorkingDir + "\\Program Files\\WD.txt").exists()) {
                            this.filePathWorkingDir = utility.readFile("WD", filePathWorkingDir);
                        }

                        if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists()) {
                            PensionPlanName = utility.readFile("PN", filePathWorkingDir);
                        }
                        if (new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists()) {
                            this.PensionPlanStartDate = utility.readFile("SD", filePathWorkingDir);
                        }

                        if (new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                            this.PensionPlanEndDate = utility.readFile("ED", filePathWorkingDir);
                        }

                    } catch (IOException e1) {
                        e1.printStackTrace();
                    }
                    JOptionPane.showMessageDialog(null, "Refreshed Successfully ", "Success", JOptionPane.PLAIN_MESSAGE);
                }
            }
        }

        if (e.getSource().equals(CheckPlanEntry)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                if (PensionPlanName == null && PensionPlanStartDate == null && PensionPlanEndDate == null) {
                    if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please refresh the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure you Input the Plan Requiremenets Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                } else {
                    if (PensionPlanName == null && PensionPlanStartDate == null && PensionPlanEndDate == null) {
                        if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                            JOptionPane.showMessageDialog(null, "Please refresh the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                        } else {
                            JOptionPane.showMessageDialog(null, "Please ensure you Input the Plan Requiremenets Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                        }
                    } else {
                        JOptionPane.showMessageDialog(null, "Now Performing Employee Plan Entry Check, Press Ok to Continue", "Plan Entry Check", JOptionPane.PLAIN_MESSAGE);
                        String result = null;
                        try {
                            ArrayList al = validationChecks.check_Plan_EntryDate_empDATE(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                            result = validationChecks.getResult();
                            new AlToTable(al, "Check Plan Entry Date");
                            resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                        } catch (IOException e1) {
                            e1.printStackTrace();
                        }

                    }
                }
            }
        }

        if (e.getSource().equals(CheckEmployeePS)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                try {
                    if (PensionPlanStartDate == null && PensionPlanEndDate == null) {
                        JOptionPane.showMessageDialog(null, "Please Ensure you Input the Start Date and End Date", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else if (PensionPlanStartDate == null) {
                        JOptionPane.showMessageDialog(null, "Please Ensure you Input the Plan Start Date", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else if (PensionPlanEndDate == null) {
                        JOptionPane.showMessageDialog(null, "Please Ensure you Input the Plan End Date", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else {
                        if (PensionPlanName == null && PensionPlanStartDate == null && PensionPlanEndDate == null) {
                            if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                                JOptionPane.showMessageDialog(null, "Please refresh the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                            } else {
                                JOptionPane.showMessageDialog(null, "Please ensure you Input the Plan Requiremenets Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                            }
                        } else {
                            JOptionPane.showMessageDialog(null, "Now Performing Pensionable Salary and Contributing Check, Press Ok to Continue", "Pensionable Check", JOptionPane.PLAIN_MESSAGE);

                            //  ArrayList al= validationChecks.Check_FivePercent_PS(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                            ArrayList al = validationChecks.check_FivePercent_PensionableSalary(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                            String result = validationChecks.getResult();
                            new AlToTable(al, "Check Pensionable Salary");
                            resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                        }
                    }
                } catch (IOException e1) {
                    e1.printStackTrace();
                }

            }
        }


        if (e.getSource().equals(CheckAge)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                //  if (planEntryAge != null) {
                if (PensionPlanName == null && PensionPlanStartDate == null && PensionPlanEndDate == null) {
                    if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please refresh the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure you Input the Plan Requiremenets Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                } else {
                    planEntryAge = JOptionPane.showInputDialog(this, "Please Input the Required Age it takes a Member to be eligible for the Plan", "Age", JOptionPane.PLAIN_MESSAGE);
                    int Age = Integer.parseInt(planEntryAge);
                    if (Age >= 18 && Age <= 70) {
                        JOptionPane.showMessageDialog(null, "Please Wait while the Members' Ages are Checked at their Plan Entry Date", "Age Check", JOptionPane.PLAIN_MESSAGE);
                        String result = null;
                        try {
                            ArrayList al = validationChecks.check_Age(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir, Age);
                            result = validationChecks.getResult();
                            new AlToTable(al, "Check Age");
                            resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                        } catch (IOException e1) {
                            e1.printStackTrace();
                            //     }
                        }
                    } else {
                        JOptionPane.showMessageDialog(null, "Important Notice: You have to be at least 18 to be Working in Jamaica", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                }
            }
        }


        if (e.getSource().equals(CheckDateofBirth)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                if (PensionPlanName == null && PensionPlanStartDate == null && PensionPlanEndDate == null) {
                    if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please refresh the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure you Input the Plan Requiremenets Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Now Performing Date of Birth Check, Press Ok to Continue", "Date of Birth Check", JOptionPane.PLAIN_MESSAGE);
                    String result = null;
                    try {
                        ArrayList al = validationChecks.check_DateofBirth(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                        result = validationChecks.getResult();
                        new AlToTable(al, "Check Date of Birth");
                        resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                    } catch (IOException e1) {
                        e1.printStackTrace();
                    }

                }
            }
        }

        if (e.getSource().equals(CheckDuplicate)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else if (!new File(filePathWorkingDir + "\\Valuation Data.xlsx").exists()) {
                JOptionPane.showMessageDialog(null, "Please ensure The Valuation Data is present in your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                if (PensionPlanName == null && PensionPlanStartDate == null && PensionPlanEndDate == null) {
                    if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please refresh the Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure you Input the Plan Requiremenets Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Now Performing Duplicate Check, Press Ok to Continue", "Duplicate Check", JOptionPane.PLAIN_MESSAGE);
                    String result = null;
                    try {
                        //  ArrayList<String > al = validationChecks.Check_For_Duplicates(filePathWorkingDir);
                        ArrayList<String> al = validationChecks.check_For_Duplicates(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                        result = validationChecks.getResult();
                        resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                        new AlToTable(al, "View Duplicates");
                    } catch (IOException e1) {
                        e1.printStackTrace();
                    }
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


        if (e.getSource().equals(MenuItemWorkingDir)) {
            filePathWorkingDir = utility.getWorkingDir();
            if (filePathWorkingDir != null) {
                JOptionPane.showMessageDialog(null, "You have set the working directory to: " + filePathWorkingDir, "Success", JOptionPane.PLAIN_MESSAGE);

                utility.writeDefaultsToFile("WD", filePathWorkingDir, filePathWorkingDir);

            }
        }

/*        if(e.getSource().equals(MenuItemCreateSeperatedTemplate)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please Ensure you Set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                TemplateSheets.Create_Template_Active_Terminee_Sheet(filePathWorkingDir);
                if (new File(filePathWorkingDir + "//Template_Separated.xlsx").exists()) {
                    JOptionPane.showMessageDialog(null, "The Template Workbook for Separated Active and Terminee Members was created Successfully", "Success", JOptionPane.PLAIN_MESSAGE);
                }
            }
        }*/
        //FEES
        if (e.getSource().equals(MenuItemCreateFeesActiveSheetTemplate)) {
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
                    TemplateSheets.Create_Template_Fees_Active_Sheet(PensionPlanStartDate, PensionPlanEndDate, PensionPlanName, filePathWorkingDir);
                    //    TemplateSheets.Create_Template_Active_Sheet(utility.readFile("SD",filePathWorkingDir), utility.readFile("ED",filePathWorkingDir), utility.readFile("PN",filePathWorkingDir), filePathWorkingDir);
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
                if (new File(filePathWorkingDir + "//Templates//Template_Active_Sheet.xlsx").exists()) {
                    JOptionPane.showMessageDialog(null, "The Active Sheet Template was created Successfully", "Success", JOptionPane.PLAIN_MESSAGE);
                }
            }
        }

        //NO FEES
        if (e.getSource().equals(MenuItemCreateNoFeesActiveSheetTemplate)) {
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
                if (new File(filePathWorkingDir + "//Templates//Template_Active_Sheet.xlsx").exists()) {
                    JOptionPane.showMessageDialog(null, "The Active Sheet Template was created Successfully", "Success", JOptionPane.PLAIN_MESSAGE);
                }
            }
        }


       //create template sheet for active members
        if (e.getSource().equals(MenuItemCreateNoFeesTermineeSheetTemplate)) {
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
                    TemplateSheets.Create_Template_Terminee_Sheet(PensionPlanStartDate, PensionPlanEndDate, PensionPlanName, filePathWorkingDir);
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
                if (new File(filePathWorkingDir + "//Templates//Template_Terminee_Sheet.xlsx").exists()) {
                    JOptionPane.showMessageDialog(null, "The Terminee Sheet  Template was created Successfully", "Success", JOptionPane.PLAIN_MESSAGE);
                }
            }//end
        }

        //create FEES template sheet for Terminee members
        if (e.getSource().equals(MenuItemCreateFeesTermineeSheetTemplate)) {
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
                    TemplateSheets.Create_Template_Fees_Terminee_Sheet(PensionPlanStartDate, PensionPlanEndDate, PensionPlanName, filePathWorkingDir);
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
                if (new File(filePathWorkingDir + "//Templates//Template_Terminee_Sheet.xlsx").exists()) {
                    JOptionPane.showMessageDialog(null, "The Terminee Sheet  Template was created Successfully", "Success", JOptionPane.PLAIN_MESSAGE);
                }
            }
        }//end


        if (e.getSource().equals(MenuItemCreateActiveSheet)) {
            String result = null;
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                if (new File(filePathWorkingDir + "\\Templates\\Template_Active_Sheet.xlsx ").exists()) {

                    if (PensionPlanEndDate != null && PensionPlanStartDate != null) {

                            excelReader.Write_Members_To_Active_Sheet(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                        try {
                            excelReader.Write_Members_Monetary_Values(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                            excelReader.Create_Active_Acc_Balances(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                               excelReader.WriteActivesTotalRow(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                        } catch (IOException e1) {
                            e1.printStackTrace();
                        }

                        JOptionPane.showMessageDialog(null, "The Active Sheet was created Successfully", "Success", JOptionPane.PLAIN_MESSAGE);
                    } else if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() || new File(filePathWorkingDir + "\\.txt").exists() || new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please Refresh The Plan Requirement Data, then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else if (!new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() || !new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() || !new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please Ensure you Input all The Plan Requirement Data, then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }

                } else {
                    JOptionPane.showMessageDialog(null, "Please Ensure you Create the Template Sheet for the Active Sheet", "Notice", JOptionPane.PLAIN_MESSAGE);
                }

            }
        }


        if (e.getSource().equals(MenuItemCreateFeesActiveSheet)) {
            String result = null;
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                if (new File(filePathWorkingDir + "\\Templates\\Template_Active_Sheet.xlsx ").exists()) {

                    if (PensionPlanEndDate != null && PensionPlanStartDate != null) {


                        try {
                            excelReader.Write_Members_To_Active_Sheet(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                            excelReader.Write_Members_Monetary_Fees_Values(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                           excelReader.Create_Fees_Active_Acc_Balances(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                           excelReader.WriteFeesActivesTotalRow(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                        } catch (IOException e1) {
                            e1.printStackTrace();
                        }
                        resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                        JOptionPane.showMessageDialog(null, "The Active Sheet was created Successfully", "Success", JOptionPane.PLAIN_MESSAGE);
                    } else if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() || new File(filePathWorkingDir + "\\.txt").exists() || new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please Refresh The Plan Requirement Data, then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else if (!new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() || !new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() || !new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please Ensure you Input all The Plan Requirement Data, then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }

                } else {
                    JOptionPane.showMessageDialog(null, "Please Ensure you Create the Template Sheet for the Active Sheet", "Notice", JOptionPane.PLAIN_MESSAGE);
                }

            }
        }


        if (e.getSource().equals(MenuItemCreateFeesTermineeSheet)) {
            String result = null;
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                if (new File(filePathWorkingDir + "\\Templates\\Template_Terminee_Sheet.xlsx ").exists()) {
                    if (PensionPlanEndDate != null && PensionPlanStartDate != null) {
                        result = excelReader.Create_Terminee_Sheet(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                        if (new File(filePathWorkingDir + "\\Terminees_Sheet.xlsx").exists()) {
                            try {
                             //   excelReader.Create_Fees_Terminee_Contribution(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                                excelReader.Write_Terminee_Members_Monetary_Fees_Values(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                               excelReader.Create_Fees_Terminee_Acc_Balances(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                            excelReader.WriteFeesTermineeTotalRow(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                            } catch (IOException e1) {
                                e1.printStackTrace();
                            }
                        //    resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                            JOptionPane.showMessageDialog(null, "The Terminee Sheet was created Successfully", "Success", JOptionPane.PLAIN_MESSAGE);
                        } else {
                            JOptionPane.showMessageDialog(null, "Please Ensure the Date of Refunds were Inputted for the Terminated Members", "Notice", JOptionPane.PLAIN_MESSAGE);
                        }
                    } else if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() || new File(filePathWorkingDir + "\\.txt").exists() || new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please Refresh The Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else if (!new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() || !new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() || !new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please Ensure you Input all The Plan Requirement Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }

                } else {
                    JOptionPane.showMessageDialog(null, "Please Ensure you Create the Template Sheet for the Terminee Sheet", "Notice", JOptionPane.PLAIN_MESSAGE);
                }

            }
        }


        if (e.getSource().equals(MenuItemCreateTermineeSheet)) {
            String result = null;
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                if (new File(filePathWorkingDir + "\\Templates\\Template_Terminee_Sheet.xlsx ").exists()) {
                    if (PensionPlanEndDate != null && PensionPlanStartDate != null) {
                        result = excelReader.Create_Terminee_Sheet(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                        if (new File(filePathWorkingDir + "\\Terminees_Sheet.xlsx").exists()) {
                            try {
                              //  excelReader.Create_Terminee_Contribution(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                                excelReader.Write_Terminee_Members_Monetary_Values(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                                excelReader.Create_Terminee_Acc_Balances(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                                excelReader.WriteTermineeTotalRow(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                            } catch (IOException e1) {
                                e1.printStackTrace();
                            }
                    //        resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                            JOptionPane.showMessageDialog(null, "The Terminee Sheet was created Successfully", "Success", JOptionPane.PLAIN_MESSAGE);
                        } else {
                            JOptionPane.showMessageDialog(null, "Please Ensure the Date of Refunds were Inputted for the Terminated Members", "Notice", JOptionPane.PLAIN_MESSAGE);
                        }
                    } else if (new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() || new File(filePathWorkingDir + "\\.txt").exists() || new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please Refresh The Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                    } else if (!new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists() || !new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() || !new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                        JOptionPane.showMessageDialog(null, "Please Ensure you Input all The Plan Requirement Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
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
                if(PensionPlanStartDate==null || PensionPlanEndDate == null || PensionPlanName==null) {
                    JOptionPane.showMessageDialog(null, "Please Refresh The Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                }else{
                if (new File(filePathWorkingDir + "\\Input Sheet.xlsx").exists()) {

                    ArrayList<String> al = null;
                    try {
                        al = excelReader.View_Actives_Members(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                    } catch (IOException e1) {
                        e1.printStackTrace();
                    }
                    if (!al.isEmpty()) {
                        String result = excelReader.getResult();
                        JOptionPane.showMessageDialog(null, "Please wait for the list of the Active Members as at " + PensionPlanEndDate, "Success", JOptionPane.PLAIN_MESSAGE);
                        resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);

                        new AlToTable(al, "View Active");
                    } else {
                        JOptionPane.showMessageDialog(null, "Based on your Query, there were no Retired Members found as at " + PensionPlanEndDate, "Success", JOptionPane.PLAIN_MESSAGE);
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Please ensure the Input Sheet is present in your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
                }
            }

        }
        }

        if (e.getSource().equals(MenuItemViewTermineeMember)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                if (PensionPlanStartDate == null || PensionPlanEndDate == null || PensionPlanName == null) {
                    JOptionPane.showMessageDialog(null, "Please Refresh The Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                } else {
                    if (new File(filePathWorkingDir + "\\Input Sheet.xlsx").exists()) {
                        ArrayList<String> al = excelReader.View_Terminee_Members(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                        String result = excelReader.getResult();
                        if (!al.isEmpty()) {
                            JOptionPane.showMessageDialog(null, "Please wait for the list of the Terminee Members as at " + PensionPlanEndDate, "Success", JOptionPane.PLAIN_MESSAGE);
                            resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                            new AlToTable(al, "View Terminee");
                        } else {
                            JOptionPane.showMessageDialog(null, "Based on your Query, there were no Retired Members found as at " + PensionPlanEndDate, "Success", JOptionPane.PLAIN_MESSAGE);
                        }
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure the Input Sheet is present in your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                }
            }
        }

        if (e.getSource().equals(MenuItemViewTerminatedMember)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
             //   if (!new File(filePathWorkingDir + "\\Program Files\\PN.txt").exists()&&!new File(filePathWorkingDir + "\\Program Files\\SD.txt").exists() && new File(filePathWorkingDir + "\\Program Files\\ED.txt").exists()) {
                if (PensionPlanStartDate == null || PensionPlanEndDate == null || PensionPlanName == null) {
                    JOptionPane.showMessageDialog(null, "Please Refresh The Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                } else {
                    if (new File(filePathWorkingDir + "\\Input Sheet.xlsx").exists()) {

                        ArrayList<String> al = excelReader.View_Terminated_Members(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                        String result = excelReader.getResult();
                        if (!al.isEmpty()) {
                            JOptionPane.showMessageDialog(null, "Please wait for the list of the Terminated Members as at " + PensionPlanEndDate, "Success", JOptionPane.PLAIN_MESSAGE);
                            resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                            new AlToTable(al, "View Terminated");
                        } else {
                            JOptionPane.showMessageDialog(null, "Based on your Query, there were no Terminated Members found as at " + PensionPlanEndDate, "Success", JOptionPane.PLAIN_MESSAGE);
                        }
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure the Input Sheet is Present in the Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                }
            }
        }

     /*   if (e.getSource().equals(MenuItemViewAllMembers)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {

                if (new File(filePathWorkingDir + "\\Seperated Members.xlsx").exists()) {

                   // String result = excelReader.View_Actives_Members(filePathWorkingDir, PensionPlanEndDate);
                //    result += excelReader.View_Terminee_Members(filePathWorkingDir, PensionPlanEndDate);
                    JOptionPane.showMessageDialog(null, "Please wait for the list of all the Members as at " + PensionPlanEndDate, "Success", JOptionPane.PLAIN_MESSAGE);
                 //   resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                } else {
                    JOptionPane.showMessageDialog(null, "Please ensure you Create the Workbook with Active and Terminee Members separated", "Notice", JOptionPane.PLAIN_MESSAGE);
                }
            }
        }*/


        if (e.getSource().equals(MenuItemViewRetiredMember)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                if (PensionPlanStartDate == null || PensionPlanEndDate == null || PensionPlanName == null) {
                    JOptionPane.showMessageDialog(null, "Please Refresh The Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                } else {
                    if (new File(filePathWorkingDir + "\\Input Sheet.xlsx").exists()) {

                        ArrayList al = excelReader.View_Retired_Members(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                        String result = excelReader.getResult();

                        if (!al.isEmpty()) {
                            JOptionPane.showMessageDialog(null, "Please wait for the list of the Retired Members as at " + PensionPlanEndDate, "Success", JOptionPane.PLAIN_MESSAGE);
                            new AlToTable(al, "View Retired");
                            resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                        } else {
                            JOptionPane.showMessageDialog(null, "Based on your Query, there were no Retired Members found as at " + PensionPlanEndDate, "Success", JOptionPane.PLAIN_MESSAGE);
                        }
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure the Input Sheet is present in the Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                }
            }
        }

        if (e.getSource().equals(MenuItemViewDeceasedMember)) {
            if (filePathWorkingDir == null) {
                JOptionPane.showMessageDialog(null, "Please ensure you set your Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
            } else {
                if (PensionPlanStartDate == null || PensionPlanEndDate == null || PensionPlanName == null) {
                    JOptionPane.showMessageDialog(null, "Please Refresh The Plan Requirements Data, Then try again", "Notice", JOptionPane.PLAIN_MESSAGE);
                } else {
                    if (new File(filePathWorkingDir + "\\Input Sheet.xlsx").exists()) {
                        ArrayList al = excelReader.View_Deceased_Members(PensionPlanStartDate, PensionPlanEndDate, filePathWorkingDir);
                        String result = excelReader.getResult();
                        if (!al.isEmpty()) {
                            JOptionPane.showMessageDialog(null, "Please wait for the list of the Deceased Members as at " + PensionPlanEndDate, "Success", JOptionPane.PLAIN_MESSAGE);
                            resultsWindow.appendToPane(resultsWindow, result + "\n", LINES, true);
                            new AlToTable(al, "View Deceased");
                        } else {
                            JOptionPane.showMessageDialog(null, "Based on your Query, there were no Deceased Members found as at " + PensionPlanEndDate, "Success", JOptionPane.PLAIN_MESSAGE);
                        }
                    } else {
                        JOptionPane.showMessageDialog(null, "Please ensure the Input Sheet is present in the Working Directory", "Notice", JOptionPane.PLAIN_MESSAGE);
                    }
                }
            }
        }
    }

    private void registerListener() {
        MenuItemCreateTableSummaryofActiveMembership.addActionListener(this);
        MenuItemCreateIncExpTable.addActionListener(this);
        MenuItemCreateBalSheetTemplate.addActionListener(this);

      //  MenuItemCreateTableSummaryofActiveMembership.addActionListener(this);
        MenuItemCreateTableMovementsinActiveMemberships.addActionListener(this);
        MenuItemCreateTableAnalysisofFundYield.addActionListener(this);
        MenuItemCreateTableGainsLosses.addActionListener(this);


        //MEMBERS LISTENERS
        MenuItemCreateIncExpTemplate.addActionListener(this);
        //MenuItemCreateSeperatedTemplate.addActionListener(this);
        MenuItemViewActiveMember.addActionListener(this);
        MenuItemViewTermineeMember.addActionListener(this);
        MenuItemViewAllMembers.addActionListener(this);
        MenuItemViewRetiredMember.addActionListener(this);
        MenuItemViewDeceasedMember.addActionListener(this);
     //   MenuItemViewDeferredMember.addActionListener(this);
        MenuItemViewTerminatedMember.addActionListener(this);

        MenuItemClearScreen.addActionListener(this);
        MenuItemRefreshData.addActionListener(this);
        MenuItemViewPlanDeatils.addActionListener(this);
        MenuItemSaveData.addActionListener(this);
        MenuItemDeleteData.addActionListener(this);

        CheckDuplicate.addActionListener(this);
        CheckAge.addActionListener(this);
        CheckDateofBirth.addActionListener(this);
        CheckEmployeePS.addActionListener(this);
        CheckPlanEntry.addActionListener(this);
        CheckAll.addActionListener(this);

//        LoadValDataWorkBook.addActionListener(this);
    //    MenuItemSeperateMembers.addActionListener(this);
        MenuItemPensionPlanName.addActionListener(this);
        MenuItemStartDate.addActionListener(this);
        MenuItemEndDate.addActionListener(this);
        MenuItemLoadOutputTemplate.addActionListener(this);
        MenuItemWorkingDir.addActionListener(this);

        MenuItemCreateNoFeesActiveSheetTemplate.addActionListener(this);
        MenuItemCreateNoFeesTermineeSheetTemplate.addActionListener(this);

        MenuItemCreateFeesActiveSheetTemplate.addActionListener(this);
        MenuItemCreateFeesTermineeSheetTemplate.addActionListener(this);

        MenuItemCreateActiveSheet.addActionListener(this);
        MenuItemCreateTermineeSheet.addActionListener(this);

        MenuItemCreateFeesActiveSheet.addActionListener(this);
        MenuItemCreateFeesTermineeSheet.addActionListener(this);

        MenuItemInterestRates.addActionListener(this);
        MenuItemCreateBalSheetTable.addActionListener(this);

        MenuItemCreateTemplateSummaryofActiveMembership.addActionListener(this);
        MenuItemCreateTemplateMovementsinActiveMemberships.addActionListener(this);
        MenuItemCreateTemplateAnalysisofFundYield.addActionListener(this);
        MenuItemCreateTemplateGainsLosses.addActionListener(this);
    }

}


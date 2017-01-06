import javax.swing.*;
import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.util.*;
import javax.swing.table.*;

public class AlToTable extends JFrame {

    private final static String[] header = {"Employee ID", "Last Name", "First Name","DOB","Employment Date","Plan Entry Date","Status Date","Status"};

    AlToTable(ArrayList<String> al) {
        super("List of Active Members");
        MyModel mm = new MyModel(al, header);
        JTable table = new JTable(mm);
        add(new JScrollPane(table));
        setSize(800, 600);
        setVisible(true);
        setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        JTableHeader anHeader = table.getTableHeader();
        anHeader.setForeground(new Color(0).black);
        anHeader.setBackground(new Color(0).gray);
        anHeader.setFont(new Font("Serif", Font.BOLD, 18));
        table.setFont(new Font("San-Serif", Font.PLAIN, 16));
     //   table.setRowHeight(table.getRowHeight());

      table.addMouseListener(new MouseAdapter() {
            public void mouseClicked(MouseEvent e) {
                if (e.getClickCount() == 2) { // check if a double click
                    // your code here
                //  String val=  this.getPropertyFromRow((String)(table.getValueAt(table.getSelectedRow(),0)));
                //    String remarks = table.getValueAt(getSelectedRow(), 3).toString();
                    JOptionPane.showMessageDialog(null, "Please ensure you Create the Workbook with Active and Terminee Members separated", "Notice", JOptionPane.PLAIN_MESSAGE);

                }
            }
        });
    }

    AlToTable(ArrayList<String> al,String type) {
        //  super("List of Terminee Members");
        if(type.equals("View Terminee"))
            setTitle("List of Terminee Members");
        else if(type.equals("View Deferred"))
            setTitle("List of Deferred Members");
        else if(type.equals("View Terminated"))
            setTitle("List of Terminated Members");
        else if(type.equals("View Deceased"))
            setTitle("List of Deceased Members");
        else if(type.equals("View Retired"))
            setTitle("List of Retired Members");
        else if(type.equals("View Active"))
            setTitle("List of Active Members");
        else if(type.equals("View Duplicates"))
        {
            setTitle("List of Duplicate Members");
       //  header = new String[]{"Employee ID", "Last Name", "First Name", "DOB"};
          //  MyModel mm = new MyModel(al, header);
        }


        MyModel mm = new MyModel(al, header);
        JTable table = new JTable(mm);
        add(new JScrollPane(table));
        setSize(800, 600);
        setVisible(true);
        setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        JTableHeader anHeader = table.getTableHeader();
        anHeader.setForeground(new Color(0).black);
        anHeader.setBackground(new Color(0).gray);
        anHeader.setFont(new Font("Serif", Font.BOLD, 18));
        table.setFont(new Font("San-Serif", Font.PLAIN, 16));
        //   table.setRowHeight(table.getRowHeight());

    }


    class MyModel extends AbstractTableModel {

        private ArrayList<String> al;
        private String[] header;

        MyModel(ArrayList<String> al, String[] header) {
            this.al = al;
            this.header = header;
        }

        public int getColumnCount() {
            return header.length;
        }

        public int getRowCount() {
            return al.size();
        }

        public Object getValueAt(int rowIndex, int columnIndex) {
            String[] token = al.get(rowIndex).split(",");
            return token[columnIndex];
        }

        public String getColumnName(int col) {
            return header[col];
        }

    }

/*    public static void main(String[] args) {
        ArrayList<String> al = new ArrayList<String>();
        al.add("PBL,59,M,hbf");
        al.add("Madona,20,F,ff");
        al.add("teQuiero,???,M,vfv");
        new AlToTable(al);
    }*/
}

import java.awt.*;
import javax.swing.*;
import javax.swing.text.*;

public class ResultsWindow extends JTextPane {    //setting GUI properties such as colors and borders

    public ResultsWindow () {
        setBackground(new Color(255, 255, 243));
        setForeground(new Color(86, 98, 112));
        setBorder(javax.swing.BorderFactory.createEmptyBorder(5, 8, 5, 8));

   //     set
    }

    public void appendToPane (JTextPane tp, String msg, Color c, boolean isBold) { //append attributes and properties to pane
        setComponentOrientation(ComponentOrientation.LEFT_TO_RIGHT);
       StyledDocument doc = tp.getStyledDocument();
        SimpleAttributeSet center = new SimpleAttributeSet();
        StyleConstants.setAlignment(center, StyleConstants.ALIGN_CENTER);
        doc.setParagraphAttributes(0, doc.getLength(), center, false);

        StyleContext sc = StyleContext.getDefaultStyleContext();
        AttributeSet aSet = sc.addAttribute(SimpleAttributeSet.EMPTY, StyleConstants.Foreground, c);

        aSet = sc.addAttribute(aSet, StyleConstants.FontFamily, "serif"); //Consales font will be used
        aSet = sc.addAttribute(aSet, StyleConstants.Alignment, StyleConstants.ALIGN_JUSTIFIED);
        aSet = sc.addAttribute(aSet, StyleConstants.FontSize, 21);    //font size of 23 will be used

        if (isBold) aSet = sc.addAttribute(aSet, StyleConstants.Bold, Boolean.TRUE);   //text will be bold
        else aSet = sc.addAttribute(aSet, StyleConstants.Bold, Boolean.FALSE);
        //set correct length of characters on window
        int len = tp.getDocument().getLength();
        tp.setCaretPosition(len);
        tp.setCharacterAttributes(aSet, false);
      // tp.replaceSelection(msg+"\n");
       // tp.insertIcon(new ImageIcon("C:\\Users\\akonowalchuk\\GFRAM\\page1.png"));
        tp.setText(msg+"\n");
      //  tp.setEditable(false);

    }

    public void ClearScreen(JTextPane tp){
       tp.setText("");
    }

}
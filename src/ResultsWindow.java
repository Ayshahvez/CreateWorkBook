import java.awt.Color;
import javax.swing.JTextPane;
import javax.swing.text.*;

public class ResultsWindow extends JTextPane {    //setting GUI properties such as colors and borders
    public ResultsWindow () {
        setBackground(new Color(255, 255, 243));
        setForeground(new Color(86, 98, 112));
        setBorder(javax.swing.BorderFactory.createEmptyBorder(5, 8, 5, 8));
    }

    public void appendToPane (JTextPane tp, String msg, Color c, boolean isBold) { //append attributes and properties to pane
        StyleContext sc = StyleContext.getDefaultStyleContext();
        AttributeSet aSet = sc.addAttribute(SimpleAttributeSet.EMPTY, StyleConstants.Foreground, c);

        aSet = sc.addAttribute(aSet, StyleConstants.FontFamily, "Consolas"); //Consales font will be used
        aSet = sc.addAttribute(aSet, StyleConstants.Alignment, StyleConstants.ALIGN_JUSTIFIED);
        aSet = sc.addAttribute(aSet, StyleConstants.FontSize, 21);    //font size of 21 will be used

        if (isBold) aSet = sc.addAttribute(aSet, StyleConstants.Bold, Boolean.TRUE);   //text will be bold
        else aSet = sc.addAttribute(aSet, StyleConstants.Bold, Boolean.FALSE);
        //set correct length of characters on window
        int len = tp.getDocument().getLength();
        tp.setCaretPosition(len);
        tp.setCharacterAttributes(aSet, false);
     //   tp.replaceSelection(" ");
        tp.replaceSelection(msg+"\n");

    }
}
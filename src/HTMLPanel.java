import javax.swing.*;
import javax.swing.text.html.HTMLEditorKit;
import javax.swing.text.html.StyleSheet;
import java.awt.*;
import java.net.URL;

public class HTMLPanel extends JPanel {
    private HTMLEditorKit kit;
    private JEditorPane HTMLPane = new JEditorPane();

    public HTMLPanel(){
        kit = new HTMLEditorKit();
        HTMLPane.setEditable(false);
        JScrollPane htmlScrollPane = new JScrollPane(HTMLPane);
        HTMLEditorKit kit = new HTMLEditorKit();
        HTMLPane.setEditorKit(kit);
        setLayout(new BorderLayout(1,1));
        add(htmlScrollPane, BorderLayout.CENTER);
    }

    protected void setText(String text){
        HTMLPane.setText(text);
    }

    public String getText(){
        return HTMLPane.getText();
    }

    protected void setStyleSheet(URL url){
        StyleSheet ss = new StyleSheet();
        ss.importStyleSheet(url);
        HTMLEditorKit kit = (HTMLEditorKit)HTMLPane.getEditorKit();
        kit.setStyleSheet(ss);
        HTMLPane.setEditorKit(kit);
    }
}

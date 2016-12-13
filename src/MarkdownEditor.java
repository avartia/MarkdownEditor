/** TODO:  A left side panel shows the structure of the document. Structure is defined as <h?>.
 TODO: An click on the item in the structure panel brings that <h?> at the first line in the editor panel.
 TODO: It is able to deal with large document, up to 1MB bytes.
 TODO: HTML tags are recognizable.
 TODO: User can specify his own .css to generate the HTML output.
 TODO: It is able to generate .docx file.
 */

import javafx.util.Pair;
import org.markdown4j.Markdown4jProcessor;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import javax.swing.*;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.BadLocationException;
import javax.swing.text.DefaultHighlighter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.net.URL;
import java.util.HashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


public class MarkdownEditor extends JFrame{
    private JTextArea Editor = new JTextArea();
    // private JTextArea Structure = new JTextArea();
    // private JEditorPane HTMLPane = new JEditorPane();
    private HTMLPanel PreviewPanel;
    private Markdown4jProcessor processor = new Markdown4jProcessor();
    // html cache
    private String htmlStringCache;
    // markdown cache
    private String mdStringCache;
    // opening markdown filename
    private String mdFile = "";
    private String mdPath = "";
    // css string
    private String cssStr = "";
    // list string
    // structure index -- Pair(start line, end line)
    private HashMap<Integer, Pair<Integer, Integer> > structList = new HashMap<>();
    private JPanel StructurePanel;
    private JList jlst;
    // list model for JList
    private DefaultListModel listModel;
    // auto preview setting
    private boolean autoPreview = false;
    // file origin data
    private String originData = "";
    public MarkdownEditor(){
        JMenuBar menuBar = new JMenuBar();
        setJMenuBar(menuBar);

        // create the File menu
        JMenu fileMenu = new JMenu("File");
        menuBar.add(fileMenu);

        // open a md file
        JMenuItem openItem = new JMenuItem("Open");
        fileMenu.add(openItem);
        openItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                OpenFile();
            }
        });

        // new a md file
        JMenuItem newItem = new JMenuItem("new");
        fileMenu.add(openItem);
        newItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                CloseFile();
            }
        });
        // save a md file
        JMenuItem saveItem = new JMenuItem("Save MD");
        fileMenu.add(saveItem);
        saveItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                SaveMDFile();
            }
        });

        // import a css file
        JMenuItem CSSItem = new JMenuItem("import css");
        fileMenu.add(CSSItem);
        CSSItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                ImportCSS();
            }
        });
        // save a HTML file
        JMenuItem saveHTMLItem = new JMenuItem("Save as HTML");
        fileMenu.add(saveHTMLItem);
        saveHTMLItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                SaveAsHTML();
            }
        });

        // save a DOCX file
        JMenuItem saveDOCXItem = new JMenuItem("Save as docx");
        fileMenu.add(saveDOCXItem);
        saveDOCXItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                SaveAsDOCX();
            }
        });


        JMenuItem closeItem = new JMenuItem("Close");
        fileMenu.add(closeItem);
        closeItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                CloseFile();
            }
        });
        // create the run menu
        JMenu runMenu = new JMenu("Run");
        menuBar.add(runMenu);

        JMenuItem runHTML = new JMenuItem("Preview in HTML");
        runMenu.add(runHTML);
        runHTML.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                ShowHTML();
            }
        });

        runMenu.add(runHTML);
        runHTML.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                ShowHTML();
            }
        });

        // add a toolBar
        JToolBar toolBar = new JToolBar("toolBar");


        // open btn
        JButton Btn = new JButton(new ImageIcon("img/open.png"));Btn.setToolTipText("Open a md file");Btn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                OpenFile();
            }
        });
        toolBar.add(Btn);

        // new btn
        Btn = new JButton( new ImageIcon("img/new.png"));
        Btn.setToolTipText("new file");
        Btn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                CloseFile();
            }
        });
        toolBar.add(Btn);
        // save btn
        Btn = new JButton( new ImageIcon("img/save.png"));Btn.setToolTipText("Save a md file");Btn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                SaveMDFile();
            }
        });
        toolBar.add(Btn);

        // preview btn
        Btn = new JButton( new ImageIcon("img/preview.png"));
        Btn.setToolTipText("Preview in html");
        Btn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                ShowHTML();
            }
        });
        toolBar.add(Btn);

        // close btn
        Btn = new JButton( new ImageIcon("img/close.png"));
        Btn.setToolTipText("close file");
        Btn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                CloseFile();
            }
        });
        toolBar.add(Btn);
        // structure panel
        StructurePanel = new JPanel();
        listModel = new DefaultListModel();
        jlst = new JList(listModel);
        listModel.addElement("structure <h?> here...");
        jlst.addListSelectionListener(new ListSelectionListener() {
            @Override
            public void valueChanged(ListSelectionEvent e) {
                try {
                    int[] indices = jlst.getSelectedIndices();
                    Pair<Integer, Integer> posPair = null;
                    if(indices.length != 0)
                        posPair = structList.get(indices[0]);
                    if(posPair != null) {
                        int startPos = posPair.getKey();
                        int endPos = posPair.getValue();
                        Editor.setCaretPosition(startPos);
                        // background highlighter
                        DefaultHighlighter BackHighlighter = (DefaultHighlighter) Editor.getHighlighter();
                        DefaultHighlighter.DefaultHighlightPainter BackPainter = new DefaultHighlighter.DefaultHighlightPainter(new Color(168, 179, 154));
                        // clear highlight
                        BackHighlighter.removeAllHighlights();
                        // set line backgroud highlight
                        BackHighlighter.addHighlight(startPos, endPos, BackPainter);
                        Editor.setCaretPosition(startPos);
                    }
                }catch (BadLocationException be){

                }
            }
        });
        JScrollPane StructureScrollPane = new JScrollPane(jlst);
        StructurePanel.setLayout(new BorderLayout(1,1));
        StructurePanel.add(StructureScrollPane, BorderLayout.CENTER);

        // editor panel
        JPanel EditorPanel = new JPanel();
        Editor.setText("Enter your markdown code here...");
        Editor.setFont(new Font("Serif", Font.PLAIN, 14));
        JScrollPane EditorScrollPane = new JScrollPane(Editor);
        EditorPanel.setLayout(new BorderLayout(1, 1));
        EditorPanel.add(EditorScrollPane, BorderLayout.CENTER);

        // text panel
        JPanel textPanel = new JPanel();
        textPanel.setLayout(new GridLayout(1, 2, 5, 5));

        // HTML preview panel
        PreviewPanel = new HTMLPanel();

        // set layout
        setLayout(new BorderLayout(5,5));
        add(toolBar, BorderLayout.PAGE_START);
        add(StructurePanel, BorderLayout.WEST);
        textPanel.add(EditorPanel);
        textPanel.add(PreviewPanel);
        add(textPanel, BorderLayout.CENTER);
    }

    public static void main(String[] args){
        MarkdownEditor frame = new MarkdownEditor();
        frame.pack();
        frame.setTitle("MarkdownEditor");
        frame.setSize(800, 600);
        frame.setLocationRelativeTo(null);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setVisible(true);
    }
    // translate md into html
    private void ConvertTOHTML(){
        try {
            mdStringCache = Editor.getText();
            htmlStringCache = processor.process(mdStringCache);
        }catch (IOException e){

        }
    }

    // use regex to get header
    private void GetHeader(){
        structList.clear();
        listModel.clear();
        String regex = "(#{1,5} +[^\\n]+\\n)|([^\\n]*\\n[-=]+[ \\f\\r\\t\\v]*\\n)|([^\\n]*<h\\d>[^\\n]+\\n)";
        Pattern p = Pattern.compile(regex);
        Matcher m = p.matcher(mdStringCache);
        int startIndex = 0, endIndex = 0;
        try {
            while (m.find()) {
                String parseResult = processor.process(mdStringCache.substring(m.start(), m.end()));
                if(((startIndex = parseResult.indexOf("<h1>"))!=-1) && ((endIndex = parseResult.indexOf("</h1>"))!=-1)){
                    structList.put(listModel.size(), new Pair<>(m.start(), m.end()));
                    listModel.addElement("" + parseResult.substring(startIndex+4, endIndex));
                }else if(((startIndex = parseResult.indexOf("<h2>"))!=-1) && ((endIndex = parseResult.indexOf("</h2>"))!=-1)){
                    structList.put(listModel.size(), new Pair<>(m.start(), m.end()));
                    listModel.addElement("--" + parseResult.substring(startIndex+4, endIndex));
                }else if(((startIndex = parseResult.indexOf("<h3>"))!=-1) && ((endIndex = parseResult.indexOf("</h3>"))!=-1)){
                    structList.put(listModel.size(), new Pair<>(m.start(), m.end()));
                    listModel.addElement("----" + parseResult.substring(startIndex+4, endIndex));
                }else if(((startIndex = parseResult.indexOf("<h4>"))!=-1) && ((endIndex = parseResult.indexOf("</h4>"))!=-1)){
                    structList.put(listModel.size(), new Pair<>(m.start(), m.end()));
                    listModel.addElement("------" + parseResult.substring(startIndex+4, endIndex));
                }else if(((startIndex = parseResult.indexOf("<h5>"))!=-1) && ((endIndex = parseResult.indexOf("</h5>"))!=-1)){
                    structList.put(listModel.size(), new Pair<>(m.start(), m.end()));
                    listModel.addElement("--------" + parseResult.substring(startIndex+4, endIndex));
                }
            }
            // update structlist
            jlst.updateUI();
        }catch (Exception e){
            e.printStackTrace();
        }

    }
    // show html
    private void ShowHTML(){
        try {
            ConvertTOHTML();
            PreviewPanel.setText(htmlStringCache);
            GetHeader();
        }catch (Exception e){

        }
    }

    // open file
    private void OpenFile(){
        if(Editor.getText() != originData){
            int dialogButton = JOptionPane.YES_NO_CANCEL_OPTION;
            int toSave = JOptionPane.showConfirmDialog(this, "Do you want to save this file?", "warning", dialogButton);
            if(toSave == JOptionPane.YES_OPTION){
                SaveMDFile();
            }else if(toSave == JOptionPane.CANCEL_OPTION){
                return;
            }
        }
        JFileChooser fc = new JFileChooser("examples");
        FileNameExtensionFilter fef = new FileNameExtensionFilter("Markdown text file", "md");
        fc.setFileFilter(fef);

        int returnVal = fc.showOpenDialog(this);

        // open md file
        if(returnVal == JFileChooser.APPROVE_OPTION){
            // save file into mdStringCache
            File file = fc.getSelectedFile();
            try {
                // up to 2mb file
                byte[] byteCache = new byte[2048*1024];
                FileInputStream fileIn = new FileInputStream(file);
                // read into buffer cache
                int bytes = fileIn.read(byteCache);
                // set markdown text
                if(bytes >= 0)
                    Editor.setText( mdStringCache = new String(byteCache, 0, bytes));
                originData = mdStringCache;
                mdFile = file.getName();
                mdPath = file.getPath();
            }catch (IOException e){

            }
        }
    }

    // save markdown file
    private void SaveMDFile(){
        try {
            // if has not been written before
            if (mdFile == "") {
                JFileChooser fc = new JFileChooser("examples");
                FileNameExtensionFilter fef = new FileNameExtensionFilter("Markdown text file", "md");
                fc.setFileFilter(fef);
                int returnVal = fc.showSaveDialog(this);

                // open md file
                if (returnVal == JFileChooser.APPROVE_OPTION) {
                    // save file into mdStringCache
                    File file = fc.getSelectedFile();
                    WriteToFile(file);
                }
            } else {
                WriteToFile(new File(mdPath));
            }
        }catch (IOException e) {

        }
    }

    // save html file
    private void SaveAsHTML(){
        try {
            JFileChooser fc = new JFileChooser("examples");
            FileNameExtensionFilter fef = new FileNameExtensionFilter("html file", "html");
            fc.setFileFilter(fef);
            int returnVal = fc.showSaveDialog(this);

            if (returnVal == JFileChooser.APPROVE_OPTION) {
                File file = fc.getSelectedFile();
                FileOutputStream fileOut = new FileOutputStream(file);
                // read into buffer cache
                ConvertTOHTML();
                // add css into html
                String outHTML = "<html>\n" + "<head>\n" + "<style type=\"text/css\">\n"
                        + cssStr  + "\n" + "</style>\n" + "</head>\n" + "<body>"+
                        htmlStringCache + "</body>" + "</html>\n";
                fileOut.write(outHTML.getBytes());
            }
        }catch (IOException e) {

        }
    }

    // save docx file
    private void SaveAsDOCX(){
        try {
            JFileChooser fc = new JFileChooser("examples");
            FileNameExtensionFilter fef = new FileNameExtensionFilter("doc/docx word file", "doc", "docx");
            fc.setFileFilter(fef);
            int returnVal = fc.showSaveDialog(this);
            // convert to word
            if (returnVal == JFileChooser.APPROVE_OPTION) {
                File file = fc.getSelectedFile();
                // read into buffer cache
                ConvertTOHTML();
                // add css into html
                String outDOCX = "<html>\n" + "<head>\n" + "<style type=\"text/css\">\n"
                        + cssStr  + "\n" + "</style>\n" + "</head>\n" + "<body>"+
                        htmlStringCache + "</body>" + "</html>\n";
                if(file.getPath().endsWith(".doc")){
                    FileOutputStream fileOut = new FileOutputStream(file);
                    fileOut.write(outDOCX.getBytes());
                }else if(file.getPath().endsWith(".docx")){
                    File tmpFile = new File("tmp.html");
                    FileOutputStream fileOut = new FileOutputStream(tmpFile);
                    fileOut.write(outDOCX.getBytes());
                    fileOut.close();
                    HtmlToWord(tmpFile.getAbsolutePath(), file.getPath());
                    tmpFile.delete();
                }
            }
        }catch (Exception e) {

        }
    }


    // HTML to word
    public static void HtmlToWord(String html, String wordFile) {

        ActiveXComponent app = new ActiveXComponent("Word.Application"); // 启动word

        try {

            app.setProperty("Visible", new Variant(false));

            Dispatch wordDoc = app.getProperty("Documents").toDispatch();

            wordDoc = Dispatch.invoke(wordDoc, "Add", Dispatch.Method, new Object[0], new int[1]).toDispatch();

            Dispatch.invoke(app.getProperty("Selection").toDispatch(), "InsertFile", Dispatch.Method, new Object[] { html, "", new Variant(false), new Variant(false), new Variant(false) }, new int[3]);

            Dispatch.call(wordDoc, "SaveAs", wordFile);

            Dispatch.call(wordDoc, "Close", new Variant(false));

        } catch (Exception e) {

        } finally {
            app.invoke("Quit", new Variant[] {});
        }
    }
    // write markdown code to file
    private void WriteToFile(File file) throws IOException{
        FileOutputStream fileOut = new FileOutputStream(file);
        // read into buffer cache
        ConvertTOHTML();
        fileOut.write(mdStringCache.getBytes());
    }

    // close markdown file
    private void CloseFile(){
        if(Editor.getText() != originData){
            int dialogButton = JOptionPane.YES_NO_CANCEL_OPTION;
            int toSave = JOptionPane.showConfirmDialog(this, "Do you want to save this file?", "warning", dialogButton);
            if(toSave == JOptionPane.YES_OPTION){
                SaveMDFile();
            }else if(toSave == JOptionPane.CANCEL_OPTION){
                return;
            }
        }
        htmlStringCache = "";
        mdFile = "";
        mdPath = "";
        Editor.setText("");
        ShowHTML();
    }

    // import a css file
    private void ImportCSS(){
        try {
            JFileChooser fc = new JFileChooser("examples");
            FileNameExtensionFilter fef = new FileNameExtensionFilter("css file", "css");
            fc.setFileFilter(fef);
            int returnVal = fc.showOpenDialog(this);

            if (returnVal == JFileChooser.APPROVE_OPTION) {
                File file = fc.getSelectedFile();
                URL url = file.toURI().toURL();
                if (url != null) {
                    FileInputStream fileIn = new FileInputStream(file);
                    byte[] cssCache = new byte[10240];
                    int bytes = fileIn.read(cssCache);
                    if(bytes > 0){
                        // set css string
                        cssStr = new String(cssCache, 0, bytes);
                    }
                    PreviewPanel.setStyleSheet(url);
                    ShowHTML();
                }
            }
        }catch (Exception e){
        }
    }
}
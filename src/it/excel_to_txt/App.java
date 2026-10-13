package it.excel_to_txt;

import javax.swing.*;
import javax.swing.border.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.awt.dnd.*;
import java.awt.event.*;
import java.io.File;
import java.util.*;
import java.util.List;
import java.util.concurrent.ExecutionException;

/** Interfaccia Swing completa con estrazione colonne e join Excel opzionale. */
public class App extends JFrame {
    private static final float SF;
    static { SF = Math.max(1.0f, Toolkit.getDefaultToolkit().getScreenResolution()/96.0f); }
    private static int sc(int v){ return Math.round(v*SF); }
    private static Font font(int style,int size){ return new Font("Segoe UI",style,sc(size)); }
    private static final Color BG=new Color(0xF8FAFC), PANEL=new Color(0xF1F5F9), ACCENT=new Color(0x3B82F6),
        ACCENT_DARK=new Color(0x1E40AF), FG=new Color(0x475569), SUCCESS=new Color(0x16A34A),
        ERROR=new Color(0xDC2626), WARNING=new Color(0xD97706), BORDER=new Color(0xCBD5E1);
    private static final String DEFAULT_COLUMNS="Numero\nDescrizione\nBreve descrizione";
    private static final String DEFAULT_MAIN_KEYS="Richiesto per\nReparto\nUbicazione";
    private static final String DEFAULT_SECONDARY_KEYS="Nome utente\nReparto\nUbicazione";
    private static final int FORM_WIDTH=610, JOIN_INNER_WIDTH=568;
    private static final String DEFAULT_JOIN_OUTPUT="Nome utente\nNumero dipendente\nTitolo\nReparto";

    private JButton processBtn;
    private JTextPane statusPane;
    private JScrollPane statusScroll;
    private JTextArea columnsArea,mainKeysArea,secondaryKeysArea,joinOutputArea;
    private JCheckBox fallbackCheckBox;
    private DropPanel mainDropPanel,secondaryDropPanel;
    private String currentFile,secondaryFile;

    public App(){
        super("Excel -> TXT"); setDefaultCloseOperation(EXIT_ON_CLOSE); setMinimumSize(new Dimension(sc(620),sc(650)));
        JPanel content=new JPanel(){
            @Override public Dimension getPreferredSize(){
                Dimension d=super.getPreferredSize();
                d.width=sc(658);
                return d;
            }
        };
        content.setLayout(new BoxLayout(content,BoxLayout.Y_AXIS)); content.setBackground(BG);
        content.setBorder(new EmptyBorder(sc(18),sc(24),sc(20),sc(24)));
        JLabel title=label("Excel -> TXT",22,Font.BOLD,ACCENT_DARK); title.setAlignmentX(CENTER_ALIGNMENT); content.add(title); content.add(vgap(12));
        content.add(sectionTitle("Documento principale")); content.add(vgap(6));
        mainDropPanel=new DropPanel("Trascina qui il file Excel principale (.xlsx / .xls)","oppure clicca per sfogliare",true);
        content.add(mainDropPanel); content.add(vgap(10));
        columnsArea=addArea(content,"Colonne principali da estrarre (una per riga, nell'ordine desiderato):",DEFAULT_COLUMNS,105);
        content.add(vgap(10)); content.add(buildJoinPanel()); content.add(vgap(12));
        processBtn=button("Elabora"); processBtn.setEnabled(false); processBtn.setAlignmentX(CENTER_ALIGNMENT); processBtn.addActionListener(e->onProcess()); content.add(processBtn); content.add(vgap(9));
        statusPane=new JTextPane();
        statusPane.setEditable(false);
        statusPane.setContentType("text/html");
        statusPane.setFont(font(Font.PLAIN,12));
        statusPane.setBackground(Color.WHITE);
        statusPane.setBorder(new EmptyBorder(sc(7),sc(9),sc(7),sc(9)));

        statusScroll=new JScrollPane(statusPane);
        statusScroll.setAlignmentX(CENTER_ALIGNMENT);
        statusScroll.setPreferredSize(new Dimension(sc(FORM_WIDTH),sc(125)));
        statusScroll.setMinimumSize(new Dimension(sc(FORM_WIDTH),sc(125)));
        statusScroll.setMaximumSize(new Dimension(sc(FORM_WIDTH),sc(125)));
        statusScroll.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
        statusScroll.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_AS_NEEDED);
        statusScroll.getVerticalScrollBar().setUnitIncrement(sc(16));
        statusScroll.setBorder(BorderFactory.createTitledBorder(
                BorderFactory.createLineBorder(BORDER),"Stato e warning",
                TitledBorder.LEFT,TitledBorder.TOP,font(Font.BOLD,12),new Color(0x0F172A)));
        setStatus("In attesa del documento principale...",FG);
        content.add(statusScroll);
        content.setMaximumSize(new Dimension(sc(658), Integer.MAX_VALUE));

        JPanel centered = new JPanel(new GridBagLayout());
        centered.setBackground(BG);
        GridBagConstraints gbc=new GridBagConstraints();
        gbc.gridx=0; gbc.gridy=0; gbc.weightx=1.0; gbc.weighty=1.0;
        gbc.anchor=GridBagConstraints.NORTH; gbc.fill=GridBagConstraints.NONE;
        centered.add(content,gbc);

        JScrollPane scroll=new JScrollPane(centered);
        scroll.setBorder(null);
        scroll.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
        scroll.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_AS_NEEDED);
        scroll.getVerticalScrollBar().setUnitIncrement(sc(18));
        scroll.getViewport().setBackground(BG);
        setContentPane(scroll);
        setSize(sc(690),Math.min(sc(900),GraphicsEnvironment.getLocalGraphicsEnvironment().getMaximumWindowBounds().height)); setLocationRelativeTo(null);
    }
    private JPanel buildJoinPanel(){
        JPanel body=new JPanel();
        body.setLayout(new BoxLayout(body,BoxLayout.Y_AXIS));
        body.setBackground(PANEL);
        body.setBorder(new EmptyBorder(sc(10),sc(10),sc(10),sc(10)));

        JLabel optional=label("Opzionale: senza documento secondario l'app continua a funzionare come prima.",12,Font.PLAIN,FG);
        optional.setAlignmentX(LEFT_ALIGNMENT);
        optional.setPreferredSize(new Dimension(sc(JOIN_INNER_WIDTH),sc(22)));
        optional.setMaximumSize(new Dimension(sc(JOIN_INNER_WIDTH),sc(22)));
        body.add(optional); body.add(vgap(8));

        secondaryDropPanel=new DropPanel("Trascina qui il documento secondario per la join","ad esempio: sys_user.xlsx",false);
        secondaryDropPanel.setPreferredSize(new Dimension(sc(JOIN_INNER_WIDTH),sc(100)));
        secondaryDropPanel.setMaximumSize(new Dimension(sc(JOIN_INNER_WIDTH),sc(100)));
        secondaryDropPanel.setAlignmentX(LEFT_ALIGNMENT);
        body.add(secondaryDropPanel); body.add(vgap(9));

        JPanel keys=new JPanel(new GridLayout(1,2,sc(10),0));
        keys.setOpaque(false); keys.setAlignmentX(LEFT_ALIGNMENT);
        keys.setPreferredSize(new Dimension(sc(JOIN_INNER_WIDTH),sc(108)));
        keys.setMaximumSize(new Dimension(sc(JOIN_INNER_WIDTH),sc(108)));
        JPanel left=areaPanel("Chiavi nel documento principale:",DEFAULT_MAIN_KEYS,82);
        JPanel right=areaPanel("Chiavi nel documento secondario:",DEFAULT_SECONDARY_KEYS,82);
        mainKeysArea=(JTextArea)((JScrollPane)left.getComponent(2)).getViewport().getView();
        secondaryKeysArea=(JTextArea)((JScrollPane)right.getComponent(2)).getViewport().getView();
        keys.add(left); keys.add(right); body.add(keys); body.add(vgap(9));

        JPanel output=areaPanel("Colonne del documento secondario da aggiungere all'output:",DEFAULT_JOIN_OUTPUT,92);
        output.setAlignmentX(LEFT_ALIGNMENT);
        output.setPreferredSize(new Dimension(sc(JOIN_INNER_WIDTH),sc(118)));
        output.setMaximumSize(new Dimension(sc(JOIN_INNER_WIDTH),sc(118)));
        JScrollPane outputScroll=(JScrollPane)output.getComponent(2);
        outputScroll.setPreferredSize(new Dimension(sc(JOIN_INNER_WIDTH),sc(92)));
        outputScroll.setMaximumSize(new Dimension(sc(JOIN_INNER_WIDTH),sc(92)));
        joinOutputArea=(JTextArea)outputScroll.getViewport().getView();
        body.add(output); body.add(vgap(7));

        fallbackCheckBox=new JCheckBox("Abilita join con fallback sull'ultima chiave",true);
        fallbackCheckBox.setOpaque(false); fallbackCheckBox.setFont(font(Font.PLAIN,12)); fallbackCheckBox.setForeground(FG);
        fallbackCheckBox.setAlignmentX(LEFT_ALIGNMENT);
        fallbackCheckBox.setMaximumSize(new Dimension(sc(JOIN_INNER_WIDTH),sc(28)));
        fallbackCheckBox.setToolTipText("Se la join completa fallisce, riprova senza l'ultima chiave, solo se la corrispondenza resta univoca.");
        body.add(fallbackCheckBox);

        JLabel help=label("<html>Le chiavi sono accoppiate riga per riga. Il fallback ignora soltanto l'ultima chiave<br>e solo in assenza di ambiguita.</html>",11,Font.PLAIN,FG);
        help.setAlignmentX(LEFT_ALIGNMENT); help.setMaximumSize(new Dimension(sc(JOIN_INNER_WIDTH),sc(38)));
        body.add(help); body.add(vgap(7));

        JButton clear=new JButton("Rimuovi documento secondario");
        clear.setFont(font(Font.PLAIN,11)); clear.setAlignmentX(LEFT_ALIGNMENT);
        clear.addActionListener(e->clearSecondaryFile()); body.add(clear);

        JPanel wrapper=new JPanel(new BorderLayout());
        wrapper.setBackground(BG);
        wrapper.setBorder(BorderFactory.createTitledBorder(BorderFactory.createLineBorder(BORDER),"Join con documento secondario",TitledBorder.LEFT,TitledBorder.TOP,font(Font.BOLD,13),new Color(0x0F172A)));
        wrapper.add(body,BorderLayout.CENTER);
        wrapper.setAlignmentX(CENTER_ALIGNMENT);
        wrapper.setPreferredSize(new Dimension(sc(FORM_WIDTH),sc(470)));
        wrapper.setMaximumSize(new Dimension(sc(FORM_WIDTH),sc(470)));
        return wrapper;
    }
    private JTextArea addArea(JPanel parent,String text,String value,int h){ JPanel p=areaPanel(text,value,h); parent.add(p); return (JTextArea)((JScrollPane)p.getComponent(2)).getViewport().getView(); }
    private JPanel areaPanel(String text,String value,int h){ JPanel p=new JPanel(); p.setLayout(new BoxLayout(p,BoxLayout.Y_AXIS)); p.setOpaque(false); p.setAlignmentX(CENTER_ALIGNMENT); p.setPreferredSize(new Dimension(sc(FORM_WIDTH),sc(h+26))); p.setMaximumSize(new Dimension(sc(FORM_WIDTH),sc(h+26))); JLabel l=label(text,12,Font.PLAIN,FG); l.setAlignmentX(LEFT_ALIGNMENT); p.add(l); p.add(vgap(4)); JTextArea a=new JTextArea(value); a.setFont(font(Font.PLAIN,13)); a.setLineWrap(false); a.setBorder(new EmptyBorder(sc(6),sc(8),sc(6),sc(8))); JScrollPane s=new JScrollPane(a); s.setPreferredSize(new Dimension(sc(FORM_WIDTH),sc(h))); s.setMaximumSize(new Dimension(sc(FORM_WIDTH),sc(h))); s.setBorder(BorderFactory.createLineBorder(BORDER)); s.setAlignmentX(LEFT_ALIGNMENT); p.add(s); return p; }
    private JLabel sectionTitle(String s){ JLabel l=label(s,15,Font.BOLD,new Color(0x0F172A)); l.setAlignmentX(CENTER_ALIGNMENT); l.setPreferredSize(new Dimension(sc(FORM_WIDTH),sc(22))); l.setMaximumSize(new Dimension(sc(FORM_WIDTH),sc(22))); return l; }
    private static JLabel label(String s,int size,int style,Color c){ JLabel l=new JLabel(s); l.setFont(font(style,size)); l.setForeground(c); return l; }
    private static Component vgap(int n){ return Box.createRigidArea(new Dimension(1,sc(n))); }
    private JButton button(String text){
        JButton b=new JButton(text){
            @Override protected void paintComponent(Graphics g){
                Graphics2D bg=(Graphics2D)g.create();
                bg.setRenderingHint(RenderingHints.KEY_ANTIALIASING,RenderingHints.VALUE_ANTIALIAS_ON);
                Color fill;
                if(!isEnabled()) fill=new Color(0xCBD5E1);
                else if(getModel().isPressed()) fill=ACCENT_DARK;
                else if(getModel().isRollover()) fill=new Color(0x2563EB);
                else fill=ACCENT;
                bg.setColor(fill);
                bg.fillRoundRect(0,0,getWidth(),getHeight(),sc(8),sc(8));
                bg.dispose();

                Graphics2D textGraphics=(Graphics2D)g.create();
                textGraphics.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING,RenderingHints.VALUE_TEXT_ANTIALIAS_LCD_HRGB);
                super.paintComponent(textGraphics);
                textGraphics.dispose();
            }
        };
        b.setFont(font(Font.BOLD,11));
        b.setForeground(Color.WHITE);
        b.setContentAreaFilled(false);
        b.setOpaque(false);
        b.setBorderPainted(false);
        b.setFocusPainted(false);
        b.setRolloverEnabled(true);
        b.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        b.setPreferredSize(new Dimension(sc(110),sc(30)));
        b.setMaximumSize(new Dimension(sc(110),sc(30)));
        return b;
    }
    private boolean validExcel(File f){ String n=f.getName().toLowerCase(Locale.ROOT); return n.endsWith(".xlsx")||n.endsWith(".xls"); }
    private void choose(boolean main){ JFileChooser c=new JFileChooser(); c.setDialogTitle("Seleziona file Excel"); c.setFileFilter(new FileNameExtensionFilter("File Excel (*.xlsx, *.xls)","xlsx","xls")); if(c.showOpenDialog(this)==JFileChooser.APPROVE_OPTION){ if(main) handleMain(c.getSelectedFile()); else handleSecondary(c.getSelectedFile()); } }
    private void handleMain(File f){ if(!validExcel(f)){currentFile=null;processBtn.setEnabled(false);setStatus("File principale non supportato. Usa .xlsx o .xls",ERROR);return;} currentFile=f.getAbsolutePath();mainDropPanel.update("Selezionato: "+f.getName(),"Documento principale pronto");processBtn.setEnabled(true);setStatus("Documento principale pronto.",FG); }
    private void handleSecondary(File f){ if(!validExcel(f)){secondaryFile=null;setStatus("Documento secondario non supportato. Usa .xlsx o .xls",ERROR);return;} secondaryFile=f.getAbsolutePath();secondaryDropPanel.update("Selezionato: "+f.getName(),"Join attiva");setStatus("Documento secondario pronto: la join verra applicata.",FG); }
    private void clearSecondaryFile(){secondaryFile=null;secondaryDropPanel.update("Trascina qui il documento secondario per la join","ad esempio: sys_user.xlsx");setStatus("Documento secondario rimosso. Elaborazione standard attiva.",FG);}
    private List<String> lines(JTextArea a){List<String> r=new ArrayList<String>();for(String s:a.getText().split("\\R")){s=s.trim();if(!s.isEmpty())r.add(s);}return r;}
    private void onProcess(){
        if(currentFile==null)return; final List<String> cols=lines(columnsArea); if(cols.isEmpty()){setStatus("Inserisci almeno una colonna principale.",ERROR);return;}
        final String main=currentFile,second=secondaryFile; final List<String> mk=lines(mainKeysArea),sk=lines(secondaryKeysArea),out=lines(joinOutputArea); final boolean fb=fallbackCheckBox.isSelected();
        processBtn.setEnabled(false); setStatus("Elaborazione in corso...",FG);
        SwingWorker<Result,Void> w=new SwingWorker<Result,Void>(){protected Result doInBackground()throws Exception{String txt=Logic.processExcel(main,cols,second,mk,sk,out,fb);String warn=Logic.consumeWarnings();return new Result(Logic.saveOutput(main,txt),warn);}protected void done(){try{Result r=get();if(r.warning.isEmpty())setStatus("File salvato:<br>"+esc(r.path),SUCCESS);else setStatus("File salvato correttamente:<br>"+esc(r.path)+"<br><br>"+esc(r.warning).replace("\n","<br>"),WARNING);}catch(InterruptedException ex){Thread.currentThread().interrupt();setStatus("Elaborazione interrotta.",ERROR);}catch(ExecutionException ex){Throwable c=ex.getCause();while(c!=null&&c.getCause()!=null)c=c.getCause();setStatus("Errore:<br>"+esc(c==null?"sconosciuto":c.getMessage()),ERROR);}finally{processBtn.setEnabled(true);}}}; w.execute();
    }
    private static String esc(String s){return s==null?"":s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;");}
    private void setStatus(String msg,Color c){
        String hex=String.format("#%02x%02x%02x",c.getRed(),c.getGreen(),c.getBlue());
        statusPane.setText("<html><body style='margin:0;font-family:Segoe UI;font-size:"+sc(12)
                +"pt;color:"+hex+";'>"+msg.replace("\n","<br>")+"</body></html>");
        statusPane.setCaretPosition(0);
        statusPane.revalidate();
        statusPane.repaint();
        if(statusScroll!=null){
            statusScroll.getVerticalScrollBar().setValue(0);
            statusScroll.revalidate();
            statusScroll.repaint();
        }
    }
    private static final class Result{final String path,warning;Result(String p,String w){path=p;warning=w;}}
    private final class DropPanel extends JPanel{
        private final JLabel first,second; private final boolean main;
        DropPanel(String a,String b,boolean m){main=m;setLayout(new BoxLayout(this,BoxLayout.Y_AXIS));setBackground(BG);setBorder(BorderFactory.createLineBorder(ACCENT,sc(2)));setPreferredSize(new Dimension(sc(610),sc(m?118:100)));setMaximumSize(new Dimension(sc(610),sc(m?118:100)));setAlignmentX(CENTER_ALIGNMENT);setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));add(Box.createVerticalGlue());JLabel icon=label("FILE",15,Font.BOLD,ACCENT_DARK);icon.setAlignmentX(CENTER_ALIGNMENT);add(icon);first=label(a,14,Font.PLAIN,FG);first.setAlignmentX(CENTER_ALIGNMENT);add(first);second=label(b,12,Font.PLAIN,new Color(0x94A3B8));second.setAlignmentX(CENTER_ALIGNMENT);add(second);add(Box.createVerticalGlue());addMouseListener(new MouseAdapter(){public void mouseClicked(MouseEvent e){choose(main);}public void mouseEntered(MouseEvent e){setBorder(BorderFactory.createLineBorder(ACCENT_DARK,sc(2)));}public void mouseExited(MouseEvent e){setBorder(BorderFactory.createLineBorder(ACCENT,sc(2)));}});new DropTarget(this,new DropTargetAdapter(){public void drop(DropTargetDropEvent e){try{e.acceptDrop(DnDConstants.ACTION_COPY);@SuppressWarnings("unchecked") List<File> fs=(List<File>)e.getTransferable().getTransferData(DataFlavor.javaFileListFlavor);if(!fs.isEmpty()){if(main)handleMain(fs.get(0));else handleSecondary(fs.get(0));}e.dropComplete(true);}catch(Exception x){e.dropComplete(false);App.this.setStatus("Errore nel trascinamento: "+esc(x.getMessage()),App.ERROR);}}});}
        void update(String a,String b){first.setText(a);second.setText(b);}
    }
}

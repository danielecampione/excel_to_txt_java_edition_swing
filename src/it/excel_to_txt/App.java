package it.excel_to_txt;

import javax.swing.*;
import javax.swing.border.*;
import java.awt.*;
import java.awt.datatransfer.*;
import java.awt.dnd.*;
import java.awt.event.*;
import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * App.java
 * --------
 * Interfaccia grafica Swing con supporto drag-and-drop.
 * Permette di specificare le colonne Excel da estrarre tramite textarea.
 *
 * HiDPI: tutti i valori pixel passano attraverso sc() / scf() che applicano
 * il fattore di scala calcolato dal DPI di sistema (96 dpi = scala 1.0×).
 */
public class App extends JFrame {

    // ── Supporto HiDPI ───────────────────────────────────────────────────────

    /** Fattore di scala rispetto ai 96 dpi di riferimento. Es. 1.5 a 144 dpi. */
    private static final float SF;
    static {
        float screenDpi = Toolkit.getDefaultToolkit().getScreenResolution();
        SF = Math.max(1.0f, screenDpi / 96.0f);
    }

    /** Converte pixel logici (96 dpi) in pixel fisici. */
    private static int sc(int px) {
        return Math.round(px * SF);
    }

    /** Crea un Font con pt scalato al DPI corrente. */
    private static Font scf(String name, int style, int ptLogical) {
        return new Font(name, style, sc(ptLogical));
    }

    /** Applica rendering hint di qualità a un Graphics2D (anti-alias, LCD, ecc.). */
    private static void applyHints(Graphics2D g2) {
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING,      RenderingHints.VALUE_ANTIALIAS_ON);
        g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_LCD_HRGB);
        g2.setRenderingHint(RenderingHints.KEY_RENDERING,         RenderingHints.VALUE_RENDER_QUALITY);
        g2.setRenderingHint(RenderingHints.KEY_STROKE_CONTROL,    RenderingHints.VALUE_STROKE_PURE);
        g2.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
    }

    // ── Palette ──────────────────────────────────────────────────────────────
    private static final Color BG          = new Color(0xF8FAFC);
    private static final Color ACCENT      = new Color(0x3B82F6);
    private static final Color ACCENT_DARK = new Color(0x1E40AF);
    private static final Color FG_DIM      = new Color(0x475569);
    private static final Color SUCCESS     = new Color(0x16A34A);
    private static final Color ERROR       = new Color(0xDC2626);

    // Colonne predefinite (stesso comportamento originale)
    private static final String DEFAULT_COLUMNS = "Numero\nDescrizione\nBreve descrizione";

    // ── Componenti ───────────────────────────────────────────────────────────
    private final JPanel    dropPanel;
    private final JButton   processBtn;
    private final JLabel    statusLbl;
    private final JTextArea columnsArea;

    private String currentFile = null;

    // ── Costruttore ──────────────────────────────────────────────────────────

    public App() {
        super("Excel → TXT");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setResizable(false);
        getContentPane().setBackground(BG);
        setLayout(new BoxLayout(getContentPane(), BoxLayout.Y_AXIS));

        // -- Titolo
        JLabel title = new JLabel("Excel → TXT", SwingConstants.CENTER);
        title.setFont(scf("Segoe UI", Font.BOLD, 20));
        title.setForeground(ACCENT_DARK);
        title.setAlignmentX(Component.CENTER_ALIGNMENT);
        title.setBorder(new EmptyBorder(sc(20), sc(24), sc(10), sc(24)));
        add(title);

        // -- Drop panel
        dropPanel = buildDropPanel();
        JPanel dropWrapper = new JPanel(new FlowLayout(FlowLayout.CENTER, 0, 0));
        dropWrapper.setBackground(BG);
        dropWrapper.setBorder(new EmptyBorder(0, sc(24), sc(8), sc(24)));
        dropWrapper.add(dropPanel);
        add(dropWrapper);

        // -- Sezione colonne
        columnsArea = buildColumnsSection();

        // -- Bottone Elabora
        processBtn = buildProcessButton();
        JPanel btnWrapper = new JPanel(new FlowLayout(FlowLayout.CENTER));
        btnWrapper.setBackground(BG);
        btnWrapper.add(processBtn);
        add(btnWrapper);

        // -- Etichetta stato
        statusLbl = new JLabel("In attesa di un file…", SwingConstants.CENTER);
        statusLbl.setFont(scf("Segoe UI", Font.PLAIN, 11));
        statusLbl.setForeground(FG_DIM);
        statusLbl.setAlignmentX(Component.CENTER_ALIGNMENT);
        statusLbl.setBorder(new EmptyBorder(0, sc(24), sc(18), sc(24)));
        add(statusLbl);

        // -- Drop target sul pannello
        registerDropTarget();

        pack();
        setSize(sc(470), sc(530));
        setLocationRelativeTo(null);
    }

    // ── Costruzione pannello drop ─────────────────────────────────────────────

    private JPanel buildDropPanel() {
        JPanel panel = new JPanel();
        panel.setPreferredSize(new Dimension(sc(420), sc(110)));
        panel.setBackground(BG);
        panel.setBorder(BorderFactory.createLineBorder(ACCENT, Math.max(1, sc(2))));
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));

        JLabel icon = new JLabel("📂", SwingConstants.CENTER);
        icon.setFont(scf("Segoe UI Emoji", Font.PLAIN, 28));
        icon.setAlignmentX(Component.CENTER_ALIGNMENT);
        icon.setBorder(new EmptyBorder(sc(12), 0, sc(4), 0));
        panel.add(icon);

        JLabel hint = new JLabel("Trascina qui il file Excel (.xlsx / .xls)", SwingConstants.CENTER);
        hint.setFont(scf("Segoe UI", Font.PLAIN, 11));
        hint.setForeground(FG_DIM);
        hint.setAlignmentX(Component.CENTER_ALIGNMENT);
        panel.add(hint);

        JLabel hint2 = new JLabel("oppure clicca per sfogliare", SwingConstants.CENTER);
        hint2.setFont(scf("Segoe UI", Font.PLAIN, 10));
        hint2.setForeground(new Color(0x94A3B8));
        hint2.setAlignmentX(Component.CENTER_ALIGNMENT);
        hint2.setBorder(new EmptyBorder(sc(2), 0, 0, 0));
        panel.add(hint2);

        panel.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        panel.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) { openFileChooser(); }
            @Override
            public void mouseEntered(MouseEvent e) {
                panel.setBorder(BorderFactory.createLineBorder(ACCENT_DARK, Math.max(1, sc(2))));
            }
            @Override
            public void mouseExited(MouseEvent e) {
                panel.setBorder(BorderFactory.createLineBorder(ACCENT, Math.max(1, sc(2))));
            }
        });

        return panel;
    }

    // ── Costruzione sezione colonne ───────────────────────────────────────────

    private JTextArea buildColumnsSection() {
        // Label descrittiva
        JLabel lbl = new JLabel("Colonne da estrarre (una per riga, nell'ordine desiderato):");
        lbl.setFont(scf("Segoe UI", Font.PLAIN, 11));
        lbl.setForeground(FG_DIM);
        lbl.setAlignmentX(Component.CENTER_ALIGNMENT);
        lbl.setBorder(new EmptyBorder(sc(4), sc(24), sc(4), sc(24)));
        add(lbl);

        // Textarea
        JTextArea area = new JTextArea(DEFAULT_COLUMNS);
        area.setFont(scf("Segoe UI", Font.PLAIN, 12));
        area.setForeground(new Color(0x0F172A));
        area.setBackground(Color.WHITE);
        area.setLineWrap(false);
        area.setWrapStyleWord(false);
        area.setBorder(new EmptyBorder(sc(6), sc(8), sc(6), sc(8)));

        JScrollPane scroll = new JScrollPane(area);
        scroll.setPreferredSize(new Dimension(sc(420), sc(110)));
        scroll.setBorder(BorderFactory.createLineBorder(new Color(0xCBD5E1), Math.max(1, sc(1))));

        JPanel areaWrapper = new JPanel(new FlowLayout(FlowLayout.CENTER, 0, 0));
        areaWrapper.setBackground(BG);
        areaWrapper.setBorder(new EmptyBorder(0, sc(24), sc(8), sc(24)));
        areaWrapper.add(scroll);
        add(areaWrapper);

        return area;
    }

    // ── Costruzione bottone ───────────────────────────────────────────────────

    private JButton buildProcessButton() {
        JButton btn = new JButton("Elabora") {
            @Override
            protected void paintComponent(Graphics g) {
                Graphics2D g2d = (Graphics2D) g;

                // Imposta rendering hints prima di super.paintComponent()
                // così anche il testo del bottone beneficia di AA e LCD.
                applyHints(g2d);

                // Sfondo con angoli arrotondati scalati
                Graphics2D g2bg = (Graphics2D) g.create();
                applyHints(g2bg);
                if (isEnabled()) {
                    g2bg.setColor(getModel().isPressed() ? ACCENT_DARK : ACCENT);
                } else {
                    g2bg.setColor(new Color(0xCBD5E1));
                }
                g2bg.fillRoundRect(0, 0, getWidth(), getHeight(), sc(8), sc(8));
                g2bg.dispose();

                super.paintComponent(g);
            }
        };
        btn.setFont(scf("Segoe UI", Font.BOLD, 11));
        btn.setForeground(Color.WHITE);
        btn.setContentAreaFilled(false);
        btn.setBorderPainted(false);
        btn.setFocusPainted(false);
        btn.setEnabled(false);
        btn.setPreferredSize(new Dimension(sc(110), sc(30)));
        btn.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        btn.addActionListener(e -> onProcess());
        return btn;
    }

    // ── Drag-and-drop ────────────────────────────────────────────────────────

    private void registerDropTarget() {
        new DropTarget(dropPanel, new DropTargetAdapter() {
            @Override
            public void dragEnter(DropTargetDragEvent dtde) {
                dropPanel.setBorder(BorderFactory.createLineBorder(ACCENT_DARK, Math.max(1, sc(3))));
            }
            @Override
            public void dragExit(DropTargetEvent dte) {
                dropPanel.setBorder(BorderFactory.createLineBorder(ACCENT, Math.max(1, sc(2))));
            }
            @Override
            public void drop(DropTargetDropEvent dtde) {
                dropPanel.setBorder(BorderFactory.createLineBorder(ACCENT, Math.max(1, sc(2))));
                try {
                    dtde.acceptDrop(DnDConstants.ACTION_COPY);
                    @SuppressWarnings("unchecked")
                    List<File> files = (List<File>)
                            dtde.getTransferable().getTransferData(DataFlavor.javaFileListFlavor);
                    if (!files.isEmpty()) handleFile(files.get(0));
                    dtde.dropComplete(true);
                } catch (Exception ex) {
                    dtde.dropComplete(false);
                    setStatus("❌ Errore nel drop: " + ex.getMessage(), ERROR);
                }
            }
        });
    }

    // ── File chooser (click) ──────────────────────────────────────────────────

    private void openFileChooser() {
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Seleziona file Excel");
        chooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter(
                "File Excel (*.xlsx, *.xls)", "xlsx", "xls"
        ));
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            handleFile(chooser.getSelectedFile());
        }
    }

    // ── Gestione file selezionato ─────────────────────────────────────────────

    private void handleFile(File file) {
        String name = file.getName().toLowerCase();
        if (!name.endsWith(".xlsx") && !name.endsWith(".xls")) {
            setStatus("❌ File non supportato. Trascina un file .xlsx o .xls", ERROR);
            currentFile = null;
            processBtn.setEnabled(false);
            return;
        }
        currentFile = file.getAbsolutePath();
        setStatus("📄 Selezionato: " + file.getName() + "\nPronto per l'elaborazione.", FG_DIM);
        processBtn.setEnabled(true);
    }

    // ── Legge le colonne dalla textarea ──────────────────────────────────────

    private List<String> parseColumns() {
        List<String> cols = new ArrayList<>();
        for (String line : columnsArea.getText().split("\n")) {
            String trimmed = line.trim();
            if (!trimmed.isEmpty()) cols.add(trimmed);
        }
        return cols;
    }

    // ── Elaborazione ─────────────────────────────────────────────────────────

    private void onProcess() {
        if (currentFile == null) return;

        List<String> columns = parseColumns();
        if (columns.isEmpty()) {
            setStatus("❌ Inserisci almeno un nome colonna.", ERROR);
            return;
        }

        setStatus("⏳ Elaborazione in corso…", FG_DIM);
        processBtn.setEnabled(false);

        final String filePath = currentFile;

        SwingWorker<String, Void> worker = new SwingWorker<String, Void>() {
            @Override
            protected String doInBackground() throws Exception {
                String content = Logic.processExcel(filePath, columns);
                return Logic.saveOutput(filePath, content);
            }
            @Override
            protected void done() {
                try {
                    String outPath = get();
                    setStatus("✅ File salvato:\n" + outPath, SUCCESS);
                } catch (Exception ex) {
                    Throwable cause = ex.getCause() != null ? ex.getCause() : ex;
                    setStatus("❌ Errore:\n" + cause.getMessage(), ERROR);
                } finally {
                    currentFile = null;
                }
            }
        };
        worker.execute();
    }

    // ── Aggiorna etichetta di stato ───────────────────────────────────────────

    private void setStatus(String msg, Color color) {
        String html = "<html><center>" + msg.replace("\n", "<br>") + "</center></html>";
        statusLbl.setText(html);
        statusLbl.setForeground(color);
    }
}

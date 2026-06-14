package it.excel_to_txt;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;
import java.util.*;

/**
 * Logic.java
 * ----------
 * Logica di elaborazione: lettura Excel e scrittura file di testo.
 * Le colonne da estrarre sono passate come lista ordinata dall'esterno.
 */
public class Logic {

    // ── Trova il primo foglio che non finisce con "bkp" ──────────────────────

    private static Sheet getFirstNonBkpSheet(Workbook wb) {
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            String name = wb.getSheetName(i);
            if (!name.toLowerCase().endsWith("bkp")) {
                return wb.getSheetAt(i);
            }
        }
        throw new IllegalArgumentException(
                "Tutti i fogli hanno il suffisso 'bkp': nessun foglio utilizzabile."
        );
    }

    // ── Legge una cella come stringa ─────────────────────────────────────────

    private static String cellStr(Cell cell) {
        if (cell == null) return "";
        DataFormatter fmt = new DataFormatter();
        return fmt.formatCellValue(cell).trim();
    }

    // ── Elabora il file Excel e restituisce il testo formattato ──────────────
    //    columns: lista ordinata dei nomi colonna da estrarre

    public static String processExcel(String path, List<String> columns) throws IOException {
        path = path.trim();
        File file = new File(path);

        try (InputStream is = new FileInputStream(file);
             Workbook wb = path.toLowerCase().endsWith(".xlsx")
                     ? new XSSFWorkbook(is)
                     : new HSSFWorkbook(is)) {

            Sheet sheet = getFirstNonBkpSheet(wb);

            // Prima riga = intestazioni
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                throw new IllegalArgumentException("Il foglio è vuoto.");
            }

            // Mappa nome colonna → indice fisico
            Map<String, Integer> colIndex = new LinkedHashMap<>();
            for (int c = 0; c <= headerRow.getLastCellNum(); c++) {
                Cell hCell = headerRow.getCell(c);
                if (hCell != null) {
                    String name = cellStr(hCell);
                    colIndex.put(name, c);
                }
            }

            // Verifica che tutte le colonne richieste esistano
            List<String> missing = new ArrayList<>();
            for (String col : columns) {
                if (!colIndex.containsKey(col)) missing.add(col);
            }
            if (!missing.isEmpty()) {
                throw new IllegalArgumentException(
                        "Colonne mancanti nel foglio '" + sheet.getSheetName() + "': " + missing
                );
            }

            // Costruisce i blocchi rispettando l'ordine delle colonne indicate
            List<String> blocks = new ArrayList<>();
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                // Raccoglie i valori nell'ordine richiesto
                List<String> values = new ArrayList<>();
                boolean allEmpty = true;
                for (String col : columns) {
                    String val = cellStr(row.getCell(colIndex.get(col)));
                    values.add(val);
                    if (!val.isEmpty()) allEmpty = false;
                }

                // Salta righe completamente vuote
                if (allEmpty) continue;

                // Concatena i valori separati da doppio a capo
                StringBuilder block = new StringBuilder();
                for (int i = 0; i < values.size(); i++) {
                    block.append(values.get(i));
                    if (i < values.size() - 1) block.append("\n\n");
                }
                block.append("\n");
                blocks.add(block.toString());
            }

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < blocks.size(); i++) {
                sb.append(blocks.get(i));
                sb.append("\n---\n");
                if (i < blocks.size() - 1) sb.append("\n");
            }
            return sb.toString();
        }
    }

    // ── Salva il file di output ───────────────────────────────────────────────

    public static String saveOutput(String excelPath, String content) throws IOException {
        String base = excelPath.replaceAll("\\.[^.]+$", "");
        File outFile = new File(base + "_output.txt");

        // Fallback: stessa cartella dello script se non scrivibile
        if (!outFile.getParentFile().canWrite()) {
            outFile = new File(System.getProperty("user.dir"),
                    new File(base).getName() + "_output.txt");
        }

        try (OutputStreamWriter w = new OutputStreamWriter(
                new FileOutputStream(outFile), "UTF-8")) {
            w.write(content);
        }

        return outFile.getAbsolutePath();
    }
}

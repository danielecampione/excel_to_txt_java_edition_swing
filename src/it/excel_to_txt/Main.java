package it.excel_to_txt;

import javax.swing.*;

/**
 * Main.java
 * ---------
 * Entry point: avvia la GUI.
 *
 * Uso:
 *   mvn package
 *   java -jar target/excel-to-txt.jar
 *
 * Dipendenze (gestite da Maven):
 *   org.apache.poi : poi-ooxml (per .xlsx)
 *   org.apache.poi : poi       (per .xls)
 *
 * HiDPI (Java 8 / Windows):
 *   "sun.java2d.dpiaware" deve essere impostato PRIMA di qualsiasi
 *   inizializzazione AWT. Dice a Windows di NON applicare il bitmap-
 *   upscaling sfocato; il ridimensionamento corretto è gestito da App.java
 *   tramite il fattore di scala calcolato dal DPI di sistema.
 *   In alternativa si può passare -Dsun.java2d.dpiaware=true alla JVM.
 */
public class Main {
    public static void main(String[] args) {
        // HiDPI – deve precedere qualsiasi init AWT/Swing
        System.setProperty("sun.java2d.dpiaware", "true");

        SwingUtilities.invokeLater(() -> {
            try {
                // Usa il look-and-feel di sistema se disponibile
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception ignored) {
                // Fallback al look-and-feel di default
            }
            new App().setVisible(true);
        });
    }
}

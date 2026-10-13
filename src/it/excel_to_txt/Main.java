package it.excel_to_txt;
import javax.swing.*;
public class Main {
    public static void main(String[] args){
        System.setProperty("sun.java2d.dpiaware","true");
        SwingUtilities.invokeLater(() -> { try{UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());}catch(Exception ignored){} new App().setVisible(true); });
    }
}

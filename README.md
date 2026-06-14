# Excel to TXT - Java Edition Swing

Un'applicazione Java, che usa il Toolkit UI Swing, leggera ed efficiente progettata per convertire fogli di calcolo Excel (`.xlsx` e `.xls`) in file di testo piano (`.txt`). Questo strumento è ideale per l'esportazione di dati, la preparazione di log o la formattazione di dataset per elaborazioni successive.

---

## 🚀 Funzionalità

- **Supporto Multi-Formato:** Compatibile sia con il formato moderno Excel (`.xlsx`) sia con quello legacy (`.xls`).
- **Efficienza:** Ottimizzato per gestire file di grandi dimensioni riducendo al minimo l'uso della memoria.

---

## 🛠️ Prerequisiti

Prima di iniziare, assicurati di avere installato sul tuo sistema:

- **Java Development Kit (JDK):** Versione 8 o superiore (consigliata JDK 11 o 17).

---

## 📦 Dipendenze

Il progetto utilizza le librerie **Apache POI** per la lettura dei file Excel.

---

## 📝 Esempio di Codice (Core Logic)

Ecco una rapida panoramica di come avviene la conversione all'interno del codice:

```java
import org.apache.poi.ss.usermodel.*;
import java.io.*;

public class ExcelToTxtConverter {
    public static void convert(String excelPath, String txtPath, String delimiter) throws Exception {
        Workbook workbook = WorkbookFactory.create(new File(excelPath));
        BufferedWriter writer = new BufferedWriter(new FileWriter(txtPath));

        // Prende il primo foglio di lavoro
        Sheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            StringBuilder rowText = new StringBuilder();
            for (int i = 0; i < row.getLastCellNum(); i++) {
                Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                rowText.append(cell.toString()).append(delimiter);
            }
            writer.write(rowText.toString().trim());
            writer.newLine();
        }

        writer.close();
        workbook.close();
        System.out.println("Conversione completata con successo!");
    }
}
```

---

## 🤝 Contributi

I contributi sono sempre benvenuti!

---

## 📄 Licenza

Questo progetto è distribuito sotto la licenza GNU GENERAL PUBLIC LICENSE. Consulta il file `LICENSE` per ulteriori dettagli.

---

## 👤 Autore

**Daniele Campione**
- GitHub: [@danielecampione](https://github.com/danielecampione)

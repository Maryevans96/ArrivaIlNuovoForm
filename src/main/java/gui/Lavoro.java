package gui;

import model.Allegati;
import model.PdfFiller;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class Lavoro extends JFrame {

    private JPanel contentPane;
    private JTextField numeroOdsField, dataOdsField, scadenzaOdsField, viaField,
            danneggianteField, descrizioneInterventoField, inizioLavoriField, fineLavoriField;

    private JButton scaricaButton, compilaButton, caricaExcelButton, prossimoButton, precedenteButton, pulisciCampiButton;
    private JLabel infoExcelLabel;

    private List<Allegati> listaDatiExcel = new ArrayList<>();
    private int indiceCorrente = -1;
    private String lastCompiledFilePath;
    private SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");

    public Lavoro() {
        super("Compilatore PDF da Excel - SCHEDA CRISTOFOROECO");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        contentPane = new JPanel(new BorderLayout(10, 10));
        contentPane.setBorder(new EmptyBorder(15, 15, 15, 15));

        // --- Pannello Superiore: Caricamento e Navigazione ---
        JPanel topPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 10));
        topPanel.setBorder(BorderFactory.createTitledBorder("Gestione Excel"));

        caricaExcelButton = new JButton("Carica Excel");
        precedenteButton = new JButton("<< Precedente");
        prossimoButton = new JButton("Prossimo >>");

        precedenteButton.setEnabled(false);
        prossimoButton.setEnabled(false);
        infoExcelLabel = new JLabel("Nessun file caricato");

        caricaExcelButton.addActionListener(e -> importaExcel());
        prossimoButton.addActionListener(e -> mostraProssimoDato());
        precedenteButton.addActionListener(e -> mostraDatoPrecedente());

        topPanel.add(caricaExcelButton);
        topPanel.add(precedenteButton);
        topPanel.add(prossimoButton);
        topPanel.add(infoExcelLabel);
        contentPane.add(topPanel, BorderLayout.NORTH);

        // --- Pannello Centrale: Dati Input ---
        JPanel dataInputPanel = new JPanel(new GridLayout(0, 2, 5, 5));
        dataInputPanel.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createTitledBorder("Dati per la Compilazione"),
                new EmptyBorder(10, 10, 10, 10)
        ));

        numeroOdsField = new JTextField(20);
        dataOdsField = new JTextField(20);
        scadenzaOdsField = new JTextField(20);
        viaField = new JTextField(20);
        danneggianteField = new JTextField(20);
        descrizioneInterventoField = new JTextField(20);
        inizioLavoriField = new JTextField(20);
        fineLavoriField = new JTextField(20);

        addLabelAndField(dataInputPanel, "Numero O.d.S.:", numeroOdsField);
        addLabelAndField(dataInputPanel, "Data O.d.S. (gg/mm/aaaa):", dataOdsField);
        addLabelAndField(dataInputPanel, "Scadenza O.d.S. (gg/mm/aaaa):", scadenzaOdsField);
        addLabelAndField(dataInputPanel, "Via:", viaField);
        addLabelAndField(dataInputPanel, "Danneggiante:", danneggianteField);
        addLabelAndField(dataInputPanel, "Descrizione Intervento:", descrizioneInterventoField);
        addLabelAndField(dataInputPanel, "Data Inizio Lavori (gg/mm/aaaa):", inizioLavoriField);
        addLabelAndField(dataInputPanel, "Data Fine Lavori (gg/mm/aaaa):", fineLavoriField);

        contentPane.add(dataInputPanel, BorderLayout.CENTER);

        // --- Pannello Inferiore: Bottoni Azione ---
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 10, 10));
        compilaButton = new JButton("Compila PDF");
        compilaButton.addActionListener(e -> compilePdf());
        scaricaButton = new JButton("Salva PDF Compilato");
        scaricaButton.setEnabled(false);
        scaricaButton.addActionListener(e -> downloadPdf());
        pulisciCampiButton = new JButton("Pulisci Campi");
        pulisciCampiButton.addActionListener(e -> clearFields());
        JButton esciButton = new JButton("Esci");
        esciButton.addActionListener(e -> System.exit(0));

        buttonPanel.add(compilaButton);
        buttonPanel.add(scaricaButton);
        buttonPanel.add(pulisciCampiButton);
        buttonPanel.add(esciButton);
        contentPane.add(buttonPanel, BorderLayout.SOUTH);

        add(contentPane);
        pack();
        setLocationRelativeTo(null);
    }

    private void addLabelAndField(JPanel panel, String label, JTextField field) {
        panel.add(new JLabel(label));
        panel.add(field);
    }

    private void importaExcel() {
        String desktopPath = System.getProperty("user.home") + File.separator + "Desktop";
        JFileChooser fileChooser = new JFileChooser(desktopPath);
        fileChooser.setDialogTitle("Seleziona il file Excel");
        fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("File Excel (.xlsx)", "xlsx"));

        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File file = fileChooser.getSelectedFile();
            try (FileInputStream fis = new FileInputStream(file);
                 Workbook workbook = new XSSFWorkbook(fis)) {

                Sheet sheet = workbook.getSheetAt(0);
                listaDatiExcel.clear();
                DataFormatter formatter = new DataFormatter();

                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    // Controllo: se la riga è nulla o la cella C (Numero) è vuota, salta
                    if (row == null || row.getCell(2, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) == null) continue;

                    String descrizioneBase = formatter.formatCellValue(row.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK));
                    String note = formatter.formatCellValue(row.getCell(12, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK));
                    String descrizioneCompleta = descrizioneBase;
                    if (note != null && !note.trim().isEmpty()) {
                        descrizioneCompleta += " - " + note;
                    }

                    Date dOds = getCellValueAsDate(row.getCell(3));
                    Date dInizio, dFine;

                    if (note != null && note.toUpperCase().contains("PRONTO INTERVENTO")) {
                        dInizio = dOds;
                        dFine = dOds;
                    } else {
                        dInizio = getCellValueAsDate(row.getCell(10));
                        dFine = getCellValueAsDate(row.getCell(11));
                    }

                    Allegati dati = new Allegati(
                            formatter.formatCellValue(row.getCell(2)),  // Numero ODS (C)
                            dOds,                                       // Data ODS (D)
                            getCellValueAsDate(row.getCell(4)),         // Scadenza (E)
                            formatter.formatCellValue(row.getCell(5)),  // Via (F)
                            formatter.formatCellValue(row.getCell(6)),  // Danneggiante (G)
                            descrizioneCompleta,                        // Descrizione (H+M)
                            dInizio,                                    // Inizio (K o D)
                            dFine                                       // Fine (L o D)
                    );

                    listaDatiExcel.add(dati);
                }

                if (!listaDatiExcel.isEmpty()) {
                    indiceCorrente = 0;
                    popolaCampi(listaDatiExcel.get(0));
                    aggiornaStatoBottoni();
                }

            } catch (Exception ex) {
                JOptionPane.showMessageDialog(this, "Errore: " + ex.getMessage());
                ex.printStackTrace();
            }
        }
    }

    private Date getCellValueAsDate(Cell cell) {
        if (cell == null || cell.getCellType() == CellType.BLANK) return null;

        // 1. Prova se è un formato data numerico (standard Excel)
        if (cell.getCellType() == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue();
            }
        }

        // 2. Prova se è una stringa (testo scritto gg/mm/aaaa)
        if (cell.getCellType() == CellType.STRING) {
            String val = cell.getStringCellValue().trim();
            if (!val.isEmpty()) {
                try {
                    return sdf.parse(val);
                } catch (ParseException e) {
                    System.err.println("Errore conversione stringa in data: " + val);
                }
            }
        }
        return null;
    }

    private void mostraProssimoDato() {
        if (indiceCorrente < listaDatiExcel.size() - 1) {
            indiceCorrente++;
            popolaCampi(listaDatiExcel.get(indiceCorrente));
            aggiornaStatoBottoni();
        }
    }

    private void mostraDatoPrecedente() {
        if (indiceCorrente > 0) {
            indiceCorrente--;
            popolaCampi(listaDatiExcel.get(indiceCorrente));
            aggiornaStatoBottoni();
        }
    }

    private void aggiornaStatoBottoni() {
        precedenteButton.setEnabled(indiceCorrente > 0);
        prossimoButton.setEnabled(indiceCorrente < listaDatiExcel.size() - 1);
        infoExcelLabel.setText("Riga " + (indiceCorrente + 1) + " di " + listaDatiExcel.size());
    }

    private void popolaCampi(Allegati a) {
        numeroOdsField.setText(a.getNumeroOds());
        dataOdsField.setText(a.getDataOds() != null ? sdf.format(a.getDataOds()) : "");
        scadenzaOdsField.setText(a.getScadenzaOds() != null ? sdf.format(a.getScadenzaOds()) : "");
        viaField.setText(a.getVia());
        danneggianteField.setText(a.getDanneggiante());
        descrizioneInterventoField.setText(a.getDescrizioneIntervento());
        inizioLavoriField.setText(a.getInizioLavori() != null ? sdf.format(a.getInizioLavori()) : "");
        fineLavoriField.setText(a.getFineLavori() != null ? sdf.format(a.getFineLavori()) : "");

        scaricaButton.setEnabled(false);
        lastCompiledFilePath = null;
    }

    private void compilePdf() {
        try {
            Allegati dati = new Allegati(
                    numeroOdsField.getText(),
                    parseDate(dataOdsField.getText()),
                    parseDate(scadenzaOdsField.getText()),
                    viaField.getText(),
                    danneggianteField.getText(),
                    descrizioneInterventoField.getText(),
                    parseDate(inizioLavoriField.getText()),
                    parseDate(fineLavoriField.getText())
            );

            File tempOutputFile = File.createTempFile("compiled_cristoforo_", ".pdf");
            tempOutputFile.deleteOnExit();

            PdfFiller filler = new PdfFiller();
            filler.fillPdfSpecificFields(tempOutputFile.getAbsolutePath(), dati);

            lastCompiledFilePath = tempOutputFile.getAbsolutePath();
            scaricaButton.setEnabled(true);

            JOptionPane.showMessageDialog(this, "PDF generato correttamente!", "Successo", JOptionPane.INFORMATION_MESSAGE);

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Errore nella generazione: " + ex.getMessage(), "Errore", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void downloadPdf() {
        if (lastCompiledFilePath == null) return;
        JFileChooser fileChooser = new JFileChooser(System.getProperty("user.home") + File.separator + "Desktop");
        String defaultName = numeroOdsField.getText().replaceAll("[^a-zA-Z0-9.-]", "_") + ".pdf";
        fileChooser.setSelectedFile(new File(defaultName));
        if (fileChooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            try {
                java.nio.file.Files.copy(new File(lastCompiledFilePath).toPath(),
                        fileChooser.getSelectedFile().toPath(),
                        java.nio.file.StandardCopyOption.REPLACE_EXISTING);
                JOptionPane.showMessageDialog(this, "File salvato con successo!");
            } catch (IOException e) {
                JOptionPane.showMessageDialog(this, "Errore salvataggio: " + e.getMessage());
            }
        }
    }

    private Date parseDate(String dateString) throws ParseException {
        if (dateString == null || dateString.trim().isEmpty()) return null;
        return sdf.parse(dateString);
    }

    private void clearFields() {
        numeroOdsField.setText(""); dataOdsField.setText(""); scadenzaOdsField.setText("");
        viaField.setText(""); danneggianteField.setText(""); descrizioneInterventoField.setText("");
        inizioLavoriField.setText(""); fineLavoriField.setText("");
        scaricaButton.setEnabled(false);
        lastCompiledFilePath = null;
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new Lavoro().setVisible(true));
    }
}
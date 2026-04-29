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
import java.util.stream.Collectors;

public class Lavoro extends JFrame {

    private JPanel contentPane;
    private JTextField numeroOdsField, dataOdsField, scadenzaOdsField, viaField,
            danneggianteField, descrizioneInterventoField, inizioLavoriField, fineLavoriField;
    private JTextField cercaOdsField;

    private JButton scaricaButton, compilaButton, caricaExcelButton, prossimoButton, precedenteButton, pulisciCampiButton, cercaButton;
    private JCheckBox soloProntoInterventoCheckBox;
    private JLabel infoExcelLabel;

    private List<Allegati> listaDatiExcel = new ArrayList<>();
    private List<Allegati> listaAttuale = new ArrayList<>();
    private int indiceCorrente = -1;
    private String lastCompiledFilePath;
    private SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");

    public Lavoro() {
        super("Compilatore PDF da Excel");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        setPreferredSize(new Dimension(1000, 750));
        setMinimumSize(new Dimension(900, 650));

        contentPane = new JPanel(new BorderLayout(15, 15));
        contentPane.setBorder(new EmptyBorder(20, 20, 20, 20));

        JPanel topContainer = new JPanel(new GridLayout(2, 1, 5, 5));
        topContainer.setBorder(BorderFactory.createTitledBorder("Strumenti di Navigazione e Ricerca"));

        JPanel loadPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 15, 5));
        caricaExcelButton = new JButton("Carica File Excel");
        precedenteButton = new JButton("<< Precedente");
        prossimoButton = new JButton("Prossimo >>");
        soloProntoInterventoCheckBox = new JCheckBox("Filtra Pronto Intervento");
        infoExcelLabel = new JLabel("Nessun file caricato");

        loadPanel.add(caricaExcelButton);
        loadPanel.add(precedenteButton);
        loadPanel.add(prossimoButton);
        loadPanel.add(soloProntoInterventoCheckBox);
        loadPanel.add(infoExcelLabel);

        JPanel searchPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 15, 5));
        searchPanel.add(new JLabel("Cerca Numero O.d.S.:"));
        cercaOdsField = new JTextField(20);
        cercaButton = new JButton("Cerca e Vai");
        searchPanel.add(cercaOdsField);
        searchPanel.add(cercaButton);

        topContainer.add(loadPanel);
        topContainer.add(searchPanel);
        contentPane.add(topContainer, BorderLayout.NORTH);

        JPanel dataInputPanel = new JPanel(new GridLayout(0, 2, 10, 15));
        dataInputPanel.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createTitledBorder("Dati Estrazione"),
                new EmptyBorder(15, 15, 15, 15)
        ));

        java.awt.Font labelFont = new java.awt.Font("SansSerif", java.awt.Font.BOLD, 14);

        numeroOdsField = createStyledTextField();
        dataOdsField = createStyledTextField();
        scadenzaOdsField = createStyledTextField();
        viaField = createStyledTextField();
        danneggianteField = createStyledTextField();
        descrizioneInterventoField = createStyledTextField();
        inizioLavoriField = createStyledTextField();
        fineLavoriField = createStyledTextField();

        addLabelAndField(dataInputPanel, "Numero O.d.S.:", numeroOdsField, labelFont);
        addLabelAndField(dataInputPanel, "Data O.d.S. (gg/mm/aaaa):", dataOdsField, labelFont);
        addLabelAndField(dataInputPanel, "Scadenza O.d.S. (gg/mm/aaaa):", scadenzaOdsField, labelFont);
        addLabelAndField(dataInputPanel, "Via:", viaField, labelFont);
        addLabelAndField(dataInputPanel, "Danneggiante:", danneggianteField, labelFont);
        addLabelAndField(dataInputPanel, "Descrizione Intervento:", descrizioneInterventoField, labelFont);
        addLabelAndField(dataInputPanel, "Data Inizio Lavori (gg/mm/aaaa):", inizioLavoriField, labelFont);
        addLabelAndField(dataInputPanel, "Data Fine Lavori (gg/mm/aaaa):", fineLavoriField, labelFont);

        contentPane.add(dataInputPanel, BorderLayout.CENTER);

        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 15, 10));
        compilaButton = new JButton("Genera PDF");
        scaricaButton = new JButton("Salva sul Desktop");
        pulisciCampiButton = new JButton("Pulisci Campi");
        JButton esciButton = new JButton("Esci");

        precedenteButton.setEnabled(false);
        prossimoButton.setEnabled(false);
        scaricaButton.setEnabled(false);
        soloProntoInterventoCheckBox.setEnabled(false);

        buttonPanel.add(compilaButton);
        buttonPanel.add(scaricaButton);
        buttonPanel.add(pulisciCampiButton);
        buttonPanel.add(esciButton);
        contentPane.add(buttonPanel, BorderLayout.SOUTH);

        caricaExcelButton.addActionListener(e -> importaExcel());
        prossimoButton.addActionListener(e -> mostraProssimoDato());
        precedenteButton.addActionListener(e -> mostraDatoPrecedente());
        soloProntoInterventoCheckBox.addActionListener(e -> applicaFiltro());
        cercaButton.addActionListener(e -> cercaOds());
        compilaButton.addActionListener(e -> compilePdf());
        scaricaButton.addActionListener(e -> downloadPdf());
        pulisciCampiButton.addActionListener(e -> clearFields());
        esciButton.addActionListener(e -> System.exit(0));

        add(contentPane);
        pack();
        setLocationRelativeTo(null);
    }

    private JTextField createStyledTextField() {
        JTextField tf = new JTextField();
        tf.setFont(new java.awt.Font("Monospaced", java.awt.Font.PLAIN, 15));
        return tf;
    }

    private void addLabelAndField(JPanel panel, String labelText, JTextField field, java.awt.Font font) {
        JLabel l = new JLabel(labelText);
        l.setFont(font);
        panel.add(l);
        panel.add(field);
    }

    private void cercaOds() {
        String query = cercaOdsField.getText().trim();
        if (query.isEmpty() || listaAttuale.isEmpty()) return;

        for (int i = 0; i < listaAttuale.size(); i++) {
            if (listaAttuale.get(i).getNumeroOds().equalsIgnoreCase(query)) {
                indiceCorrente = i;
                popolaCampi(listaAttuale.get(i));
                aggiornaStatoBottoni();
                return;
            }
        }
        JOptionPane.showMessageDialog(this, "Numero ODS '" + query + "' non trovato.");
    }

    private void importaExcel() {
        String desktopPath = System.getProperty("user.home") + File.separator + "Desktop";
        JFileChooser fileChooser = new JFileChooser(desktopPath);
        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            try (FileInputStream fis = new FileInputStream(fileChooser.getSelectedFile());
                 Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheetAt(0);
                listaDatiExcel.clear();
                DataFormatter formatter = new DataFormatter();

                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;

                    // Recupero campi chiave
                    String odsNum = formatter.formatCellValue(row.getCell(2)).trim();
                    String via = formatter.formatCellValue(row.getCell(5)).trim();
                    String dannegg = formatter.formatCellValue(row.getCell(6)).trim();

                    // MODIFICA: Salta solo se la riga è completamente priva di dati essenziali
                    if (odsNum.isEmpty() && via.isEmpty() && dannegg.isEmpty()) continue;

                    // Se l'ODS manca, creiamo un ID basato sulla riga per permettere la visualizzazione
                    if (odsNum.isEmpty()) {
                        odsNum = "RIGA-" + (i + 1);
                    }

                    String note = formatter.formatCellValue(row.getCell(12, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)).trim();
                    Date dOds = getCellValueAsDate(row.getCell(3));

                    // Gestione automatica date per Pronto Intervento
                    Date dInizio, dFine;
                    if (note.toUpperCase().contains("PRONTO INTERVENTO")) {
                        dInizio = dOds;
                        dFine = dOds;
                    } else {
                        dInizio = getCellValueAsDate(row.getCell(10));
                        dFine = getCellValueAsDate(row.getCell(11));
                    }

                    listaDatiExcel.add(new Allegati(
                            odsNum,
                            dOds,
                            getCellValueAsDate(row.getCell(4)),
                            via,
                            dannegg,
                            formatter.formatCellValue(row.getCell(7)).trim() + (note.isEmpty() ? "" : " - " + note),
                            dInizio,
                            dFine
                    ));
                }

                if (!listaDatiExcel.isEmpty()) {
                    soloProntoInterventoCheckBox.setEnabled(true);
                    applicaFiltro();
                    JOptionPane.showMessageDialog(this, "Caricate con successo " + listaDatiExcel.size() + " righe.");
                } else {
                    JOptionPane.showMessageDialog(this, "Nessun dato trovato nel file.");
                }

            } catch (Exception ex) {
                JOptionPane.showMessageDialog(this, "Errore caricamento: " + ex.getMessage());
                ex.printStackTrace();
            }
        }
    }

    private void applicaFiltro() {
        if (soloProntoInterventoCheckBox.isSelected()) {
            listaAttuale = listaDatiExcel.stream()
                    .filter(a -> a.getDescrizioneIntervento().toUpperCase().contains("PRONTO INTERVENTO"))
                    .collect(Collectors.toList());
            if (listaAttuale.isEmpty()) {
                JOptionPane.showMessageDialog(this, "Nessun Pronto Intervento trovato.");
                soloProntoInterventoCheckBox.setSelected(false);
                listaAttuale = new ArrayList<>(listaDatiExcel);
            }
        } else {
            listaAttuale = new ArrayList<>(listaDatiExcel);
        }
        indiceCorrente = listaAttuale.isEmpty() ? -1 : 0;
        if (indiceCorrente != -1) popolaCampi(listaAttuale.get(0));
        aggiornaStatoBottoni();
    }

    private void aggiornaStatoBottoni() {
        precedenteButton.setEnabled(indiceCorrente > 0);
        prossimoButton.setEnabled(indiceCorrente < listaAttuale.size() - 1);
        String modo = soloProntoInterventoCheckBox.isSelected() ? " [MODALITÀ P.I.]" : " [TUTTI]";
        infoExcelLabel.setText("Record: " + (indiceCorrente + 1) + " / " + listaAttuale.size() + modo);
    }

    private void mostraProssimoDato() {
        if (indiceCorrente < listaAttuale.size() - 1) {
            indiceCorrente++;
            popolaCampi(listaAttuale.get(indiceCorrente));
            aggiornaStatoBottoni();
        }
    }

    private void mostraDatoPrecedente() {
        if (indiceCorrente > 0) {
            indiceCorrente--;
            popolaCampi(listaAttuale.get(indiceCorrente));
            aggiornaStatoBottoni();
        }
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
    }

    private Date getCellValueAsDate(Cell cell) {
        if (cell == null || cell.getCellType() == CellType.BLANK) return null;
        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) return cell.getDateCellValue();
        if (cell.getCellType() == CellType.STRING) {
            try { return sdf.parse(cell.getStringCellValue().trim()); } catch (ParseException e) { }
        }
        return null;
    }

    private Date parseDate(String s) throws ParseException {
        return (s == null || s.trim().isEmpty()) ? null : sdf.parse(s);
    }

    private void clearFields() {
        numeroOdsField.setText(""); dataOdsField.setText(""); scadenzaOdsField.setText("");
        viaField.setText(""); danneggianteField.setText(""); descrizioneInterventoField.setText("");
        inizioLavoriField.setText(""); fineLavoriField.setText("");
        scaricaButton.setEnabled(false);
    }

    private void compilePdf() {
        try {
            Allegati dati = new Allegati(numeroOdsField.getText(), parseDate(dataOdsField.getText()),
                    parseDate(scadenzaOdsField.getText()), viaField.getText(), danneggianteField.getText(),
                    descrizioneInterventoField.getText(), parseDate(inizioLavoriField.getText()), parseDate(fineLavoriField.getText()));
            File temp = File.createTempFile("preview_", ".pdf");
            temp.deleteOnExit();
            new PdfFiller().fillPdfSpecificFields(temp.getAbsolutePath(), dati);
            lastCompiledFilePath = temp.getAbsolutePath();
            scaricaButton.setEnabled(true);
            JOptionPane.showMessageDialog(this, "Compilazione completata!");
        } catch (Exception ex) { JOptionPane.showMessageDialog(this, "Errore: " + ex.getMessage()); }
    }

    private void downloadPdf() {
        if (lastCompiledFilePath == null) {
            JOptionPane.showMessageDialog(this, "Nessun PDF generato da salvare.");
            return;
        }

        // 1. Ottieni il percorso del Desktop
        String desktopPath = System.getProperty("user.home") + File.separator + "Desktop";

        // 2. Pulisci il nome del file dai caratteri vietati (es. ODS 123/A -> ODS 123_A)
        String numeroOds = numeroOdsField.getText().trim().isEmpty() ? "Documento" : numeroOdsField.getText();
        String safeName = numeroOds.replaceAll("[\\\\/:*?\"<>|]", "_");

        // 3. Crea il file di destinazione finale
        File destinazione = new File(desktopPath + File.separator + safeName + ".pdf");

        try {
            // Controllo se il file esiste già per non sovrascrivere senza avvisare (opzionale)
            if (destinazione.exists()) {
                int response = JOptionPane.showConfirmDialog(this,
                        "Il file " + safeName + ".pdf esiste già sul desktop. Sovrascrivere?",
                        "Conferma sovrascrittura",
                        JOptionPane.YES_NO_OPTION);
                if (response != JOptionPane.YES_OPTION) return;
            }

            // 4. Copia effettiva del file dal percorso temporaneo al desktop
            java.nio.file.Files.copy(
                    new File(lastCompiledFilePath).toPath(),
                    destinazione.toPath(),
                    java.nio.file.StandardCopyOption.REPLACE_EXISTING
            );

            JOptionPane.showMessageDialog(this, "File salvato correttamente sul Desktop:\n" + destinazione.getName());

        } catch (IOException e) {
            JOptionPane.showMessageDialog(this, "Errore durante il salvataggio automatico: " + e.getMessage());
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new Lavoro().setVisible(true));
    }
}
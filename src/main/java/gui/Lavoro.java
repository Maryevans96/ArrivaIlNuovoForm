package gui;

import model.Allegati;
import model.PdfFiller;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class Lavoro extends JFrame {

    private JPanel contentPane;
    private JTextField numeroOdsField, dataOdsField, scadenzaOdsField, viaField,
            danneggianteField, descrizioneInterventoField, inizioLavoriField, fineLavoriField;

    private JButton scaricaButton;
    private JButton compilaButton;
    private JButton pulisciCampiButton;

    private String lastCompiledFilePath;
    private SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");

    public Lavoro() {
        super("Compilatore PDF - SCHEDA CRISTOFOROECO");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        contentPane = new JPanel(new BorderLayout(10, 10));
        contentPane.setBorder(new EmptyBorder(15, 15, 15, 15));

        // Pannello Titolo (sostituisce la selezione modello che non serve più)
        JPanel headerPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        headerPanel.setBorder(BorderFactory.createTitledBorder("Modello in uso"));
        headerPanel.add(new JLabel("Utilizzando: SCHEDA CRISTOFOROECO.pdf"));
        contentPane.add(headerPanel, BorderLayout.NORTH);

        // Pannello Input Dati
        JPanel dataInputPanel = new JPanel(new GridLayout(0, 2, 5, 5));
        dataInputPanel.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createTitledBorder("Dati per la Compilazione"),
                new EmptyBorder(10, 10, 10, 10)
        ));

        // Inizializzazione campi
        numeroOdsField = new JTextField(20);
        dataOdsField = new JTextField(20);
        scadenzaOdsField = new JTextField(20);
        viaField = new JTextField(20);
        danneggianteField = new JTextField(20);
        descrizioneInterventoField = new JTextField(20);
        inizioLavoriField = new JTextField(20);
        fineLavoriField = new JTextField(20);

        // Aggiunta componenti al pannello
        addLabelAndField(dataInputPanel, "Numero O.d.S.:", numeroOdsField);
        addLabelAndField(dataInputPanel, "Data O.d.S. (gg/mm/aaaa):", dataOdsField);
        addLabelAndField(dataInputPanel, "Scadenza O.d.S. (gg/mm/aaaa):", scadenzaOdsField);
        addLabelAndField(dataInputPanel, "Via:", viaField);
        addLabelAndField(dataInputPanel, "Danneggiante:", danneggianteField);
        addLabelAndField(dataInputPanel, "Descrizione Intervento:", descrizioneInterventoField);
        addLabelAndField(dataInputPanel, "Data Inizio Lavori (gg/mm/aaaa):", inizioLavoriField);
        addLabelAndField(dataInputPanel, "Data Fine Lavori (gg/mm/aaaa):", fineLavoriField);

        contentPane.add(dataInputPanel, BorderLayout.CENTER);

        // Pannello Bottoni
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

    private void compilePdf() {
        try {
            // Raccolta dati
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

            // Creazione file temporaneo per l'output
            File tempOutputFile = File.createTempFile("compiled_cristoforo_", ".pdf");
            tempOutputFile.deleteOnExit();

            // Chiamata al nuovo PdfFiller (che ora sa già quale template usare)
            PdfFiller filler = new PdfFiller();
            filler.fillPdfSpecificFields(tempOutputFile.getAbsolutePath(), dati);

            lastCompiledFilePath = tempOutputFile.getAbsolutePath();
            scaricaButton.setEnabled(true);

            JOptionPane.showMessageDialog(this, "PDF compilato con successo!", "Successo", JOptionPane.INFORMATION_MESSAGE);

        } catch (ParseException ex) {
            JOptionPane.showMessageDialog(this, "Errore formato data (gg/mm/aaaa).", "Errore Input", JOptionPane.ERROR_MESSAGE);
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this, "Errore file: " + ex.getMessage(), "Errore I/O", JOptionPane.ERROR_MESSAGE);
        }
    }

    // ... (Il metodo downloadPdf(), parseDate() e clearFields() rimangono identici a prima)
    // Li ometto per brevità, ma mantienili nel tuo file.

    private void downloadPdf() {
        if (lastCompiledFilePath == null) return;
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setSelectedFile(new File(numeroOdsField.getText().replaceAll("[^a-zA-Z0-9.-]", "_") + ".pdf"));
        if (fileChooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            try {
                java.nio.file.Files.copy(new File(lastCompiledFilePath).toPath(),
                        fileChooser.getSelectedFile().toPath(),
                        java.nio.file.StandardCopyOption.REPLACE_EXISTING);
                JOptionPane.showMessageDialog(this, "File salvato correttamente!");
            } catch (IOException e) {
                JOptionPane.showMessageDialog(this, "Errore nel salvataggio.");
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
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new Lavoro().setVisible(true));
    }
}
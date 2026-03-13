package model;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;

public class PdfFiller {

    // Nome del file fisso nelle risorse
    private static final String TEMPLATE_NAME = "SCHEDA CRISTOFOROECO.pdf";

    public void fillPdfSpecificFields(String outputPath, Allegati datiAllegato) throws IOException {
        PDDocument pdfDocument = null;

        // Caricamento del file dalle risorse del progetto
        try (InputStream is = getClass().getClassLoader().getResourceAsStream(TEMPLATE_NAME)) {
            if (is == null) {
                throw new IOException("Errore: Il file " + TEMPLATE_NAME + " non è stato trovato nella cartella resources.");
            }

            pdfDocument = PDDocument.load(is);

            if (pdfDocument.getNumberOfPages() < 2) {
                System.err.println("Il documento PDF non ha almeno due pagine.");
                return;
            }

            PDPage firstPage = pdfDocument.getPage(0);
            PDPage secondPage = pdfDocument.getPage(1);

            // Apertura degli stream di contenuto
            try (PDPageContentStream contentStreamFirstPage = new PDPageContentStream(pdfDocument, firstPage, PDPageContentStream.AppendMode.APPEND, true, true);
                 PDPageContentStream contentStreamSecondPage = new PDPageContentStream(pdfDocument, secondPage, PDPageContentStream.AppendMode.APPEND, true, true)) {

                SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");

                // --- Compilazione Prima Pagina ---
                addTextToPdf(contentStreamFirstPage, datiAllegato.getNumeroOds(), 470, 683, 10);
                addTextToPdf(contentStreamFirstPage, (datiAllegato.getDataOds() != null) ? sdf.format(datiAllegato.getDataOds()) : "", 267, 697, 10);
                addTextToPdf(contentStreamFirstPage, (datiAllegato.getScadenzaOds() != null) ? sdf.format(datiAllegato.getScadenzaOds()) : "", 267, 683, 10);
                addTextToPdf(contentStreamFirstPage, datiAllegato.getVia(), 220, 670, 10);
                addTextToPdf(contentStreamFirstPage, datiAllegato.getDanneggiante(), 165, 645, 10);
                addTextToPdf(contentStreamFirstPage, datiAllegato.getDescrizioneIntervento(), 195, 560, 10);
                addTextToPdf(contentStreamFirstPage, (datiAllegato.getInizioLavori() != null) ? sdf.format(datiAllegato.getInizioLavori()) : "", 200, 223, 10);
                addTextToPdf(contentStreamFirstPage, (datiAllegato.getFineLavori() != null) ? sdf.format(datiAllegato.getFineLavori()) : "", 470, 223, 10);

                // --- Compilazione Seconda Pagina ---
                addTextToPdf(contentStreamSecondPage, datiAllegato.getNumeroOds(), 60, 337, 10);
                addTextToPdf(contentStreamSecondPage, (datiAllegato.getDataOds() != null) ? sdf.format(datiAllegato.getDataOds()) : "", 125, 337, 10);
                addTextToPdf(contentStreamSecondPage, datiAllegato.getVia(), 190, 337, 10);
                addTextToPdf(contentStreamSecondPage, datiAllegato.getDescrizioneIntervento(), 320, 337, 10);
                addTextToPdf(contentStreamSecondPage, (datiAllegato.getInizioLavori() != null) ? sdf.format(datiAllegato.getInizioLavori()) : "", 490, 277, 10);
                addTextToPdf(contentStreamSecondPage, (datiAllegato.getFineLavori() != null) ? sdf.format(datiAllegato.getFineLavori()) : "", 490, 252, 10);
            }

            // Salvataggio del file compilato nel percorso di output specificato
            pdfDocument.save(outputPath);

        } finally {
            if (pdfDocument != null) {
                pdfDocument.close();
            }
        }
    }

    private void addTextToPdf(PDPageContentStream contentStream, String text, float x, float y, float fontSize) throws IOException {
        if (text == null || text.isEmpty()) return;
        contentStream.beginText();
        contentStream.setFont(PDType1Font.HELVETICA, fontSize);
        contentStream.newLineAtOffset(x, y);
        contentStream.showText(text);
        contentStream.endText();
    }
}
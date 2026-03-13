package model;

import java.util.Date;

public class Allegati {
    String numeroOds;
    Date dataOds;
    Date scadenzaOds;
    String via;
    String danneggiante;
    String descrizioneIntervento;
    Date inizioLavori;
    Date fineLavori;

    public Allegati(String numeroOds, Date dataOds, Date scadenzaOds, String via, String danneggiante, String descrizioneIntervento,
                    Date inizioLavori, Date fineLavori) {

        this.numeroOds = numeroOds;
        this.dataOds = dataOds;
        this.scadenzaOds = scadenzaOds;
        this.via = via;
        this.danneggiante = danneggiante;
        this.descrizioneIntervento = descrizioneIntervento;
        this.inizioLavori = inizioLavori;
        this.fineLavori = fineLavori;
    }

    public String getNumeroOds() {
        return numeroOds;
    }

    public void setNumeroOds(String numeroOds) {
        this.numeroOds = numeroOds;
    }

    public Date getDataOds() {
        return dataOds;
    }

    public void setDataOds(Date dataOds) {
        this.dataOds = dataOds;
    }

    public Date getScadenzaOds() {
        return scadenzaOds;
    }

    public void setScadenzaOds(Date scadenzaOds) {
        this.scadenzaOds = scadenzaOds;
    }

    public String getVia() {
        return via;
    }

    public void setVia(String via) {
        this.via = via;
    }

    public String getDanneggiante() {
        return danneggiante;
    }

    public void setDanneggiante(String danneggiante) {
        this.danneggiante = danneggiante;
    }

    public String getDescrizioneIntervento() {
        return descrizioneIntervento;
    }

    public void setDescrizioneIntervento(String descrizioneIntervento) {
        this.descrizioneIntervento = descrizioneIntervento;
    }

    public Date getInizioLavori() {
        return inizioLavori;
    }

    public void setInizioLavori(Date inizioLavori) {
        this.inizioLavori = inizioLavori;
    }

    public Date getFineLavori() {
        return fineLavori;
    }

    public void setFineLavori(Date fineLavori) {
        this.fineLavori = fineLavori;
    }
}

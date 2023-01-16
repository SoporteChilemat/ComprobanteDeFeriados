/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package Clases;

/**
 *
 * @author DELL
 */
public class Personas {

    String rut;
    String fecha;
    String fechaComprobante;
    int saldoHab;
    int diasProg;
    String fechaInicio;

    public String getFechaInicio() {
        return fechaInicio;
    }

    public void setFechaInicio(String fechaInicio) {
        this.fechaInicio = fechaInicio;
    }

    public String getRut() {
        return rut;
    }

    public void setRut(String rut) {
        this.rut = rut;
    }

    public String getFecha() {
        return fecha;
    }

    public void setFecha(String fecha) {
        this.fecha = fecha;
    }

    public int getSaldoHab() {
        return saldoHab;
    }

    public void setSaldoHab(int saldoHab) {
        this.saldoHab = saldoHab;
    }

    public int getDiasProg() {
        return diasProg;
    }

    public void setDiasProg(int diasProg) {
        this.diasProg = diasProg;
    }

    public String getFechaComprobante() {
        return fechaComprobante;
    }

    public void setFechaComprobante(String fechaComprobante) {
        this.fechaComprobante = fechaComprobante;
    }
}

/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Project/Maven2/JavaApp/src/main/java/${packagePath}/${mainClassName}.java to edit this template
 */
package Logica;

import Clases.Personas;
import Clases.PersonasInicio;
import Ventanas.VentanaPrincipal;
import io.github.bonigarcia.wdm.WebDriverManager;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Consumer;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.io.FileUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

/**
 *
 * @author sopor
 */
public class Feriados {

// declaraciones globales de variables    
    public static WebDriver driver;
    public static WebDriverWait wait;
    public static ArrayList<Personas> arrPersonas = new ArrayList<>();
    public static boolean aviso = true;

    public static void principal() throws IOException {

        //System.out.println(transformarExcel.size());
        VentanaPrincipal.leerydescargarbutton.setEnabled(true);//boton que da inicio al programa se encuentra activado

        CompletableFuture.runAsync(() -> { // utilización de hilos
//            arrPersonas = new ArrayList<>(); // creacion del array global
//            arrPersonasInicio = new ArrayList<>();

            try { // detecta y controla cualquier incidencia
                ArrayList<PersonasInicio> arrPersonasListado = transformarExcelaArray(System.getProperty("user.dir") + "//LISTADO FINAL.xlsx");
                ArrayList<Personas> descargarLectura = descargaLectura(arrPersonas, arrPersonasListado);// creo un arrayList que será lo que retorna el metodo descargaLectura
                ArrayList<Personas> manejoFecha = manejoFecha(descargarLectura);
                if (aviso == false) {
                    VentanaPrincipal.info.setText("Proceso terminado!");
                    VentanaPrincipal.leerydescargarbutton.setEnabled(true);
                }
                if (aviso == true) {
                    crearExcel(manejoFecha);//metodo para crear y llenar el excel
                }
            } catch (IOException ex) {
                Logger.getLogger(Feriados.class.getName()).log(Level.SEVERE, null, ex);
            }
        });

    }
//----------------------------------------------------------------------------------------------------------------------------------------

    public static int devuelveAño(String fecha) {
        String[] split = fecha.split("/");
        String name = split[2];
        int parseInt = Integer.parseInt(name);
        return parseInt;
    }

//------------------------------------------------------------------------------------------------------------------------------------------
    public static void crearExcel(ArrayList<Personas> arrPersonas) throws FileNotFoundException, IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet spreadsheet = workbook.createSheet("Facturas");

        XSSFFont headerFont = workbook.createFont();
        headerFont.setColor(IndexedColors.WHITE.index);
        CellStyle headerCellStyle = spreadsheet.getWorkbook().createCellStyle();
        // fill foreground color ...
        headerCellStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
        // and solid fill pattern produces solid grey cell fill
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerCellStyle.setFont(headerFont);

        Map<String, Object[]> data = new TreeMap<>();
        data.put("1", new Object[]{"Rut", "Fecha", "Saldo Habil", "Dias Programados", "Fecha Comprobante"});

        // Iterate over data and write to sheet 
        VentanaPrincipal.info.setText("Generando excel");
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset) {
            // this creates a new row in the sheet 
            Row row = spreadsheet.createRow(rownum++);
            Object[] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr) {
                // this line creates a cell in the next column of that row 
                Cell cell = row.createCell(cellnum++);
                // if rownum is 1 (first row was created before) then set header CellStyle
                if (rownum == 1) {
                    cell.setCellStyle(headerCellStyle);
                }
                if (obj instanceof String) {
                    cell.setCellValue((String) obj);
                } else if (obj instanceof Integer) {
                    cell.setCellValue((Integer) obj);
                }
            }
        }
        int filaInicio = 1;
        for (int f = 0; f < arrPersonas.size(); f++) {
            Row fila = spreadsheet.createRow(filaInicio);
            filaInicio++;
            Personas get = arrPersonas.get(f);

            for (int c = 0; c < 5; c++) {

//                Cell celda = fila.createCell(c);
//                celda.setCellValue(String.valueOf(c));
                if (c == 0) {
                    Cell celda = fila.createCell(c);
                    celda.setCellValue(String.valueOf(get.getRut()));
                } else if (c == 1) {
                    Cell celda = fila.createCell(c);
                    celda.setCellValue(String.valueOf(get.getFecha()));

                } else if (c == 2) {
                    Cell celda = fila.createCell(c);
                    celda.setCellValue(String.valueOf(get.getSaldoHab()));

                } else if (c == 3) {
                    Cell celda = fila.createCell(c);
                    celda.setCellValue(String.valueOf(get.getDiasProg()));

                } else if (c == 4) {
                    Cell celda = fila.createCell(c);
                    celda.setCellValue(String.valueOf(get.getFechaComprobante()));
                }

            }
        }

        FileOutputStream out = new FileOutputStream(new File("ListadoFacturas.xlsx"));
        workbook.write(out);

        VentanaPrincipal.leerydescargarbutton.setEnabled(true);
        out.close();

        VentanaPrincipal.info.setText("Excel Generado!");
    }
//---------------------------------------------------------------------------------------------------------------------------

    public static ArrayList<Personas> descargaLectura(ArrayList<Personas> arrPersonas, ArrayList<PersonasInicio> arrPersonasListado) throws IOException {
        VentanaPrincipal.leerydescargarbutton.setEnabled(false);
        VentanaPrincipal.progresonum.setText("");
        driver = chromeOption();
        String[] split = null;

        wait = new WebDriverWait(driver, Duration.ofSeconds(5));

        String baseURL = "https://web.nubox.com/Login/Account/Login?ReturnUrl=%2fSistemaLogin";
        String rut = "171193346";
        String password = "302905";
        driver.get(baseURL);

        VentanaPrincipal.info.setText("Ingresando a Nubox!");

        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"ae740e71936fa3eec403935de72a7aa3a68bbe7\"]")));
        driver.findElement(By.xpath("//*[@id=\"ae740e71936fa3eec403935de72a7aa3a68bbe7\"]")).sendKeys(rut);
        driver.findElement(By.xpath("//*[@id=\"d70911c2de484460cf9f927ee6c6166585718189\"]")).sendKeys(password);
        driver.findElement(By.xpath("//*[@id=\"nuboxForm\"]/input[2]")).click();

        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"row2treeGrid\"]")));
        driver.findElement(By.xpath("//*[@id=\"row2treeGrid\"]")).click();

        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/table/tbody/tr[1]/td/table[1]/tbody/tr/td[2]/img")));
        driver.findElement(By.xpath("//*[@id=\"BarraMenu2\"]/a")).click();

        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"SubBarraMenu2\"]/tbody/tr/td[2]/a")));
        driver.findElement(By.xpath("//*[@id=\"SubBarraMenu2\"]/tbody/tr/td[2]/a")).click();

        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"SubBarraMenu2_2\"]/tbody/tr/td[7]/a")));
        driver.findElement(By.xpath("//*[@id=\"SubBarraMenu2_2\"]/tbody/tr/td[7]/a")).click();

        int length = 0;
        for (int q = 0; q < arrPersonasListado.size(); q++) {

            if (aviso == false) {
                VentanaPrincipal.info.setText("Proceso interrumpido !");
                VentanaPrincipal.leerydescargarbutton.setEnabled(true);
                break;
            }
            PersonasInicio loquito = arrPersonasListado.get(q);

            try {
                try {
                    driver.switchTo().frame("Contenido");
                } catch (Exception ex) {
//                    Logger.getLogger(Feriados.class.getName()).log(Level.SEVERE, null, ex);
                }

                if (q == 0) {
                    wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"Funcionario\"]")));
                    String text = driver.findElement(By.xpath("//*[@id=\"Funcionario\"]")).getText();
                    //*[@id="CodigoFuncionario"]
                    //*[@id="Funcionario"]

                    System.out.println("texteate " + text);

                    split = text.split("\n");

                    VentanaPrincipal.progresonashe.setMaximum(arrPersonasListado.size());
                    VentanaPrincipal.info.setText("Leyendo informacion");
                }

                String trim = split[q].trim();
                System.out.println("i " + q);
                System.out.println("-----------------------------------> " + trim);

                length = split.length;
                System.out.println("length " + length);
                //--------------------------------- Progreso en numeros---------------------------------------------                  
                String progresonum = String.valueOf(arrPersonasListado.size());

                VentanaPrincipal.progresonum.setText(q + 1 + " / " + progresonum);
                ArrayList<Integer> indicesIguales = new ArrayList<>();
                ArrayList<String> fechasInicio = new ArrayList<>();
                //------------------------FOR PARA ---------------------------------------------------------------------
                for (int k = 0; k < split.length; k++) {
                    String nombrex = split[k];
                    String nombretrim = nombrex.trim();
                    String nombres = loquito.getNombres();

//                    System.out.println("nombres " + nombres);
                    String[] split1 = nombres.split(" ");

                    String trim1 = "";
                    try {
                        trim1 = split1[0].trim();
                    } catch (Exception ex) {

                    }
                    String trim2 = "";
                    try {
                        trim2 = split1[1].trim();
                    } catch (Exception ex) {

                    }

                    String completName = loquito.getApellido1() + " " + trim1 + " " + trim2;
                    completName = completName.trim();

                    String fechaInicio = loquito.getFechaInicio();

                    if (completName.equals(nombretrim)) {
                        indicesIguales.add(k);
                        System.out.println("k " + k);
                        fechasInicio.add(fechaInicio);
                        System.out.println("nombre lista nubox " + nombretrim);
                        System.out.println("nombre lista final " + completName);
                        System.out.println("fechaInicio " + fechaInicio);
                    }
                }

                System.out.println("indicesIguales.size() " + indicesIguales.size());
                System.out.println("fechasInicio.size() " + fechasInicio.size());

                indicesIguales.stream().forEach((Integer indice) -> {
                    System.out.println(indice);

                    System.out.println("AHB!");
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div/div[1]/form/div/table/tbody[1]/tr/td[2]/table/tbody/tr/td[2]/select")));
                    Select select = new Select(wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div/div[1]/form/div/table/tbody[1]/tr/td[2]/table/tbody/tr/td[2]/select"))));
                    select.selectByIndex(indice);

                    System.out.println("F");
                    wait = new WebDriverWait(driver, Duration.ofSeconds(3));
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"Footer\"]/table/tbody/tr/td[3]/table/tbody/tr/td[2]")));
                    driver.findElement(By.xpath("//*[@id=\"Footer\"]/table/tbody/tr/td[3]/table/tbody/tr/td[2]")).click();

                    try {
                        wait = new WebDriverWait(driver, Duration.ofSeconds(3));
                        wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("pdfviewer")));
                        String attribute = driver.findElement(By.id("pdfviewer")).getAttribute("src");
                        driver.get(attribute);
                        driver.navigate().back();
                        String title = driver.getTitle();
                        System.out.println("title " + title);
                        boolean boolx = false;
                        boolean boolx1 = false;
                        boolean boolx2 = false;
                        File[] listOfFiles = null;
                        File folder = new File(System.getProperty("user.dir") + "\\PDF");
                        while (boolx == false || boolx1 == true || boolx2 == true) {
                            try {
                                folder = new File(System.getProperty("user.dir") + "\\PDF");
                                listOfFiles = folder.listFiles();
                                for (int j = 0; j < listOfFiles.length; j++) {
                                    String absolutePath = listOfFiles[j].getAbsolutePath();
//                            System.out.println("absolutePath " + absolutePath);
                                    boolx = listOfFiles[j].getAbsolutePath().contains(".pdf");
                                    boolx1 = listOfFiles[j].getAbsolutePath().contains(".crdownload");
                                    boolx2 = listOfFiles[j].getAbsolutePath().contains(".tmp");
                                }
                            } catch (Exception ex) {
                                Logger.getLogger(Feriados.class.getName()).log(Level.SEVERE, null, ex);
                            }
                        }

                        File absoluteFile = listOfFiles[0].getAbsoluteFile();

                        PDDocument document = null;
                        boolean bool = true;
                        while (bool) {
                            try {
                                System.out.println("absoluteFile " + absoluteFile);
                                document = PDDocument.load(absoluteFile);
                                bool = false;
                            } catch (Exception ex) {
//                                Logger.getLogger(Feriados.class.getName()).log(Level.SEVERE, null, ex);
                            }
                        }
                        String text = "";
                        while (text.equals("")) {
                            try {
                                PDFTextStripper stripper = new PDFTextStripper();
                                text = stripper.getText(document);
                                document.close();
                            } catch (Exception ex) {
//                                Logger.getLogger(Feriados.class.getName()).log(Level.SEVERE, null, ex);
                            }
                        }
                        //                        System.out.println(text);
                        String[] split1 = text.trim().split("\n");
                        String fecha = "";
                        String rutt = "";
                        String fechaComprobante = "";
                        for (int j = 0; j < split1.length - 13; j++) {
                            String name = split1[j];
                            if (name.contains("Fecha:")) {
                                System.out.println("name " + name.trim());
                                fecha = name.trim();

                                int indexOf = fecha.indexOf("Fecha: ");
                                String substring = fecha.substring(indexOf + 7);
                                System.out.println(substring);
                                fecha = substring;

                            }
                            if (name.contains("Rut:")) {
                                System.out.println("name " + name.trim());
                                rutt = name.trim();

                                int indexOf = rutt.indexOf("Rut: ");
                                String substring = rutt.substring(indexOf + 5);
                                rutt = substring;
                                System.out.println(substring);
                            }
                        }
                        int length1 = split1.length;
                        String name = split1[length1 - 14];
                        System.out.println("name " + name);
                        String[] ult = name.trim().split(" ");
                        int saldoHab = 0;
                        int diasProg = 0;
                        int largo = ult.length;
                        saldoHab = Integer.parseInt(ult[largo - 2]);
                        diasProg = Integer.parseInt(ult[largo - 1]);
                        fechaComprobante = ult[largo - 8];
                        Personas persona = new Personas();
                        persona.setRut(rutt);
                        persona.setFecha(fecha);
                        persona.setDiasProg(diasProg);
                        persona.setSaldoHab(saldoHab);
                        persona.setFechaComprobante(fechaComprobante);
                        persona.setFechaInicio(fechasInicio.get(0));
                        System.out.println("-------------------------------");
                        System.out.println(persona.getRut());
                        System.out.println(persona.getFecha());
                        System.out.println(persona.getDiasProg());
                        System.out.println(persona.getSaldoHab());
                        System.out.println(persona.getFechaComprobante());
                        System.out.println(persona.getFechaInicio());
                        System.out.println("-------------------------------");
                        arrPersonas.add(persona);
                        System.out.println("Saldo Habil " + saldoHab);
                        System.out.println(diasProg + "Dias Programados ");
                        System.out.println("fecha de comprobante" + fechaComprobante);
                        VentanaPrincipal.info.setText(rutt);

                        boolean boolc = true;
                        while (folder.listFiles().length != 0 || boolc) {
                            try {
                                System.out.println("folder " + folder);
                                FileUtils.cleanDirectory(folder);
                                boolc = false;
                            } catch (Exception ex) {
                                Logger.getLogger(Feriados.class.getName()).log(Level.SEVERE, null, ex);
                            }
                        }
                    } catch (Exception ex) {
                        if (!ex.toString().contains("by By.id: pdfviewer ")) {
                            Logger.getLogger(Feriados.class.getName()).log(Level.SEVERE, null, ex);
                        }
                    }
                });
            } catch (Exception ex) {

                Logger.getLogger(Feriados.class.getName()).log(Level.SEVERE, null, ex);

//                if (ex.toString().contains("By.xpath: //*[@id=\"Funcionario\"]")) {
//                    driver.quit();
//                    System.exit(0);
//                }

                q++;

                File folder = new File(System.getProperty("user.dir") + "\\PDF");
                while (folder.listFiles().length != 0) {
                    System.out.println("folder " + folder);
                    FileUtils.cleanDirectory(folder);
                }
            }

            VentanaPrincipal.progresonashe.setValue(q);
            System.out.println("-------------------------> " + q);

            //break;
        }

        driver.quit();

        return arrPersonas;
    }
    //---------------------------------------------------------------------------------------------------------------------------

    //-----------------------------------------------------------------------------------------------------------------------
    public static ArrayList<PersonasInicio> transformarExcelaArray(String path) throws FileNotFoundException, IOException {
        System.out.println("path " + path);
        ArrayList<PersonasInicio> arrPersonasInicio = new ArrayList<>();
        InputStream ExcelFileToRead = new FileInputStream(new File(path));

        XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
        XSSFSheet sheet = wb.getSheetAt(0);
        Row row;
        Cell cell;
        Iterator rows = sheet.rowIterator();
        rows.next();

        while (rows.hasNext()) {
            int cont = 0;
            PersonasInicio personaInicio = new PersonasInicio();
            row = (Row) rows.next();
            Iterator cells = row.cellIterator();
            while (cells.hasNext()) {
                System.out.println(cont);
                cell = (Cell) cells.next();
                switch (cont) {
                    case 1: {
                        if (cell.getCellType() == CellType.FORMULA) {
                            if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
                                personaInicio.setNombres(String.valueOf(cell.getNumericCellValue()));
//                                arrPersonasInicio.add(String.valueOf(cell.getNumericCellValue()));
//                                System.out.println("1 " + String.valueOf(cell.getNumericCellValue()));
                            } else if (cell.getCachedFormulaResultType() == CellType.STRING) {
                                personaInicio.setNombres(cell.getStringCellValue());
//                                arrPersonasInicio.add(cell.getStringCellValue());
//                                System.out.println("2 " + cell.getStringCellValue());
                            }
                        } else {
                            try {
                                personaInicio.setNombres(cell.getStringCellValue());
//                                arrPersonasInicio.add(cell.getStringCellValue());
//                                System.out.println("3 " + cell.getStringCellValue());
                            } catch (Exception e) {
                                personaInicio.setNombres(String.valueOf(cell.getNumericCellValue()));
//                                arrPersonasInicio.add(String.valueOf(cell.getNumericCellValue()));
//                                System.out.println("4 " + String.valueOf(cell.getNumericCellValue()));
                            }
                        }
                    }

                    case 2: {
                        if (cell.getCellType() == CellType.FORMULA) {
                            if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
                                personaInicio.setApellido1(String.valueOf(cell.getNumericCellValue()));
//                                arrPersonasInicio.add(String.valueOf(cell.getNumericCellValue()));
//                                System.out.println("1 " + String.valueOf(cell.getNumericCellValue()));
                            } else if (cell.getCachedFormulaResultType() == CellType.STRING) {
                                personaInicio.setApellido1(cell.getStringCellValue());
//                                arrPersonasInicio.add(cell.getStringCellValue());
//                                System.out.println("2 " + cell.getStringCellValue());
                            }
                        } else {
                            try {
                                personaInicio.setApellido1(cell.getStringCellValue());
//                                arrPersonasInicio.add(cell.getStringCellValue());
//                                System.out.println("3 " + cell.getStringCellValue());
                            } catch (Exception e) {
                                personaInicio.setApellido1(String.valueOf(cell.getNumericCellValue()));
//                                arrPersonasInicio.add(String.valueOf(cell.getNumericCellValue()));
//                                System.out.println("4 " + String.valueOf(cell.getNumericCellValue()));
                            }
                        }
                    }
                    case 3: {
                        if (cell.getCellType() == CellType.FORMULA) {
                            if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
                                personaInicio.setApellido2(String.valueOf(cell.getNumericCellValue()));
//                                arrPersonasInicio.add(String.valueOf(cell.getNumericCellValue()));
//                                System.out.println("1 " + String.valueOf(cell.getNumericCellValue()));
                            } else if (cell.getCachedFormulaResultType() == CellType.STRING) {
                                personaInicio.setApellido2(cell.getStringCellValue());
//                                arrPersonasInicio.add(cell.getStringCellValue());
//                                System.out.println("2 " + cell.getStringCellValue());
                            }
                        } else {
                            try {
                                personaInicio.setApellido2(cell.getStringCellValue());
//                                arrPersonasInicio.add(cell.getStringCellValue());
//                                System.out.println("3 " + cell.getStringCellValue());
                            } catch (Exception e) {
                                personaInicio.setApellido2(String.valueOf(cell.getNumericCellValue()));
//                                arrPersonasInicio.add(String.valueOf(cell.getNumericCellValue()));
//                                System.out.println("4 " + String.valueOf(cell.getNumericCellValue()));
                            }
                        }
                    }

                    case 4: {
                        if (cell.getCellType() == CellType.FORMULA) {
                            if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
                                personaInicio.setFechaInicio(String.valueOf(cell.getNumericCellValue()));
//                                arrPersonasInicio.add(String.valueOf(cell.getNumericCellValue()));
//                                System.out.println("1 " + String.valueOf(cell.getNumericCellValue()));
                            } else if (cell.getCachedFormulaResultType() == CellType.STRING) {
                                personaInicio.setFechaInicio(cell.getStringCellValue());
//                                arrPersonasInicio.add(cell.getStringCellValue());
//                                System.out.println("2 " + cell.getStringCellValue());
                            }
                        } else {
                            try {
                                personaInicio.setFechaInicio(cell.getStringCellValue());
//                                arrPersonasInicio.add(cell.getStringCellValue());
//                               System.out.println("3 " + cell.getStringCellValue());
                            } catch (Exception e) {
                                double numericCellValue = cell.getNumericCellValue();
                                Date fecha = DateUtil.getJavaDate((double) numericCellValue);
                                String format = new SimpleDateFormat("dd/MM/yyyy").format(fecha);
                                personaInicio.setFechaInicio(format);
//                                arrPersonasInicio.add(String.valueOf(cell.getNumericCellValue()));
//                                System.out.println("4 " + String.valueOf(cell.getNumericCellValue()));
                            }
                        }
                    }

                }

                cont++;
            }

            arrPersonasInicio.add(personaInicio); //agregar persona al arreglo
        }

        ExcelFileToRead.close();

        for (int i = 0; i < arrPersonasInicio.size(); i++) {
            PersonasInicio get = arrPersonasInicio.get(i);
            System.out.println(get.getNombres());
            System.out.println(get.getApellido1());
            System.out.println(get.getApellido2());
            System.out.println(get.getFechaInicio());

        }
        System.out.println(arrPersonasInicio.size() + "--------------------------------------------------------------------------------");

        return arrPersonasInicio;
    }
//---------------------------------------------------------------------------------------------------------------------------

    public static WebDriver chromeOption() {
        File fileCarpeta = new File(System.getProperty("user.dir") + "\\Facturas");
        if (!fileCarpeta.exists()) {
            new File(System.getProperty("user.dir") + "\\Facturas").mkdir();
        }

        ChromeOptions options = new ChromeOptions();
        HashMap<String, Object> chromeOptionsMap = new HashMap<String, Object>();
        chromeOptionsMap.put("plugins.plugins_disabled", new String[]{"Chrome PDF Viewer"});
        chromeOptionsMap.put("plugins.always_open_pdf_externally", true);
        chromeOptionsMap.put("download.default_directory", System.getProperty("user.dir") + "\\PDF");
        options.setExperimentalOption("prefs", chromeOptionsMap);
        //options.addArguments("--headless");

        WebDriverManager.chromedriver().setup();
        WebDriver driver = new ChromeDriver(options);
        driver.manage().window().maximize();

        return driver;
    }
//------------------------------------------------------------------------------------------------------------------------------

    public static ArrayList<Personas> manejoFecha(ArrayList<Personas> arrPersonas) {

        arrPersonas.stream().forEach(new Consumer<Personas>() {
            @Override
            public void accept(Personas persona) {
                String fechaActual = persona.getFecha();
                String fechaComprobante = persona.getFechaComprobante();
                String fechaInicio = persona.getFechaInicio();

                System.out.println("fechaActual " + fechaActual);
                System.out.println("fechaComprobante " + fechaComprobante);
                System.out.println("fechaInicio " + fechaInicio);
                System.out.println("Saldo0------------------------------> " + persona.getSaldoHab());

                int fechaActualInt = devuelveAño(fechaActual);
                int fechaComprobanteInt = devuelveAño(fechaComprobante);

                int name = fechaActualInt - fechaComprobanteInt;
                int namex = name + 1;
                int saldox = persona.getSaldoHab();

                System.out.println("namex " + namex);

                String[] split = fechaInicio.split("/");
                String dia = split[0].trim();
                int diaInt = Integer.parseInt(dia);
                System.out.println("dia inicio" + dia);
                String mes = split[1].trim();
                int mesInt = Integer.parseInt(mes);
                System.out.println("mes inicio " + mes);

                String[] splitComprobante = fechaComprobante.split("/");
                String diaComprobante = splitComprobante[0].trim();
                int diaIntComprobante = Integer.parseInt(diaComprobante);
                System.out.println("diacomprobante" + diaIntComprobante);
                String mesComprobante = splitComprobante[1].trim();
                int mesIntComprobante = Integer.parseInt(mesComprobante);
                System.out.println("mes comprobante" + mesIntComprobante);

                String[] splitActual = fechaActual.split("/");
                String diaActual = splitActual[0].trim();
                int diaIntActual = Integer.parseInt(diaActual);
                System.out.println("dia Actual" + diaIntActual);
                String mesActual = splitActual[1].trim();
                int mesIntActual = Integer.parseInt(mesActual);
                System.out.println("mes Actual" + mesIntActual);
                int auxiliarcito = 0;

                for (int i = 0; i < namex; i++) {
                    System.out.println("AH");
                    if (i == 0) {
                        if (mesIntComprobante < mesInt) {
                            saldox = saldox + 15;
                            auxiliarcito = auxiliarcito + 1;
                            System.out.println("entra al primer año y suma 15" + auxiliarcito);
                        }
                        if ((mesIntComprobante == mesInt) && (diaIntComprobante <= diaInt)) {
                            saldox = saldox + 15;
                            auxiliarcito = auxiliarcito + 1;
                            System.out.println("entra al primer año y suma 15" + auxiliarcito);
                        }
                    } else if (i == (namex - 1)) {
                        if ((mesInt < mesIntActual)) {
                            saldox = saldox + 15;
                            auxiliarcito = auxiliarcito + 1;
                            System.out.println("entra al ultimo año y suma 15" + auxiliarcito);
                        }
                        if ((mesInt == mesIntActual) && (diaInt <= diaIntActual)) {
                            saldox = saldox + 15;
                            auxiliarcito = auxiliarcito + 1;
                            System.out.println("entra al ultimo año y suma 15" + auxiliarcito);
                        }
                    } else {
                        saldox = saldox + 15;
                        auxiliarcito = auxiliarcito + 1;
                        System.out.println("entra a lso de almedio y va sumando 15" + auxiliarcito);
                    }

                    auxiliarcito = 0;

                }

                persona.setSaldoHab(saldox);
                System.out.println(" ---------------------------------------------->saldo " + persona.getSaldoHab());
            }
        });

        return arrPersonas;
    }
}

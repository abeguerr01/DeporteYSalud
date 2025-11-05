package com.mycompany.di_p3_maven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Font; 
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.ss.usermodel.charts.ChartAxis;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.ChartLegend;
import org.apache.poi.ss.usermodel.charts.DataSources;
import org.apache.poi.ss.usermodel.charts.LineChartData;
import org.apache.poi.ss.usermodel.charts.LegendPosition;
import org.apache.poi.ss.usermodel.charts.ValueAxis;
import org.apache.poi.ss.usermodel.charts.AxisPosition;
import org.apache.poi.ss.usermodel.charts.AxisCrosses;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.ChartUtils;
import org.jfree.data.category.DefaultCategoryDataset;
import org.apache.poi.ss.usermodel.CellType;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.io.image.ImageData;




public class Control extends javax.swing.JFrame {

    private static final java.util.logging.Logger logger = java.util.logging.Logger.getLogger(Control.class.getName());

    public Control() {
        Components();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">                          
    private void Components() {
        java.awt.GridBagConstraints gridBagConstraints;

        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jTFpeso = new javax.swing.JTextField();
        jLabel3 = new javax.swing.JLabel();
        jTFaltura = new javax.swing.JTextField();
        jLimc = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jTFcalcon = new javax.swing.JTextField();
        jLabel6 = new javax.swing.JLabel();
        jLbalcal = new javax.swing.JLabel();
        jBguardar = new javax.swing.JButton();
        jTFcalquem = new javax.swing.JTextField();
        jBimc = new javax.swing.JButton();
        jBcalcal = new javax.swing.JButton();
        jLdepu = new javax.swing.JLabel();

        setMinimumSize(new java.awt.Dimension(900, 400));
        getContentPane().setLayout(new java.awt.GridBagLayout());

        jLabel1.setFont(new java.awt.Font("Segoe UI", 0, 48)); // NOI18N
        jLabel1.setText("Control");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 0;
        gridBagConstraints.gridy = 0;
        gridBagConstraints.gridwidth = 4;
        gridBagConstraints.ipady = 50;
        getContentPane().add(jLabel1, gridBagConstraints);

        jLabel2.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jLabel2.setText("Peso en kg: ");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 0;
        gridBagConstraints.gridy = 1;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        getContentPane().add(jLabel2, gridBagConstraints);

        jTFpeso.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jTFpeso.setToolTipText("");
        jTFpeso.setName(""); // NOI18N
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 1;
        gridBagConstraints.gridy = 1;
        gridBagConstraints.ipadx = 80;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        gridBagConstraints.insets = new java.awt.Insets(0, 0, 0, 50);
        getContentPane().add(jTFpeso, gridBagConstraints);

        jLabel3.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jLabel3.setText("Altura en m: ");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 0;
        gridBagConstraints.gridy = 2;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        getContentPane().add(jLabel3, gridBagConstraints);

        jTFaltura.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 1;
        gridBagConstraints.gridy = 2;
        gridBagConstraints.ipadx = 80;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        gridBagConstraints.insets = new java.awt.Insets(0, 0, 0, 50);
        getContentPane().add(jTFaltura, gridBagConstraints);

        jLimc.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jLimc.setText("IMC");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 1;
        gridBagConstraints.gridy = 3;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        gridBagConstraints.insets = new java.awt.Insets(0, 0, 0, 50);
        getContentPane().add(jLimc, gridBagConstraints);

        jLabel5.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jLabel5.setText("Calorias consumidas: ");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 2;
        gridBagConstraints.gridy = 1;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        getContentPane().add(jLabel5, gridBagConstraints);

        jTFcalcon.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 3;
        gridBagConstraints.gridy = 1;
        gridBagConstraints.ipadx = 80;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        getContentPane().add(jTFcalcon, gridBagConstraints);

        jLabel6.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jLabel6.setText("Calorias quemadas:");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 2;
        gridBagConstraints.gridy = 2;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        getContentPane().add(jLabel6, gridBagConstraints);

        jLbalcal.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jLbalcal.setText("Balance de calorias");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 3;
        gridBagConstraints.gridy = 3;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        getContentPane().add(jLbalcal, gridBagConstraints);

        jBguardar.setText("Guardar informacion");
        jBguardar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jBguardarActionPerformed(evt);
            }
        });
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 0;
        gridBagConstraints.gridy = 4;
        gridBagConstraints.gridwidth = 4;
        getContentPane().add(jBguardar, gridBagConstraints);

        jTFcalquem.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jTFcalquem.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTFcalquemActionPerformed(evt);
            }
        });
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 3;
        gridBagConstraints.gridy = 2;
        gridBagConstraints.ipadx = 80;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        getContentPane().add(jTFcalquem, gridBagConstraints);

        jBimc.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jBimc.setText("Mostrar IMC:");
        jBimc.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jBimcActionPerformed(evt);
            }
        });
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 0;
        gridBagConstraints.gridy = 3;
        getContentPane().add(jBimc, gridBagConstraints);

        jBcalcal.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jBcalcal.setText("Mostrar calorias");
        jBcalcal.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jBcalcalActionPerformed(evt);
            }
        });
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 2;
        gridBagConstraints.gridy = 3;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        getContentPane().add(jBcalcal, gridBagConstraints);
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 0;
        gridBagConstraints.gridy = 5;
        gridBagConstraints.gridwidth = 4;
        getContentPane().add(jLdepu, gridBagConstraints);

        pack();
    }// </editor-fold>
    
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {
        java.awt.GridBagConstraints gridBagConstraints;

        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jTFpeso = new javax.swing.JTextField();
        jLabel3 = new javax.swing.JLabel();
        jTFaltura = new javax.swing.JTextField();
        jLimc = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jTFcalcon = new javax.swing.JTextField();
        jLabel6 = new javax.swing.JLabel();
        jLbalcal = new javax.swing.JLabel();
        jBguardar = new javax.swing.JButton();
        jTFcalquem = new javax.swing.JTextField();
        jBimc = new javax.swing.JButton();
        jBcalcal = new javax.swing.JButton();
        jLdepu = new javax.swing.JLabel();

        setMinimumSize(new java.awt.Dimension(900, 400));
        getContentPane().setLayout(new java.awt.GridBagLayout());

        jLabel1.setFont(new java.awt.Font("Segoe UI", 0, 48)); // NOI18N
        jLabel1.setText("Control");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 0;
        gridBagConstraints.gridy = 0;
        gridBagConstraints.gridwidth = 4;
        gridBagConstraints.ipady = 50;
        getContentPane().add(jLabel1, gridBagConstraints);

        jLabel2.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jLabel2.setText("Peso en kg: ");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 0;
        gridBagConstraints.gridy = 1;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        getContentPane().add(jLabel2, gridBagConstraints);

        jTFpeso.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jTFpeso.setToolTipText("");
        jTFpeso.setName(""); // NOI18N
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 1;
        gridBagConstraints.gridy = 1;
        gridBagConstraints.ipadx = 80;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        gridBagConstraints.insets = new java.awt.Insets(0, 0, 0, 50);
        getContentPane().add(jTFpeso, gridBagConstraints);

        jLabel3.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jLabel3.setText("Altura en m: ");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 0;
        gridBagConstraints.gridy = 2;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        getContentPane().add(jLabel3, gridBagConstraints);

        jTFaltura.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 1;
        gridBagConstraints.gridy = 2;
        gridBagConstraints.ipadx = 80;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        gridBagConstraints.insets = new java.awt.Insets(0, 0, 0, 50);
        getContentPane().add(jTFaltura, gridBagConstraints);

        jLimc.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jLimc.setText("IMC");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 1;
        gridBagConstraints.gridy = 3;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        gridBagConstraints.insets = new java.awt.Insets(0, 0, 0, 50);
        getContentPane().add(jLimc, gridBagConstraints);

        jLabel5.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jLabel5.setText("Calorias consumidas: ");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 2;
        gridBagConstraints.gridy = 1;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        getContentPane().add(jLabel5, gridBagConstraints);

        jTFcalcon.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 3;
        gridBagConstraints.gridy = 1;
        gridBagConstraints.ipadx = 80;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        getContentPane().add(jTFcalcon, gridBagConstraints);

        jLabel6.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jLabel6.setText("Calorias quemadas:");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 2;
        gridBagConstraints.gridy = 2;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        getContentPane().add(jLabel6, gridBagConstraints);

        jLbalcal.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jLbalcal.setText("Balance de calorias");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 3;
        gridBagConstraints.gridy = 3;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        getContentPane().add(jLbalcal, gridBagConstraints);

        jBguardar.setText("Guardar informacion");
        jBguardar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jBguardarActionPerformed(evt);
            }
        });
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 0;
        gridBagConstraints.gridy = 4;
        gridBagConstraints.gridwidth = 4;
        getContentPane().add(jBguardar, gridBagConstraints);

        jTFcalquem.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jTFcalquem.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTFcalquemActionPerformed(evt);
            }
        });
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 3;
        gridBagConstraints.gridy = 2;
        gridBagConstraints.ipadx = 80;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        getContentPane().add(jTFcalquem, gridBagConstraints);

        jBimc.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jBimc.setText("Mostrar IMC:");
        jBimc.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jBimcActionPerformed(evt);
            }
        });
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 0;
        gridBagConstraints.gridy = 3;
        getContentPane().add(jBimc, gridBagConstraints);

        jBcalcal.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jBcalcal.setText("Mostrar calorias");
        jBcalcal.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jBcalcalActionPerformed(evt);
            }
        });
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 2;
        gridBagConstraints.gridy = 3;
        gridBagConstraints.anchor = java.awt.GridBagConstraints.LINE_START;
        getContentPane().add(jBcalcal, gridBagConstraints);
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 0;
        gridBagConstraints.gridy = 5;
        gridBagConstraints.gridwidth = 4;
        getContentPane().add(jLdepu, gridBagConstraints);

        pack();
    }// </editor-fold>//GEN-END:initComponents
    
    private void generarGraficoBalanceCalorico() {
        final String NOMBRE_ARCHIVO_EXCEL = "control_datos.xlsx";
        final String NOMBRE_ARCHIVO_GRAFICO = "balance_calorico.jpg";
        final int FILA_ENCABEZADO = 4; // Fila donde empieza el encabezado (0-based)

        File archivoExcel = new File(NOMBRE_ARCHIVO_EXCEL);
        if (!archivoExcel.exists()) {
            jLdepu.setText("Error: El archivo Excel no existe para generar el gráfico.");
            return;
        }

        Workbook libroExcel = null;
        try (FileInputStream fis = new FileInputStream(archivoExcel)) {
            libroExcel = new XSSFWorkbook(fis);
            Sheet hoja = libroExcel.getSheet("Datos");
            if (hoja == null || hoja.getLastRowNum() < FILA_ENCABEZADO + 1) {
                jLdepu.setText("Error: No hay suficientes datos para generar el gráfico.");
                return;
            }

            DefaultCategoryDataset dataset = new DefaultCategoryDataset();
            
            final int COL_FECHA = 0;
            final int COL_CONSUMIDAS = 5;
            final int COL_QUEMADAS = 6;
            final int COL_BALANCE = 7;
            
            for (int r = FILA_ENCABEZADO + 1; r <= hoja.getLastRowNum(); r++) {
                Row fila = hoja.getRow(r);
                if (fila != null) {
                    Cell celdaFecha = fila.getCell(COL_FECHA);
                    Cell celdaConsumidas = fila.getCell(COL_CONSUMIDAS);
                    Cell celdaQuemadas = fila.getCell(COL_QUEMADAS);
                    Cell celdaBalance = fila.getCell(COL_BALANCE);

                    if (celdaFecha != null && celdaConsumidas != null && celdaQuemadas != null && celdaBalance != null
                        && celdaConsumidas.getCellType() == CellType.NUMERIC
                        && celdaQuemadas.getCellType() == CellType.NUMERIC
                        && celdaBalance.getCellType() == CellType.NUMERIC) {
                        
                        String fecha = celdaFecha.getStringCellValue();
                        double consumidas = celdaConsumidas.getNumericCellValue();
                        double quemadas = celdaQuemadas.getNumericCellValue();
                        double balance = celdaBalance.getNumericCellValue();
                        
                        dataset.addValue(consumidas, "Consumidas", fecha);
                        dataset.addValue(quemadas, "Quemadas", fecha);
                        dataset.addValue(balance, "Balance", fecha);
                    }
                }
            }
            
            JFreeChart barChart = ChartFactory.createBarChart(
                "Control de Calorías por Día", // Título
                "Fecha",                      // Etiqueta del Eje X
                "Cantidad de Calorías",       // Etiqueta del Eje Y
                dataset,                      // Datos
                PlotOrientation.VERTICAL,     // Orientación
                true, true, false             // Mostrar Leyenda (true), Tooltips, URLs
            );

            int ancho = 1000; // Un poco más ancho para las tres barras
            int alto = 600;
            File archivoGrafico = new File(NOMBRE_ARCHIVO_GRAFICO);

            ChartUtils.saveChartAsJPEG(archivoGrafico, barChart, ancho, alto);
            
            jLdepu.setText("Datos guardados y gráfico generado con éxito como '" + NOMBRE_ARCHIVO_GRAFICO + "'.");

        } catch (IOException e) {
            jLdepu.setText("Error de I/O al leer el Excel o escribir el gráfico: " + e.getMessage());
            e.printStackTrace();
        } catch (Exception e) {
            jLdepu.setText("Error al generar el gráfico: " + e.getMessage());
            e.printStackTrace();
        } finally {
            if (libroExcel != null) {
                try {
                    libroExcel.close();
                } catch (IOException e) {
                    System.err.println("Error al cerrar el Workbook: " + e.getMessage());
                }
            }
        }
    }
    
    private void generarInformePDF() {
        try {
        String rutaExcel = "control_datos.xlsx"; // no se usa aquí pero queda por coherencia
        String rutaImagen = "balance_calorico.jpg";
        String rutaPDF = "informe_control.pdf";

        PdfWriter writer = new PdfWriter(rutaPDF);
        PdfDocument pdfDoc = new PdfDocument(writer);
        Document document = new Document(pdfDoc);

        document.add(new Paragraph("Informe de Control de Salud y Deporte")
                .setBold()
                .setFontSize(18));

        String fechaGeneracion = new java.text.SimpleDateFormat("dd/MM/yyyy HH:mm:ss")
                .format(new java.util.Date());
        document.add(new Paragraph("Fecha de generación: " + fechaGeneracion));
        document.add(new Paragraph(" ")); // Espacio

        document.add(new Paragraph("Último registro guardado:").setBold());

        String pesoStr = (jTFpeso != null) ? jTFpeso.getText() : "";
        String alturaStr = (jTFaltura != null) ? jTFaltura.getText() : "";
        String imcStr = (jLimc != null) ? jLimc.getText() : "";
        String calConsStr = (jTFcalcon != null) ? jTFcalcon.getText() : "";
        String calQuemStr = (jTFcalquem != null) ? jTFcalquem.getText() : "";
        String balanceStr = (jLbalcal != null) ? jLbalcal.getText() : "";

        document.add(new Paragraph("Peso: " + pesoStr + " kg"));
        document.add(new Paragraph("Altura: " + alturaStr + " m"));
        document.add(new Paragraph("IMC: " + imcStr));
        document.add(new Paragraph("Calorías Consumidas: " + calConsStr));
        document.add(new Paragraph("Calorías Quemadas: " + calQuemStr));
        document.add(new Paragraph("Balance Calórico: " + balanceStr));
        document.add(new Paragraph(" ")); // Espacio

        java.io.File imgFile = new java.io.File(rutaImagen);
        if (imgFile.exists()) {
            ImageData imageData = ImageDataFactory.create(rutaImagen);
            com.itextpdf.layout.element.Image pdfImg = new com.itextpdf.layout.element.Image(imageData);
            pdfImg.setAutoScale(true);

            document.add(new Paragraph("Gráfico de Balance Calórico:").setBold());
            document.add(pdfImg);
        } else {
            document.add(new Paragraph("⚠ No se encontró el gráfico: " + rutaImagen));
        }

        document.close();


        } catch (Exception ex) {
            ex.printStackTrace();
            javax.swing.JOptionPane.showMessageDialog(this,
                    "Error al generar el PDF: " + ex.getMessage());
        }
    }
    
    private void jBguardarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jBguardarActionPerformed
        if (jTFpeso.getText().isEmpty() || jTFaltura.getText().isEmpty() || jTFcalcon.getText().isEmpty() || jTFcalquem.getText().isEmpty()) {
            jLdepu.setText("Error: Rellena todos los campos.");
            return;
        }

        final String NOMBRE_ARCHIVO = "control_datos.xlsx";
        final String RUTA_IMAGEN = "src/main/resources/img/fondoMenu.jpg";
        
        final int FILA_ENCABEZADO = 4; 
        
        File archivo = new File(NOMBRE_ARCHIVO);

        Workbook libroExcel = null;
        Sheet hoja = null;
        double pesoD, alturaD, imc, calbal;
        int calconI, calquemI;

        int numeroFilaParaDatos; // Fila (0-based) donde se insertará la *nueva* línea de datos

        try {
            pesoD = Double.parseDouble(jTFpeso.getText());
            alturaD = Double.parseDouble(jTFaltura.getText());
            imc = pesoD / (alturaD * alturaD);
            calconI = Integer.parseInt(jTFcalcon.getText());
            calquemI = Integer.parseInt(jTFcalquem.getText());
            calbal = calconI - calquemI;

            if (archivo.exists()) {
                try (FileInputStream fis = new FileInputStream(archivo)) {
                    libroExcel = new XSSFWorkbook(fis);
                }
                hoja = libroExcel.getSheet("Datos");
                if (hoja == null) {
                    hoja = libroExcel.createSheet("Datos");
                }

                numeroFilaParaDatos = hoja.getLastRowNum() + 1; 
                
                if (numeroFilaParaDatos <= FILA_ENCABEZADO) {
                    numeroFilaParaDatos = FILA_ENCABEZADO + 1;
                }

            } else {
                libroExcel = new XSSFWorkbook();
                hoja = libroExcel.createSheet("Datos");

                try {
                    InputStream is = Files.newInputStream(Paths.get(RUTA_IMAGEN));
                    byte[] bytesImagen = IOUtils.toByteArray(is);
                    int idImagen = libroExcel.addPicture(bytesImagen, Workbook.PICTURE_TYPE_PNG);
                    is.close();

                    CreationHelper helper = libroExcel.getCreationHelper();
                    Drawing<?> drawing = hoja.createDrawingPatriarch(); 
                    ClientAnchor anchor = helper.createClientAnchor();
                    anchor.setCol1(0); 
                    anchor.setRow1(0);
                    anchor.setCol2(4);
                    anchor.setRow2(FILA_ENCABEZADO);
                    drawing.createPicture(anchor, idImagen);

                } catch (IOException e) {
                    System.err.println("Advertencia: No se pudo cargar o insertar la imagen '" + RUTA_IMAGEN + "'.");
                }
                
                CellStyle estiloEncabezado = libroExcel.createCellStyle();
                estiloEncabezado.setFillForegroundColor(IndexedColors.GREEN.getIndex());
                estiloEncabezado.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                Font font = libroExcel.createFont();
                font.setBold(true);
                estiloEncabezado.setFont(font);
                estiloEncabezado.setAlignment(HorizontalAlignment.CENTER);

                Row encabezado = hoja.createRow(FILA_ENCABEZADO); 

                String[] nombresEncabezados = {
                    "Fecha", "Hora", "Peso (kg)", "Altura (m)", "IMC", 
                    "Calorias Consumidas", "Calorias Quemadas", "Balance Calorico"
                };

                for (int i = 0; i < nombresEncabezados.length; i++) {
                    Cell cell = encabezado.createCell(i);
                    cell.setCellValue(nombresEncabezados[i]);
                    cell.setCellStyle(estiloEncabezado);
                }

                for (int i = 0; i < nombresEncabezados.length; i++) {
                     hoja.autoSizeColumn(i);
                }

                numeroFilaParaDatos = FILA_ENCABEZADO + 1; // La primera fila de datos es la 5 (índice 5)
            }

            Row nuevaFila = hoja.createRow(numeroFilaParaDatos); 

            SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
            SimpleDateFormat timeFormat = new SimpleDateFormat("HH:mm:ss");
            Date ahora = new Date();

            nuevaFila.createCell(0).setCellValue(dateFormat.format(ahora)); // Col A
            nuevaFila.createCell(1).setCellValue(timeFormat.format(ahora)); // Col B
            nuevaFila.createCell(2).setCellValue(pesoD);                    // Col C
            nuevaFila.createCell(3).setCellValue(alturaD);                  // Col D
            nuevaFila.createCell(4).setCellValue(imc);                      // Col E
            nuevaFila.createCell(5).setCellValue(calconI);                   // Col F
            nuevaFila.createCell(6).setCellValue(calquemI);                  // Col G
            nuevaFila.createCell(7).setCellValue(calbal);                   // Col H

            
            
            int primeraFilaDatos = FILA_ENCABEZADO + 1; 
            int ultimaFilaDatos = numeroFilaParaDatos;  

            if (ultimaFilaDatos >= primeraFilaDatos) {
                
                Drawing<?> drawingPatriarch = hoja.getDrawingPatriarch();
                if (drawingPatriarch == null) {
                    drawingPatriarch = hoja.createDrawingPatriarch();
                }

                XSSFDrawing drawing = (XSSFDrawing) drawingPatriarch;

                ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 8, 4, 20, 24);

                XSSFChart chart = drawing.createChart(anchor);
                chart.setTitleText("Evolución de Peso, Altura e IMC");
                
                ChartLegend legend = chart.getOrCreateLegend();
                legend.setPosition(LegendPosition.TOP_RIGHT);

                ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
                ChartDataSource<String> categories = DataSources.fromStringCellRange(hoja, 
                    new CellRangeAddress(primeraFilaDatos, ultimaFilaDatos, 0, 0)); // Col A

                ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
                leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
                
                ValueAxis rightAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.RIGHT);
                rightAxis.setCrosses(AxisCrosses.MAX); 

                LineChartData dataPrincipal = chart.getChartDataFactory().createLineChartData();

                ChartDataSource<Number> seriesPeso = DataSources.fromNumericCellRange(hoja, 
                    new CellRangeAddress(primeraFilaDatos, ultimaFilaDatos, 2, 2));
                org.apache.poi.ss.usermodel.charts.LineChartSeries series1 = dataPrincipal.addSeries(categories, seriesPeso);
                series1.setTitle("Peso (kg)"); 

                ChartDataSource<Number> seriesIMC = DataSources.fromNumericCellRange(hoja, 
                    new CellRangeAddress(primeraFilaDatos, ultimaFilaDatos, 4, 4));
                org.apache.poi.ss.usermodel.charts.LineChartSeries series3 = dataPrincipal.addSeries(categories, seriesIMC);
                series3.setTitle("IMC");
                
                LineChartData dataSecundario = chart.getChartDataFactory().createLineChartData();
                
                ChartDataSource<Number> seriesAltura = DataSources.fromNumericCellRange(hoja, 
                    new CellRangeAddress(primeraFilaDatos, ultimaFilaDatos, 3, 3));
                org.apache.poi.ss.usermodel.charts.LineChartSeries series2 = dataSecundario.addSeries(categories, seriesAltura);
                series2.setTitle("Altura (m)");

                chart.plot(dataPrincipal, bottomAxis, leftAxis);
                chart.plot(dataSecundario, bottomAxis, rightAxis);
            }


            try (FileOutputStream fos = new FileOutputStream(archivo)) {
                libroExcel.write(fos);
            }

            jLdepu.setText("Datos guardados en '" + NOMBRE_ARCHIVO + "' con éxito.");
            
            generarGraficoBalanceCalorico();
            generarInformePDF();

        } catch (NumberFormatException e) {
            jLdepu.setText("Error de formato: Asegúrate de que los campos tienen formato numérico válido.");
        } catch (IOException e) {
            jLdepu.setText("Error al manejar el archivo Excel: " + e.getMessage());
            System.out.println("Error: " + e.getMessage());
            e.printStackTrace();
        } finally {
            if (libroExcel != null) {
                try {
                    libroExcel.close();
                } catch (IOException e) {
                    System.err.println("Error al cerrar el Workbook: " + e.getMessage());
                }
            }
        }
    }//GEN-LAST:event_jBguardarActionPerformed

    private void jBimcActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jBimcActionPerformed
        String pesoS = jTFpeso.getText();
        String alturaS = jTFaltura.getText();
        
        if (pesoS.isEmpty() || alturaS.isEmpty()) {
            jLimc.setText("Error: Rellena peso y altura");
            return;
        }
        
        try {
            double pesoD = Double.parseDouble(pesoS);
            double alturaD = Double.parseDouble(alturaS);

            if (alturaD == 0) {
                jLimc.setText("Error: Altura no puede ser 0");
                return;
            }

            double imc = pesoD/(alturaD*alturaD);
            String imcS = String.format("%.2f", imc); //Limitar a 2 decimales

            if(imc < 18.5){
                jLimc.setText(imcS+" - Infrapeso");
            }else if(imc >= 18.5 && imc <= 25){
                jLimc.setText(imcS+" - Peso normal");
            }else{
                jLimc.setText(imcS+" - Sobrepeso");
            }
        } catch (NumberFormatException e) {
            jLimc.setText("Error: Formato numérico incorrecto");
        }

    }//GEN-LAST:event_jBimcActionPerformed

    private void jBcalcalActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jBcalcalActionPerformed
        String calconS = jTFcalcon.getText();
        String calquemS = jTFcalquem.getText();
        
        if (calconS.isEmpty() || calquemS.isEmpty()) {
            jLbalcal.setText("Error: Rellena las calorías");
            return;
        }
        
        try {
            int calconI = Integer.parseInt(calconS);
            int calquemI = Integer.parseInt(calquemS);

            double calbal = calconI - calquemI;

            jLbalcal.setText("Balance: "+calbal);
        } catch (NumberFormatException e) {
            jLbalcal.setText("Error: Formato numérico incorrecto");
        }
    }//GEN-LAST:event_jBcalcalActionPerformed

    private void jTFcalquemActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTFcalquemActionPerformed
    }//GEN-LAST:event_jTFcalquemActionPerformed

    public static void main(String args[]) {
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ReflectiveOperationException | javax.swing.UnsupportedLookAndFeelException ex) {
            logger.log(java.util.logging.Level.SEVERE, null, ex);
        }

        java.awt.EventQueue.invokeLater(() -> new Control().setVisible(true));
    }

    private javax.swing.JButton jBcalcal;
    private javax.swing.JButton jBguardar;
    private javax.swing.JButton jBimc;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLbalcal;
    private javax.swing.JLabel jLdepu;
    private javax.swing.JLabel jLimc;
    private javax.swing.JTextField jTFaltura;
    private javax.swing.JTextField jTFcalcon;
    private javax.swing.JTextField jTFcalquem;
    private javax.swing.JTextField jTFpeso;
}

//@author Daniel.Meza
import java.awt.Component;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Date;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.SwingWorker;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableColumn;
import javax.swing.table.TableColumnModel;

import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Font;
import com.lowagie.text.HeaderFooter;
import com.lowagie.text.PageSize;
import com.lowagie.text.Paragraph;
import com.lowagie.text.Phrase;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;
import java.io.FileNotFoundException;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.List;
import javax.swing.JTextField;
import org.joda.time.Chronology;
import org.joda.time.DateTimeField;
import org.joda.time.LocalDate;
import org.joda.time.chrono.ISOChronology;
 
public class GenerarReportes extends JPanel{
    
       private static final long serialVersionUID = 1L;
   	   
       private JLabel etiquetaInstrumento,etiquetaSubAplicacion,etiquetaAño,etiquetaMes,etiquetaNoAplicacion,etiquetaFi,etiquetaDiaFi,etiquetaMesFi,etiquetaAñoFi,
                      etiquetaFf,etiquetaDiaFf,etiquetaMesFf,etiquetaAñoFf,etiquetaFiltros;                           
       private JTextField campoNoAplicacion;
       private JButton botonGenerarReporte,botonImprimirReporte;        
       private JScrollPane panelTabla;
       private static JPanel panelFiltros,panelFiltroAplicacion,panelFiltroFechas,panelFiltroInstrumento;
       private JTable tabla;                          
       private GridBagConstraints gbc;
       private JComboBox<String> comboTipoInstr,comboNombres_cortos,comboAños,comboMes,comboDiaFi,comboMesFi,comboAñoFi,comboDiaFf,comboMesFf,comboAñoFf,
                                 comboFiltros;
       private ArrayList<Object[]> datosReporte;     
              
       private final String[] tiposInstrumento = {"AC286","ACRESEC","ACRETSU","ACUERDO","ALI","CEAACES","CONALEP","DGESPE","ECCYPEC","ECELE","ECODEMS","EGAL",
                                                  "EGEL","EGETSU","EPROM","ESPECIALES","EUC","EUCCA","EXANI","EXTRA","IFE","LEPRE_LEPRI","MCEF","Metropolitano",                                                  
                                                  "MINNESOTA","OLIMPIADA","PILOTO","PREESCOLAR_BACH","PREESCOLAR_LIC","SEISP","SSP","TRIF","UPN"
                                                 };
                  
       private final String[] nombresColumnas = {"Num app","Tipo inst","Nombre","Fecha App","Fecha de Proc","Imag Reg","Imag Res","Reg","Reg bpm","Reg mc",
                                                 "Aplicados","Aplicados bpm","Aplicados mc","Estado","Institucion","Observacion"};
              
       private final String[] años    = {"2012","2013"}; 
       private final String[] meses   = {"Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"};
       private final String[] filtros = {"Numero de aplicacion","Tipo de instrumento","Rango de Fechas"};
       
       private String name,añoFiSeleccionado,añoFfSeleccionado;
       final int nombresCantidad = nombresColumnas.length - 1;
              
       @SuppressWarnings("LeakingThisInConstructor")
       public GenerarReportes(String nombre){
             
    	      name = nombre;
    	      
    	      setLayout(new GridBagLayout());
    	      
              gbc = new GridBagConstraints();
              
              panelFiltros = new JPanel();
              
              etiquetaFiltros = new JLabel("Filtro por : ");              
              comboFiltros = new JComboBox<>(filtros);
              comboFiltros.addActionListener(new ActionListener() {

                           @Override
                           public void actionPerformed(ActionEvent e) {
                                
                                  SwingWorker<Void,Void> sw;
                                  sw = new SwingWorker<Void, Void>() {

                                      @Override
                                      protected Void doInBackground() throws Exception {
                                          
                                                try{
                                                    
                                                    String comando = GenerarReportes.this.comboFiltros.getActionCommand();
                                                    
                                                    if( comando.equals("comboBoxChanged") ){ 
                                                        
                                                        String filtroSeleccionado = (String)GenerarReportes.this.comboFiltros.getSelectedItem();
                                                        
                                                        if( filtroSeleccionado.equals(filtros[0]) ){
                                                            GenerarReportes.this.remove(panelFiltroFechas);
                                                            GenerarReportes.this.remove(panelFiltroInstrumento);
                                                            GenerarReportes.this.add(panelFiltroAplicacion);
                                                        }
                                                        
                                                        if( filtroSeleccionado.equals(filtros[1]) ){
                                                            GenerarReportes.this.remove(panelFiltroAplicacion);
                                                            GenerarReportes.this.remove(panelFiltroFechas);
                                                            GenerarReportes.this.add(panelFiltroInstrumento);
                                                        }
                                                        
                                                        if( filtroSeleccionado.equals(filtros[2]) ){
                                                            GenerarReportes.this.remove(panelFiltroAplicacion);
                                                            GenerarReportes.this.remove(panelFiltroInstrumento);
                                                            GenerarReportes.this.add(panelFiltroFechas);
                                                        }
                                                        
                                                        GenerarReportes.this.revalidate();
                                                        GenerarReportes.this.repaint();
                                                        
                                                    }                                                                                                               
                                                                                                  
                                                }catch(Exception e){ e.printStackTrace();}
                                                
                                          
                                                return null;
                                           
                                      }
                                      
                                      
                                  };
                                  
                                  sw.execute();
                                  
                           }
              });
                               
              panelFiltros.add(etiquetaFiltros);
              panelFiltros.add(comboFiltros);
              
              gbc = new GridBagConstraints();
              gbc.gridx = 0;
              gbc.gridy = 0;              
              gbc.weightx = 0.1;              
              gbc.anchor  = GridBagConstraints.EAST;
              gbc.insets = new Insets(5,5,5,5);
              add(panelFiltros,gbc);
                      
              panelFiltroAplicacion = new JPanel(new GridBagLayout());              
              panelFiltroAplicacion.setSize(500, 200);
              
              etiquetaNoAplicacion = new JLabel("Numero de aplicacion : ");
              
              gbc.gridx = 0;
              gbc.gridy = 0;              
              gbc.weightx = 0.1;
              gbc.weighty = 0.1;
              gbc.anchor  = GridBagConstraints.WEST;
              gbc.insets = new Insets(5,5,5,5);
              panelFiltroAplicacion.add(etiquetaNoAplicacion,gbc);                                     
              
              campoNoAplicacion = new JTextField(9);
              gbc = new GridBagConstraints();
              gbc.gridx = 1;
              gbc.gridy = 0;              
              gbc.weightx = 0.1;
              gbc.weighty = 0.1;
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.WEST;
              panelFiltroAplicacion.add(campoNoAplicacion,gbc);                                     
              
              panelFiltroInstrumento = new JPanel(new GridBagLayout());
              
              etiquetaAño = new JLabel("Año : ");              
              gbc = new GridBagConstraints();
              gbc.gridx = 0;
              gbc.gridy = 1;  
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.WEST;
              panelFiltroInstrumento.add(etiquetaAño,gbc);
              
              comboAños = new JComboBox<>(años);
              gbc = new GridBagConstraints();
              gbc.gridx = 0;
              gbc.gridy = 1;                            
              gbc.insets  = new Insets(5,5,5,5);
              //gbc.anchor  = GridBagConstraints.EAST;
              panelFiltroInstrumento.add(comboAños,gbc);
              
              etiquetaInstrumento = new JLabel("Instrumento : ");
              gbc = new GridBagConstraints();              
              gbc.gridy = 1;
              gbc.gridx = 1;
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.WEST;
              panelFiltroInstrumento.add(etiquetaInstrumento,gbc);                                     
              
              comboTipoInstr = new JComboBox<>(tiposInstrumento);
              comboTipoInstr.addActionListener(
                      new ActionListener() {
                          @Override
                          public void actionPerformed(ActionEvent e) {                                                         
                          
                       	         SwingWorker swbe; 
                                 swbe = new SwingWorker() {
                                     
                                 @Override
                                 protected Object doInBackground() throws Exception {
                                                                                             
                                           panelFiltroAplicacion.remove(comboNombres_cortos);
                        
                                           String itExamen = (String)comboTipoInstr.getSelectedItem();                                                
                        
                                           gbc = new GridBagConstraints();
                        
                                           comboNombres_cortos = new JComboBox<>(traeNombresCortos(itExamen));                                      
                                           gbc = new GridBagConstraints();
                                           gbc.gridx = 4;
                                           gbc.gridy = 1;                                                                                       
                                           gbc.gridwidth = 3;
                                           gbc.insets = new Insets(5,5,5,5);
                                           gbc.anchor = GridBagConstraints.WEST;                                                          
                                           panelFiltroAplicacion.add(comboNombres_cortos,gbc);                                               
                                           panelFiltroAplicacion.revalidate();
                                           panelFiltroAplicacion.repaint();   
                                               
                                           return null;
                              
                                 }
                                                                                                                                                        
                                 private String[] traeNombresCortos(String item){ 

                                         Connection c;
                                         Statement s;
                                         String[] nombres_cortosArray = null;

                                         try{
   
                                             Class.forName("com.mysql.jdbc.Driver");
                                             c = DriverManager.getConnection("jdbc:mysql://172.16.34.21:3306/replicasiipo","test","slipknot");
   
                                             String select = "select nom_corto from datos_examenes where tipo_instr = '" + item + "'";
                                             s = c.createStatement();
                                             ResultSet rs = s.executeQuery(select);
                                         
                                             List<String> nombres_cortos = new ArrayList<>();
                                                                                                                                                                                                   
                                             while( rs.next() ){
                                                         
                                                    String nom_corto = rs.getString(1);
                                                    if( nom_corto != null ){                                                        
                                                        nombres_cortos.add(nom_corto);
                                                    }
          
                                             }
                                                                                                                                                                                                                                            
                                             nombres_cortosArray = new String[nombres_cortos.size()];
                                             int i = 0;
                                             for(Object o : nombres_cortos.toArray()){
                                                 nombres_cortosArray[i] = (String)o;
                                                 i++;
                                             }      
                                                
                                             s.close();
                                             c.close();
                     
                                         }catch(ClassNotFoundException | SQLException e){ e.printStackTrace(); }
                                              
                                         return nombres_cortosArray;       

                                 } 
                                    
                                 };
                                       
                                 swbe.execute();
                            
                          }
                     
                      }
            
              );
              
              gbc = new GridBagConstraints();              
              gbc.gridy = 1;  
              gbc.gridx = 2;
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.WEST;
              panelFiltroInstrumento.add(comboTipoInstr,gbc);
              
    	      datosReporte = new ArrayList<>();    	                                                                                         
              
              etiquetaSubAplicacion = new JLabel("Nombre : ");
              gbc = new GridBagConstraints();
              gbc.gridx = 3;
              gbc.gridy = 1;  
              gbc.weightx = 0.5;              
              gbc.insets  = new Insets(5,5,5,5);              
              gbc.anchor  = GridBagConstraints.WEST;
              panelFiltroInstrumento.add(etiquetaSubAplicacion,gbc); 
              
              comboNombres_cortos = new JComboBox<>();
              gbc = new GridBagConstraints();
              gbc.gridx = 4;
              gbc.gridy = 1;  
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor = GridBagConstraints.WEST;              
              panelFiltroInstrumento.add(comboNombres_cortos,gbc);                                          
              
              panelFiltroFechas = new JPanel(new GridBagLayout());
                            
              etiquetaFi = new JLabel("Fecha inicial");
              gbc = new GridBagConstraints();
              gbc.gridx = 0;
              gbc.gridy = 2;    
              gbc.weightx = 0.1;
              gbc.weighty = 0.1;
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.WEST;
              panelFiltroFechas.add(etiquetaFi,gbc);   
              
              etiquetaAñoFi = new JLabel("Año : ");
              gbc = new GridBagConstraints();
              gbc.gridx = 1;
              gbc.gridy = 2;    
              gbc.weightx = 0.1;
              gbc.weighty = 0.1;
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.WEST;
              panelFiltroFechas.add(etiquetaAñoFi,gbc);   
              
              comboAñoFi = new JComboBox<>(años);
              comboAñoFi.addActionListener(new ActionListener() {

                         @Override
                         public void actionPerformed(ActionEvent e) {
                                 
                                JComboBox event = (JComboBox)e.getSource();
                                añoFiSeleccionado = ((String)event.getSelectedItem()).trim();
                          
                         }
                         
              });
              
              gbc = new GridBagConstraints();
              gbc.gridx = 1;
              gbc.gridy = 2;    
              gbc.weightx = 0.1;
              gbc.weighty = 0.1;
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.EAST;
              panelFiltroFechas.add(comboAñoFi,gbc);                             
              
              etiquetaMesFi = new JLabel("Mes : ");
              gbc = new GridBagConstraints();
              gbc.gridx = 3;
              gbc.gridy = 2;                     
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.WEST;
              panelFiltroFechas.add(etiquetaMesFi,gbc);   
              
              comboMesFi = new JComboBox<>(meses);
              comboMesFi.addActionListener(new ActionListener(){
                	
                         @Override
                         public void actionPerformed(ActionEvent e) {                                                             
                                 
                                SwingWorker<Void,Void> sw;
                                sw = new SwingWorker<Void,Void>(){                                                                                
                                     
                                @Override
                                protected Void doInBackground() {                                                                                    
                                                    
                                          try{
                                              
                                              System.out.println("en el fondo");
                                              String comando = GenerarReportes.this.comboAñoFi.getActionCommand();
                                              System.out.println("comando " + comando);

                                              int año = 2012;                                                    
                                                    
                                              if( comando.equals("comboBoxChanged") ){                                                                                                                
                                                        
                                                  for( int i = 0; i <= años.length - 1; i++ ){                                                             
                                                       if( añoFiSeleccionado.equals(años[i])){ año = i + 1; }
                                                  }
                                                    
                                                  int mes = comboMesFi.getSelectedIndex() + 1;
                                                   
                                                  System.out.println(año + " " + mes);
                                
                                                  Chronology cronologia = ISOChronology.getInstance();
                                                  DateTimeField dtf = cronologia.dayOfMonth();
                                                  LocalDate ld = new LocalDate(año,mes,1);
                                
                                                  int dias = dtf.getMaximumValue(ld); 
                                                  System.out.println("La cantidad de dias del mes " + mes  + " del año " + añoFiSeleccionado + " son : " + dias);
                                                  String cadenasDias [] = new String[dias];
                                                   
                                                  for( int i = 0; i < dias; i++ ){                                                                                                            
                                                       cadenasDias[i] = String.valueOf( i + 1 );                                                      
                                                  }                                                                                                                                                                                                                            
                                                        
                                                  GenerarReportes.panelFiltroFechas.remove(GenerarReportes.this.comboAñoFi);
                                                    
                                                  GenerarReportes.this.comboAñoFi = new JComboBox<>(cadenasDias);
                                                  gbc = new GridBagConstraints();
                                                  gbc.gridx = 6;
                                                  gbc.gridy = 2;    
                                                  gbc.weightx = 0.1;                            
                                                  gbc.insets = new Insets(5,5,5,5);
                                                  gbc.anchor  = GridBagConstraints.EAST;
                                                  GenerarReportes.panelFiltroFechas.add(GenerarReportes.this.comboAñoFi,gbc);
                                                        
                                         }}catch(Exception e){ e.printStackTrace();}
                               
                                         return null;
                                         
                                }
                                         
                                @Override
                                public void done(){                                                                                               
                                                
                                       GenerarReportes.panelFiltroAplicacion.revalidate();
                                       GenerarReportes.panelFiltroAplicacion.repaint();
                                              
                                }
                                         
                           };
                                
                           sw.execute();
                                                          
                     }    
                         
              });
              
              gbc = new GridBagConstraints();
              gbc.gridx = 4;
              gbc.gridy = 2;    
              gbc.weightx = 0.1;              
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.EAST;
              panelFiltroFechas.add(comboMesFi,gbc); 
              
              etiquetaDiaFi = new JLabel("Dia : ");
              gbc = new GridBagConstraints();
              gbc.gridx = 5;
              gbc.gridy = 2;    
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.WEST;
              panelFiltroFechas.add(etiquetaDiaFi,gbc);   
              
              comboAñoFi = new JComboBox<>();
              gbc = new GridBagConstraints();
              gbc.gridx = 6;
              gbc.gridy = 2;    
              gbc.weightx = 0.1;                            
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.EAST;
              panelFiltroFechas.add(comboAñoFi,gbc);   
              
              etiquetaFf = new JLabel("Fecha final");
              gbc = new GridBagConstraints();
              gbc.gridx = 0;
              gbc.gridy = 3;    
              gbc.weightx = 0.1;
              gbc.weighty = 0.1;
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.WEST;
              panelFiltroFechas.add(etiquetaFf,gbc);   
              
              etiquetaAñoFf = new JLabel("Año : ");
              gbc = new GridBagConstraints();
              gbc.gridx = 1;
              gbc.gridy = 3;    
              gbc.weightx = 0.1;
              gbc.weighty = 0.1;
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.WEST;
              panelFiltroFechas.add(etiquetaAñoFf,gbc);   
              
              comboAñoFf = new JComboBox<>(años);
              comboAñoFf.addActionListener(new ActionListener() {

                         @Override
                         public void actionPerformed(ActionEvent e) {
                                 
                                JComboBox event = (JComboBox)e.getSource();
                                añoFfSeleccionado = ((String)event.getSelectedItem()).trim();                                
                          
                         }
                         
              });
              
              gbc = new GridBagConstraints();
              gbc.gridx = 1;
              gbc.gridy = 3;    
              gbc.weightx = 0.1;
              gbc.weighty = 0.1;
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.EAST;
              panelFiltroFechas.add(comboAñoFf,gbc);                             
              
              etiquetaMesFf = new JLabel("Mes : ");
              gbc = new GridBagConstraints();
              gbc.gridx = 3;
              gbc.gridy = 3;                     
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.WEST;
              panelFiltroFechas.add(etiquetaMesFf,gbc);   
              
              comboMesFf = new JComboBox<>(meses);
              comboMesFf.addActionListener(new ActionListener(){
                	
                         @Override
                         public void actionPerformed(ActionEvent e) {                                                             
                                 
                                SwingWorker<Void,Void> sw;
                                sw = new SwingWorker<Void,Void>(){                                                                                
                                     
                                @Override
                                protected Void doInBackground() {                                                                                    
                                                    
                                          try{
                                              
                                              System.out.println("en el fondo");
                                              String comando = GenerarReportes.this.comboAñoFf.getActionCommand();
                                              System.out.println("comando " + comando);

                                              int año = 2012;                                                    
                                                    
                                              if( comando.equals("comboBoxChanged") ){                                                                                                                
                                                        
                                                  for( int i = 0; i <= años.length - 1; i++ ){                                                             
                                                       if( añoFfSeleccionado.equals(años[i])){ año = i + 1; }
                                                  }
                                                    
                                                  int mes = comboMesFf.getSelectedIndex() + 1;
                                                   
                                                  System.out.println(año + " " + mes);
                                
                                                  Chronology cronologia = ISOChronology.getInstance();
                                                  DateTimeField dtf = cronologia.dayOfMonth();
                                                  LocalDate ld = new LocalDate(año,mes,1);
                                
                                                  int dias = dtf.getMaximumValue(ld); 
                                                  System.out.println("La cantidad de dias del mes " + mes  + " del año " + añoFfSeleccionado + " son : " + dias);
                                                  String cadenasDias [] = new String[dias];
                                                   
                                                  for( int i = 0; i < dias; i++ ){                                                                                                            
                                                       cadenasDias[i] = String.valueOf( i + 1 );                                                      
                                                  }                                                                                                                                                                                                                            
                                                        
                                                  GenerarReportes.panelFiltroAplicacion.remove(GenerarReportes.this.comboAñoFf);
                                                    
                                                  GenerarReportes.this.comboAñoFf = new JComboBox<>(cadenasDias);
                                                  gbc = new GridBagConstraints();
                                                  gbc.gridx = 6;
                                                  gbc.gridy = 3;    
                                                  gbc.weightx = 0.1;                            
                                                  gbc.insets = new Insets(5,5,5,5);
                                                  gbc.anchor  = GridBagConstraints.EAST;
                                                  GenerarReportes.panelFiltroAplicacion.add(GenerarReportes.this.comboAñoFf,gbc);
                                                        
                                         }}catch(Exception e){ e.printStackTrace();}
                               
                                         return null;
                                         
                                }
                                         
                                @Override
                                public void done(){                                                                                               
                                                
                                       GenerarReportes.panelFiltroAplicacion.revalidate();
                                       GenerarReportes.panelFiltroAplicacion.repaint();
                                              
                                }
                                         
                           };
                                
                           sw.execute();
                                                           
                     }    
                         
              });
              
              gbc = new GridBagConstraints();
              gbc.gridx = 4;
              gbc.gridy = 3;    
              gbc.weightx = 0.1;              
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.EAST;
              panelFiltroFechas.add(comboMesFf,gbc); 
              
              etiquetaDiaFf = new JLabel("Dia : ");
              gbc = new GridBagConstraints();
              gbc.gridx = 5;
              gbc.gridy = 3;    
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.WEST;
              panelFiltroFechas.add(etiquetaDiaFf,gbc);   
              
              comboAñoFf = new JComboBox<>();
              gbc = new GridBagConstraints();
              gbc.gridx = 6;
              gbc.gridy = 3;    
              gbc.weightx = 0.1;                            
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.EAST;
              panelFiltroFechas.add(comboAñoFf,gbc);   
              
              botonGenerarReporte = new JButton("Generar Reporte");
              botonGenerarReporte.addActionListener(new ActionListener() {

                      @Override
                      public void actionPerformed(ActionEvent e) {
                          
                             SwingWorker<Void,Void> sw;
                             sw = new SwingWorker<Void,Void>(){
                		     
                 	     Connection c = null;
                             Statement  s = null;                  
                             ResultSet  rsMysql = null;                                 
                              
                	     @Override
                	     protected Void doInBackground() throws Exception {
                		        	                 	                	        	                                                  
                                       try{
                                    	                                                                                         
                                           Class.forName("com.mysql.jdbc.Driver");                 
                                           c = DriverManager.getConnection("jdbc:mysql://172.16.34.21:3306/ceneval","user","slipknot");
                                           s = c.createStatement();
                                           
                                           String subtipo;
                                           String select = "";
                                           
                                           System.out.println(select);
                                           
                                           if( comboNombres_cortos.getSelectedItem() != null ){ 
                                               subtipo = " and subtipo = '" + (String)comboNombres_cortos.getSelectedItem()+ "'";
                                           }else{ subtipo = ""; }
                                           
                                           select = "select * from viimagenes where year(fecha_registro) = " + comboAños.getSelectedItem() + 
                                                    " and month(fecha_registro) = " + (comboMes.getSelectedIndex() + 1) + " and tipo_aplicacion = '" + 
                                                    (String)comboTipoInstr.getSelectedItem() + "'" + subtipo + " order by estado";
                                           
                                           System.out.println(select);
                                           
                                           rsMysql = s.executeQuery(select);
                                                                    
                                           DefaultTableModel dtm = new DefaultTableModel();
                                          
                                           TableCellRenderer renderer = new JComponentTableCellRenderer();                                                                                                                                                                                                                  
                                          
                                           for(int l = 0;l <= nombresCantidad;l++){ dtm.addColumn(""); }
                                          
                                               tabla.setModel(dtm);
                                               TableColumnModel columnModel = tabla.getColumnModel();
                                          
                                               for( int k = 0; k <= nombresCantidad; k++ ){
                                                    TableColumn tcTemp = columnModel.getColumn(k);                                                              
                                                    JLabel encabezado = new JLabel(nombresColumnas[k]);
                                                    tcTemp.setHeaderRenderer(renderer);
                                                    tcTemp.setHeaderValue(encabezado);
                                               }
                                                                                             
                                               while(rsMysql.next()){               
                                                   
                                                     int numApp         = rsMysql.getInt(2);
                                                     String instr       = rsMysql.getString(3);
                                                     String nombre      = rsMysql.getString(4);
                     		                     Date alta          = rsMysql.getDate(5);
                                                     Date registro      = rsMysql.getDate(6);
                                                     int imagReg        = rsMysql.getInt(7);
                                                     int imagRes        = rsMysql.getInt(8);
                                                     int preg           = rsMysql.getInt(9);
                                                     int pregbpm        = rsMysql.getInt(10);
                                                     int pregmc         = rsMysql.getInt(11);
                                                     int pres           = rsMysql.getInt(12);
                                                     int presbpm        = rsMysql.getInt(13);
                                                     int presmc         = rsMysql.getInt(14);                                                     
                                                     String institucion = rsMysql.getString(16);
                                                     String estado      = rsMysql.getString(17);
                                                     String observacion = rsMysql.getString(18);
                                                     
                                                     System.out.println(numApp + " " + instr + " " + nombre + " " + alta.toString() + " " + registro.toString() + 
                                                                        " " + imagReg + " " + imagRes + " " + preg + " " + pregbpm + " " + pregmc + " " + pres +
                                                                        " " + presbpm + " " + presmc + " " + institucion + " " + estado + " " + observacion);
                                                     
                                            	     Object[] datos = new Object[]{ numApp,instr,nombre,alta.toString(),registro.toString(),imagReg,imagRes,preg,
                                                                                    pregbpm,pregmc,pres,presbpm,presmc,institucion,estado,observacion };
                             
                                                     datosReporte.add(datos);                             
                                                     dtm.addRow(datos);                                                                  
                                               }
                                          
                                               s.close();
                                               c.close();
                                               rsMysql.close();
                                                                                                                                                                                                                                                                                                                                                                        
                                       }catch(ClassNotFoundException | SQLException e){ e.printStackTrace(); }
                                       finally{
                                               try{
                                                   rsMysql.close();                                                        
                                                   s.close();
                                                   c.close();
                                               }catch(Exception e){ e.printStackTrace(); }
                                       }                		        	        
                                   
                                       return null;
                                       
                	     }
                             
                             @Override
                             public void done(){
                                  
                                    GenerarReportes.this.revalidate();
                                    GenerarReportes.this.repaint();                                                            
                                                                 
                             }
                		                          		                          		                          		                          		  
                	  };
                	  
                	  sw.execute();                	                      

              
                      }
              });
              
              gbc = new GridBagConstraints();
              gbc.gridx = 1;
              gbc.gridy = 5;    
              gbc.weightx = 0.1;
              gbc.weighty = 0.1;
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.WEST;
              //panelFiltroAplicacion.add(botonGenerarReporte,gbc);   
                           
              botonImprimirReporte = new JButton("Imprimir Reporte");
              botonImprimirReporte.addActionListener(new ActionListener() {

                      @Override
                      public void actionPerformed(ActionEvent e) {
              
                      }
              });
                            
              gbc = new GridBagConstraints();
              gbc.gridx = 3;
              gbc.gridy = 5;    
              gbc.weightx = 0.1;
              gbc.weighty = 0.1;
              gbc.insets = new Insets(5,5,5,5);
              gbc.anchor  = GridBagConstraints.WEST;
              //panelFiltroAplicacion.add(botonImprimirReporte,gbc);                 
                            
              tabla = new JTable();
             
              panelTabla = new JScrollPane(tabla);              
              panelTabla.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
              panelTabla.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
              
              gbc = new GridBagConstraints();
              gbc.gridx = 1;
              gbc.gridy = 0;    
              gbc.weightx = 0.1;
              gbc.weighty = 0.1;                            
              add(panelFiltroAplicacion,gbc);
              
              gbc = new GridBagConstraints();
              gbc.gridx = 0;
              gbc.gridy = 1;    
              gbc.weightx = 0.1;
              gbc.weighty = 0.1;    
              gbc.gridwidth = 9;
              gbc.insets = new Insets(0, 10, 0, 10);
              gbc.fill = GridBagConstraints.HORIZONTAL;
              gbc.anchor = GridBagConstraints.NORTH;
              add(panelTabla,gbc);                               
                                        
       }                         
       
       public void actionPerformed(ActionEvent ae){                            
                               
              if( ae.getSource() == botonImprimirReporte ){
                	                  	  
              	  /*PrinterJob pj = PrinterJob.getPrinterJob();                                                      
                    PrintRequestAttributeSet pras = new HashPrintRequestAttributeSet();
                    PageFormat pf = pj.pageDialog(pras);
                    pj.setPrintable(GenerarReportes.this, pf);
                    boolean ok = pj.printDialog(pras);
                     
                    if( ok ){
                        try{
                            pj.print(pras);                             
                        }catch(Exception e){ e.printStackTrace(); }
                    }
                      
                    datosReporte.clear();
                  */
                	  
                  try{ GeneraReportePdf(); }
                  catch(Exception e){ e.printStackTrace(); }
                      
              }
                                                                                                                     
       }
       
       public void GeneraReportePdf(){    	         	    
    	      
    	      try{
    	    	  
    	    	  Document pdf = new Document();
  		  pdf.setPageSize(PageSize.A4.rotate());
  		  PdfWriter.getInstance(pdf,new FileOutputStream("Test.pdf"));
  		             		           
  		  HeaderFooter encabezado = new HeaderFooter(new Phrase("Direccion de procesos opticos y calificacion.Validacion de posiciones e imagenes de "+
  		                                                        "lectura optica de " + comboTipoInstr.getSelectedItem() + "/" + 
  		                                                        comboNombres_cortos.getSelectedItem() + " de " + comboMes.getSelectedItem() + "-" + 
  		                                                       comboAños.getSelectedItem(), new Font(Font.TIMES_ROMAN,12f,Font.BOLD)), false);
  		          
  		  SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
  		  String fCadena = sdf.format(new Date());
  		  HeaderFooter pie = new HeaderFooter(new Phrase("Usuario: " + name + " " + fCadena + "         ",
  		                                      new Font(Font.TIMES_ROMAN,10f,Font.BOLD)),true);
  		           
  		  pdf.setHeader(encabezado);
  		  pdf.setFooter(pie);
  		           
  		  pdf.open();
  		  Paragraph parrafo = new Paragraph();
  		            		            		          
  		  Font fuenteDatos = new Font(Font.TIMES_ROMAN,7f);
		  fuenteDatos.setStyle(Font.NORMAL);  		          
		          
		  float[] anchosCelda = {0.06f,0.05f,0.05f,0.05f,0.05f,0.05f,0.05f,0.05f,0.05f,0.05f,0.05f,0.05f,0.39f};
		  PdfPTable tablaPdf;		          
                  tablaPdf = new PdfPTable(anchosCelda);
		  tablaPdf.setWidthPercentage(100);
		          
		  String[] encabezados = {"Aplicacion","Tipo","Fecha registro","Fecha alta","Imagenes","Preg","Preg BPM","Preg Mcontrol","Pres",
  		                          "Pres BPM ","Pres Mcontrol","Estado","Institucion"};
    
                  Font fuenteEncabezados = new Font(Font.TIMES_ROMAN,8f);                  
                  fuenteEncabezados.setStyle(Font.BOLD);
                  
                  for( int i = 0; i <= (encabezados.length - 1); i++ ){
                       Phrase fraseEncabezados = new Phrase();
                       fraseEncabezados.setFont(fuenteEncabezados);
          	       fraseEncabezados.add(encabezados[i]);
          	       PdfPCell celda = new PdfPCell(fraseEncabezados);
                       celda.setFixedHeight(20);  	    	      	            	    	
          	       tablaPdf.addCell(celda);                	   
                  }
		          
                  for( Object[] ao: datosReporte ){  		        	     		        	   
      	               for( Object dato: ao ){      	            	    
      	            	    Phrase frase = new Phrase();
      	            	    frase.setFont(fuenteDatos);
      	            	    if( dato instanceof String ){      	    
      	            	    	frase.add(dato);
      	            	    	PdfPCell celda = new PdfPCell(frase);
      	                        celda.setFixedHeight(20);  	    	      	            	    	
      	            	    	tablaPdf.addCell(celda);
      	            	    }else{ 	      	  
      	            	    	  frase.add(String.valueOf(dato));
      	            	     	  PdfPCell celda = new PdfPCell(frase);
      	            	     	  celda.setFixedHeight(20);        	            	    
        	          	  tablaPdf.addCell(celda);
        	            }      	            	  
      	               }      	                     	               
      	               
                  } 
  		          
  		  pdf.add(tablaPdf);  		             		                       		           
  		  pdf.close();
  		           
    	      }catch(FileNotFoundException | DocumentException e){ e.printStackTrace(); }    	    	      	             	    	  	       	    	       	              	    	 
    	    
       }              	  
                 
}

class JComponentTableCellRenderer implements TableCellRenderer {
    
      @Override
      public Component getTableCellRendererComponent(JTable table, Object value, 
             boolean isSelected, boolean hasFocus, int row, int column) {
             return (JComponent)value;
      }
      
}

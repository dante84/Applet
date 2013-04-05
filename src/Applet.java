
// @author Daniel.Meza

import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.io.File;
import java.security.AccessController;
import java.security.PrivilegedAction;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Locale;
import java.util.Map;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTabbedPane;
import javax.swing.JTextField;
import javax.swing.SwingWorker;
import javax.jnlp.ExtendedService;
import javax.swing.JFrame;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
 
public class Applet extends JPanel {
     
       private JLabel eRutaExcel,eMes,eTExamen,eSTExamen,eExcel;
       private JPanel panelAplicacion,panelExcel;              
       private JTextField cRuta,cExcel;
       private JButton bExaminar,bProcesarAplicaciones,bExcel;
       private JFileChooser fileChooser;
       private ArrayList<Object> aplicaciones,fechas,instituciones,registro,respuesta;
       private File aExcel;
       private ExtendedService extendedService;
       private JComboBox<String> comboMes,comboTExamen,comboSTExamen;
       private GridBagConstraints gbc;
       
       private final String[] meses   =     {"Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"};
       
       private final String[] tExamen =     {"AC286",
                                             "AC286CCC",
                                             "AC286CMT",
                                             "AC286CCE",
                                             "AC286CCS",
                                             "ACREDITA",
                                             "ACREL",
                                             "ACRETSU",
                                             "BULATS",
                                             "DELEGACIONES",
                                             "DGEP",
                                             "DGESPE",
                                             "ECCYPEC",
                                             "ECELE",
                                             "EGEL",
                                             "EGETSU",
                                             "EGREB",
                                             "EGREMS",
                                             "ENAMS",
                                             "ENLACE",
                                             "EPROM",
                                             "EUC",
                                             "EUCCA",
                                             "EXANI",
                                             "EXIL",
                                             "EXTRA-ES",
                                             "GESE",
                                             "IEE-CEF",
                                             "IFE",
                                             "ISE",
                                             "MINNESOTA",
                                             "PPD",
                                             "TKT",
                                             "UPN",
                                             "UPN_LE",
                                             "LEPEPMI"};
              
       private final String[][] stExamen =  {{""},
                                             {""},
                                             {""},
                                             {""},
                                             {""},
                                             {"ACREDITA_BACH","ACREDITA_SEC"},
                                             {"ACREL_DII","ACREL_EIN","ACREL_EPRE","ACREL_EPRIM","ACREL_MODA"},
                                             {"ACRETSU_CI","ACRETSU_PFP","ACRETSU_PI"},
                                             {""},
                                             {"IZTACALCO","TLAHUAC"},
                                             {""},
                                             {"EGC","EXI"},
                                             {""},
                                             {""},
                                             {""},
                                             {""},
                                             {""},
                                             {""},
                                             {""},
                                             {""},
                                             {""},
                                             {"EUC_ENFER","EUC_EO","EUC_ODON","EUC_PSI","EUC_QUICLI","EUC_TENFER"},
                                             {"EUCCA_ACCIDEN","EUCCA_AUD_ACCIDEN","EUCCA_AUD_DANOS","EUCCA_AUD_FIANZAS","EUCCA_AUDIT","EUCCA_AUD_RENTAS","EUCCA_AUD_VIDA","EUCCA_DANOS","EUCCA_FIANZAS","EUCCA_RENTAS","EUCCA_VIDA"},
                                             {"E2E","EXANI_I","EXANI_II","EXANI_III","PREEXANI_I","PREEXANI_II"},
                                             {""},
                                             {"EXTRA_ES_BAS","EXTRA_ES_EXP","EXTRA_ES_MET","EXTRA_ES_MUES"},
                                             {""},
                                             {""},
                                             {""},
                                             {""},
                                             {""},
                                             {""},
                                             {""},
                                             {""},
                                             {""},
                                             {""}};                      
       
       private Map<String,Integer> mapaUpn;
       
       private Applet(){
                    
               mapaUpn = new HashMap<>();      
               mapaUpn.put("UPN", 1);
               mapaUpn.put("UPN_LE", 2);
               mapaUpn.put("LEPEPMI", 3);
               
               setLayout(new GridBagLayout());
               
               gbc = new GridBagConstraints();
               
               panelAplicacion = new JPanel(new GridBagLayout());
               panelExcel = new JPanel(new GridBagLayout());
               
               setSize(1400,450);
               
               eRutaExcel = new JLabel("Ruta :");                              
               gbc.gridx = 0;
               gbc.gridy = 0;               
               gbc.weightx = 0.1;               
               gbc.insets = new Insets(5,5,5,5);               
               panelAplicacion.add(eRutaExcel,gbc);
               
               cRuta = new JTextField(40);
               gbc = new GridBagConstraints();
               gbc.gridx = 1;
               gbc.gridy = 0;
               gbc.insets = new Insets(5,5,5,5);
               gbc.anchor = GridBagConstraints.CENTER;
               gbc.gridwidth = 3;
               gbc.weightx = 0.2;               
               panelAplicacion.add(cRuta,gbc);                              
               
               bExaminar = new JButton("Examinar");
                              
               bExaminar.addActionListener(
                         new ActionListener() {
                             @Override
                             public void actionPerformed(ActionEvent e) {
                                 
                                    fileChooser = new JFileChooser();
                                    fileChooser.setMultiSelectionEnabled(false);
                                    
                                    fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                                    
                                    int valor = fileChooser.showOpenDialog(null);
                                    
                                    if( valor == JFileChooser.APPROVE_OPTION ){
                                        File f = fileChooser.getSelectedFile();
                                        cRuta.setText(f.getAbsolutePath());
                                    }
                                    
                             }
                             
                       }
                      
               );
               
               gbc = new GridBagConstraints();
               gbc.gridx = 4;
               gbc.gridy = 0;         
               gbc.weightx = 0.1;               
               gbc.anchor = GridBagConstraints.WEST;
               gbc.insets = new Insets(5,5,5,5);
               panelAplicacion.add(bExaminar,gbc);     

               gbc = new GridBagConstraints();
               gbc.gridx = 0;
               gbc.gridy = 1;
               gbc.insets = new Insets(5,5,5,5);
               eMes = new JLabel("Mes :");                                            
               panelAplicacion.add(eMes,gbc);
               
               gbc = new GridBagConstraints();
               gbc.gridx = 1;
               gbc.gridy = 1;
               gbc.anchor = GridBagConstraints.WEST;
               gbc.insets = new Insets(5,5,5,5);
               comboMes = new JComboBox<>(meses);
               panelAplicacion.add(comboMes,gbc);
              
               gbc = new GridBagConstraints();
               gbc.gridx = 2;
               gbc.gridy = 1;               
               gbc.anchor = GridBagConstraints.WEST;               
               gbc.insets = new Insets(5,5,5,5);
               eTExamen = new JLabel("Examen :");
               panelAplicacion.add(eTExamen,gbc);
               
               comboTExamen = new JComboBox<>(tExamen);
               comboTExamen.setSelectedIndex(0);
               comboTExamen.addActionListener(
                            new ActionListener() {
                                @Override
                                public void actionPerformed(ActionEvent e) {                                                                                                                    	                            	                                          
                                                                                                              
                                       panelAplicacion.remove(comboSTExamen);
                                       
                                       int itExamen = comboTExamen.getSelectedIndex();                                           
                                       System.out.println("indice = " + itExamen);
                                       
                                       gbc = new GridBagConstraints();
                                       
                                       if( itExamen == 0  || itExamen == 1  || itExamen == 2  || itExamen == 3  || itExamen == 4  || 
                                           itExamen == 7  || itExamen == 8  || itExamen == 9  || itExamen == 10 || itExamen == 11 || 
                                           itExamen == 12 || itExamen == 14 || itExamen == 22 || itExamen == 24 || itExamen == 25 || 
                                           itExamen == 26 || itExamen == 27 || itExamen == 28 || itExamen == 29 ){                                                                                                                                
                                           
                                           gbc.gridx = 4;
                                           gbc.gridy = 1;
                                           comboSTExamen = new JComboBox<>();
                                           gbc.insets = new Insets(5,5,5,5);
                                           gbc.anchor = GridBagConstraints.WEST;               
                                           panelAplicacion.add(comboSTExamen,gbc);                                             
                                           panelAplicacion.revalidate();
                                           panelAplicacion.repaint();              
                                           
                                       }else{       
                                           
                                             System.out.println("Longitud = " + stExamen[itExamen].length);
                                             comboSTExamen = new JComboBox<>(stExamen[itExamen]);                                      
                                             gbc = new GridBagConstraints();
                                             gbc.gridx = 4;
                                             gbc.gridy = 1;                                            
                                             gbc.insets = new Insets(5,5,5,5);
                                             gbc.anchor = GridBagConstraints.WEST;               
                                             panelAplicacion.add(comboSTExamen,gbc);                                               
                                             panelAplicacion.revalidate();
                                             panelAplicacion.repaint();   
                                             
                                       }
                                                                                                       
                                }
                          
                            }
                 
               );
                                 
               gbc = new GridBagConstraints();
               gbc.gridx = 2;
               gbc.gridy = 1;
               gbc.anchor = GridBagConstraints.CENTER;
               gbc.weightx = 0.1;
               gbc.insets = new Insets(5,5,5,5);
               panelAplicacion.add(comboTExamen,gbc);
               
               comboSTExamen = new JComboBox<>(stExamen[comboTExamen.getSelectedIndex()]);                                                    
               gbc = new GridBagConstraints();
               gbc.gridx = 4;
               gbc.gridy = 1;                                            
               gbc.insets = new Insets(5,5,5,5);
               gbc.anchor = GridBagConstraints.WEST;                                                            
              
               eSTExamen = new JLabel("Subtipo :");
               gbc = new GridBagConstraints();
               gbc.gridx = 3;
               gbc.gridy = 1;
               gbc.anchor = GridBagConstraints.WEST;               
               gbc.insets = new Insets(5,5,5,5);               
               panelAplicacion.add(eSTExamen,gbc);
               
               comboSTExamen = new JComboBox<>();
               gbc = new GridBagConstraints();
               gbc.gridx = 4;
               gbc.gridy = 1;               
               gbc.insets = new Insets(5,5,5,5);
               gbc.anchor = GridBagConstraints.WEST;               
               panelAplicacion.add(comboSTExamen,gbc);  
                         
               eExcel = new JLabel("Excel :");
               gbc = new GridBagConstraints();
               gbc.gridx = 0;
               gbc.gridy = 0;               
               gbc.insets = new Insets(5,5,5,0);
               gbc.anchor = GridBagConstraints.WEST;               
               panelExcel.add(eExcel,gbc);
               
               cExcel = new JTextField(40);  
               gbc = new GridBagConstraints();
               gbc.gridx = 1;
               gbc.gridy = 0;
               gbc.gridwidth = 2;
               gbc.anchor = GridBagConstraints.WEST;               
               gbc.insets = new Insets(5,5,5,5);                                           
               panelExcel.add(cExcel,gbc);                             
               
               bExcel = new JButton("Examinar");                           
               bExcel.addActionListener(
                      new ActionListener() {
                          @Override
                          public void actionPerformed(ActionEvent e) {                                
                                 SwingWorker swbe = new SwingWorker() {
                                             @Override
                                             protected Object doInBackground() throws Exception {
                                               
                                                       fileChooser = new JFileChooser();
                                                       fileChooser.setMultiSelectionEnabled(false);
                                    
                                                       fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
                                    
                                                       int valor = fileChooser.showOpenDialog(null);
                                     
                                                       if( valor == JFileChooser.APPROVE_OPTION){ 
                                                           
                                                           aExcel = (File)AccessController.doPrivileged(new PrivilegedAction() {
                                                                           @Override
                                                                           public Object run(){
                                                                                  File inputFile1 = fileChooser.getSelectedFile();
                                                                                  return inputFile1;
                                                                           }
                                                                           
                                                                        }
                                                                   
                                                                    );                                                                                                        
                                        
                                                           cExcel.setText(aExcel.getAbsolutePath());
                                        
                                                       }
                                                       
                                                       return null;
                                    
                                             }
                              
                                 }; 
                                 
                                 swbe.execute();
                                                                                         
                          }                                                   
                             
                      }
                     
               );                                                                  
               
               gbc = new GridBagConstraints();
               gbc.gridx = 3;
               gbc.gridy = 0;
               gbc.insets = new Insets(5,5,5,5);
               panelExcel.add(bExcel,gbc);                                           
                              
               bProcesarAplicaciones = new JButton("Procesar Aplicaciones");
               bProcesarAplicaciones.addActionListener(new ActionListener() {

                       @Override
                       public void actionPerformed(ActionEvent e) {
                                                                                            
                              SwingWorker<Void,Void> chamber = new SwingWorker<Void,Void>() {

                                          @Override
                                          protected Void doInBackground() {                                                                                                
                                                    
                                                    String cte = (String)comboTExamen.getSelectedItem();
                                                    String cste = (String)comboSTExamen.getSelectedItem();                                                                                
                                       
                                                    int cmi = comboMes.getSelectedIndex();
                                                    String rutaExcel = cExcel.getText().trim();
                                                    String rutaDats = cRuta.getText().trim();
                                       
                                                    aplicaciones  = new ArrayList<>();
                                                    fechas        = new ArrayList<>();
                                                    instituciones = new ArrayList<>();
                                                    registro      = new ArrayList<>();
                                                    respuesta     = new ArrayList<>();
                                                    
                                                    if( rutaExcel.equals("") ){
                                                        JOptionPane.showMessageDialog(
                                                                    null,
                                                                    "Debes especificar un archivo excel.",
                                                                    "Especificar excel",
                                                                    JOptionPane.WARNING_MESSAGE);
                                                        return null;
                                                    }
                                                    
                                                    if( rutaDats.equals("") ){
                                                        JOptionPane.showMessageDialog(
                                                                    null,
                                                                    "La ruta es invalida, verifica.",
                                                                    "Ruta invalida",
                                                                    JOptionPane.WARNING_MESSAGE);
                                                        return null;
                                                    }                                                                                                       
                                                    
                                                    try{
                                                        
                                                        System.out.println("Dentro del try antes del workbook " + aExcel.getAbsolutePath());                                                        
                                                        Workbook wb = WorkbookFactory.create(aExcel);                                                        
                                                                                                                
                                                        Sheet hoja = wb.getSheetAt(0);      
                                                        
                                                        System.out.println("Nombre hoja " + hoja.getSheetName());
                                                        
                                                        Iterator<Row> rowIt = hoja.rowIterator();                                    
                                                        rowIt.next();
                                       
                                                        DateFormat df = new SimpleDateFormat("dd-MMM-yy",Locale.ENGLISH);                    
                                       
                                                        System.out.println("cte " + cte + " cste " + cste);
                                                        
                                                        
                                                        if(cste == null){                                                                                                                                                                                                                                                                                                                                                                 
                                                            
                                                           int h = 1;
                                                           
                                                           for( Iterator<Row> it = rowIt; it.hasNext(); ){
                      
                                                                Row r = it.next();
                      
                                                                Cell cFechaInicio        = r.getCell(1); 
                                                                Cell cTipoAplicacion     = r.getCell(11);
                                                                Cell cSTipoAplcacion     = r.getCell(13);
                                                                Cell cInstitucion        = r.getCell(32);
                                                                Cell noRegistradosCell   = r.getCell(21);
                                                                Cell noRespuestaCell     = r.getCell(22);
                                                     
                                                                String scTipoAplicacion  = cTipoAplicacion.getStringCellValue().trim();                                                  
                                                                String scSTipoAplicacion = cSTipoAplcacion.getStringCellValue().trim();
                                                                String scInstitucion     = cInstitucion.getStringCellValue().trim();
                                                                double noRegistrados     = noRegistradosCell.getNumericCellValue();
                                                                double noRespuesta       = noRespuestaCell.getNumericCellValue();
                                                                
                                                                //System.out.println(scTipoAplicacion + " " + scSTipoAplicacion + " " + scInstitucion);
                                                  
                                                                String valor = cFechaInicio.getStringCellValue().trim();                                                                                                            
                                                       
                                                                if( valor.length() < 7 ){ continue; }
                      
                                                                Date fechaExcel = df.parse(valor);
                                                                Calendar c = Calendar.getInstance();
                                                                c.setTime(fechaExcel);
                                                                int fem = c.get(Calendar.MONTH);                                                                                                                    
                                                  
                                                                int itExamen = comboTExamen.getSelectedIndex();                                                                                                                       
                                                                
                                                                if( itExamen == 27 || itExamen == 28 || itExamen == 29 ){
                                                                
                                                                    //System.out.println( itExamen + " " + fechaExcel + " " + fem + " " + cmi + " " + cte + " " + 
                                                                    //                    scTipoAplicacion + " " + " " + scSTipoAplicacion );                                                                                                                                
                                                                
                                                                    if( (fem == cmi) && (cte.equals(scTipoAplicacion)) && (!scSTipoAplicacion.startsWith("[EUC")) && 
                                                                        (itExamen == 0  || itExamen == 1  || itExamen == 2  || itExamen == 3  || itExamen == 4  || 
                                                                         itExamen == 7  || itExamen == 8  || itExamen == 9  || itExamen == 10 || itExamen == 11 || 
                                                                         itExamen == 12 || itExamen == 14 || itExamen == 22 || itExamen == 23 || itExamen == 24 || 
                                                                         itExamen == 25 || itExamen == 26 ) ){
                                                                                                                            
                                                                         Cell cApp   = r.getCell(0);
                                                                         Object oapp = cApp.getStringCellValue();                                                                                                                                                                                            
                                                                                                            
                                                                         if( oapp != null ){                                                      
                                                                             System.out.println( h + " -  " + oapp + " " + fechaExcel + " " + noRegistrados + " " + 
                                                                                                 noRespuesta + " " + scInstitucion  );                                                                                                                                 
                                                                             h++;
                                                                             aplicaciones.add(oapp);   
                                                                             fechas.add(fechaExcel);
                                                                             registro.add(noRegistrados);
                                                                             respuesta.add(respuesta);
                                                                             instituciones.add(scInstitucion);
                                                                         }
                                                     
                                                                     }
                                                                
                                                                }                                                                                                                               
                                                           
                                                           }
                                                        
                                                        }else{                                                                                                                          
                                                              
                                                              int i = 1;
                                                              for(Iterator<Row> it1 = rowIt; it1.hasNext();){
                          
                                                                  Row r1 = it1.next();
                                                                       
                                                                  Cell cFechaInicio1        = r1.getCell(1); 
                                                                  Cell cTipoAplicacion1     = r1.getCell(11);
                                                                  Cell cSTipoAplicacion1    = r1.getCell(13);
                                                                  Cell noRegistradosCell1   = r1.getCell(21);
                                                                  Cell noRespuestaCell1     = r1.getCell(22);
                                                                  Cell cInstitucion1        = r1.getCell(32);
                                                     
                                                                  String scTipoAplicacion1  = cTipoAplicacion1.getStringCellValue().trim();                                                  
                                                                  String scSTipoAplicacion1 = cSTipoAplicacion1.getStringCellValue().trim();
                                                                  String scInstitucion1     = cInstitucion1.getStringCellValue().trim();                                                                                                                                                                                      
                                                                  String valor1             = cFechaInicio1.getStringCellValue().trim();                                                                                                                                                          
                                                                  double noRegistrados1     = noRegistradosCell1.getNumericCellValue();
                                                                  double noRespuesta1       = noRespuestaCell1.getNumericCellValue();
                        
                                                                  if( valor1.length() < 7 ){ continue; }                                                                                                                                    
                      
                                                                      Date fechaExcel1 = df.parse(valor1);
                                                                      Calendar c1      = Calendar.getInstance();
                                                                      c1.setTime(fechaExcel1);
                                                                      int fem1         = c1.get(Calendar.MONTH);                                                                                                                    
                                                                                                                                            
//                                                                    System.out.println(fechaExcel + " " + fem + " " + cmi + " " + cte + " " + scTipoAplicacion + " "
//                                                                                       + cste + " " + scSTipoAplicacion);
//                                                                  
                                                                      if( (fem1 == cmi) && (cte.equals(scTipoAplicacion1)) && (cste.equals(scSTipoAplicacion1))){
                                                      
                                                                           Cell cApp   = r1.getCell(0);
                                                                           Object oapp = cApp.getStringCellValue();                                                     
                                                     
                                                                           if( oapp != null ){                                                                  
                                                                               System.out.println( i + " - " + oapp + " " + fechaExcel1 + " " + noRegistrados1 + 
                                                                                                   " " + noRespuesta1 + " " + scInstitucion1  );                                                                                                                                 
                                                                               i++;
                                                                               aplicaciones.add(oapp);
                                                                               fechas.add(fechaExcel1);
                                                                               registro.add(noRegistrados1);
                                                                               respuesta.add(respuesta);
                                                                               instituciones.add(scInstitucion1);                                                                         
                                                                           }
                                                                           
                                                                      }
                                                                      
                                                              }
                                                     
                                                           }                                                                                                                                                                                                                                                                                                                                                                                                                                                                     
                                                   
                                                 }catch(Exception e){ e.printStackTrace(); }                                                                                                                                                                      
                                                                                                                               
                                                 return null; 
                                                 
                                          }
                                          
                                          protected void tamales(){
                                                                                                                                                                                         
                                          }
                                                                                    
                              };
                                                                                                             
                              chamber.execute();
                              
                         }
                                                                               
                    }
                       
               );
              
               gbc = new GridBagConstraints();
               gbc.gridx = 4;
               gbc.gridy = 2;
               gbc.insets = new Insets(5,5,5,5);
               panelAplicacion.add(bProcesarAplicaciones,gbc);
               
               gbc = new GridBagConstraints();
               gbc.gridx = 0;
               gbc.gridy = 0;
               gbc.anchor = GridBagConstraints.WEST;
               gbc.insets = new Insets(5,5,5,5);
               add(panelExcel,gbc);                              
               
               gbc = new GridBagConstraints();
               gbc.gridx = 0;
               gbc.gridy = 1;
               gbc.insets = new Insets(5,5,5,5);
               add(panelAplicacion,gbc);
                         
       }                                                               
                 
       public static void main(String args[]){
             
              JTabbedPane tabPane = new JTabbedPane();  
              tabPane.addTab("Consolidacion",new Applet());
              JFrame frame = new JFrame("Consolidacion de datos");                                    
              frame.setSize(550,450);
              frame.add(tabPane);                            
              frame.setResizable(false);
              frame.pack();                                    
              frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
              frame.setLocationByPlatform(true);
              frame.setVisible(true);
                                                  
       }
      
}

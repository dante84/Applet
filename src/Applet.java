
// @author Daniel.Meza

import java.awt.Component;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.io.File;
import java.io.IOException;
import java.security.AccessController;
import java.security.PrivilegedAction;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
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
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

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
       private ArrayList<Object> aplicaciones,fechas,instituciones,registro,respuesta,cve_instr;
       private File aExcel;
       private ExtendedService extendedService;
       private JComboBox<String> comboMes,comboTipoInstr,comboNombres_cortos;
       private GridBagConstraints gbc;
       
       private final String[] meses   =     {"Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"};
       
       private final String[] tiposInstrumento = {"AC286","ACRESEC","ACRETSU","ACUERDO","ALI","CEAACES","CONALEP","DGESPE","ECCYPEC","ECELE","ECODEMS","EGAL",
                                                  "EGEL","EGETSU","EPROM","ESPECIALES","EUC","EUCCA","EXANI","EXTRA","IFE","LEPRE_LEPRI","MCEF","Metropolitano",                                                  
                                                  "MINNESOTA","OLIMPIADA","PILOTO","PREESCOLAR_BACH","PREESCOLAR_LIC","SEISP","SSP","TRIF","UPN"
                                                 };
              
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
                                             {"ECCYPEC","ECCYPEC-CL","ECCYPEC-PC"},
                                             {"ECELE - A2","ECELE - B1","ECELE - B2","ECELE - C1"},
                                             {"ACREL - EPRIM","ACREL_EIN","ACREL_EPRE"},
                                             {"ADMINISTRACION","CONTADURÍA","CIENCIAS FARMACÉUTICAS","DERECHO","LICENCIADO EN ENFERMERIA","TÉCNICO EN ENFERMERIA",
                                              "INFORMÁTICA","INGENIERÍA DE SOFTWARE","CIENCIAS COMPUTACIONALES","INGENIERÍA COMPUTACIONAL",
                                              "CIENCIAS AGRONÓMICAS PERFIL FITOTECNIA","CIENCIAS AGRONÓMICAS PERFIL ZOOTÉCNIA",
                                              "CIENCIAS AGRONÓMICAS PERFIL AGROINDUSTRIA","EXIL-CBI","INGENIERÍA CIVIL","INGENIERÍA ELÉCTRICA",
                                              "INGENIERÍA ELECTRÓNICA","INGENIERÍA INDUSTRIAL","INGENIERÍA QUÍMICA","MEDICINA GENERAL",
                                              "MEDICINA VETERINARIA Y ZOOTÉCNIA","ODONTOLOGÍA","PEDAGOGÍA","PSICOLOGÍA CLÍNICA","PSICOLOGÍA EDUCATIVA",
                                              "PSICOLOGÍA INDUSTRIAL","PSICOLOGÍA SOCIAL","TURISMO - GESTIÓN EMPRESARIAL","TURISMO - PLANIFICACIÓN Y DESARROLLO",
                                              "COLEGIO DE CONTADORES","ACTUARÍA","INGENIERÍA MECÁNICA","INGENIERÍA MECÁNICA ELÉCTRICA","MERCADOTECNIA",
                                              "COMERCIO/NEGOCIOS INTERNACIONALES","CIENCIAS QUÍMICAS","ECONOMÍA","REGISTRO EN LÍNEA","NUTRICIÓN",
                                              "CIENCIAS DE LA COMUNICACIÓN","QUÍMICA INDUSTRIAL","QUÍMICA EN ALIMENTOS","PSICOLOGÍA","EUCP-E ( TECNICO )","TURISMO",
                                              "EUCP-E ( LICENCIATURA )","BIOLOGÍA","QUIMICA CLINICA","INGENIERIA MECATRÓNICA","TRABAJO SOCIAL","CIENCIAS AGRÍCOLAS",
                                              "QUIMICA","EXAMEN UNIFORME DE CERTIFICACION","EXAMEN UNICO DE CERTIFICACION PARA PROFESIONALES EN ORTODONCIA",
                                              "EXAMEN UNIFORME DE CERTIFICACION DE LA CONTADURIA PUBLICA",
                                              "EXAMEN UNIFORME DE CERTIFICACION DE LA CONTABILIDAD Y AUDITORIA GUBERNAMENTAL",
                                              "EXAMEN UNIFORME DE CERTIFICACION DE LA CONTABILIDAD GUBERNAMENTAL","EXAMEN UNIFORME DE CERTIFICACION EN FISCAL",
                                              "INGENIERIA EN ALIMENTOS","EXAMEN DE CERTIFICACION POR DISCIPLINAS DE LA CONTADURÍA - CONTABILIDAD"                                              ,
                                              "EXAMEN DE CERTIFICACION POR DISCIPLINAS DE LA CONTADURÍA - FINANZAS"
                                              },
                                             {"EGETSU ADMINISTRACIÓN","EGETSU ADMINISTRACIÓN Y EVALUACIÓN DE PROYECTOS","EGETSU AGROBIOTECNOLOGÍA",
                                              "EGETSU BIOTECNOLOGÍA","EGETSU COMERCIALIZACIÓN","EGETSU CONTABILIDAD CORPORATIVA",
                                              "EGETSU ELECTRÓNICA Y AUTOMATIZACIÓN","EGETSU ELECTRICIDAD Y ELECTRÓNICA INDUSTRIAL","EGETSU INFORMÁTICA",
                                              "EGETSU MANTENIMIENTO INDUSTRIAL","EGETSU MECÁNICA","EGETSU MECÁNICA Y PRODUCTICA","EGETSU METALICA Y AUTOPARTES",
                                              "EGETSU OFIMÁTICA","EGETSU TECNOLOGÍA AMBIENTAL","EGETSU PROCESOS AGROINDUSTRIALES",
                                              "EGETSU PROCESOS DE PRODUCCIÓN TEXTIL","EGETSU PROCESOS DE PRODUCCIÓN","EGETSU TECNOLOGÍA DE ALIMENTOS",
                                              "EGETSU TELEMÁTICA","EGETSU TURISMO","EGETSU MECATRÓNICA","EGETSU ADMINISTRACIÓN DE EMPRESAS TURÍSTICAS",
                                              "EGETSU ADMINISTRACIÓN DE AUTOPARTES Y LOGÍSTICA","EGETSU ARTES GRÁFICAS","EGETSU CONTADURÍA",
                                              "EGETSU CLASIFICACIÓN ARANCELARIA Y DESPACHO ADUANERO","EGETSU INFORMÁTICA ADMINISTRATIVA","EGETSU IDIOMAS",
                                              "EGETSU METÁLICA Y AUTOPARTES","EGETSU PARAMÉDICO","EGETSU QUÍMICA INDUSTRIAL","EGETSU QUÍMICA DE MATERIALES",
                                              "EGETSU REDES, TELECOMUNICACIONES","EGETSU SISTEMAS DE GESTIÓN DE CALIDAD","EGETSU SISTEMAS INFORMÁTICOS","EGETSU CI",
                                              "EGETSU GASTRONOMIA","EGETSU MANTENIMIENTO A MAQUINARIA PESADA","EGETSU NEGOCIOS INTERNACIONALES",
                                              "EGETSU SALUD PÚBLICA","EGETSU DISEÑO Y PRODUCCION INDUSTRIAL","EGETSU SEGURIDAD ALIMENTARIA"},
                                             {""},
                                             {""},
                                             {""},
                                             {"EXAMEN GENERAL PARA LA PROMOCION DE LA POLICIA FEDERAL POLICIA DE REACCION N.C",
                                              "EXAMEN GENERAL PARA LA PROMOCION DE LA POLICIA FEDERAL POLICIA DE REACCION N.B",
                                              "EXAMEN GENERAL PARA LA PROMOCION DE LA POLICIA FEDERAL POLICIA DE REACCION N.A",
                                              "EXAMEN GENERAL PARA LA PROMOCION DE LA POLICIA FEDERAL POLICIA DE INV. NIVEL A",
                                              "EXAMEN GENERAL PARA LA PROMOCION DE LA POLICIA FEDERAL POLICIA DE INV. NIVEL B",
                                              "EXAMEN GENERAL PARA LA PROMOCION DE LA POLICIA FEDERAL POLICIA DE INV. NIVEL C",
                                              "EXAMEN GENERAL PARA LA PROMOCION DE LA POLICIA FEDERAL POLICIA DE PREVENCION N.A",
                                              "EXAMEN GENERAL PARA LA PROMOCION DE LA POLICIA FEDERAL POLICIA DE PREVENCION N.B",
                                              "EXAMEN GENERAL PARA LA PROMOCION DE LA POLICIA FEDERAL POLICIA DE PREVENCION N.C"},
                                             {"ESPECIALES"},
                                             {"EUC_ENFER","EUC_EO","EUC_ODON","EUC_PSI","EUC_QUICLI","EUC_TENFER"},
                                             {"EXAMEN UNICO DE CERTIFICACION DEL COLEGIO DE ACTUARIOS AUDITORIA-VIDA",
                                              "EXAMEN UNICO DE CERTIFICACION DEL COLEGIO DE ACTUARIOS VIDA",
                                              "EXAMEN UNICO DE CERTIFICACION DEL COLEGIO DE ACTUARIOS ACCIDENTES",
                                              "EXAMEN UNICO DE CERTIFICACION DEL COLEGIO DE ACTUARIOS AUDITORIA-DAÑOS",
                                              "EXAMEN UNICO DE CERTIFICACION DEL COLEGIO DE ACTUARIOS AUDITORIA-FIANZAS",
                                              "EXAMEN UNICO DE CERTIFICACION DEL COLEGIO DE ACTUARIOS FIANZAS",
                                              "EXAMEN UNICO DE CERTIFICACION DEL COLEGIO DE ACTUARIOS AUDITORIA GENERAL",
                                              "EXAMEN UNICO DE CERTIFICACION DEL COLEGIO DE ACTUARIOS DAÑOS",
                                              "EXAMEN UNICO DE CERTIFICACION DEL COLEGIO DE ACTUARIOS RENTAS VITALICIAS",
                                              "EXAMEN UNICO DE CERTIFICACION DEL COLEGIO DE ACTUARIOS AUDITORIA-ACCIDENTES"},
                                             {"E2E","EXANI_I","EXANI_II","EXANI_III","PREEXANI_I","PREEXANI_II","EXANI I - PILOTO METROPOLITANO"},
                                             {""},
                                             {"EXAMEN TRANSVERSAL POR CAMPOS DE CONOCIMIENTO PARA LA LIC. ES-BAS",
                                              "EXAMEN TRANSVERSAL POR CAMPOS DE CONOCIMIENTO PARA LA LIC. ES-BAS-MUES",
                                              "EXAMEN TRANSVERSAL POR CAMPOS DE CONOCIMIENTO PARA LA LIC. ES-BAS-MET",
                                              "EXAMEN TRANSVERSAL POR CAMPOS DE CONOCIMIENTO PARA LA LIC. ES-BAS-INF-MUES",
                                              "EXAMEN TRANSVERSAL POR CAMPOS DE CONOCIMIENTO PARA LA LIC. ES-INF-MUES-EXP",
                                              "EXAMEN TRANSVERSAL POR CAMPOS DE CONOCIMIENTO PARA LA LIC. ES-INF-MUES",
                                              "EXAMEN TRANSVERSAL POR CAMPOS DE CONOCIMIENTO PARA LA LIC. ES-MET",
                                              "EXAMEN TRANSVERSAL POR CAMPOS DE CONOCIMIENTO PARA LA LIC. ES-MUES",
                                              "EXAMEN TRANSVERSAL POR CAMPOS DE CONOCIMIENTO PARA LA LIC. ES-EXP"},
                                             {""},
                                             {""},
                                             {"IFE","IFE-MPIE","IFE-PFDP",
                                              "Prueba de Ingreso al Servicio Profesional Electoral del IFE - Habilidades Intelectuales y Ciencias Sociales"},
                                             {"EGC-LEPRE Y LEPRI"},
                                             {"METROPOLITANO (COMIPEMS)"},
                                             {""},
                                             {"MINESSOTA"},
                                             {"OLIMPIADA DE HABILIDADES ACADEMICAS INFANTILES Y JUVENILES"},
                                             {"DIAGNÓSTICO DE COMPETENCIAS DE PROFESORES PARA LA EDUCACIÓN BÁSICA INTERCULTURAL"},
                                             {"PREESCOLAR-BACHILLERATO"},
                                             {"PREESCOLAR-LICENCIATURA"},
                                             {"SEISP"},
                                             {"POLICIA MINISTERIAL SERVICIO","POLICIA MINISTERIAL INGRESO","MINISTERIO PÚBLICO SERVICIO",
                                              "MINISTERIO PÚBLICO INGRESO","PERITO SERVICIO","PERITO INGRESO","POLICIA PREVENTIVO SERVICIO",
                                              "POLICIA PREVENTIVO INGRESO","CUSTODIO PENITENCIARIO SERVICIO","CUSTODIO PENITENCIARIO INGRESO",
                                              "POLICIA PREVENTIVO DEL D.F.","SECRETARÍO DEL MINISTERIO PÚBLICO SERVICIO","SECRETARÍO DEL MINISTERIO PÚBLICO INGRESO",
                                              "ACTUARIO SERVICIO","ACTUARIO INGRESO","DEFENSOR DE OFICIO SERVICIO","POLICIA FEDERAL PREVENTIVO"},
                                             {"EXAMEN TEORICO DEL TRIBUNAL FEDERAL DE JUSTICIA FISCAL Y ADMINISTRATIVA"},
                                             {""},
                                             {""},
                                             {""},
                                             {""},
                                             {""},
                                             {""}                                             
                                           };                      
       
       private int[][] claves_examen = { 
                                         {156},
                                         {305},
                                         {303},
                                         {304},
                                         {301},
                                         {324,325,326},
                                         {381,244,354,243,},
                                         {263,0,327},
                                         {},
                                         {},
                                         {},
                                         {},
                                         {},
                                         {},
                                         {239,237,238},
                                         {243,244,354},
                                         {1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,153,154,155,157,158,160,170,186,190,213,
                                          214,215,224,227,230,226,232,233,234,225,322,300,194,195,382,383,384,385,392,393,394},
                                         {51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,171,172,173,174,175,176,177,178,179,180,181,182,183,184,185,
                                          199,258,259,260,261,262,295},
                                         {},
                                         {},
                                         {},
                                         {},
                                         {328,329,330,331,332,333,334,335,336},
                                         {231},
                                         {},
                                         {337,338,339,340,341,342,343,344,345,379},
                                         {0,35,36,47,48,49,356},
                                         {},
                                         {255,264,265,266,267,268,386,387,388},
                                         {},
                                         {323,253,254,357},
                                         {188},
                                         {189},
                                         {149},
                                         {150},
                                         {321},
                                         {257},
                                         {200},
                                         {201},
                                         {165},
                                         {138,139,140,141,142,143,144,145,146,147,148,161,162,163,164,166,167},
                                         {358},
                                         {},
                                         {},
                                         {50},
                                         {},
                                         {},
                                         {192}
                                       };              
       
       private Applet() throws IOException, InvalidFormatException{                                   
               
               setLayout(new GridBagLayout());
               
               gbc = new GridBagConstraints();
               
               panelAplicacion = new JPanel(new GridBagLayout());
               panelExcel = new JPanel(new GridBagLayout());
               
               setSize(700,400);
               
               eRutaExcel = new JLabel("Ruta :");                              
               gbc.gridx = 0;
               gbc.gridy = 0;               
               gbc.weightx = 0.1;  
               gbc.anchor = GridBagConstraints.WEST;
               gbc.insets = new Insets(5,5,5,5);               
               panelAplicacion.add(eRutaExcel,gbc);
               
               cRuta = new JTextField(40);
               gbc = new GridBagConstraints();
               gbc.gridx = 1;
               gbc.gridy = 0;
               gbc.insets = new Insets(5,5,5,5);
               gbc.anchor = GridBagConstraints.WEST;
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
               gbc.anchor = GridBagConstraints.WEST;
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
                              
               comboTipoInstr = new JComboBox<>(tiposInstrumento);
               comboTipoInstr.setSelectedIndex(0);
               comboTipoInstr.addActionListener(
                            new ActionListener() {
                                @Override
                                public void actionPerformed(ActionEvent e) {                                                                                                                    	                            	                                          
                                                                                                              
                                       SwingWorker swbe; 
                                    swbe = new SwingWorker() {
                                    @Override
                                    protected Object doInBackground() throws Exception {
                                                                                             
                                              panelAplicacion.remove(comboNombres_cortos);
                        
                                              String itExamen = (String)comboTipoInstr.getSelectedItem();
                                              System.out.println("indice = " + itExamen);
                        
                                              gbc = new GridBagConstraints();
                        
                                              comboNombres_cortos = new JComboBox<>(traeNombresCortos(itExamen));                                      
                                              gbc = new GridBagConstraints();
                                              gbc.gridx = 1;
                                              gbc.gridy = 2;                                            
                                              gbc.gridwidth = 3;
                                              gbc.insets = new Insets(5,5,5,5);
                                              gbc.anchor = GridBagConstraints.WEST;               
                                              panelAplicacion.add(comboNombres_cortos,gbc);                                               
                                              panelAplicacion.revalidate();
                                              panelAplicacion.repaint();   
                                               
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
                                                         
                                                       String nom_corto  = rs.getString(1);
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
                     
                                            }catch(Exception e){ e.printStackTrace(); }
                                              
                                            return nombres_cortosArray;       

                                    } 
                                    
                        };
                                       
                                       swbe.execute();
                                       
                                }
                            
                            } 
                 
               );
                                 
               gbc = new GridBagConstraints();
               gbc.gridx = 2;
               gbc.gridy = 1;
               gbc.anchor = GridBagConstraints.CENTER;
               gbc.weightx = 0.1;
               gbc.insets = new Insets(5,5,5,5);
               panelAplicacion.add(comboTipoInstr,gbc);
                              
               gbc = new GridBagConstraints();
               gbc.gridx = 4;
               gbc.gridy = 1;                                            
               gbc.insets = new Insets(5,5,5,5);
               gbc.anchor = GridBagConstraints.WEST;                                                            
              
               eSTExamen = new JLabel("Subtipo :");
               gbc = new GridBagConstraints();
               gbc.gridx = 0;
               gbc.gridy = 2;
               gbc.anchor = GridBagConstraints.WEST;               
               gbc.insets = new Insets(5,5,5,5);               
               panelAplicacion.add(eSTExamen,gbc);
               
               comboNombres_cortos = new JComboBox<>();
               gbc = new GridBagConstraints();
               gbc.gridx = 1;
               gbc.gridy = 2;               
               gbc.insets = new Insets(5,5,5,5);
               gbc.anchor = GridBagConstraints.WEST;               
               panelAplicacion.add(comboNombres_cortos,gbc);  
                         
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
                                                                                            
                              SwingWorker<Void,Void> chamber;
                              chamber = new SwingWorker<Void,Void>() {
                                                                
                              Workbook wb;
                              
                              Statement s;
                              String cte = (String)comboTipoInstr.getSelectedItem();
                              String cste = (String)comboNombres_cortos.getSelectedItem();                                                                                

                              @Override
                              protected Void doInBackground() {                                                                                                                         
                                                     
                                        String rutaExcel = cExcel.getText().trim();
                                        String rutaDats = cRuta.getText().trim();
             
                                        aplicaciones  = new ArrayList<>();
                                        fechas        = new ArrayList<>();
                                        instituciones = new ArrayList<>();
                                        registro      = new ArrayList<>();
                                        respuesta     = new ArrayList<>();
                                        cve_instr     = new ArrayList<>();
                          
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
                                                                                        
                                            obtenDatos();
                                            
                                            
                                       }catch(Exception e){ e.printStackTrace(); }                                                                                                                                                                      
                                                                                                     
                                       return null; 
                       
                              }
                
                              private void obtenDatos() throws Exception{
                       
                                      Statement statement = conectaBase();                        
                                      String select = "select cve_instr from datos_examenes";
                                      int h = 1;
                                                              
                                      System.out.println("Dentro del try antes del workbook " + aExcel.getAbsolutePath());                                                                                      
                                                                        
                                      try{  wb = WorkbookFactory.create(aExcel); }
                                      catch(Exception e){ e.printStackTrace(); }
                              
                                      Sheet hoja = wb.getSheetAt(0);      
                              
                                      System.out.println("Nombre hoja " + hoja.getSheetName());
                              
                                      Iterator<Row> rowIt = hoja.rowIterator();                                    
                                      rowIt.next();
             
                                      DateFormat df = new SimpleDateFormat("dd-MMM-yy",Locale.ENGLISH);                    
             
                                      System.out.println("cte " + cte + " cste " + cste);             
                        
                                      try{
                                                   
                                          select += " where  = '" + cste + "'" ;
                                          ResultSet rs = statement.executeQuery(select);
                                          
                                          while ( rs.next() ){
                                              
                                                  int datoCve_instr = rs.getInt(1);
                                               
                                                  for( Iterator<Row> it = rowIt; it.hasNext(); ){
  
                                                       Row r = it.next();
 
                                                       Cell cFechaInicio        = r.getCell(1); 
                                                       Cell cTipoAplicacion     = r.getCell(11);
                                                       Cell cSTipoAplcacion     = r.getCell(13);
                                                       Cell cInstitucion        = r.getCell(32);
                                                       Cell noRegistradosCell   = r.getCell(21);
                                                       Cell noRespuestaCell     = r.getCell(22);
                                                       Cell cClave_instr        = r.getCell(12);
                            
                                                       String scTipoAplicacion  = cTipoAplicacion.getStringCellValue().trim();                                                  
                                                       String scSTipoAplicacion = cSTipoAplcacion.getStringCellValue().trim();
                                                       String scInstitucion     = cInstitucion.getStringCellValue().trim();
                                                       String scClave_instr      = cClave_instr.getStringCellValue().trim();
                                                       double noRegistrados     = noRegistradosCell.getNumericCellValue();
                                                       double noRespuesta       = noRespuestaCell.getNumericCellValue();                                                                                  
                         
                                                       String valor = cFechaInicio.getStringCellValue().trim();
                              
                                                       if( valor.length() < 7 ){ continue; }

                                                       Date fechaExcel = df.parse(valor);
                                                       Calendar c = Calendar.getInstance();
                                                       c.setTime(fechaExcel);
                                                       int fem = c.get(Calendar.MONTH);                                                                                                                    
                                                       int cmi = comboMes.getSelectedIndex();
                                                                                                                                                                              
                                                       if( fem == cmi && cte.equals(scTipoAplicacion) ){                                                                                                                                  
                                               
                                                           Cell cApp   = r.getCell(0);
                                                           Object oapp = cApp.getStringCellValue();                                                                                                                                                                                            
                                                                                   
                                                           if( oapp != null ){                                                      
                                                               System.out.println( h + " -  " + oapp + " " + fechaExcel + " " + noRegistrados + " " + 
                                                                                   noRespuesta + " " + scInstitucion + " " + scClave_instr );                                                                                                                                 
                                                               h++;
                                               
                                                               aplicaciones.add(oapp);   
                                                               fechas.add(fechaExcel);
                                                               registro.add(noRegistrados);
                                                               respuesta.add(respuesta);
                                                               instituciones.add(scInstitucion);
                                                               cve_instr.add(scClave_instr);
                                                   
                                                           }
                                                                                                                                                                                                             
                                                       }                                                                                                                                     
                                  
                                                  }
                                                  
                                          }
                                          
                                  }catch(Exception e){ e.printStackTrace(); }
                                      
                              }
                
                              private Statement conectaBase(){
                    
                                      Connection c;
                                      Statement s = null;
                     
                                      try{
                
                                          Class.forName("com.mysql.jdbc.Driver");
                                          c = DriverManager.getConnection("jdbc:mysql://172.16.34.21:3306/replicasiipo","test","slipknot");
                                          s = c.createStatement();
                            
                                      }catch(Exception e ){ e.printStackTrace(); }
                        
                                      return s;
                      
                              }
                                                          
                              };
                                                                                                             
                              chamber.execute();
                              
                       }
                                                                               
                    }
                       
               );
              
               gbc = new GridBagConstraints();
               gbc.gridx = 4;
               gbc.gridy = 2;
               gbc.weightx = 0.1;
               gbc.weighty = 0.1;
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
             
              try{
                  
                  JTabbedPane tabPane = new JTabbedPane();                                                          
                                      
                  JPanel contenedor = new JPanel();
                  contenedor.setLayout(new GridBagLayout());
              
                  GridBagConstraints gbc = new GridBagConstraints();                            
                  gbc.anchor = GridBagConstraints.WEST;
                  contenedor.add(new Applet(),gbc);
              
                  gbc = new GridBagConstraints();              
                  gbc.weighty = 0.1;
                  gbc.gridx = 0;
                  gbc.gridy = 1;
                  contenedor.add(new TablaResultados(),gbc);
              
                  tabPane.addTab("Consolidacion",contenedor);
              
                  JFrame frame = new JFrame("Consolidacion de datos");                                    
                  frame.setSize(1000,450);
              
                  frame.setContentPane(tabPane);
                  //frame.setResizable(false);                                             
                  frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
                  frame.setLocationByPlatform(true);
                  frame.setVisible(true);
                  
              }catch(Exception e){ e.printStackTrace(); }
                                                  
       }
      
}

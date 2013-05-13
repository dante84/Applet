
// @author Daniel.Meza

import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.HeadlessException;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileFilter;
import java.io.FileInputStream;
import java.io.IOException;
import java.security.AccessController;
import java.security.PrivilegedAction;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

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
import javax.swing.JCheckBox;
import javax.swing.JFrame;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
 
public class Applet extends JPanel {
     
       private static JFrame frame;                                           
       private JLabel eRutaExcel,eMes,eTExamen,eSTExamen,eExcel,eProcesando;
       private JPanel panelAplicacion,panelExcel,panelDown;              
       private JTextField cRuta,cExcel;
       private JButton bExaminar,bProcesarAplicaciones,bExcel,bSalvarDatos;
       private JFileChooser fileChooser;
       private ArrayList<Object> cve_instr,appDatMControlNoDat,alAplicacionesDatsErraticos = new ArrayList<>();  
       private Map<Object,Date> fechas;
       private Map<Object,String> registro,respuesta,instituciones,mapaTipoAplicacion,mapaSubtipoAplicacion;
       private Map<Object,Object> aplicaciones,imagEncR,imagEncS,mapaValoresMControlR,mapaValoresMControlS,mapaPosicionesRegistro = new HashMap<>(),
                                  mapaPosicionesRegistroBPM = new HashMap<>(),mapaPosicionesRegistroMcontrol = new HashMap<>(),
                                  mapaPosicionesRespuesta = new HashMap<>(),mapaPosicionesRespuestaBPM = new HashMap<>(),
                                  mapaPosicionesRespuestaMcontrol = new HashMap<>();
       private Map<Object,Object> mapaAplicacionesSinDatif = new HashMap<>(),aplicacionesInexistentes,
                                  mapaAplicacionesPosicionesDesfazadas = new HashMap<>(),
                                  mapaCantidadImagenes = new HashMap<>();
       private ArrayList<Object[]> alResultados = new ArrayList<>();
       private File aExcel;
       private ExtendedService extendedService;
       private JComboBox<String> comboMes,comboTipoInstr,comboNombres_cortos;
       private JTable tresultados;
       private DefaultTableModel dftm;
       private JScrollPane sResultados;
       private GridBagConstraints gbc;      
       private SimpleDateFormat sdf;
       
       private int posicionesExcel = 0,ndr = -1,nds = -1;
       private int numeroPosiciones = 0;    
       
       private ConexionBase conexionBase;
       private final String localhost = "127.0.0.1";
       private final String remoto = "172.16.34.21";
       
       private final String[] meses   = {"Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"};
       private final String[] months  = {"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"};
       
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
       
       private Applet() throws IOException,InvalidFormatException{                                   
               
               setLayout(new GridBagLayout());
               
               conexionBase = new ConexionBase();
               
               gbc = new GridBagConstraints();
               
               eProcesando = new JLabel("En espera");
               
               tresultados = new JTable(new DefaultTableModel());
               dftm = new DefaultTableModel();
               dftm.addColumn("No");
               dftm.addColumn("No aplicacion");
               dftm.addColumn("Existe");
               dftm.addColumn("Datif");
               dftm.addColumn("No Imag Reg");
               dftm.addColumn("No Imag Res");
               dftm.addColumn("No PReg");
               dftm.addColumn("No PReg BPM"); 
               dftm.addColumn("No PReg MControl");
               dftm.addColumn("No PRes");
               dftm.addColumn("No PRes BPM"); 
               dftm.addColumn("No PRes MControl");
               dftm.addColumn("Estado");
               dftm.addColumn("Observaciones");
               
               tresultados.setModel(dftm);                                      
               sResultados = new JScrollPane(tresultados);     
               
               gbc = new GridBagConstraints();         
               gbc.gridx = 0;
               gbc.gridy = 2;          
               gbc.weightx = 0.1;
               gbc.fill = GridBagConstraints.HORIZONTAL;
               gbc.insets = new Insets(5,5,5,5);                                                                                                                                    
               add(sResultados,gbc);
                                                     
               panelAplicacion = new JPanel(new GridBagLayout());
               panelExcel = new JPanel(new GridBagLayout());
               panelDown = new JPanel(new GridBagLayout());
               setSize(700,400);
               
               eRutaExcel = new JLabel("Ruta :");                              
               gbc.gridx = 0;
               gbc.gridy = 0;               
               gbc.weightx = 0.1;  
               gbc.anchor = GridBagConstraints.WEST;
               gbc.insets = new Insets(5,5,5,5);               
               panelAplicacion.add(eRutaExcel,gbc);
               
               cRuta = new JTextField(40);
               cRuta.setEditable(false);
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
                                                   //c = DriverManager.getConnection("jdbc:mysql://172.16.34.21:3306/replicasiipo","test","slipknot");
                                                   //c = conexionBase.getC(remoto,"replicasiipo","test","slipknot");   
                                                   c = conexionBase.getC(localhost,"replicasiipo","test","slipknot");
                                                   
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
               gbc.insets = new Insets(5,5,5,5);
               gbc.anchor = GridBagConstraints.WEST;               
               panelExcel.add(eExcel,gbc);
               
               cExcel = new JTextField(40);  
               cExcel.setEditable(false);
               gbc = new GridBagConstraints();
               gbc.gridx = 1;
               gbc.gridy = 0;
               gbc.gridwidth = 2;
               gbc.anchor = GridBagConstraints.WEST;               
               gbc.insets = new Insets(5,15,5,8);                                           
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
                                                     
                              Applet.this.bProcesarAplicaciones.setEnabled(false);
                              
                              SwingWorker<Void,Void> chamber;
                              chamber = new SwingWorker<Void,Void>() {
                                                                
                              Workbook wb;
                              
                              Statement s;
                              String cte = (String)comboTipoInstr.getSelectedItem();
                              String cste = (String)comboNombres_cortos.getSelectedItem();   
                              
                              int h = 1;
                              
                              @Override
                              protected Void doInBackground() {                                                                                                                         
                                                     
                                        Applet.this.eProcesando.setText("Procesando");
                                        
                                        String rutaExcel = cExcel.getText().trim();
                                        String rutaDats = cRuta.getText().trim();
             
                                        aplicaciones          = new HashMap<>();
                                        fechas                = new HashMap<>();
                                        instituciones         = new HashMap<>();
                                        registro              = new HashMap<>();
                                        respuesta             = new HashMap<>();
                                        cve_instr             = new ArrayList<>();
                                        mapaTipoAplicacion    = new HashMap<>();
                                        mapaSubtipoAplicacion = new HashMap<>();
                          
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
                                            cuentaImagenes();  
                                            cuentaPosiciones();
                                            
                                       }catch(Exception e){ e.printStackTrace(); }                                                                                                                                                                      
                                                                                                     
                                       return null; 
                       
                              }
                
                              private void obtenDatos() {
                       
                                      Connection con;
                                      con = conexionBase.getC(localhost,"replicasiipo","test","slipknot");
                                      //con = conexionBase.getC(remoto,"replicasiipo","test","slipknot");
                                                                                                                                                                                                                  
                                      try{  wb = WorkbookFactory.create(aExcel); }
                                      catch(Exception e){ e.printStackTrace(); }
                              
                                      Sheet hoja = wb.getSheetAt(0);      
                              
                                      
                                      Iterator<Row> rowIt = hoja.rowIterator();                                    
                                      rowIt.next();
             
                                      sdf = new SimpleDateFormat("yyyy-MM-dd",Locale.ENGLISH);                    
                                                 
                                      try{
                                                   
                                          Statement statement = con.createStatement();//conectaBase();                        
                                          String select = "select cve_instr from datos_examenes";                                      
                                          
                                          select += " where nom_corto = '" + cste + "'" ;
                                           
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
                                                       int scClave_instr        = Integer.parseInt( cClave_instr.getStringCellValue());
                                                       double noRegistrados     = noRegistradosCell.getNumericCellValue();
                                                       double noRespuesta       = noRespuestaCell.getNumericCellValue();                                                                                  
                         
                                                       String valor = cFechaInicio.getStringCellValue().trim();                                                       
                              
                                                       if( valor.length() < 7 ){ continue; }
                                                       
                                                       String mes = valor.substring(3,6);
                                                       String month = "";
                                                       
                                                       for(int i = 0; i <= months.length - 1; i++){
                                                           if( months[i].equals(mes) ){
                                                               month = String.valueOf(i + 1);
                                                               if(month.length() - 1 == 0){
                                                                  month = "0" + month;
                                                               }
                                                           }
                                                       }
                                                       
                                                       String fecha = "20" + valor.substring(7,9) + "-" +  month  + "-" + valor.subSequence(0,2);                                                     
                                                       
                                                       Date fechaExcel = sdf.parse(fecha);
                                                       Calendar c = Calendar.getInstance();
                                                       c.setTime(fechaExcel);
                                                       int fem = c.get(Calendar.MONTH);                                                                                                                    
                                                       int cmi = comboMes.getSelectedIndex();
                                                                                                                                                                              
                                                       if( fem == cmi && scClave_instr == datoCve_instr ){                                                                                                                                  
                                               
                                                           Cell cApp   = r.getCell(0);
                                                           Object oapp = cApp.getStringCellValue();                                                                                                                                                                                            
                                                                                   
                                                           if( oapp != null ){                                 
                                                               
                                                               h++;
                                               
                                                               System.out.println(h + " " + oapp + " " + scClave_instr + " " + valor + " " + fecha + " " + 
                                                                                  fechaExcel);
                                                               aplicaciones.put(oapp,scClave_instr);   
                                                               fechas.put(oapp,fechaExcel);
                                                               registro.put(oapp,String.valueOf(noRegistrados));
                                                               respuesta.put(oapp,String.valueOf(noRespuesta));
                                                               instituciones.put(oapp,scInstitucion);
                                                               cve_instr.add(scClave_instr);                                                               
                                                               mapaTipoAplicacion.put(oapp,scTipoAplicacion);
                                                               mapaSubtipoAplicacion.put(oapp,scSTipoAplicacion);
                                                   
                                                           }
                                                                                                                                                                                                             
                                                       }                                                                                                                                     
                                  
                                                  }
                                                  
                                          }                                                                                    
                                                                                                                              
                                  }catch(SQLException | NumberFormatException | ParseException e){ e.printStackTrace(); }                                                                                                                                                   
                                      
                              }
                              
                              private void cuentaImagenes(){
                                     
                                      String ruta = cRuta.getText().trim();                                      
                                      aplicacionesInexistentes = new HashMap<>();                                      
                                      imagEncR = new HashMap<>();
                                      imagEncS = new HashMap<>();
                                                                            
                                      int imagenesExistenR = 0;
                                      int imagenesExistenS = 0;
                                      int k = 1;
                                      
                                      Set<Object> ks = aplicaciones.keySet();
                                      Iterator<Object> it = ks.iterator();
                                      
                                      try{
                                                                                                                    
                                          while( it.hasNext() ){
                                          
                                                 Applet.this.eProcesando.setText("Procesando imagenes - aplicacion " + k + " de " + h);
                                                 Object o = it.next();
                                                 String numeroAplicacion = (String)o;                                                  
                                                 
                                                 File appDir = new File(ruta + "\\" + numeroAplicacion);
                                                 boolean existe = appDir.exists();                                                                                                  
                                                 
                                                 if( !existe ){ aplicacionesInexistentes.put(o,o); }                                                 
                                                 else{
                                                      boolean esDir = appDir.isDirectory();                       	  
                                                      if( esDir ){                                                                                                                                  
                                                          File[] archivos = appDir.listFiles();                                                                                                                          
                                                          for( File f : archivos ){                                                                                                      
                                                               String nombreArchivo = f.getName();  
                                                           
                                                               if( nombreArchivo.matches("\\d{6}\\_[Rr]\\d{3}\\.[t][i][f]") ){                                                                    
                                                                   imagenesExistenR++; 
                                                               }                                                                                                                                                                                                                                                                                                                    
                                                               if( nombreArchivo.matches("\\d{6}\\_[Ss]\\d{3}\\.[t][i][f]") ){ 
                                                                   imagenesExistenS++; 
                                                               }                                                              
                                                              
                                                          }
                                                                                                                                                                             
                                                      }                                           
                                                                                      
                                                 }
                                                 
                                                 System.out.println(o  + " " + imagenesExistenR + " " + imagenesExistenS);
                                                 imagEncR.put(o,imagenesExistenR);
                                                 imagEncS.put(o,imagenesExistenS);                                                                                                                   
                                                 
                                                 imagenesExistenR = 0;
                                                 imagenesExistenS = 0;
                                     
                                                 k++;
                                                 
                                          }
                                          
                                      }catch(Exception e){ e.printStackTrace(); }    
                              
                              }
                              
                              private void cuentaPosiciones(){
                                                                       
                                      String rutaDatif = cRuta.getText().trim();  
                                      appDatMControlNoDat = new ArrayList<>();
                                      mapaValoresMControlR = new HashMap<>();
                                      mapaValoresMControlS = new HashMap<>();                                                                            
                                      
                                      Set<Object> ks = aplicaciones.keySet();
                                      Iterator<Object> it = ks.iterator();
                                      
                                      int k = 1;
                                      while( it.hasNext() ){
                                          
                                             Applet.this.eProcesando.setText("Procesando dats - aplicacion " + k + " de " + h);
                                             Object o = it.next();                                                                                     
                                             
                                             if( aplicacionesInexistentes.containsKey(o) ){ continue; }                                             
                                             
                                             String aplicacion = (String)o;
                                             rutaDatif += "\\" + aplicacion + "\\DATIF";
                                             File datif = new File(rutaDatif);                                                                                          
                                           
                                             boolean existeDatif = datif.exists();                                                                                     
                                             
                                             if( existeDatif ){
                                                 String Datif = datif.getAbsolutePath();                     
                                                 File[] archivos = datif.listFiles(                            		   
                         	 	                new FileFilter() {
                     			                    @Override
                                                            public boolean accept(File pathname) {                                     
                                                                   if( pathname.getName().endsWith(".dat") ){ return true; }                                          
                                                                   return false;                                        
                                                            }
                   
                                                        }
                                                      
                                                 );                                                                                            
                                                                                            
                                                 int r = -1;
                                                 int S = -1;
                                                 int la = (archivos.length - 1);                          
                                                                                                                    
                                                 if( la == -1 && 
                                                     ( ( registro.containsKey(o)  && Double.valueOf( registro.get(o)  ) > 0  ) || 
                                                     ( ( respuesta.containsKey(o) && Double.valueOf( respuesta.get(o) ) > 0) ) ) ) {      
                                                     System.out.println("No hay dats " + o);
                                                     appDatMControlNoDat.add(o);
                                                     continue;
                                                 }
                                              
                                                 for( int m = 0; m <= la; m++ ){
                              
                                                      String nombreArchivo = archivos[m].getName();                                                                                                            
                                                      String subNombreArchivo = "";
                               
                                                      for( int i = 0; i <= nombreArchivo.length() - 5; i++ ){
                                                           subNombreArchivo += nombreArchivo.charAt(i);
                                                      }                                                                                                        
                                                              
                                                      char ci = nombreArchivo.charAt(0);
                                                                                  
                                                      if( subNombreArchivo.matches("[Rr]\\d{9}[Xx][_\\d]") || subNombreArchivo.matches("[Ss]\\d{9}[Xx][_\\d]") ){ 
                                   
                                                          String c = "";
                                                          c += ci;
                                      
                                                          if( c.matches("[RrSs]") ){                                                       
                                                              if( "r".equals(c) || "R".equals(c) ){                                                                 
                                                                  r++;                                 
                                                              }   
                                                              if( "s".equals(c) || "S".equals(c) ){                                                               
                                                                  S++;
                                                              }
                                                          }                                                                                          
                                   
                                                      }else{                                     
                                                            alAplicacionesDatsErraticos.add(o);
                                                            continue;
                                                      }   
                                                                                                                                     
                                                 }
                                                                                                 
                                                 if( respuesta.containsKey(o) && S == -1 && Double.valueOf( respuesta.get(o)) > 0 ){
                                                     System.out.println("No hay dats de respuestas en " + o);
                                                     appDatMControlNoDat.add(o);
                                                 }
                                                                                                  
                                                 if( registro.containsKey(o) && r == -1 && Double.valueOf(registro.get(o)) > 0){
                                                     System.out.println("No hay dats de registros en " + o);
                                                     appDatMControlNoDat.add(o);
                                                 }
                                                 
                                                 if( la == -1 ){ mapaAplicacionesSinDatif.put(o,o); }                     
                                                 else{     
                                                      int i = 0;
                                                      while( i <=  r ){                               
                                                             String nombreArchivo = archivos[i].getName();                                                                                                       
                                                             mapaValoresMControlR.put(o,leeArchivo(nombreArchivo,rutaDatif,(String)o,i,r,"R"));
                                                             i++;
                                                      }                                                            
                              
                                                      while( i <= la ){                               
                                                             String nombreArchivo = archivos[i].getName();                                                                                                                                                                                    
                                                             mapaValoresMControlS.put(o,leeArchivo(nombreArchivo,rutaDatif,(String)o,i,la,"S"));
                                                             i++;
                                                      }
                              
                                                 }
                                              
                                           }else{ mapaAplicacionesSinDatif.put(o, o); }                                                                                     
                                                                                      
                                           rutaDatif = cRuta.getText().trim();
                                           k++;
                                                                                              
                                      }
                                   
                              }
                              
                              @SuppressWarnings("resource")
                              public int leeArchivo(String nombreArchivo,String rutaDatif,Object f,int i,int noArchivos,String tipo){               
        
                                     String linea = "";                                                 
                                     int temp;                                                                                                                                                                                                         
              
                                     try{         
                                                      
                                         File f1 = new File(rutaDatif + "\\" + nombreArchivo);                  
                                         FileInputStream fis = new FileInputStream(f1);                       
                     
                                         while(true){
                   
                                               temp = fis.read();                                                                   
                                               
                                               if( temp == -1 ){                              
                                                   ndr++;
                                                   break;
                                              }
                     
                                              int digitoSub;
                                              linea += (char)temp;                                                                                                          
                     
                                              if( temp == '\n' ){ 
                                                        
                                                  String sub = linea.substring(3,9);                            
                                                  digitoSub  = Integer.parseInt(sub);                               
                         
                                                  if( digitoSub == 0 ){ numeroPosiciones--; }
                         
                                                  numeroPosiciones++;                                                                                
                                                         
                                                  if( digitoSub != numeroPosiciones ){ 
                            	                      mapaAplicacionesPosicionesDesfazadas.put(f,f);
                                                  }
                                                                                         
                                                  linea = "";
                         
                                              }                                                        
                     
                                         }                                                                                                         
                                         
                                         posicionesExcel = mcExcelPosiciones((String)f, "2012",tipo);                                                                                          
                                      
                                         if( i == noArchivos ){ 
                   
                                             int posiciones = revisaBpmPosiciones((String)f, "2012",tipo);
                                                                                                                                                                           
                                             if( tipo.equals("R") ){                                   	 
                                                 mapaPosicionesRegistro.put(f,numeroPosiciones);
                                                 mapaPosicionesRegistroBPM.put(f,posiciones);
                                                 mapaPosicionesRegistroMcontrol.put(f,posicionesExcel);
                                             }else{
                                                   mapaPosicionesRespuesta.put(f, numeroPosiciones);
                                                   mapaPosicionesRespuestaBPM.put(f, posiciones);
                                                   mapaPosicionesRespuestaMcontrol.put(f, posicionesExcel);
                                             }  
                         
                                             numeroPosiciones = 0;
                                                    
                                         }                                   
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
                                     }catch(IOException | NumberFormatException e){ e.printStackTrace(); }                              
                                                  
                                     return posicionesExcel;
              
                              }
                              
                              public int mcExcelPosiciones(String app,String año,String tipo){ 
   
                                     int posiciones = 0;                                                                               
        
                                     try{
            	                                          
                                         Sheet hoja = wb.getSheetAt(0);                  
                                         Iterator<Row> rowIt = hoja.rowIterator();                                                 
                                         
                                         rowIt.next();
                                         for(Iterator<Row> it = rowIt; it.hasNext(); ){
                                             Row r = it.next();
                                             Cell cAplicacion = r.getCell(0);
                                             String cvc = cAplicacion.getStringCellValue().trim();                                             
                                             if( cvc.matches("^[0-9]+$") ){
                                                 if( Integer.parseInt(cvc) == Integer.parseInt(app) ){                                                     
                                                     Cell cPosiciones;
                                                     if( tipo.equals("R") ){ cPosiciones = r.getCell(21); }
                                                     else{ cPosiciones = r.getCell(22);}
                                                     posiciones += cPosiciones.getNumericCellValue();                     
                                                 }
                                             }
                                             
                                         }
         
                                     }catch(Exception e){ e.printStackTrace(); }                                          
                                     
                                     return posiciones;
     
                              }    
                              
                              public int revisaBpmPosiciones(String aplicacion,String año,String tipo){
     
                                     Connection c;
                                     Statement s;
                                     ResultSet rs;                                                        
     
                                     int posiciones = 0;
      
                                     try{
     
                                         Class.forName("oracle.jdbc.OracleDriver");                   
                                         c = DriverManager.getConnection("jdbc:oracle:thin:@10.10.2.10:1521:ceneval","dpoc","bpm_DPOC");
                        
                                         s = c.createStatement();
                  
                                         String select = "";
                  
                                         if( tipo.equals("R") ){ 
                                             select += "select \"Registrado_desglose\",\"Registrado\" from dpoc where NUM_APLIC = '" + 
                                             aplicacion + "' and extract(year from \"fecha_de_inicio\") ='" + año + "'";
                                         }else{
                	                        select += "select \"Aplicados_desglose\",\"Aplicados\" from dpoc where NUM_APLIC = '" + aplicacion + "' and " + 
                                                          " to_char(\"fecha_de_inicio\",'YYYY') = '" + año + "'";
                                         }
                  
                                         rs = s.executeQuery(select);
                                     
                                         int i = 0;                      
             
                                         while( rs.next() ){
                                                i++;
                                                if( i > 1 ){                             
                                                    posiciones =  rs.getInt(1);                             
                                                    break;
                                                }else{ posiciones = rs.getInt(2); }                             
                    
                                         }         
         
                                     }catch(ClassNotFoundException | SQLException e){ e.printStackTrace(); }
     
                                     return posiciones;
   
                              }

                              public int revisaBpmAplicados(String aplicacion,String año){
  
                                     Connection c;
                                     Statement s; 
                                     ResultSet rs;                            
     
                                     int aplicados = 0;
     
                                     try{
     
                                         Class.forName("oracle.jdbc.OracleDriver");                   
                                         c = DriverManager.getConnection("jdbc:oracle:thin:@10.10.2.10:1521:ceneval","dpoc","bpm_DPOC");
                                         s = c.createStatement();                                    
                 
                                         rs = s.executeQuery("select \"Aplicados_desglose\",\"Aplicados\" from dpoc where NUM_APLIC = '" + aplicacion + "' and " + 
                                                             " to_char(\"fecha_de_inicio\",'YYYY') = '" + año + "'");              
     
                                         int i = 0;                      
                                         while( rs.next() ){                  
                                                i++;
                                                if( i > 1 ){                             
                                                    aplicados = rs.getInt(1);
                                                    break;
                                                }else{ aplicados = rs.getInt(2); }    
                                         }
         
                                     }catch(ClassNotFoundException | SQLException e){ e.printStackTrace(); }
     
                                     return aplicados;
   
                              }             
                                              
                              
                              @Override
                              public void done(){                                                                                                                                                   
                                     
                                     Set<Object> ks = aplicaciones.keySet();
                                     Set<Object> setAne = aplicacionesInexistentes.keySet();    
                                     Set<Object> setAsd = mapaAplicacionesSinDatif.keySet();                  
                                     Iterator<Object> it = ks.iterator();
                                     Object[] aResultados;
                                     
                                     boolean s1 = false;
                                     boolean s2 = false;   
                                     
                                     boolean imagResDat = false;
                                     boolean imagRegDat = false;
                                                    
                                     int i = 0;
                                     while( it.hasNext() ){
            	                              
                                            i++;
       	                                    ArrayList<Object> ao =  new ArrayList<>();   
                                            ArrayList<Object> aoq = new ArrayList<>();
                                            Object o = it.next();
                                            ao.add(i);
       	                                    ao.add(o);   
                                            aoq.add(o);
                                            aoq.add(mapaTipoAplicacion.get(o));
                                            aoq.add(mapaSubtipoAplicacion.get(o));
                                            aoq.add(sdf.format(fechas.get(o)));
                                            aoq.add(sdf.format(new Date()));
                                           	               	         
        	                            int tamañoNoEncontradas = aplicacionesInexistentes.size() - 1;
       	                                    int tamañoSinDatif = mapaAplicacionesSinDatif.size() - 1;
       	                                    int tamañoSinDats = appDatMControlNoDat.size() - 1;
       	                                    int tamañoCantidadImagenesR = imagEncR.size() - 1;
                                            int tamañoCantidadImagenesS = imagEncS.size() - 1;
       	                                    int tamañoPosicionesRegistro = mapaPosicionesRegistro.size() -1;
       	                                    int tamañoPosicionesRegistroBPM = mapaPosicionesRegistroBPM.size() -1;
       	                                    int tamañoPosicionesRegistroMControl = mapaPosicionesRegistroMcontrol.size() -1;
       	                                    int tamañoPosicionesRespuesta = mapaPosicionesRespuesta.size() - 1;
       	                                    int tamañoPosicionesRespuestaBPM = mapaPosicionesRespuesta.size() - 1;
       	                                    int tamañoPosicionesRespuestaMControl = mapaPosicionesRespuesta.size() - 1;
                                            int tamañoAplicacionesDatsErraticos = alAplicacionesDatsErraticos.size() - 1;       	                	                	                                                            
                                                                                        
        	                            ao.add(agregaDato(aplicacionesInexistentes, tamañoNoEncontradas,o,false,"Inexistentes"));        	 		
                                            if( !aplicacionesInexistentes.containsKey(o) ){
                                                Object osd = agregaDato(mapaAplicacionesSinDatif, tamañoSinDatif,o,false,"SinDatif");
           	                                ao.add(osd);           	  	       	                                                                                              
                                                Object ir = agregaDato(imagEncR, tamañoCantidadImagenesR,o,true,"ImagenesR");
           	                                ao.add(ir);           	  
                                                aoq.add(ir);
                                                Object is = agregaDato(imagEncS, tamañoCantidadImagenesS,o,true,"ImagenesS");
                                                ao.add(is);           	  
                                                aoq.add(is);
                                                Object pr = agregaDato(mapaPosicionesRegistro, tamañoPosicionesRegistro,o,true,"PR");
              	                                ao.add(pr);                              	 
                                                aoq.add(pr);
                                                Object prb = agregaDato(mapaPosicionesRegistroBPM, tamañoPosicionesRegistroBPM,o,true,"PRB");                                                
              	                                ao.add(prb);         	  	                	 
                                                aoq.add(prb);
                                                Object prm = agregaDato(mapaPosicionesRegistroMcontrol, tamañoPosicionesRegistroMControl,o,true,"PRM");
                                                ao.add(prm);                        	 
                                                aoq.add(prm);
                                                Object pres = agregaDato(mapaPosicionesRespuesta, tamañoPosicionesRespuesta, o, true,"PS");
                                                ao.add(pres);             	  		 
                                                aoq.add(pres);
                                                Object presb = agregaDato(mapaPosicionesRespuestaBPM, tamañoPosicionesRespuestaBPM, o, true,"PSB");
                                                ao.add(presb);                                    	       	                	    	                    	       	                	         
                                                aoq.add(presb);
                                                Object presm = agregaDato(mapaPosicionesRespuestaMcontrol, tamañoPosicionesRespuestaMControl, o, true,"PSM");
                                                ao.add(presm);                                                                  
                                                aoq.add(presm);                                                
                                                aoq.add(cRuta.getText().trim());
                                                aoq.add(instituciones.get(o));
       	        	                                                                                 
                                                if( mapaPosicionesRegistro.containsKey(o) && mapaPosicionesRegistroMcontrol.containsKey(o) ){                                             
                                                    
                                                    if( (int)mapaPosicionesRegistro.get(o) != (int)mapaPosicionesRegistroMcontrol.get(o) ){ 
                                                        s1 = true; 
              	                                    }                            
                                                    
                                                    if( (int)imagEncR.get(o) > 0 ){
                                                            
                                                    }
                                                    
                                                }
                   	       	                    	     
                                                if( mapaPosicionesRespuesta.containsKey(o) && mapaPosicionesRespuestaMcontrol.containsKey(o) ){
                                                    if( (int)mapaPosicionesRespuesta.get(o) != (int)mapaPosicionesRespuestaMcontrol.get(o) ){
                                                        s2 = true; 
                                                    }
                                                }                                                                                                
                                              
                                                System.out.println(o + " " + setAne.contains(o) + " " + setAsd.contains(o) + " " +  s1 + " " + s2 + " " +
                                                                   alAplicacionesDatsErraticos.contains(o) + " " + appDatMControlNoDat.contains(o));
                                                if( setAne.contains(o) || setAsd.contains(o) || s1 || s2 || alAplicacionesDatsErraticos.contains(o) || 
                                                    appDatMControlNoDat.contains(o) ){                                                                           
                                                    ao.add("Verificar");                        
                                                    aoq.add("Verificar");
                                                }else{                                                                              
                                                      ao.add("Correcto");                          
                                                      aoq.add("Correcto");                          
                                                }                     
        	  	 
        	                                aResultados = ao.toArray();
        	                                alResultados.add(aoq.toArray());
        	                                dftm.addRow(aResultados);        	                                                                                                                                                                                                                                                                                         
                                                
                                            }else{
                                                   aoq.add(cRuta.getText().trim());
                                                   aoq.add(instituciones.get(o));
       	        	                    
                                                  if( mapaPosicionesRegistro.containsKey(o) && mapaPosicionesRegistroMcontrol.containsKey(o) ){                                             
                                                      if( (int)mapaPosicionesRegistro.get(o) != (int)mapaPosicionesRegistroMcontrol.get(o) ){ 
                                                          s1 = true; 
             	                                      }                       
                                                  }
                   	       	                    	     
                                                  if( mapaPosicionesRespuesta.containsKey(o) && mapaPosicionesRespuestaMcontrol.containsKey(o) ){
                                                      if( (int)mapaPosicionesRespuesta.get(o) != (int)mapaPosicionesRespuestaMcontrol.get(o) ){
                                                           s2 = true; 
                                                      }
                                                  }
                                              
                                                  if( setAne.contains(o) || setAsd.contains(o) || s1 || s2 || alAplicacionesDatsErraticos.contains(o) || 
                                                      appDatMControlNoDat.contains(o) ){                                                                             
                                                      ao.add("Verificar");                        
                                                      aoq.add("Verificar");                        
                                                  }else{                                                                                
                                                        ao.add("Correcto");                          
                                                        aoq.add("Correcto");                          
                                                  }                
                                                
                                                  aResultados = ao.toArray();
                                                  alResultados.add(aoq.toArray());
        	                                  dftm.addRow(aResultados);        	                                                                                                                                                                                                
                                                
                                            }
                                            
                                            s1 = false;
                                            s2 = false;   
        	     
                                      }    
                                                     
                                      tresultados.setModel(dftm);                                      
                                      
                                      tresultados.getColumn("Observaciones").setCellRenderer(new BotonRender());
                                      tresultados.getColumn("Observaciones").setCellEditor(new ButtonEditor(new JCheckBox(),frame));
 
                                      sResultados = new JScrollPane(tresultados);     
         
                                      gbc = new GridBagConstraints();         
                                      gbc.gridx = 0;
                                      gbc.gridy = 2;          
                                      gbc.weightx = 0.1;
                                      gbc.fill = GridBagConstraints.HORIZONTAL;
                                      gbc.insets = new Insets(5,5,5,5);                                      
                                                                                              
                                      Applet.this.add(sResultados,gbc);
                                                                            
                                      Applet.this.eProcesando.setText("Terminado");
                                      Applet.this.bProcesarAplicaciones.setEnabled(true);
                                      Applet.this.bSalvarDatos.setEnabled(true);
                                      
                                      Applet.this.revalidate();
                                      Applet.this.repaint();                                                           
                                                                        
                                 }                                                               
                                                               
                                 public Object agregaDato(Map<Object,Object> alObjetos,int tamaño,Object o,boolean real,String vieneDe){	        	            	    
    	                                       
    	                                Set<Object> set     = alObjetos.keySet();    	         	        
    	                                Iterator<Object> si = set.iterator();                                        
                                        
    	                                while( si.hasNext() ){
                                            
                                               Object objeto = alObjetos.get(o);
                                               boolean existe = si.next().equals(o);                                               
                                               
    	       	                               if( existe ){    	       	    	                                                   
    	       	     	                           if( real ){           
                                                       if( objeto != null ){ return objeto; }
                                                       else{ return 0; } 
                                                   }else{ return "No"; }    	       	    	
    	       	                               }else{    	       	    	
     	       	                                     if( real ){ 
                                                         if( objeto != null ){ return objeto; }
                                                         else{ return 0; }
                                                     }
     	       	                                     else{ return "Si"; }     	       	    	
    	       	                               }
                                               
    	                                }
    	         	    
    	                                return "Si";
	      
                                 }
                                 
                                 public JScrollPane agregarBotObs( int canFilas ){
                                       
                                        JScrollPane panel = new JScrollPane();                                        
                                                                                
                                        for( int i = 0; i <= canFilas; i++ ){
                                             JButton bo = new JButton("Agregar observacion");
                                             panel.add(bo);
                                        }
                                        
                                        return panel;
                                       
                                 }
                                                          
                              };
                                                                                                             
                              chamber.execute();
                              
                       }
                                                                               
                    }
                       
               );
               
               bSalvarDatos = new JButton("Guardar");
               bSalvarDatos.setEnabled(false);
               bSalvarDatos.addActionListener(new ActionListener() {

                   @Override
                   public void actionPerformed(ActionEvent e) {
                         
                          SwingWorker worker;
                          worker = new SwingWorker() {

                                  @Override
                                  protected Void doInBackground(){                                                                                  
                                            
                                            Connection c = conexionBase.getC(localhost,"ceneval","user","slipknot");
                                            //Connection c = conexionBase.getC(remoto,"ceneval","user","slipknot");
                                            
                                            Statement  s = null;
                                                                                                                                 
                                            try{                                                                                                 
                                                                                                 
                                                String select =  "select no_aplicacion from viimagenes where no_aplicacion = '";
                                                for( Object[] datosArreglo : alResultados ){                                                                                                     
                                                     Object no_aplicacion = datosArreglo[0];
                                                     select += no_aplicacion + "'";
                                                     System.out.println(select);
                                                     s = c.createStatement();    
                                                     ResultSet  rs = s.executeQuery(select);                                                     
                                                     
                                                     if( !rs.isBeforeFirst() ){
                                                         
                                                         String insert = "insert into viimagenes(no_aplicacion,instrumento,nombre,fecha_alta," +
                                                                         "fecha_registro,imag_reg,imag_res,pregistro,pregistrobpm,pregistromcontrol," + 
                                                                         "prespuesta,prespuestabpm,prespuestamcontrol,ruta,institucion,estado) values('";
                                                                                            
                                                         int i = 0;
                                                         int longitud = datosArreglo.length - 1;
                                                         for( Object dato: datosArreglo ){
                                                          
                                                              if( i == longitud ){ insert += dato + "')"; }
                                                              else{ insert += dato + "','"; }
                                                              i++;                                                                                                                                                                                                                                              
                                                               
                                                         }                                                                                                                  
                                                           
                                                         System.out.println(insert);
                                                         s.executeUpdate(insert);                                                             
                                                         
                                                     }else{
                                                           
                                                           s  = c.createStatement();                                                               
                                                           
                                                           while( rs.next() ){
                                                                  System.out.println(no_aplicacion + " existe");
                                                                  String update = "update viimagenes set ";
                                                                  String[] campos = {"instrumento","nombre","fecha_alta","fecha_registro","imag_reg","imag_res",
                                                                                     "pregistro","pregistrobpm","pregistromcontrol","prespuesta","prespuestabpm",
                                                                                     "prespuestamcontrol","ruta","institucion","estado"};
                                                            
                                                                  int longitud = datosArreglo.length - 1;
                                                                  System.out.println("longitud " + longitud);
                                                                  for( int i = 1; i <= datosArreglo.length - 1; i++ ){
                                                                       if( i == longitud ){                                                                      
                                                                           update += campos[i - 1] + " = '" + datosArreglo[i] + "' where no_aplicacion = '" +
                                                                                     no_aplicacion + "'";                                                                                                                                                                                                                                                                                                                                                                                        
                                                                       }else{
                                                                             update += campos[i - 1] + " = '" + datosArreglo[i] + "',";
                                                                       }
                                                                  }
                                                             
                                                                  System.out.println(update);
                                                                  s.executeUpdate(update);                                                                                                                                                                                                                                             
                                                               
                                                           }                                                                                                                       
                                                                                                                                                                                 
                                                     }                                                                                                          
                                                                                                          
                                                     select = "select no_aplicacion from viimagenes where no_aplicacion = '";
                                                     rs.close();
                                                
                                                }                                                                                             
                                                                                                
                                                JOptionPane.showMessageDialog(
                                                            null,
                                                            "Datos salvados correctamente",
                                                            "Exito",
                                                            JOptionPane.INFORMATION_MESSAGE);
                                                
                                                
                                            }catch(SQLException e){ 
                                                
                                                   int codigoError = e.getErrorCode();
                                                           
                                                   e.printStackTrace();
                                                   
                                                   if( codigoError > 0 ){
                                                       JOptionPane.showMessageDialog(
                                                                   null,
                                                                   "Codigo de error " + codigoError + 
                                                                   ".Reportalo al administrador del sistema",
                                                                   "Error",
                                                                   JOptionPane.ERROR_MESSAGE);                                                                                                                                                                             
                                                   }
                                                   
                                            }
                                            
                                            finally{ 
                                                    try{                                                        
                                                        s.close();
                                                     c.close();
                                                    }catch(SQLException e){ e.printStackTrace(); }
                                            }
                                      
                                            return null;
                                      }
                                  
                                      @Override
                                      public void done(){
                                             Applet.this.eProcesando.setText("En espera");
                                      }
                            
                                };
                                  
                                worker.execute();
                                  
                          };
                          
                          
                          
               });                                     
                             
               gbc = new GridBagConstraints();
               gbc.anchor = GridBagConstraints.SOUTHEAST;
               gbc.weightx = 0.1;
               gbc.weighty = 0.1;
               gbc.insets = new Insets(5,5,5,5);
               panelDown.add(eProcesando,gbc); 
               
               gbc = new GridBagConstraints();
               gbc.anchor = GridBagConstraints.SOUTHEAST;
               gbc.weightx = 0.1;
               gbc.weighty = 0.1;
               gbc.insets = new Insets(5,5,5,5);
               panelDown.add(bSalvarDatos,gbc);                              
                             
               gbc = new GridBagConstraints();
               gbc.gridx = 4;
               gbc.gridy = 2;
               gbc.weightx = 0.1;
               gbc.weighty = 0.1;
               gbc.insets = new Insets(5,5,5,5);
               panelAplicacion.add(bProcesarAplicaciones,gbc);
               
               JPanel panelUp = new JPanel();
               panelUp.add(panelExcel);
               panelUp.add(panelAplicacion);
               
               gbc = new GridBagConstraints();               
               gbc.anchor = GridBagConstraints.WEST;
               add(panelUp,gbc);                                                                           
               
               gbc = new GridBagConstraints();                    
               gbc.gridy = 4;
               gbc.anchor = GridBagConstraints.SOUTH;
               gbc.fill = GridBagConstraints.HORIZONTAL;               
               gbc.insets = new Insets(5,5,5,5);
               gbc.weighty = 0.5;
               add(panelDown,gbc);
                         
       }                
                 
       public static void main(String args[]){
             
              try{
                  
                  JTabbedPane tabPane = new JTabbedPane();                                                          
                                      
                  JPanel contenedor = new JPanel();
                  contenedor.setLayout(new GridBagLayout());
              
                  GridBagConstraints gbc = new GridBagConstraints();                                              
                  contenedor.add(new Applet(),gbc);
                            
                  tabPane.addTab("Consolidacion",contenedor);
                  tabPane.addTab("Generar Reportes",new GenerarReportes(""));
                   
                  frame = new JFrame("Consolidacion de datos");
                  frame.setExtendedState(JFrame.MAXIMIZED_BOTH);              
                  frame.setContentPane(tabPane);                  
                  frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);                  
                  frame.setVisible(true);
                  
              }catch(IOException | InvalidFormatException | HeadlessException e){ e.printStackTrace(); }
                                                  
       }
      
}

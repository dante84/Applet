
import javax.swing.JPanel;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;

// @author Daniel.Meza
public class TablaResultados extends JPanel{
    
       JTable tResultados;
       public TablaResultados(){
                           
	      tResultados = new JTable(new DefaultTableModel());
              DefaultTableModel dftm = new DefaultTableModel();
              dftm.addColumn("No aplicacion");
              dftm.addColumn("Existe");
              dftm.addColumn("Datif");
              dftm.addColumn("No Imagenes");
              dftm.addColumn("No PRegistro ");
              dftm.addColumn("No PRegistro BPM"); 
              dftm.addColumn("No PRegistro MControl");
              dftm.addColumn("No PRespuesta ");
              dftm.addColumn("No PRespuesta BPM"); 
              dftm.addColumn("No PRespuesta MControl");
              dftm.addColumn("Estado");
             
              tResultados.setModel(dftm);
              
              add(tResultados);
              
       }
      
}

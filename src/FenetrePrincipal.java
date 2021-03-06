/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import javax.swing.BorderFactory;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JPanel;
import javax.swing.JScrollBar;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.WindowConstants;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


//MODIFICATION TEST GIT


/**
 *
 * @author Fakri
 */
public class FenetrePrincipal extends javax.swing.JFrame {
	static String[][] data ;
	static int rowNum , colNum ;
	
    // Variables declaration - do not modify                     
    private JLabel jLabel1,jLabelBanniere;
    private JMenu jMenu1,jMenu2;
    private JMenuBar jMenuBar1;
    private JPanel jPanel4,jPanelBanniere,jPanelCentre,jPanelOption;
    private JPanel jPanelBouton1,jPanelBouton2;
    private JScrollBar jScrollBar1;
    private JScrollPane jScrollPane1 ; 
    private JTable jTable1;
    private JButton ajouterBouton,statBouton,pdfBouton,exportBouton;

    // End of variables declaration   
	
	
	
	
	
	
    /**
     * Creates new form NewJFrame
     */
    public FenetrePrincipal() {
        initComponents();
        
       // jLabel1.setBorder(BorderFactory.createTitledBorder("Option 1"));
        jPanelBouton1.setBorder(BorderFactory.createTitledBorder("Option 1"));
        jPanelBouton2.setBorder(BorderFactory.createTitledBorder("Option 2"));
        jPanelCentre.setBorder(BorderFactory.createTitledBorder("Liste Etudiant"));
        jPanelBanniere.setBorder(BorderFactory.createTitledBorder(""));
        ImageIcon icone =  new ImageIcon("../InterfaceProjetGestion/src/image1.png"); 
        jLabelBanniere = new JLabel(icone);        
        jLabelBanniere.setSize(jPanelBanniere.getWidth(),jPanelBanniere.getHeight());
        jPanelBanniere.add(jLabelBanniere);
        //jPanelBanniere.setBackground(Color.BLUE);
        jPanelBanniere.repaint();
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">                          
    private void initComponents() {

        jScrollBar1 = new JScrollBar();
        jPanelOption = new JPanel();
        jPanelBanniere = new JPanel();
        jLabelBanniere = new JLabel();
        jPanel4 = new JPanel();
        jLabel1 = new JLabel();
        jPanelCentre = new JPanel();
        jScrollPane1 = new JScrollPane();
        jTable1 = new JTable();
        jMenuBar1 = new JMenuBar();
        jMenu1 = new JMenu();
        jMenu2 = new JMenu();
        ajouterBouton = new JButton("Ajouter Etudiant");
        statBouton = new JButton   ("   Satistiques   ");
        jPanelBouton1 = new JPanel();
        jPanelBouton2 = new JPanel();
        pdfBouton = new JButton("   Exporter Pdf   ");
        exportBouton = new JButton ("Exporter Fichier ");
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        getContentPane().setLayout(new javax.swing.BoxLayout(getContentPane(), javax.swing.BoxLayout.LINE_AXIS));
        jLabelBanniere.setText("");

        javax.swing.GroupLayout jPanelBanniereLayout = new javax.swing.GroupLayout(jPanelBanniere);
        jPanelBanniere.setLayout(jPanelBanniereLayout);
        jPanelBanniereLayout.setHorizontalGroup(
            jPanelBanniereLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jLabelBanniere, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        jPanelBanniereLayout.setVerticalGroup(
            jPanelBanniereLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanelBanniereLayout.createSequentialGroup()
                .addComponent(jLabelBanniere, javax.swing.GroupLayout.PREFERRED_SIZE, 50, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 0, Short.MAX_VALUE))
        );

        jPanel4.setLayout(new java.awt.GridLayout(4, 1));

        jLabel1.setText("");
        jPanel4.add(jPanelBouton1);
        jPanel4.add(jPanelBouton2);
        jPanelBouton1.add(ajouterBouton);
        jPanelBouton1.add(statBouton);
        jPanelBouton2.add(exportBouton);
        jPanelBouton2.add(pdfBouton);
        
        try {
			File excel = new File("results.xlsx");
			FileInputStream fis = new FileInputStream(excel);
			
			XSSFWorkbook wb = new XSSFWorkbook(fis); 
			XSSFSheet ws = wb.getSheet("M1 Informatique - S8 - Pre-insc");
			
			 rowNum = ws.getLastRowNum()+1 ; 
			 colNum = ws.getRow(0).getLastCellNum(); 
			
			data = new String [rowNum][colNum]; 
			
			for (int i = 0 ; i< rowNum ; i++  )
			{
				XSSFRow row = ws.getRow(i); 
				for (int j = 0; j < colNum ; j++) 
				{
					XSSFCell cell = row.getCell(j);
					String value = cellToString(cell);
					data [i][j] = value ; 
					System.out.println("The value is : " +value); 
				}
				
			}
			
		} catch (FileNotFoundException e) 
		{
			// TODO Auto-generated catch block
			System.out.println("Erreur");
			e.printStackTrace();
		} catch (IOException e) 
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        
        String[] entetes = new String[colNum];
        for (int i = 0; i < colNum ; i++) 
        {
        	entetes[i] = data[0][i];
		}
        
        jTable1.setModel(new javax.swing.table.DefaultTableModel(data, entetes
        ));
        jScrollPane1.setViewportView(jTable1);
        jTable1.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);

        javax.swing.GroupLayout jPanelCentreLayout = new javax.swing.GroupLayout(jPanelCentre);
        jPanelCentre.setLayout(jPanelCentreLayout);
        jPanelCentreLayout.setHorizontalGroup(
            jPanelCentreLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanelCentreLayout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 465, Short.MAX_VALUE)
                .addContainerGap())
        );
        jPanelCentreLayout.setVerticalGroup(
            jPanelCentreLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanelCentreLayout.createSequentialGroup()
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 481, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 0, Short.MAX_VALUE))
        );

        javax.swing.GroupLayout jPanelOptionLayout = new javax.swing.GroupLayout(jPanelOption);
        jPanelOption.setLayout(jPanelOptionLayout);
        jPanelOptionLayout.setHorizontalGroup(
            jPanelOptionLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanelBanniere, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addGroup(jPanelOptionLayout.createSequentialGroup()
                .addComponent(jPanel4, javax.swing.GroupLayout.PREFERRED_SIZE, 150, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jPanelCentre, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanelOptionLayout.setVerticalGroup(
            jPanelOptionLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanelOptionLayout.createSequentialGroup()
                .addComponent(jPanelBanniere, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanelOptionLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jPanelCentre, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jPanel4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
        );

        getContentPane().add(jPanelOption);

        jMenu1.setText("Fichier");
        jMenuBar1.add(jMenu1);
		JMenuItem importer = new JMenuItem("Importer");
		JMenuItem exporter = new JMenuItem("Exporter");
		jMenu1.add(importer);
		jMenu1.add(exporter);
		
        jMenu2.setText("Edition");
        jMenuBar1.add(jMenu2);

        setJMenuBar(jMenuBar1);
        
        importer.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				//FileExtensionFilterDemo f = new FileExtensionFilterDemo();
				showOpenFileDialog();
			}
		});

        pack();
    }// </editor-fold>                        

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(FenetrePrincipal.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(FenetrePrincipal.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(FenetrePrincipal.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(FenetrePrincipal.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new FenetrePrincipal().setVisible(true);
            }
        });
    }
    
    private static String cellToString(XSSFCell cell) {
		// TODO Auto-generated method stub
		int type ;
		Object result ;
		if (cell == null)
			return (String) (result ="");
		type = cell.getCellType() ;
		
		switch (type)
		{
		
		case 0 : 
			result = cell.getNumericCellValue();
			break;
			
		case 1 : 
			result = cell.getStringCellValue();
			break;
		
		default :
			throw new RuntimeException("There are no support for this type of cell ");
		
		}
		return result.toString();
	}
    
	 //Methode pour importer un fichier Excel
    
	   private static void showOpenFileDialog() {
	        JFileChooser fileChooser = new JFileChooser();
	        fileChooser.setCurrentDirectory(new File(System.getProperty("user.home")));
	        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
	   //     fileChooser.addChoosableFileFilter(new FileNameExtensionFilter("PDF Documents", "pdf"));
	        fileChooser.addChoosableFileFilter(new FileNameExtensionFilter("MS Office Documents", "docx", "xlsx", "pptx"));
	   //     fileChooser.addChoosableFileFilter(new FileNameExtensionFilter("Images", "jpg", "png", "gif", "bmp"));
	        fileChooser.setAcceptAllFileFilterUsed(false);
	        int result = fileChooser.showOpenDialog(fileChooser);
	        if (result == JFileChooser.APPROVE_OPTION) {
	            File selectedFile = fileChooser.getSelectedFile();
	            System.out.println("Selected file: " + selectedFile.getAbsolutePath());
	        }
	        else
	        {
	        	System.out.println("erreur fichier!");
	        	
	        }
	    }

                
}

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package abc;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

/**
 *
 * @author Zafack Billy
 */
public class InterfaceUser extends javax.swing.JDialog {

    /**
     * Creates new form InterfaceUser
     */
    public InterfaceUser(java.awt.Frame parent, boolean modal) {
        super(parent, modal);
        initComponents();
        this.annee = extractAnnee(LocalDateTime.now().toString());
        this.setLocationRelativeTo(null);
        this.setModal(true);
        this.setAlwaysOnTop(true);
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jLabel1 = new javax.swing.JLabel();
        jTextField1 = new javax.swing.JTextField();
        jFileChooser1 = new javax.swing.JFileChooser();
        jLabel2 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jLabel6 = new javax.swing.JLabel();
        jButton1 = new javax.swing.JButton();
        parcourir_fichier_source = new javax.swing.JButton();
        parcourir_dossier_destination = new javax.swing.JButton();
        nom_fichier_destination = new javax.swing.JTextField();
        label_synthese = new javax.swing.JLabel();
        label_destination = new javax.swing.JLabel();
        load = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        parcourir_fichier_synthese = new javax.swing.JButton();
        label_source1 = new javax.swing.JLabel();

        jLabel1.setText("jLabel1");

        jTextField1.setText("jTextField1");

        setDefaultCloseOperation(javax.swing.WindowConstants.DISPOSE_ON_CLOSE);

        jLabel2.setFont(new java.awt.Font("Tekton Pro Cond", 1, 36)); // NOI18N
        jLabel2.setText("SYNTHESE   HP");

        jLabel4.setFont(new java.awt.Font("Tahoma", 1, 13)); // NOI18N
        jLabel4.setText("Fichier TRP Source : ");

        jLabel5.setFont(new java.awt.Font("Tahoma", 1, 13)); // NOI18N
        jLabel5.setText("Dossier Destination:");

        jLabel6.setFont(new java.awt.Font("Tahoma", 1, 13)); // NOI18N
        jLabel6.setText("Nom Fichier Destination:");

        jButton1.setFont(new java.awt.Font("Tahoma", 1, 14)); // NOI18N
        jButton1.setText("LANCER");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        parcourir_fichier_source.setText("Parcourir ...");
        parcourir_fichier_source.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                parcourir_fichier_sourceActionPerformed(evt);
            }
        });

        parcourir_dossier_destination.setText("Parcourir ...");
        parcourir_dossier_destination.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                parcourir_dossier_destinationActionPerformed(evt);
            }
        });

        label_synthese.setText("Aucun Fichier Choisi");

        label_destination.setText("Aucun Dossier Choisi");

        load.setText(".");

        jLabel7.setFont(new java.awt.Font("Tahoma", 1, 13)); // NOI18N
        jLabel7.setText("Fichier Synthese act. : ");

        parcourir_fichier_synthese.setText("Parcourir ...");
        parcourir_fichier_synthese.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                parcourir_fichier_syntheseActionPerformed(evt);
            }
        });

        label_source1.setText("Aucun Fichier Choisi");

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(88, 88, 88)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel4)
                            .addComponent(jLabel7))
                        .addGap(18, 18, 18)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(parcourir_fichier_source, javax.swing.GroupLayout.PREFERRED_SIZE, 109, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(label_source1))
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(parcourir_fichier_synthese, javax.swing.GroupLayout.PREFERRED_SIZE, 109, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(label_synthese)))
                        .addContainerGap(67, Short.MAX_VALUE))
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel6)
                            .addComponent(jLabel5))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(parcourir_dossier_destination, javax.swing.GroupLayout.PREFERRED_SIZE, 109, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(label_destination))
                            .addComponent(nom_fichier_destination, javax.swing.GroupLayout.PREFERRED_SIZE, 135, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(0, 0, Short.MAX_VALUE))))
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(197, 197, 197)
                        .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 113, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(172, 172, 172)
                        .addComponent(jLabel2))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(203, 203, 203)
                        .addComponent(load)))
                .addGap(0, 0, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(42, 42, 42)
                .addComponent(jLabel2)
                .addGap(39, 39, 39)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel4)
                    .addComponent(parcourir_fichier_source)
                    .addComponent(label_source1))
                .addGap(34, 34, 34)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel7)
                    .addComponent(parcourir_fichier_synthese)
                    .addComponent(label_synthese))
                .addGap(35, 35, 35)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel5)
                    .addComponent(parcourir_dossier_destination)
                    .addComponent(label_destination))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 33, Short.MAX_VALUE)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel6)
                    .addComponent(nom_fichier_destination, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(28, 28, 28)
                .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 41, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(load)
                .addGap(30, 30, 30))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void parcourir_fichier_sourceActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_parcourir_fichier_sourceActionPerformed
        // TODO add your handling code here:
         // TODO add your handling code here:
        JFileChooser fc = new JFileChooser();
         int returnVal = fc.showOpenDialog(InterfaceUser.this); 
         FileNameExtensionFilter filter = new FileNameExtensionFilter(
        "Fichier Excel", "xls", "xlsx");
        fc.setFileFilter(filter);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            File file = fc.getSelectedFile();
            //This is where a real application would open the file.
            file_source_chosen=true;
            fichier_source = fc.getSelectedFile().getPath();
            parcourir_fichier_source.setText(file.getName());
            label_source1.setText("Fichier choisie");
        }  
    }//GEN-LAST:event_parcourir_fichier_sourceActionPerformed

    private void parcourir_dossier_destinationActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_parcourir_dossier_destinationActionPerformed
        // TODO add your handling code here:
        JFileChooser fc = new JFileChooser();
        fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
         int returnVal = fc.showOpenDialog(InterfaceUser.this);  
        //fc.setFileFilter(filter);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            File file = fc.getSelectedFile();
            //This is where a real application would open the file.
            dossier_destination_chosen=true;
            chemin_destination = fc.getCurrentDirectory().getAbsolutePath()+"\\"+file.getName();
            System.out.println("Chemin destinations : "+chemin_destination);
            parcourir_dossier_destination.setText(file.getName());
            label_destination.setText("Dossier Choisie");
        } 
    }//GEN-LAST:event_parcourir_dossier_destinationActionPerformed
    public boolean testFormatAnnee(String annee){
        boolean valide = true;
        if(annee.length() != 4){
            valide = false;
        }
        try{
            annee+=2;
        }catch(NumberFormatException er){
            valide = false;
            er.printStackTrace();
        }
        return valide;
    }
    public boolean testValidFileName(String name){
        if(name.contains("/") || name.contains("\\") || name.contains(":")|| name.contains("*")|| name.contains("?")|| name.contains("\"")|| name.contains("<")|| name.contains(">")|| name.contains("|")){
         return false;   
        }else{
            return true;
        }
    }
    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        // TODO add your handling code here:
         //load.setText("Execution en cours ...");
        ImageIcon loading = new ImageIcon("loading30.gif");
        //load = new JLabel("Chargement... ", loading, JLabel.CENTER);
        load.setIcon(loading);

        new SwingWorker<Void, String>() {
            @Override
            protected Void doInBackground() throws Exception {
                // Worken hard or hardly worken...
                InterfaceUser.this.jButton1.setEnabled(false);
                load.setVisible(true);
                if(nom_fichier_destination.getText().equals("")){
                    nom_fichier_destination_chosen = false;
                }else{
                    nom_fichier_destination_chosen = true;
                }
                if(!testValidFileName(nom_fichier_destination.getText())){
                    // JOptionPane.showMessageDialog(null, "Vous avez entree un nom de fichier destination invalide. Il contient un caractere invalide(/\\:*?\"<>|)", "Nom Fichier Invalide", JOptionPane.ERROR_MESSAGE);
                }else{
                    if(!file_source_chosen){
                        if(!(new File("TRP").exists())){
                            //   JOptionPane.showMessageDialog(this, "Bien vouloir mettre les TRPs source dans un dossier \"TRP\"", "Fichier TRP Source Introuvable", JOptionPane.ERROR_MESSAGE);
                            return null;
                        }else{
                            InterfaceUser.this.semaine_courante = extractSemaine(getLatestFilefromDir("TRP").getName().toString());
                            InterfaceUser.this.fichier_source = getLatestFilefromDir("TRP").getAbsolutePath();
                        }
                    }
                    if(!dossier_destination_chosen){
                        if(!(new File("SYNTHESE").exists())){
                            new File("SYNTHESE").mkdir();
                        }
                    }
                    //this.annee = this.annee_courante.getText();
                    if(!dossier_destination_chosen && !nom_fichier_destination_chosen){
                        //JOptionPane.showMessageDialog(this, "Vous n'avez ni renseignee le dossier destination ni le nom du fichier destination. \nDes valeurs par defaut seront utilise(meme dossier que l'executable de ce programme)", "Valeurs par defaut", JOptionPane.WARNING_MESSAGE);
                    }else if(!dossier_destination_chosen && nom_fichier_destination_chosen){
                        //   JOptionPane.showMessageDialog(this, "Vous n'avez pas renseignee le dossier destination. \nDes valeurs par defaut seront utilise(meme dossier que l'executable de ce programme)", "Valeurs par defaut", JOptionPane.WARNING_MESSAGE);
                    }else if(dossier_destination_chosen && !nom_fichier_destination_chosen){
                        //    JOptionPane.showMessageDialog(this, "Vous n'avez pas renseignee le nom du fichier destination. \nDes valeurs par defaut seront utilise", "Valeurs par defaut", JOptionPane.WARNING_MESSAGE);
                    }
                    //System.getProperty("user.home");
                    if(!dossier_destination_chosen && !nom_fichier_destination_chosen){
                        if(!file_synthese_chosen){
                            fichier_destination = "SyntheseHP_Semaine_"+InterfaceUser.this.semaine_courante+"_Du_"+ LocalDateTime.now().toString().replaceAll(":", "-");
                            Abcd abcd = new Abcd(fichier_source, "SYNTHESE", fichier_destination+".xlsx", annee, null);
                            try {
                                load.setText(".");
                                abcd.writeFormatALLSheet(abcd.readALL(abcd.fichier_source), "SYNTHESE/" + fichier_destination + ".xlsx");
                                load.setVisible(false);
                                if(JOptionPane.showConfirmDialog(InterfaceUser.this, "Voulez vous ouvrir le fichier destination", "Ouvrir Fichier Resultat", JOptionPane.OK_CANCEL_OPTION, JOptionPane.QUESTION_MESSAGE) == JOptionPane.OK_OPTION){
                                    Desktop.getDesktop().open(new File("SYNTHESE/"+fichier_destination+".xlsx"));
                                }
                            } catch (IOException ex) {
                                Logger.getLogger(InterfaceUser.class.getName()).log(Level.SEVERE, null, ex);
                            }
                        }else{
                            fichier_destination = "SyntheseHP_Semaine_"+InterfaceUser.this.semaine_courante+"_Du_"+ LocalDateTime.now().toString().replaceAll(":", "-");
                            Abcd abcd = new Abcd(fichier_source, "SYNTHESE", fichier_destination+".xlsx", annee, fichier_synthese);
                            try {
                                load.setText(".");
                                abcd.writeFormatALLSheet(abcd.readALL(abcd.fichier_source), "SYNTHESE/" + fichier_destination + ".xlsx");
                                load.setVisible(false);
                                if(JOptionPane.showConfirmDialog(InterfaceUser.this, "Voulez vous ouvrir le fichier destination", "Ouvrir Fichier Resultat", JOptionPane.OK_CANCEL_OPTION, JOptionPane.QUESTION_MESSAGE) == JOptionPane.OK_OPTION){
                                    Desktop.getDesktop().open(new File("SYNTHESE/"+fichier_destination+".xlsx"));
                                }
                            } catch (IOException ex) {
                                Logger.getLogger(InterfaceUser.class.getName()).log(Level.SEVERE, null, ex);
                            }
                        }
                    }else if(dossier_destination_chosen && !nom_fichier_destination_chosen){
                        if(!file_synthese_chosen){
                            fichier_destination = "SyntheseHP_Semaine_"+InterfaceUser.this.semaine_courante+"_Du_"+ LocalDateTime.now().toString().replaceAll(":", "-") ;
                            Abcd abcd = new Abcd(fichier_source, chemin_destination, fichier_destination+".xlsx", annee, null);
                            try {
                                abcd.writeFormatALLSheet(abcd.readALL(abcd.fichier_source), chemin_destination+"/"+fichier_destination +".xlsx");
                                load.setText(".");
                                load.setVisible(false);
                                if(JOptionPane.showConfirmDialog(InterfaceUser.this, "Voulez vous ouvrir le fichier destination", "Ouvrir Fichier Resultat", JOptionPane.OK_CANCEL_OPTION, JOptionPane.QUESTION_MESSAGE) == JOptionPane.OK_OPTION){
                                    Desktop.getDesktop().open(new File(chemin_destination+"/"+fichier_destination +".xlsx"));
                                }
                            } catch (IOException ex) {
                                Logger.getLogger(InterfaceUser.class.getName()).log(Level.SEVERE, null, ex);
                            }
                        }else{
                            fichier_destination = "SyntheseHP_Semaine_"+InterfaceUser.this.semaine_courante+"_Du_"+ LocalDateTime.now().toString().replaceAll(":", "-");
                            Abcd abcd = new Abcd(fichier_source, chemin_destination, fichier_destination +".xlsx", annee, fichier_synthese);
                            try {
                                abcd.writeFormatALLSheet(abcd.readALL(abcd.fichier_source), chemin_destination+"/"+fichier_destination+".xlsx");
                                load.setText(".");
                                load.setVisible(false);
                                if(JOptionPane.showConfirmDialog(InterfaceUser.this, "Voulez vous ouvrir le fichier destination", "Ouvrir Fichier Resultat", JOptionPane.OK_CANCEL_OPTION, JOptionPane.QUESTION_MESSAGE) == JOptionPane.OK_OPTION){
                                    Desktop.getDesktop().open(new File(chemin_destination+"/"+fichier_destination +".xlsx"));
                                }
                            } catch (IOException ex) {
                                Logger.getLogger(InterfaceUser.class.getName()).log(Level.SEVERE, null, ex);
                            }
                        }

                    }else if(!dossier_destination_chosen && nom_fichier_destination_chosen){
                        if(!file_synthese_chosen){
                            Abcd abcd = new Abcd(fichier_source, "SYNTHESE", nom_fichier_destination.getText()+".xlsx", annee, null);
                            try {
                                abcd.writeFormatALLSheet(abcd.readALL(abcd.fichier_source),"SYNTHESE/"+nom_fichier_destination.getText()+".xlsx");
                                load.setText(".");
                                load.setVisible(false);
                                if(JOptionPane.showConfirmDialog(InterfaceUser.this, "Voulez vous ouvrir le fichier destination", "Ouvrir Fichier Resultat", JOptionPane.OK_CANCEL_OPTION, JOptionPane.QUESTION_MESSAGE) == JOptionPane.OK_OPTION){
                                    Desktop.getDesktop().open(new File("SYNTHESE/"+nom_fichier_destination.getText()+".xlsx"));
                                }
                            } catch (IOException ex) {
                                Logger.getLogger(InterfaceUser.class.getName()).log(Level.SEVERE, null, ex);
                            }
                        }else{
                            Abcd abcd = new Abcd(fichier_source, "SYNTHESE", nom_fichier_destination.getText()+".xlsx", annee, fichier_synthese);
                            try {
                                abcd.writeFormatALLSheet(abcd.readALL(abcd.fichier_source),"SYNTHESE/"+nom_fichier_destination.getText()+".xlsx");
                                load.setText(".");
                                load.setVisible(false);
                                if(JOptionPane.showConfirmDialog(InterfaceUser.this, "Voulez vous ouvrir le fichier destination", "Ouvrir Fichier Resultat", JOptionPane.OK_CANCEL_OPTION, JOptionPane.QUESTION_MESSAGE) == JOptionPane.OK_OPTION){
                                    Desktop.getDesktop().open(new File("SYNTHESE/"+nom_fichier_destination.getText()+".xlsx"));
                                }
                            } catch (IOException ex) {
                                Logger.getLogger(InterfaceUser.class.getName()).log(Level.SEVERE, null, ex);
                            }
                        }

                    }else{
                        if(!file_synthese_chosen){
                            Abcd abcd = new Abcd(fichier_source, chemin_destination, nom_fichier_destination.getText()+".xlsx", annee, null);
                            try {
                                abcd.writeFormatALLSheet(abcd.readALL(abcd.fichier_source), chemin_destination+"/"+nom_fichier_destination.getText()+".xlsx");
                                load.setText(".");
                                load.setVisible(false);
                                if(JOptionPane.showConfirmDialog(InterfaceUser.this, "Voulez vous ouvrir le fichier destination", "Ouvrir Fichier Resultat", JOptionPane.OK_CANCEL_OPTION, JOptionPane.QUESTION_MESSAGE) == JOptionPane.OK_OPTION){
                                    Desktop.getDesktop().open(new File(chemin_destination+"/"+nom_fichier_destination.getText()+".xlsx"));
                                }
                            } catch (IOException ex) {
                                Logger.getLogger(InterfaceUser.class.getName()).log(Level.SEVERE, null, ex);
                            }
                        }else{
                            Abcd abcd = new Abcd(fichier_source, chemin_destination, nom_fichier_destination.getText()+".xlsx", annee, fichier_synthese);
                            try {
                                abcd.writeFormatALLSheet(abcd.readALL(abcd.fichier_source), chemin_destination+"/"+nom_fichier_destination.getText()+".xlsx");
                                load.setText(".");
                                load.setVisible(false);
                                if(JOptionPane.showConfirmDialog(InterfaceUser.this, "Voulez vous ouvrir le fichier destination", "Ouvrir Fichier Resultat", JOptionPane.OK_CANCEL_OPTION, JOptionPane.QUESTION_MESSAGE) == JOptionPane.OK_OPTION){
                                    Desktop.getDesktop().open(new File(chemin_destination+"/"+nom_fichier_destination.getText()+".xlsx"));
                                }
                            } catch (IOException ex) {
                                Logger.getLogger(InterfaceUser.class.getName()).log(Level.SEVERE, null, ex);
                            }
                        }
                    }
                }

                return null;
            }

            @Override
            protected void done() {
                InterfaceUser.this.jButton1.setEnabled(true);
                load.setVisible(false);
            }
        }.execute();
//
//

    }//GEN-LAST:event_jButton1ActionPerformed

    private void parcourir_fichier_syntheseActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_parcourir_fichier_syntheseActionPerformed
        //-- TODO add your handling code here:
        JFileChooser fc = new JFileChooser();
         int returnVal = fc.showOpenDialog(InterfaceUser.this); 
         FileNameExtensionFilter filter = new FileNameExtensionFilter(
        "Fichier Excel", "xls", "xlsx");
        fc.setFileFilter(filter);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            File file = fc.getSelectedFile();
            //This is where a real application would open the file.
            file_synthese_chosen=true;
             fichier_synthese = fc.getSelectedFile().getPath();
            parcourir_fichier_synthese.setText(file.getName());
            label_synthese.setText("Fichier choisie");
        } 
    }//GEN-LAST:event_parcourir_fichier_syntheseActionPerformed
  private File getLatestFilefromDir(String dirPath){
    File dir = new File(dirPath);
    File[] files = dir.listFiles();
    if (files == null || files.length == 0) {
        return null;
    }

    File lastModifiedFile = files[0];
    for (int i = 1; i < files.length; i++) {
       if (lastModifiedFile.lastModified() < files[i].lastModified()) {
           lastModifiedFile = files[i];
       }
    }
    return lastModifiedFile;
}
  //flag = true ==> Les TRPs ;; flag = flase ==> La synthese
//  public ArrayList<File> listFilesForFolder(final File folder, boolean  flag) {
//    ArrayList<File> trp = new ArrayList<File>(); 
//    for (final File fileEntry : folder.listFiles()) {
//        if (!fileEntry.isDirectory()) {
//            trp.add(fileEntry);
//        }
//    }
//    return trp;
//}
  public String extractSemaine(String filename){
      System.out.println("_____________ "+filename.substring(0, 3));
      if(filename.substring(0, 3).equals("TRP")){
      return filename.charAt(filename.indexOf("(")+1)+"";
  }else{
      return "";
    }
  }
  
  public String extractAnnee(String timestamp){
       return timestamp.substring(0, 4);
  }
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
            java.util.logging.Logger.getLogger(InterfaceUser.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(InterfaceUser.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(InterfaceUser.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(InterfaceUser.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the dialog */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                InterfaceUser dialog = new InterfaceUser(new javax.swing.JFrame(), true);
                dialog.addWindowListener(new java.awt.event.WindowAdapter() {
                    @Override
                    public void windowClosing(java.awt.event.WindowEvent e) {
                        System.exit(0);
                    }
                });
                dialog.setVisible(true);
            }
        });
    }
    
    public String semaine_courante="";
    public String dossier_TRP = "TRP";
    public String dossier_synthese = "Synthese";
    public boolean file_source_chosen = false;
    public boolean file_synthese_chosen = false;
    public String fichier_source;
    public String chemin_destination;
    public boolean dossier_destination_chosen = false;
    public boolean nom_fichier_destination_chosen = false;
    public String fichier_destination;
    public String fichier_synthese;
    public String annee;
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JFileChooser jFileChooser1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JLabel label_destination;
    private javax.swing.JLabel label_source1;
    private javax.swing.JLabel label_synthese;
    private javax.swing.JLabel load;
    private javax.swing.JTextField nom_fichier_destination;
    private javax.swing.JButton parcourir_dossier_destination;
    private javax.swing.JButton parcourir_fichier_source;
    private javax.swing.JButton parcourir_fichier_synthese;
    // End of variables declaration//GEN-END:variables
}

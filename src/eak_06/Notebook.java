package eak_06;

import java.awt.Cursor;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class Notebook extends javax.swing.JFrame {

    private static final long serialVersionUID = 1L;

    class Launch_Thread extends Thread {

        @Override
        public void run() {
            String dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
                    + System.getProperty("file.separator");
            try {
                modifData(dir + "notebook_template.xls", dir + "notebook.xls",
                        jTextField_Name.getText(),
                        jTextField_Surname.getText(),
                        jTextField_Patronymic.getText());
                Desktop.getDesktop().open(new File(dir + "notebook.xls"));
            } catch (IOException ex) {
                System.err.println("Error modifData!");
            }
            setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
        }
    }

    // Метод создания отчета
    private void modifData(String inputFileName, String outputFileName, String Date, String Day, String Time) throws IOException {

        HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(inputFileName)));
        HSSFSheet sheet = wb.getSheetAt(0);
        sheet.getRow(2).getCell(2).setCellValue(Date);
        sheet.getRow(3).getCell(2).setCellValue(Day);
        sheet.getRow(4).getCell(2).setCellValue(Time);

        try (FileOutputStream fileOut = new FileOutputStream(outputFileName)) {
            wb.write(fileOut);
        }
    }

    public Notebook() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jTextField_Name = new javax.swing.JTextField();
        jTextField_Surname = new javax.swing.JTextField();
        jTextField_Patronymic = new javax.swing.JTextField();
        jButton1 = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("notebook");
        setLocation(new java.awt.Point(1000, 1000));
        setResizable(false);
        getContentPane().setLayout(null);

        jTextField_Name.setCaretColor(new java.awt.Color(255, 0, 255));
        getContentPane().add(jTextField_Name);
        jTextField_Name.setBounds(310, 110, 210, 26);

        jTextField_Surname.setCaretColor(new java.awt.Color(255, 0, 255));
        jTextField_Surname.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_SurnameActionPerformed(evt);
            }
        });
        getContentPane().add(jTextField_Surname);
        jTextField_Surname.setBounds(310, 150, 210, 35);

        jTextField_Patronymic.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_PatronymicActionPerformed(evt);
            }
        });
        getContentPane().add(jTextField_Patronymic);
        jTextField_Patronymic.setBounds(310, 200, 210, 25);

        jButton1.setFont(new java.awt.Font("Times New Roman", 0, 14)); // NOI18N
        jButton1.setText("Экспортировать в Excel");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });
        getContentPane().add(jButton1);
        jButton1.setBounds(200, 440, 190, 30);

        jLabel1.setIcon(new javax.swing.ImageIcon(getClass().getResource("/eak_06/notebook.jpg"))); // NOI18N
        getContentPane().add(jLabel1);
        jLabel1.setBounds(0, 0, 620, 430);

        setBounds(0, 0, 710, 539);
    }// </editor-fold>//GEN-END:initComponents

    private void jTextField_PatronymicActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_PatronymicActionPerformed
    }//GEN-LAST:event_jTextField_PatronymicActionPerformed

    private void jTextField_SurnameActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_SurnameActionPerformed
    }//GEN-LAST:event_jTextField_SurnameActionPerformed

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        new Launch_Thread().start();
    }//GEN-LAST:event_jButton1ActionPerformed

    public static void main(String args[]) {
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Windows".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Notebook.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        java.awt.EventQueue.invokeLater(() -> {
            new Notebook().setVisible(true);
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JTextField jTextField_Name;
    private javax.swing.JTextField jTextField_Patronymic;
    private javax.swing.JTextField jTextField_Surname;
    // End of variables declaration//GEN-END:variables
}

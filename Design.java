package duannhatruong;

import DAOduan.DAOkehoach;
import DAOduan.DAOkehoachthiimprements;
import Service.checksv;
import java.io.File;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;

public class Design extends javax.swing.JFrame {

    private DAOkehoach lst = new DAOkehoachthiimprements();

    private checksv check = new checksv();

    private String namefile;

    public Design() {
        initComponents();
        this.setLocationRelativeTo(null);

    }
    private int count = 0;

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jButton1 = new javax.swing.JButton();
        btnloai1 = new javax.swing.JButton();
        jLabel3 = new javax.swing.JLabel();
        txtlanthi = new javax.swing.JTextField();
        jButton3 = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jButton1.setFont(new java.awt.Font("Dialog", 0, 18)); // NOI18N
        jButton1.setText("đọc kế Hoạch Thi");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        btnloai1.setFont(new java.awt.Font("Dialog", 0, 18)); // NOI18N
        btnloai1.setText("upload file điểm online/điểm quiz");
        btnloai1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnloai1ActionPerformed(evt);
            }
        });

        jLabel3.setText("Lần Thi");

        txtlanthi.setText("1");
        txtlanthi.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtlanthiActionPerformed(evt);
            }
        });

        jButton3.setFont(new java.awt.Font("Dialog", 0, 18)); // NOI18N
        jButton3.setText("upload Điểm Danh");
        jButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton3ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jLabel3)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(txtlanthi, javax.swing.GroupLayout.PREFERRED_SIZE, 120, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(18, 18, 18)
                        .addComponent(jButton1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(btnloai1)
                        .addGap(18, 18, 18)
                        .addComponent(jButton3)
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel3)
                    .addComponent(txtlanthi, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jButton1))
                .addGap(18, 18, 18)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(btnloai1, javax.swing.GroupLayout.DEFAULT_SIZE, 47, Short.MAX_VALUE)
                    .addComponent(jButton3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap(36, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        try {
            lst.docexcel(chonfile());
            lst.luudb();
            JOptionPane.showMessageDialog(this, "đọc file thành công");
        } catch (Exception ex) {
            ex.printStackTrace();
            JOptionPane.showMessageDialog(this, "đọc file thất bại");
        }
    }//GEN-LAST:event_jButton1ActionPerformed

    private void btnloai1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnloai1ActionPerformed
        if (chonfile() != null) {
            try {
                check.ktramondauvao(chonfile());
                JOptionPane.showMessageDialog(this, "đọc file thành công");
                if (savefile() != null) {
                    try {
                        String fileString = savefile();
                        check.xuatdssthi(fileString, this.txtlanthi.getText(), count);
                        JOptionPane.showMessageDialog(this, "lưu file thành công");
                    } catch (Exception ex) {
                        ex.printStackTrace();
                        JOptionPane.showMessageDialog(this, "lưu file thất bại");
                    }
                }

            } catch (Exception ex) {
                ex.printStackTrace();
                JOptionPane.showMessageDialog(this, "đọc file thất bại");
            }
            if (check.xuatdssvthi() < 27) {
                count = 2;
            } else {
                count = 3;
            }

        }


    }//GEN-LAST:event_btnloai1ActionPerformed

    private void txtlanthiActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtlanthiActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_txtlanthiActionPerformed

    private void jButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton3ActionPerformed
        if (chonfile() != null) {
            try {
                check.docfilediemdanh(chonfile());
                JOptionPane.showMessageDialog(this, "đọc file thành công");
                if (check.xuatdssvthi() < 27) {
                    count = 2;
                } else {
                    count = 3;
                }
                if (savefile() != null) {
                    try {
                        String fileString = savefile();
                        check.xuatdssthi(fileString, this.txtlanthi.getText(), count);
                        JOptionPane.showMessageDialog(this, "lưu file thành công");
                    } catch (Exception ex) {
                        ex.printStackTrace();
                        JOptionPane.showMessageDialog(this, "lưu file thất bại");
                    }
                }

            } catch (Exception ex) {
                ex.printStackTrace();
                JOptionPane.showMessageDialog(this, "đọc file thất bại");
            }
        }


    }//GEN-LAST:event_jButton3ActionPerformed

    private String chonfile() {
        JFileChooser jfc = new JFileChooser();
        FileNameExtensionFilter fnef = new FileNameExtensionFilter(".xlsx", ".xlsm", ".xls");
        jfc.setFileFilter(fnef);
        jfc.setMultiSelectionEnabled(false);
        int ktr = jfc.showDialog(this, "chọn file");
        if (ktr == JFileChooser.APPROVE_OPTION) {
            File f = jfc.getSelectedFile();
            namefile = f.getAbsolutePath();
        }
        return namefile;
    }

    private String savefile() {
        JFileChooser jfc = new JFileChooser();
        jfc.setMultiSelectionEnabled(false);
        int ktr = jfc.showDialog(this, "save");
        if (ktr == JFileChooser.APPROVE_OPTION) {
            File f = jfc.getSelectedFile();
            namefile = f.getAbsolutePath();
        }
        return namefile;
    }

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
            java.util.logging.Logger.getLogger(Design.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Design.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Design.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Design.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Design().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnloai1;
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton3;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JTextField txtlanthi;
    // End of variables declaration//GEN-END:variables
}

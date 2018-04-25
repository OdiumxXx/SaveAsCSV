package saveascsvMain;

import java.awt.BorderLayout;
import java.awt.FlowLayout;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;

@SuppressWarnings("serial")
public class OpenDialog extends JDialog {
  JLabel Label = new JLabel();  
  String filePath;
  private final JPanel contentPanel = new JPanel();

  /**
   * Launch the application.
   */
  public static void main(String[] args) {
    try {
      OpenDialog dialog = new OpenDialog();
      dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
      dialog.setVisible(true);
    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  /**
   * Create the dialog.
   */
  public OpenDialog() {
    setBounds(100, 100, 450, 300);
    getContentPane().setLayout(new BorderLayout());
    contentPanel.setLayout(new FlowLayout());
    contentPanel.setBorder(new EmptyBorder(5, 5, 5, 5));
    getContentPane().add(contentPanel, BorderLayout.CENTER);
    {
      contentPanel.add(Label);
      Label.setText("Click 'Select Spreadsheet' and select the members database");
    }
    {
      JPanel buttonPane = new JPanel();
      buttonPane.setLayout(new FlowLayout(FlowLayout.RIGHT));
      getContentPane().add(buttonPane, BorderLayout.SOUTH);
      {
        JButton okButton = new JButton("Select Spreadsheet");
        okButton.addActionListener(new ActionListener() {
          public void actionPerformed(ActionEvent arg0) {
            JFileChooser fileChooser = new JFileChooser();
            File selectedFile = null;
            fileChooser.setCurrentDirectory(new File(System.getProperty("user.dir")));
            int result = fileChooser.showOpenDialog(okButton);
            if (result == JFileChooser.APPROVE_OPTION) {
              selectedFile = fileChooser.getSelectedFile();
              filePath = selectedFile.getAbsolutePath().toString();
              System.out.println("Selected file: " + selectedFile.getAbsolutePath());
              Label.setText("<html>Selected file:<br />"+ selectedFile.getAbsolutePath()+"<p>&nbsp</p><p>&nbsp</p>Click 'Create CSV File' to create the CSV file.<p>&nbsp<p />This window will close and a new 'memdraw.csv' will be created for you<br />in the same folder you opened this application from.<p>&nbsp</p>");
            }

          }
        });
        okButton.setActionCommand("OK");

        buttonPane.add(okButton);
        getRootPane().setDefaultButton(okButton);
      }
      {
        JButton createCSVButton = new JButton("Create CSV File");
        createCSVButton.addActionListener(new ActionListener() {
          public void actionPerformed(ActionEvent arg0) {
            saveascsvMain.ApachePOIExcelRead.main(filePath);
            setVisible(false);
          }
        });
        createCSVButton.setActionCommand("Create CSV File");        
        buttonPane.add(createCSVButton);
      }
    }
  }

  public File fileChooser() {
    JFileChooser fileChooser = new JFileChooser();
    File selectedFile = null;
    fileChooser.setCurrentDirectory(new File(System.getProperty("user.home")));
    int result = fileChooser.showOpenDialog(this);
    if (result == JFileChooser.APPROVE_OPTION) {
      selectedFile = fileChooser.getSelectedFile();
      System.out.println("Selected file: " + selectedFile.getAbsolutePath());
    }

    return selectedFile;
  }


}

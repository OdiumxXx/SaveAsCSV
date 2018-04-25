package saveascsvMain;

import java.awt.BorderLayout;
import java.awt.FlowLayout;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.border.EmptyBorder;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import javax.swing.SwingConstants;
import javax.swing.JTextArea;
import java.awt.Color;
import java.awt.Component;
import javax.swing.Box;

@SuppressWarnings("serial")
public class OpenDialog extends JDialog {
  JLabel Label = new JLabel(); 
  static JTextArea textArea = new JTextArea();
  static JScrollPane sp = new JScrollPane(textArea);
  String filePath;
  private final JPanel contentPanel = new JPanel();
  private final JButton btnExit = new JButton("Close");
  private final Component horizontalStrut = Box.createHorizontalStrut(525);

  private static void updateTextArea(String insertText) {
    textArea.append(insertText);
  }

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
    setBounds(400, 400, 900, 600);
    getContentPane().setLayout(new BorderLayout());
    contentPanel.setBorder(new EmptyBorder(5, 5, 5, 5));
    getContentPane().add(contentPanel, BorderLayout.CENTER);
    contentPanel.setLayout(new BorderLayout(0, 0));
    {
      Label.setHorizontalAlignment(SwingConstants.CENTER);
      contentPanel.add(Label, BorderLayout.NORTH);
      Label.setText("<html>Click 'Select Spreadsheet' and select the members database<p>&nbsp<p /><p>&nbsp<p /></html>");
    }
    {
      textArea.setBackground(Color.WHITE);
      textArea.setWrapStyleWord(true);
      textArea.setEditable(false);
      contentPanel.add(sp, BorderLayout.CENTER);
    }


    {
      JPanel buttonPane = new JPanel();
      buttonPane.setLayout(new FlowLayout(FlowLayout.LEFT));
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
              textArea.setText("\n- Selected file: " + selectedFile.getAbsolutePath()+"\n\n");
              updateTextArea("Click 'Create CSV File' to create the CSV file.\nA new 'memdraw.csv' will be created for you in the same folder you launched this application from.\n\n");
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
          }
        });
        createCSVButton.setActionCommand("Create CSV File");        
        buttonPane.add(createCSVButton);
      }
      {
        buttonPane.add(horizontalStrut);
      }
      {
        btnExit.addActionListener(new ActionListener() {
          public void actionPerformed(ActionEvent arg0) {
            System.exit(0);
          }
        });
        buttonPane.add(btnExit);
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

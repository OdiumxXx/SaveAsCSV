package saveascsvMain;

import java.awt.BorderLayout;
import java.awt.FlowLayout;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;

import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import javax.swing.SwingConstants;
import javax.swing.JTextArea;
import java.awt.Color;
import java.awt.Component;
import javax.swing.Box;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.Cursor;
import java.awt.Insets;

@SuppressWarnings("serial")
public class OpenDialog extends JDialog {
  JLabel Label = new JLabel(); 
  static JTextArea textArea = new JTextArea();
  static JScrollPane sp = new JScrollPane(textArea);
  String filePath;
  private final JPanel contentPanel = new JPanel();
  private final JButton btnExit = new JButton("Close");
  private final Component horizontalStrut = Box.createHorizontalStrut(540);
  private final JPanel panel = new JPanel();
  private final JButton btnAdvanced = new JButton("Advanced");

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
      dialog.setLocationRelativeTo(null);
      dialog.setVisible(true);
    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  /**
   * Create the dialog.
   */
  public OpenDialog() {
    setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
    setFont(new Font("Dialog", Font.PLAIN, 16));
    setIconImage(Toolkit.getDefaultToolkit().getImage(OpenDialog.class.getResource("/saveascsvMain/pixel-ghost-blue-128x128.png")));
    setType(Type.UTILITY);
    setTitle("Save as CSV");
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
   
    // Add our toolbar
    {
      FlowLayout flowLayout = (FlowLayout) panel.getLayout();
      flowLayout.setAlignment(FlowLayout.LEFT);
      getContentPane().add(panel, BorderLayout.NORTH);
    }    
    // add our preferences button
    {
      btnAdvanced.setForeground(Color.DARK_GRAY);
      btnAdvanced.setMargin(new Insets(2, 2, 2, 2));
      panel.add(btnAdvanced);
      btnAdvanced.setFont(new Font("Tahoma", Font.PLAIN, 10));
      btnAdvanced.addActionListener(new ActionListener() {
        public void actionPerformed(ActionEvent arg0) {
          
          String m = JOptionPane.showInputDialog(null, "Advanced Use Only:\nSearch for a different Sheet Title?", "Consolidate");
          
          if (m == null || m == "" || m == "\n") {
            // do nothing
            System.out.println("Title remains: "+saveascsvMain.ApachePOIExcelRead.ourTitle);
            updateTextArea("\n*** Title remains: "+saveascsvMain.ApachePOIExcelRead.ourTitle+"\n");
          } else {
            // change title we are looking for (ourTitle)
            System.out.println("Title set to: "+m);
            updateTextArea("\n*** Title set to: "+m+"\n");
            saveascsvMain.ApachePOIExcelRead.ourTitle = m;
          }
        }
      });
    }


    {
      JPanel buttonPane = new JPanel();
      FlowLayout fl_buttonPane = new FlowLayout(FlowLayout.LEFT);
      buttonPane.setLayout(fl_buttonPane);
      getContentPane().add(buttonPane, BorderLayout.SOUTH);
      {
        JButton okButton = new JButton("Select Spreadsheet");
        okButton.setFont(new Font("Tahoma", Font.PLAIN, 12));
        okButton.addActionListener(new ActionListener() {
          public void actionPerformed(ActionEvent arg0) {
            JFileChooser fileChooser = new JFileChooser();
            FileFilter filter = new FileNameExtensionFilter("XLSX, XLS files", new String[] {"xlsx", "xls"});
            fileChooser.setFileFilter(filter);
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
        createCSVButton.setFont(new Font("Tahoma", Font.PLAIN, 12));
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
        btnExit.setHorizontalAlignment(SwingConstants.RIGHT);
        btnExit.setFont(new Font("Tahoma", Font.PLAIN, 12));
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

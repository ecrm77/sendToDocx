package sendToDocx;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.io.*;
import java.net.URI;
import java.net.URISyntaxException;

public class Form  {

    private JButton createNew_Button;
    private JButton open_button;
    public JPanel mainPanel;
    public JPanel actionPanel;
    private JButton reload_button;
    public JTextField textField;
    public String textContent;
    private JButton send_button;
    private JCheckBox alwaysOnTopCheckBox;
    private JButton remove_button;
    public String path_to_write;
    private JLabel label_path;
    private JLabel pLabel;
    public Boolean onTop = false;

    public Form() throws IOException {

        createNew_Button.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

                try {
                    boolean out= readFromDocxList();

                    if(!out){

                        JOptionPane.showMessageDialog (null, "Please remove the loaded file first", "Cannot create!", JOptionPane.ERROR_MESSAGE);

                    }else {
                        JFrame frame1 = new JFrame();
                        FileDialog fDialog = new FileDialog(frame1, "Save", FileDialog.SAVE);

                        fDialog.setFilenameFilter((dir, name) -> name.endsWith(".docx"));
                        fDialog.setFile("Untitled.docx");

                        fDialog.setVisible(true);

                        String path = fDialog.getDirectory() + fDialog.getFile();


                        if (path.equals("nullnull")) {
                            System.out.println("user cancelled!");
                        } else {

                            try {
                                XWPFDocument doc = new XWPFDocument();

                                    FileOutputStream fos = new FileOutputStream(path);
                                    doc.write(fos);
                                    fos.close();


                            } catch (IOException ioException) {
                                ioException.printStackTrace();
                            }

                            try {
                                if (path.contains(" ")){

                                    String path1 = path.replaceAll(" ", "%20");
                                    addComponents(path1);
                                } else {

                                    addComponents(path);
                                }
                                Main.frame.revalidate();

                            } catch (IOException ioException) {
                                ioException.printStackTrace();
                            }
                            try {

                                    fileWrite(path);

                            } catch (IOException ioException) {
                                ioException.printStackTrace();
                            }
                        }
                    }

                } catch (IOException ex) {
                    ex.printStackTrace();
                }
            }
        });

        open_button.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

                try {
                    boolean out= readFromDocxList();

                    if(!out){

                        JOptionPane.showMessageDialog (null, "Please remove the loaded file first", "Cannot open!", JOptionPane.ERROR_MESSAGE);

                    }else {

                        JFrame frame1 = new JFrame();
                        FileDialog fDialog = new FileDialog(frame1, "Open", FileDialog.LOAD);

                        fDialog.setFilenameFilter((dir, name) -> name.endsWith(".docx"));
                        fDialog.setVisible(true);

                        String path = fDialog.getDirectory() + fDialog.getFile();
                        if (path.equals("nullnull")) {
                            System.out.println("user cancelled the choice");
                        }
                        else {
                            System.out.println(path);
                            try {
                                if (path.contains(" ")){

                                    String path1 = path.replaceAll(" ", "%20");
                                    addComponents(path1);
                                } else {

                                    addComponents(path);
                                }
                                Main.frame.revalidate();


                            } catch (IOException ioException) {
                                ioException.printStackTrace();
                            }

                            try {

                                    fileWrite(path);


                            } catch (IOException ioException) {
                                ioException.printStackTrace();
                            }
                        }
                    }

                } catch (IOException ex) {
                    ex.printStackTrace();
                }

            }

        });

        //reload_button is hidden
        reload_button.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

                FileReader reader = null;
                try {
                    reader = new FileReader(Main.txtFile);
                } catch (FileNotFoundException fileNotFoundException) {
                    fileNotFoundException.printStackTrace();
                }
                assert reader != null;
                BufferedReader bufferedReader = new BufferedReader(reader);

                String line = null;
                String pathLine = null;

                while (true) {
                    try {
                        if ((line = bufferedReader.readLine()) == null) break;
                        pathLine = line;
                    } catch (IOException ioException) {
                        ioException.printStackTrace();
                    }

                    try {
                        addComponentsReload(pathLine);

                    } catch (IOException ioException) {
                        ioException.printStackTrace();
                    }

                    System.out.println(line);
                }
                try {
                    reader.close();
                } catch (IOException ioException) {
                    ioException.printStackTrace();
                }
                reload_button.setEnabled(false);

            }
        });

        send_button.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

                textContent = textField.getText();

                Main.getImageFromClipboard();
                try {
                    Main.writeInToDocs(path_to_write, textContent);
                    textField.setText("");
                } catch (IOException | InvalidFormatException ioException) {
                    ioException.printStackTrace();
                }
            }
        });
        alwaysOnTopCheckBox.addItemListener(new ItemListener() {
            @Override
            public void itemStateChanged(ItemEvent e) {
                onTop = e.getStateChange() == 1;
                Main.frame.setAlwaysOnTop(onTop);
                Main.frame.revalidate();
            }
        });

        File f = new File(Main.txtFile);
        long length = f.length();
        if (f.exists() & length > 1) {
          reload();
        }
    }

    public void addComponents(String path) throws IOException {

        actionPanel.setLayout(new BoxLayout(actionPanel,BoxLayout.Y_AXIS));
        Box vBox = Box.createVerticalBox();

        label_path = new JLabel(path);
        label_path.setForeground(Color.BLUE);
        path_to_write = path;

        label_path.setAlignmentX(Component.LEFT_ALIGNMENT);
        textField = new JTextField();

        textField.setMaximumSize(new Dimension(400,25));
        textField.setMinimumSize(new Dimension(400,25));
        textField.setBackground(Color.GREEN);
        textField.setAlignmentX(Component.LEFT_ALIGNMENT);

        vBox.setAlignmentY(Component.CENTER_ALIGNMENT);

        vBox.add(label_path);

        vBox.add(textField);

        Box pBox = Box.createVerticalBox();
        pLabel = new JLabel();
        pLabel.setText(" ");
        pLabel.setSize(new Dimension(100,10));
        pLabel.setAlignmentX(Component.LEFT_ALIGNMENT);
        pBox.add(pLabel);

        remove_button = new JButton("Remove");
        remove_button.setAlignmentX(Component.LEFT_ALIGNMENT);

        remove_button.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {

                    removeTxtContent();

                } catch (IOException ex) {
                    ex.printStackTrace();
                }
            }
        });

        pBox.add(remove_button);

        Box hor = Box.createHorizontalBox();
        hor.add(vBox);
        hor.add(pBox);

        actionPanel.add(hor);

        // set Cursor and create a link to path to open up in Desktop
        label_path.setCursor(new Cursor(Cursor.HAND_CURSOR));
        label_path.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                try {

                    Desktop.getDesktop().open(
                            new File(new URI(path).getPath()));
                } catch (IOException | URISyntaxException e1) {

                    e1.printStackTrace();
                }
            }
        });
    }

    public void addComponentsReload(String path) throws IOException {

        actionPanel.setLayout(new BoxLayout(actionPanel,BoxLayout.Y_AXIS));
        Box vBox = Box.createVerticalBox();

        label_path = new JLabel(path);
        label_path.setForeground(Color.BLUE);
        path_to_write = path;

        label_path.setAlignmentX(Component.LEFT_ALIGNMENT);
        textField = new JTextField();

        textField.setMaximumSize(new Dimension(400,25));
        textField.setMinimumSize(new Dimension(400,25));
        textField.setBackground(Color.GREEN);
        textField.setAlignmentX(Component.LEFT_ALIGNMENT);

        vBox.setAlignmentY(Component.CENTER_ALIGNMENT);

        vBox.add(label_path);

        vBox.add(textField);

        Box pBox = Box.createVerticalBox();

        // set pseudo label for to align with other components
        pLabel = new JLabel();
        pLabel.setText(" ");
        pLabel.setSize(new Dimension(100,10));
        pLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        pBox.add(pLabel);

        remove_button = new JButton("Remove");
        remove_button.setAlignmentX(Component.LEFT_ALIGNMENT);

        remove_button.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    
                    removeTxtContent();

                   } catch (IOException ex) {
                    ex.printStackTrace();
                }

            }
        });


        pBox.add(remove_button);

        Box hor = Box.createHorizontalBox();
        hor.add(vBox);
        hor.add(pBox);

        actionPanel.add(hor);

        label_path.setCursor(new Cursor(Cursor.HAND_CURSOR));
        label_path.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                try {
                    Desktop.getDesktop().open(
                            new File(new URI(path).getPath()));
                } catch (IOException | URISyntaxException e1) {

                    e1.printStackTrace();
                }
            }
        });
    }

    public void fileWrite(String path) throws IOException {

        File file = new File(Main.txtFile);
        FileWriter writer = new FileWriter(file,true);
        String path1;
        if (path.contains(" ")){
            path1 = path.replaceAll(" ","%20");
            writer.write(path1+"\n");
        }else {
            writer.write(path+"\n");
        }
        writer.close();
    }

    public boolean readFromDocxList() throws IOException {

        FileReader reader = new FileReader(Main.txtFile);
        BufferedReader bufferedReader = new BufferedReader(reader);

        while ((bufferedReader.readLine()) != null) {

            return false;
        }
        reader.close();
        return true;
    }

    public void reload() throws IOException {

        reload_button.doClick();

    }

    public void removeTxtContent() throws IOException {
        PrintWriter pw = new PrintWriter(Main.txtFile);
        pw.close();
        actionPanel.removeAll();
        actionPanel.revalidate();
        actionPanel.repaint();
    }
}

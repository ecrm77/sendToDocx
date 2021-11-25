package sendToDocx;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.util.Units;
import javax.imageio.ImageIO;
import javax.swing.*;
import java.awt.*;
import java.awt.datatransfer.*;
import java.awt.datatransfer.Transferable;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.awt.image.BufferedImage;
import java.io.*;

class Main extends Form {

    public static JFrame frame;
    public static String txtFile;

    public Main() throws IOException {

    }

    public static void main(String[] args) throws IOException, InvalidFormatException {

        String dirUser = System.getProperty("user.home");
        String txtFileFolder = dirUser + "/sendToDocx";
        txtFile = dirUser + "/sendToDocx/docxList.txt";
        File dir = new File(txtFileFolder);

        // If directory doesn't exist create a new directory and file
        if (!dir.exists()) {
            dir.mkdirs();
            txtFile = dirUser + "/sendToDocx/docxList.txt";
            FileWriter writer = new FileWriter(txtFile);
            writer.close();
        }
        // Load JFrame
        frame = new JFrame();
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setContentPane(new Form().mainPanel);
        frame.setTitle("SendToDocx App");
        frame.setMinimumSize(new Dimension(600,200));
        frame.setMaximumSize(new Dimension(600,200));
        frame.setLocationRelativeTo(null);
        frame.setResizable(true);
        frame.setVisible(true);

    }

    public static void getImageFromClipboard() {
        Transferable transferable = Toolkit.getDefaultToolkit().getSystemClipboard().getContents(null);
        if (transferable != null && transferable.isDataFlavorSupported(DataFlavor.imageFlavor)) {
            try {
                String dirUser = System.getProperty("user.home");
                String path = dirUser + "/sendToDocx/clipboard.png";
                String format = "png";

                Image img = (Image) transferable.getTransferData(DataFlavor.imageFlavor);

                BufferedImage buffImg = convertToBufferedImage(img);
                ImageIO.write(buffImg, format, new File(path));

                } catch (UnsupportedFlavorException | IOException e) {
                // handle this as desired
                e.printStackTrace();
            }
        } else {
            System.err.println("That wasn't an image!");
        }

    }

    public static void writeInToDocs(String fileDocx, String textTitle) throws IOException, InvalidFormatException {

        String fd = fileDocx.replaceAll("%20", " ");

        Transferable transferable = Toolkit.getDefaultToolkit().getSystemClipboard().getContents(null);

        assert transferable != null;
        // If clipboard contains image
        if (transferable.isDataFlavorSupported(DataFlavor.imageFlavor)) {

            File f = new File(fd);
            XWPFDocument doc = new XWPFDocument(OPCPackage.open(new FileInputStream(f)));

            XWPFParagraph title = doc.createParagraph();
            XWPFRun run = title.createRun();
            run.setText(textTitle);
            run.setBold(true);
            title.setAlignment(ParagraphAlignment.LEFT);

            String dirUser = System.getProperty("user.home");
            String imgName = dirUser + "/sendToDocx/clipboard.png";

            Image img = ImageIO.read(new File(imgName));
            double height = img.getHeight(null);
            double width = img.getWidth(null);

            double scaling = 1.0;
            if (width > 450) {
                scaling = 450 / width;
            }

            FileInputStream is = new FileInputStream(imgName);
            run.addBreak();
            run.addPicture(is, XWPFDocument.PICTURE_TYPE_PNG, imgName, Units.toEMU(width * scaling), Units.toEMU(height * scaling));
            run.addBreak();
            is.close();

            // write word docx to file
            FileOutputStream fos = new FileOutputStream(fd);
            doc.write(fos);
            fos.close();

        }else {
            System.err.println("That wasn't an image!");
        }

        // If clipboard contains String
        if (transferable.isDataFlavorSupported(DataFlavor.stringFlavor)) {
            String s;
            try {
                s = (String) (transferable.getTransferData(
                        DataFlavor.stringFlavor));
            } catch (UnsupportedFlavorException | IOException ee) {
                s = ee.toString();
            }
            System.out.println(s);

            File f = new File(fd);
            XWPFDocument doc = new XWPFDocument(OPCPackage.open(new FileInputStream(f)));

            XWPFParagraph title = doc.createParagraph();
            XWPFRun run = title.createRun();

            run.setText(textTitle);

            run.addBreak();
            run.setText(s);
            run.addBreak();
            title.setAlignment(ParagraphAlignment.LEFT);

            // write word docx to file
            FileOutputStream fos = new FileOutputStream(fd);
            doc.write(fos);
            fos.close();

        } else {

            System.err.println("That wasn't a String!");

        }
    }

    public static BufferedImage convertToBufferedImage(Image image) {
        BufferedImage newImage = new BufferedImage(
                image.getWidth(null), image.getHeight(null),
                BufferedImage.TYPE_INT_RGB);
        Graphics2D g = newImage.createGraphics();
        g.drawImage(image, 0, 0, null);
        g.dispose();
        return newImage;
    }
}







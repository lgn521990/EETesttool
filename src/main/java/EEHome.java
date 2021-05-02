import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

public class EEHome {

    public static int count;
    //public static XSSFWorkbook wb;
    public static void main(String[] args) throws IOException {
        count=0;

       //AddImagesToExcelFile.initializeExcel();
        final JFrame frame = new JFrame("EasyEvidencing");
        //frame.getContentPane().;

        // Set's the window to be "always on top"
        frame.setAlwaysOnTop( true );
        JLabel l1=new JLabel("WorkBook");
        l1.setBounds(1, 1, 150, 20);
        JLabel l2=new JLabel("WorkSheet");
        l2.setBounds(1, 25, 150, 20);

        JButton btn=new JButton("Capture");
        btn.setBounds(50, 50, 200, 20);

        final JTextField field = new JTextField(1);
        field.setBounds(100, 1, 150, 20);
        final JTextField field2 = new JTextField(1);
        field2.setBounds(100, 25, 150, 20);
        frame.setLayout(null);
        frame.setLocationByPlatform( true );
        btn.addActionListener(new ActionListener(){
            public void actionPerformed(ActionEvent e){
                frame.setVisible(false);
                captureScreen(field.getText(),field2.getText());
                frame.setVisible(true);
            }
        });
        frame.add(l1);
        frame.add(l2);
        frame.add(btn);
        frame.add(field);
        frame.add(field2);
        frame.pack();
        frame.setVisible(true);
       // AddImagesToExcelFile.write();
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(310, 130);
    }


    public static void captureScreen(String filepath,String sheetName)
    {

        try {
            Robot r = new Robot();
            // It saves screenshot to desired path
            String path = "F:// Shot.jpg";
            // Used to get ScreenSize and capture image
            Rectangle capture =
                    new Rectangle(Toolkit.getDefaultToolkit().getScreenSize());
            BufferedImage Image = r.createScreenCapture(capture);
            ImageIO.write(Image, "jpg", new File(path));
            System.out.println("Screenshot saved");

           AddImagesToExcelFile.addImages(count,filepath,sheetName);
            count=count+19;
        }
        catch (Exception ex) {
            System.out.println(ex);
        }
    }
}

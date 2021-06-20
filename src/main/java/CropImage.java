import javax.imageio.ImageIO;
import javax.swing.*;
import java.awt.*;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.awt.event.MouseMotionListener;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;


public class CropImage extends JFrame implements MouseListener, MouseMotionListener {
    int drag_status = 0, c1, c2, c3, c4;
    public static int count;
    public static String filepath;
    public static String sheetName;


    public void start(String filepath,String sheetName) throws AWTException, IOException {

        this.filepath=filepath;
        this.sheetName=sheetName;

        Robot r = new Robot();
        // It saves screenshot to desired path
        //String path = "Shot.jpg";
        // Used to get ScreenSize and capture image
        Rectangle capture =
                new Rectangle(Toolkit.getDefaultToolkit().getScreenSize());
        BufferedImage Image = r.createScreenCapture(capture);
        //ImageIO.write(Image, "jpg", new File(path));

        ImagePanel im = new ImagePanel(Image);//new ImagePanel("C:\\Users\\lucky\\Pictures\\img1.jpg");
        add(im);
        setAlwaysOnTop(true);
        //(new File(img)).delete();
        setSize(im.getWidth(),im.getHeight());
        setVisible(true);
        addMouseListener(this);
        addMouseMotionListener(this);
    }


    public void mouseClicked(MouseEvent arg0) {
    }

    public void mouseEntered(MouseEvent arg0) {
    }

    public void mouseExited(MouseEvent arg0) {
    }

    public void mousePressed(MouseEvent arg0) {
        repaint();
        c1 = arg0.getX();
        c2 = arg0.getY();
    }

    public void mouseReleased(MouseEvent arg0) {
        repaint();
        if (drag_status == 1) {
            c3 = arg0.getX();
            c4 = arg0.getY();
            try {
                draggedScreen();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

    }

    public void mouseDragged(MouseEvent arg0) {
        repaint();
        drag_status = 1;
        c3 = arg0.getX();
        c4 = arg0.getY();
    }

    public void mouseMoved(MouseEvent arg0) {
    }

    public void paint(Graphics g) {
        super.paint(g);
        int w = c1 - c3;
        int h = c2 - c4;
        w = w * -1;
        h = h * -1;
        if (w < 0)
            w = w * -1;
        g.drawRect(c1, c2, w, h);
    }

    public void draggedScreen()throws Exception
    {
        int w = c1 - c3;
        int h = c2 - c4;
        w = w * -1;
        h = h * -1;
        Robot robot = new Robot();
        BufferedImage img = robot.createScreenCapture(new Rectangle(c1, c2,w,h));
        dispose();

        File save_path=new File("CroppedShot.jpg");
        ImageIO.write(img, "JPG", save_path);

        System.out.println("Cropped image saved successfully.");
        AddImagesToExcelFile.addImages(count,filepath,sheetName);

        count=count+19;
    }

}
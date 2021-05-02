import java.io.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AddImagesToExcelFile {

    public static Workbook wb;
    public static Sheet sheet;
    public static boolean flag;
    public static CreationHelper helper;
    public static void initializeExcel(String filePath,String sheetName) throws IOException {
        // create a new workbook


        if(flag)
        {
            wb = new XSSFWorkbook();
            sheet=wb.createSheet(sheetName);
            EEHome.count=0;
            flag=false;
        }else{
            wb = new XSSFWorkbook(new FileInputStream(filePath));
            // create sheet
            if(wb.getSheetIndex(sheetName)==-1) {
                sheet = wb.createSheet(sheetName);
                EEHome.count=0;
            }
            else
                sheet=wb.getSheet(sheetName);
        }

         helper = wb.getCreationHelper();

    }
    public static void addImages(int count,String filePath,String SheetName) throws IOException {
        // create a new workbook
    //    Workbook wb = new XSSFWorkbook(); // or new HSSFWorkbook();

        // add jpg/jpeg and png formats picture data to this workbook.
        if(!(new File(filePath)).exists()) {
            flag=true;
            System.out.println("prior reading screenshot");
            EEHome.count=0;
        }

        initializeExcel(filePath,SheetName);
        count=(EEHome.count==0)?0:count;

        InputStream is = new FileInputStream("F:// Shot.jpg");
        byte[] bytes = IOUtils.toByteArray(is);
        int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
        is.close();

        // Create the drawing patriarch. This is the top level container for all shapes.
        Drawing<?> drawing = sheet.createDrawingPatriarch();

        // add a picture shape
        ClientAnchor anchor = helper.createClientAnchor();

        // set top-left corner of the picture,
        // subsequent call of Picture#resize() will operate relative to it
        anchor.setCol1(2);
        anchor.setRow1(count+2);
        Picture pict = drawing.createPicture(anchor, pictureIdx);

        // auto-size picture relative to its top-left corner
        pict.resize(15);

        System.out.println("writing to file");
        // save workbook
        String fileName = filePath;
        OutputStream fileOut = new FileOutputStream(new File(fileName));
        wb.write(fileOut);
        wb.close();
    }

    public static void write() throws IOException {
        String fileName = "F://excel_image.xls";

        if (wb instanceof XSSFWorkbook) {
            fileName += "x";
        }

        OutputStream fileOut = new FileOutputStream(new File(fileName));
        wb.write(fileOut);

        wb.close();
    }

}
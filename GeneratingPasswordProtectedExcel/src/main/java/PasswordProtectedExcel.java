import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class PasswordProtectedExcel {
    public static void main(String[] args) throws IOException {
        //Generating excel file with data
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("User Info");
        //Data for excel
        Object[][] bookData = {
                {"Name", "Department", "email"},
                {"Aman Ullah", "EEE", "aman@stsvinc.com"},
                {"Ismail Hossain", "EEE", "ismailhossain@example.com"},
                {"Alam al soud Ratul", "CSE", "alamalsaud@example.com"},
                {"Md. Nasir uddin", "CSE", "nasir@stsvinc.com"},
        };

        //Adding the data to the excel worksheet
        int rowCount = 0;
        for (Object[] aBook : bookData) {
            Row row = sheet.createRow(rowCount++);

            int columnCount = 0;

            for (Object field : aBook) {
                Cell cell = row.createCell(columnCount++);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
        }

        //Save the excel file in a directory
        try (FileOutputStream outputStream = new FileOutputStream("file.xlsx")) {
            workbook.write(outputStream);
        }

        //Protecting the file with password
        try(POIFSFileSystem fs=new POIFSFileSystem()) {
            EncryptionInfo info=new EncryptionInfo(EncryptionMode.agile);
            Encryptor enc=info.getEncryptor();
            enc.confirmPassword("satl");
            try(OPCPackage opc=OPCPackage.open(new File("file.xlsx"), PackageAccess.READ_WRITE);
                OutputStream os=enc.getDataStream(fs)) {
                opc.save(os);
            }
            try (FileOutputStream fos=new FileOutputStream("file.xlsx")){
                fs.writeFilesystem(fos);
            }
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}

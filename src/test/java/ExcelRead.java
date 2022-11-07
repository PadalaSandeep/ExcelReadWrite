import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;

public class ExcelRead {

    public static void main(String[] args) throws IOException {
       getTestCaseData("SampleExcel.xlsx", "testData", "tcs001");
    }

    public static HashMap<String, Object> getTestCaseData(String fileName, String sheetName, String testCaseName) throws IOException {
        File myFile = new File(fileName);
        FileInputStream fis = new FileInputStream(myFile);
        XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
        XSSFSheet mySheet = myWorkBook.getSheet(sheetName);
        HashMap<String, Object> testCaseData = new HashMap<>();
        try {
            for (int i=1; i<=mySheet.getLastRowNum(); i++){

                if(mySheet.getRow(i).getCell(0).getStringCellValue().contentEquals(testCaseName)){
                    for (int j=0; j<mySheet.getRow(0).getLastCellNum(); j++){

                        switch (mySheet.getRow(i).getCell(j).getCellType()) {
                            case STRING:
                                testCaseData.put(mySheet.getRow(0).getCell(j).getStringCellValue(), mySheet.getRow(i).getCell(j).getStringCellValue());
                                break;
                            case NUMERIC:
                                testCaseData.put(mySheet.getRow(0).getCell(j).getStringCellValue(), mySheet.getRow(i).getCell(j).getNumericCellValue());
                                break;
                            case BOOLEAN:
                                testCaseData.put(mySheet.getRow(0).getCell(j).getStringCellValue(), mySheet.getRow(i).getCell(j).getBooleanCellValue());
                                break;
                            default :
                        }
                    }
                    break;
                }
            }

        }
        catch (Exception e){
            System.out.println("Something went wrong in reading data, please check your file");
        }
        finally {
            myWorkBook.close();
            fis.close();
        }
        System.out.println("Test data being returned is "+ testCaseData);
       return testCaseData;
    }
}

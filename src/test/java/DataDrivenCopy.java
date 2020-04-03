import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.plaf.basic.BasicInternalFrameTitlePane;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class DataDrivenCopy {

    //Identify Testcases coloum by scanning the entire 1st row
//once coloumn is identified then scan entire testcase coloum to identify purcjhase testcase row
//after you grab purchase testcase row = pull all the data of that row and feed into test


    public static void main(String[] args) throws IOException {
        //scan the rows and find the testcase header row

        //get the desired sheet where we want to read data
        // declare column
        int k = 0;
        int column = 0;
        FileInputStream fis = new FileInputStream ("/Users/administrator/ExcelDriven/src/main/resources/DemoData1.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook (fis);
        int sheets = workbook.getNumberOfSheets ();
        System.out.println ("number of sheets" + sheets);
        ArrayList<String> a = new ArrayList<String> ();


        for (int i = 0; i < sheets; i++) {
            String sheetName = workbook.getSheetName (i);
            System.out.println ("sheetname" + sheetName);

            if (sheetName.equalsIgnoreCase ("TestData")) {
                XSSFSheet sheet = workbook.getSheetAt (i);
                Iterator<Row> rows = sheet.iterator (); //sheet is collection of rows
                Row firstrow = rows.next (); //row is collection of cells
                Iterator<Cell> ce = firstrow.cellIterator (); //define cell iterator

                while (ce.hasNext ()) {
                    if (ce.next ().getStringCellValue ().equalsIgnoreCase ("Testcases"))
                    //get the row number of the testcase
                    {
                        column = k;
                        System.out.println ("column is " + column);
                    }

                    k++;
                }

                    //System.out.println ("column is "+column);


                    while (rows.hasNext ()) {

                        Row r = rows.next ();

                        if (r.getCell (column).getStringCellValue ().equalsIgnoreCase ("Trunkshow")) {

                            ////after you grab purchase testcase row = pull all the data of that row and feed into test

                            Iterator<Cell> cv = r.cellIterator ();

                            while (cv.hasNext ()) {
                                Cell c = cv.next ();

                                if (c.getCellType () == Cell.CELL_TYPE_STRING) {

                                    a.add (c.getStringCellValue ());
                                } else {

                                    a.add (NumberToTextConverter.toText (c.getNumericCellValue ()));


                                }

                            }
                            System.out.println (a);
                        }
                    }

                }

            }
        }
    }


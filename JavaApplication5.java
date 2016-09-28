/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package javaapplication5;

import com.aspose.cells.Chart;
import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import java.io.IOException;

/**
 *
 * @author tummyz
 */
public class JavaApplication5 {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws Exception {
        // TODO code application logic here\
        Workbook workbook = new Workbook("/Users/tummyz/Desktop/aaa.xls"); 
        WorksheetCollection worksheets = workbook.getWorksheets(); 
        int sheets=worksheets.getCount(); 
        for(int i = 0; i < sheets; i++) 
            { 
      ImageOrPrintOptions options = new ImageOrPrintOptions(); 
      Worksheet sheet = worksheets.get(i); 
      options.setAllColumnsInOnePagePerSheet(true); 
      options.setImageFormat(ImageFormat.getJpeg()); 
      options.setHorizontalResolution(1500);
      options.setVerticalResolution(1500);
      SheetRender sr = new SheetRender(sheet, options); 
      System.out.println("Name:"+sheet.getName()); 

      try 
{ 
                for(int j = 0; j < sr.getPageCount(); j++)
				{
					sr.toImage(j, "/Users/tummyz/Desktop/imagexls/excel"+ sheet.getName() + "_" + j + ".jpg");
				} 	
} 
catch(Exception e) 
{ 
e.printStackTrace(); 
} 


    }
    


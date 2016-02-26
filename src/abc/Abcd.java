/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package abc;

import java.awt.*;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.NumberFormat;
import java.util.*;
import java.util.List;
import java.util.Map.Entry;
import java.util.logging.Logger;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.renderer.category.StandardBarPainter;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.labels.StandardCategoryItemLabelGenerator;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.chart.renderer.category.CategoryItemRenderer;
import org.jfree.chart.renderer.category.StackedBarRenderer;

/**
 *
 * @author Zafack Billy
 */
public class Abcd {
    static XSSFRow row;
    public String fichier_source;
    public String chemin_destination;
    public String fichier_destination;
    public String fichier_synthese;
    public String annee;
    public static final int NB_USINE = 6;
    public static int num_sheets = 1;
    
    
    
    public Abcd(String source, String dossier_destination, String fichier_destination, String an, String synthese){
        this.fichier_source = source;
        this.chemin_destination = dossier_destination;
        this.fichier_destination = fichier_destination;
        this.annee = an;
        this.fichier_synthese = synthese;
    }
    
    public void writeSheet(Map < String, Object[] > empinfo, String name) throws FileNotFoundException, IOException{
         //Create blank workbook
      XSSFWorkbook workbook = new XSSFWorkbook(); 
      //Create a blank sheet
      XSSFSheet spreadsheet = workbook.createSheet(" Employee Info ");
      //Create row object
      XSSFRow row;
      //This data needs to be written (Object[])
//      Map < String, Object[] > empinfo = 
//      new TreeMap < String, Object[] >();
//      empinfo.put( "1", new Object[] { 
//      "EMP ID", "EMP NAME", "DESIGNATION" });
//      empinfo.put( "2", new Object[] { 
//      "tp01", "Gopal", "Technical Manager" });
//      empinfo.put( "3", new Object[] { 
//      "tp02", "Manisha", "Proof Reader" });
//      empinfo.put( "4", new Object[] { 
//      "tp03", "Masthan", "Technical Writer" });
//      empinfo.put( "5", new Object[] { 
//      "tp04", "Satish", "Technical Writer" });
//      empinfo.put( "6", new Object[] { 
//      "tp05", "Krishna", "Technical Writer" });
      //Iterate over data and write to sheet
      Set < String > keyid = empinfo.keySet();
      int rowid = 0;
      for (String key : keyid) //__
      {
         row = spreadsheet.createRow(rowid++);
         Object [] objectArr = empinfo.get(key);
         int cellid = 0;
         for (Object obj : objectArr)
         {
            Cell cell = row.createCell(cellid++);
            cell.setCellValue(obj.toString());
         }
      }
      //Write the workbook in file system
      FileOutputStream out = new FileOutputStream(new File(name)); //+".xlsx"
      workbook.write(out);
      out.close();
      System.out.println(name+".xlsx written successfully" );
    }
    
    public boolean testPresence(String name, String jr) throws IOException{
        FileInputStream fis = new FileInputStream(new File(name)); //+".xlsx"
        boolean test = true;
        switch(jr){
          case "LUNDI":
              if(isCellContentNull("H5", 1, name)){
                  test = false;
              } 
          break;
          case "MARDI":
              if(isCellContentNull("M5", 1, name)){
                  test = false;
              } 
          break;
          case "MERCREDI":
              if(isCellContentNull("R5", 1, name)){
                  test = false;
              } 
          break;
          case "JEUDI":
              if(isCellContentNull("W5", 1, name)){
                  test = false;
              } 
          break;
          case "VENDREDI":
              if(isCellContentNull("AB5", 1, name)){
                  test = false;
              } 
          break;
          case "SAMEDI":
              if(isCellContentNull("AG5", 1, name)){
                  test = false;
              } 
          break;
          case "DIMANCHE":
             if(isCellContentNull("AL5", 1, name)){
                  test = false;
              } 
          break;
             default:
             test = true;
      } 
      fis.close();
      return test;
    }

    public void writeFormatALLSheet(ArrayList<Map < String, Object[] >> empinfo, String name) throws FileNotFoundException, IOException{
         //Create blank workbook
      XSSFWorkbook workbook = null;
      if(fichier_synthese == null){
      workbook = new XSSFWorkbook(); 
      }else{
      workbook = new XSSFWorkbook(new FileInputStream(new File(fichier_synthese))); 
      }  
      //Create a blank sheet
      
      for(int i=0; i< empinfo.size()-2; i++){
          if(i < empinfo.size()-3){
              XSSFSheet spreadsheet = null;
              if(workbook.getSheet(empinfo.get(i).get("9")[0].toString().replace("/", "-")+"-"+this.annee)==null){
                spreadsheet = workbook.createSheet(empinfo.get(i).get("9")[0].toString().replace("/", "-")+"-"+this.annee);
              }else{
                //spreadsheet = workbook.getSheet(empinfo.get(i).get("9")[0].toString().replace("/", "-")+"-"+this.annee);
                continue;
              }
      
      //Create row object
      XSSFRow row; 
//      Vector<String> factories = new Vector<String>();
//      Vector<Double> hps = new Vector<Double>();
      //////////Mise En Forme
      Set < String > keyid = empinfo.get(i).keySet();
      int rowid = 0;
      Map<String, Integer> values = new HashMap<String, Integer>();
      for (String key : keyid)
      {
         row = spreadsheet.createRow(rowid++);
         Object [] objectArr = empinfo.get(i).get(key);
         int cellid = 0; 
             String aa = null;
             Integer bb = null;
         for (Object obj : objectArr){
            Cell cell = row.createCell(cellid++);
            if(cell.getColumnIndex() == 2){
              cell.setCellValue(obj.toString());
            }
            else if((cell.getColumnIndex() == 1 || cell.getColumnIndex() == 3) && cell.getRowIndex()>1){
                double a = Double.parseDouble(obj.toString());
                if(cell.getColumnIndex() == 1 && cell.getRowIndex()>1){ 
                    //hps.add(a);
                    bb = Math.round((float)a);
                }
              cell.setCellValue(a);  
            }
            else{
                String x = obj.toString();
                if(cell.getColumnIndex() == 0 && cell.getRowIndex()>1  && cell.getRowIndex()< 2+NB_USINE){ //////----------------//////////////// CECI FIXE LE NOMBRE D'USINE(PAS BON)
                    //factories.add(x);
                    aa = x;
                }
                cell.setCellValue(x);
            }
            if(aa!=null && bb!=null){
                values.put(aa, bb);
            }    
         }
      }
      
      spreadsheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 4));
      
      CellReference cr_Y;
      Cell cell2;  
      colorHeader("B1", HSSFColor.GREEN.index, workbook, spreadsheet);
      for(int j=3;j<=8;j++){
          style(11,"B"+j,workbook,spreadsheet);
      }
      for(int j=3;j<=8;j++){
          style(11,"C"+j,workbook,spreadsheet);
      }
      for(int j=3;j<=8;j++){
          style(11,"D"+j,workbook,spreadsheet);
      }
      for(int j=3;j<=8;j++){
          style(12,"E"+j,workbook,spreadsheet);
      }
      
      style(3,"B2",workbook,spreadsheet);  
      style(4,"C2",workbook,spreadsheet); 
      style(4,"D2",workbook,spreadsheet);  
      style(6,"A3",workbook,spreadsheet);  
      style(7,"A4",workbook,spreadsheet);  
      style(7,"A5",workbook,spreadsheet);  
      style(7,"A6",workbook,spreadsheet);  
      style(7,"A7",workbook,spreadsheet);  
      style(10,"A8",workbook,spreadsheet); 
      style(8,"B8",workbook,spreadsheet);  
      style(8,"C8",workbook,spreadsheet);  
      style(8,"D8",workbook,spreadsheet);  
      style(7,"E3",workbook,spreadsheet);  
      style(7,"E4",workbook,spreadsheet);  
      style(7,"E5",workbook,spreadsheet);  
      style(7,"E6",workbook,spreadsheet);  
      style(7,"E7",workbook,spreadsheet);
      style(5,"E2",workbook,spreadsheet);  
      spreadsheet.setColumnWidth(0,5000);  
      spreadsheet.setColumnWidth(1,4000);  
      spreadsheet.setColumnWidth(2,4000);  
      spreadsheet.setColumnWidth(3,4000);  
      spreadsheet.setColumnWidth(4,8000);  
      
      cr_Y = new CellReference("A9");
      row = spreadsheet.getRow(cr_Y.getRow());
      cell2 = row.getCell(cr_Y.getCol()); 
      cell2.setCellValue("");
      drawChart(workbook, spreadsheet, 9, 0, 1, values);
      
          }else{ //////////////?????????
              Map<String, Integer> values0 = new HashMap<String, Integer>();
         Map<String, Integer> values1 = new HashMap<String, Integer>();
         Map<String, Integer> values2 = new HashMap<String, Integer>();
             // System.out.println("AM HEREEEEEEEEEEEEEEEEEEEE");
               XSSFSheet spreadsheet = null;
              if(workbook.getSheet(empinfo.get(i).get("9")[0].toString().replace("/", "-")+"-"+this.annee)==null){
                spreadsheet = workbook.createSheet(empinfo.get(i).get("9")[0].toString().replace("/", "-")+"-"+this.annee);
              }else{
                spreadsheet = workbook.getSheet(empinfo.get(i).get("9")[0].toString().replace("/", "-")+"-"+this.annee);
              }      //Create row object
      XSSFRow row; 
      Vector<String> factories = new Vector<String>();
      Vector<Double> hps = new Vector<Double>();
      
      Vector<String> factories1 = new Vector<String>();
      Vector<Double> hps1 = new Vector<Double>();
      
      Vector<String> factories2 = new Vector<String>();
      Vector<Double> hps2 = new Vector<Double>();
      //////////Mise En Forme
      Set < String > keyid = empinfo.get(i).keySet();
      int rowid = 0;
      for (String key : keyid){
         row = spreadsheet.createRow(rowid++);
         Object [] objectArr = empinfo.get(i).get(key);
         Object [] objectArr_hebdo = empinfo.get(empinfo.size()-2).get(key);
         Object [] objectArr_mensuel = empinfo.get(empinfo.size()-1).get(key);
         int cellid = 0;
         int cellid_h = 7;
         int cellid_m = 14;
         int k=0;
         
         String aa = null;
         Integer bb = null;
         String aa1 = null;
         Integer bb1 = null;
         String aa2 = null;
         Integer bb2 = null;
         for (Object obj : objectArr){
            Cell cell = row.createCell(cellid++);
            if(cell.getColumnIndex() == 2){
              cell.setCellValue(obj.toString());
            }
            else if((cell.getColumnIndex() == 1 || cell.getColumnIndex() == 3) && cell.getRowIndex()>1){
                double a = Double.parseDouble(obj.toString());
                if(cell.getColumnIndex() == 1 && cell.getRowIndex()>1){ 
                    //hps.add(a);
                    bb = Math.round((float)a); 
                }
              cell.setCellValue(a);  
            }
            else{
                String x = obj.toString();
                if(cell.getColumnIndex() == 0 && cell.getRowIndex()>1  && cell.getRowIndex()< 2+NB_USINE){ //////----------------//////////////// CECI FIXE LE NOMBRE D'USINE(PAS BON)
                    //factories.add(x);
                    aa = x;
                }
                cell.setCellValue(x);
            }
            if(aa!=null && bb!=null){
                values0.put(aa, bb);
            }  
            ////////////////////////
            /////////////////////// 
            ///////////////////////
            Cell cell1 = row.createCell(cellid_h++);
            if(cell1.getColumnIndex() == 9){
              cell1.setCellValue(objectArr_hebdo[k].toString());
            }
            else if((cell1.getColumnIndex() == 8 || cell1.getColumnIndex() == 10) && cell1.getRowIndex()>1){
                double a = Double.parseDouble(objectArr_hebdo[k].toString());
                if(cell1.getColumnIndex() == 8 && cell1.getRowIndex()>1){
                    //hps1.add(a);
                    bb1 = Math.round((float)a);
                }
              cell1.setCellValue(a);  
            }
            else{
                String x = objectArr_hebdo[k].toString();
                if(cell1.getColumnIndex() == 7 && cell1.getRowIndex()>1  && cell1.getRowIndex()< 2+NB_USINE){ //////----------------//////////////// CECI FIXE LE NOMBRE D'USINE(PAS BON)
                    //factories1.add(x);
                    aa1 = x;
                }
                cell1.setCellValue(x);
            }
            if(aa1!=null && bb1!=null){
                values1.put(aa1, bb1);
            } 
            /////////////////////////
            //////////////////////////
            ////////////////////////
            Cell cell2 = row.createCell(cellid_m++);
            if(cell2.getColumnIndex() == 16){
              cell2.setCellValue(objectArr_mensuel[k].toString());
            }
            else if((cell2.getColumnIndex() == 15 || cell2.getColumnIndex() == 17) && cell2.getRowIndex()>1){
                double a = Double.parseDouble(objectArr_mensuel[k].toString());
                if(cell2.getColumnIndex() == 15 && cell2.getRowIndex()>1){
                    //hps2.add(a);
                    bb2 = Math.round((float)a);
                }
              cell2.setCellValue(a);  
            }
            else{
                String x = objectArr_mensuel[k].toString();
                if(cell2.getColumnIndex() == 14 && cell2.getRowIndex()>1  && cell2.getRowIndex()< 2+NB_USINE){ //////----------------//////////////// CECI FIXE LE NOMBRE D'USINE(PAS BON)
                    //factories2.add(x);
                    aa2 = x;
                }
                cell2.setCellValue(x);
            }
            if(aa2!=null && bb2!=null){
                values2.put(aa2, bb2);
            } 
            k++;    
         }
      }
      
      spreadsheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 4));
      spreadsheet.addMergedRegion(new CellRangeAddress(0, 0, 8, 11));
      spreadsheet.addMergedRegion(new CellRangeAddress(0, 0, 15, 18));
       

      CellReference cr_Y;
      Cell cell2;  
      colorHeader("B1", HSSFColor.GREEN.index, workbook, spreadsheet);
      colorHeader("I1", HSSFColor.GREEN.index, workbook, spreadsheet);
      colorHeader("P1", HSSFColor.GREEN.index, workbook, spreadsheet);
     
      for(int j=3;j<=8;j++){
          style(11,"B"+j,workbook,spreadsheet);
          style(11,"I"+j,workbook,spreadsheet);
          style(11,"P"+j,workbook,spreadsheet);
      }
      for(int j=3;j<=8;j++){
          style(11,"C"+j,workbook,spreadsheet); 
          style(11,"J"+j,workbook,spreadsheet);
          style(11,"R"+j,workbook,spreadsheet);
      }
      for(int j=3;j<=8;j++){
          style(11,"D"+j,workbook,spreadsheet);
          style(11,"K"+j,workbook,spreadsheet);
          style(11,"Q"+j,workbook,spreadsheet);
      }
      for(int j=3;j<=8;j++){
          style(11,"E"+j,workbook,spreadsheet);

          style(12,"F"+j,workbook,spreadsheet); //*_*

          style(11,"L"+j,workbook,spreadsheet);

          style(12,"M"+j,workbook,spreadsheet); //*_*

          style(11,"S"+j,workbook,spreadsheet);

          style(12,"T"+j,workbook,spreadsheet); //*_*
      }
      
      style(3,"B2",workbook,spreadsheet);
      style(3,"I2",workbook,spreadsheet);
      style(3,"P2",workbook,spreadsheet);
      
      style(4,"C2",workbook,spreadsheet);
      style(4,"J2",workbook,spreadsheet);
      style(4,"R2",workbook,spreadsheet);
      
      style(4,"D2",workbook,spreadsheet);
              style(4,"E2",workbook,spreadsheet);
      style(4,"K2",workbook,spreadsheet);
              style(4,"L2",workbook,spreadsheet);
      style(4,"Q2",workbook,spreadsheet);
              style(4,"S2",workbook,spreadsheet);
      
      style(6,"A3",workbook,spreadsheet);
      style(6,"H3",workbook,spreadsheet);
      style(6,"O3",workbook,spreadsheet);
      
      style(7,"A4",workbook,spreadsheet);
      style(7,"H4",workbook,spreadsheet);
      style(7,"O4",workbook,spreadsheet);
      
      style(7,"A5",workbook,spreadsheet);
      style(7,"H5",workbook,spreadsheet);
      style(7,"O5",workbook,spreadsheet);
      
      style(7,"A6",workbook,spreadsheet);
      style(7,"H6",workbook,spreadsheet);
      style(7,"O6",workbook,spreadsheet);
      
      style(7,"A7",workbook,spreadsheet);
      style(7,"H7",workbook,spreadsheet);
      style(7,"O7",workbook,spreadsheet);
      
      style(10,"A8",workbook,spreadsheet);
      style(10,"H8",workbook,spreadsheet);
      style(10,"O8",workbook,spreadsheet);
      
      style(8,"B8",workbook,spreadsheet);
      style(8,"I8",workbook,spreadsheet);
      style(8,"P8",workbook,spreadsheet);
      
      style(8,"C8",workbook,spreadsheet);
      style(8,"J8",workbook,spreadsheet);
      style(8,"R8",workbook,spreadsheet);
      
      style(8,"D8",workbook,spreadsheet);
      style(8,"K8",workbook,spreadsheet);
      style(8,"Q8",workbook,spreadsheet);
      
      style(7,"F3",workbook,spreadsheet);
      style(7,"M3",workbook,spreadsheet);
      style(7,"T3",workbook,spreadsheet);
      
      style(7,"F4",workbook,spreadsheet);
      style(7,"M4",workbook,spreadsheet);
      style(7,"T4",workbook,spreadsheet);
      
      style(7,"F5",workbook,spreadsheet);
      style(7,"M5",workbook,spreadsheet);
      style(7,"T5",workbook,spreadsheet);
      
      style(7,"F6",workbook,spreadsheet);
      style(7,"M6",workbook,spreadsheet);
      style(7,"T6",workbook,spreadsheet);
      
      style(7,"F7",workbook,spreadsheet);
      style(7,"M7",workbook,spreadsheet);
      style(7,"T7",workbook,spreadsheet);
      
      style(10,"F8",workbook,spreadsheet);
      style(10,"M8",workbook,spreadsheet);
      style(10,"T8",workbook,spreadsheet);
        
      style(5,"F2",workbook,spreadsheet);
      style(5,"M2",workbook,spreadsheet);
      style(5,"T2",workbook,spreadsheet);

              style(8,"E8",workbook,spreadsheet);
              style(8,"S8",workbook,spreadsheet);
              style(8,"L8",workbook,spreadsheet);
      spreadsheet.setColumnWidth(0, 5000);
      spreadsheet.setColumnWidth(6,5000);
      spreadsheet.setColumnWidth(12,5000);
      
      spreadsheet.setColumnWidth(1,4000);
      spreadsheet.setColumnWidth(7,4000);
      spreadsheet.setColumnWidth(13,4000);
      
      spreadsheet.setColumnWidth(2,4000);
      spreadsheet.setColumnWidth(8,4000);
      spreadsheet.setColumnWidth(14,4000);

              ///////////////////
      spreadsheet.setColumnWidth(10,4000);
      
      spreadsheet.setColumnWidth(3,4000);
      spreadsheet.setColumnWidth(9,4000);
      spreadsheet.setColumnWidth(15,4000);
      
      spreadsheet.setColumnWidth(4,8000);
      spreadsheet.setColumnWidth(11,8000);
      spreadsheet.setColumnWidth(18,8000);

      spreadsheet.setColumnWidth(5,8000);
      spreadsheet.setColumnWidth(12,8000);
      spreadsheet.setColumnWidth(19,8000);

              spreadsheet.setColumnWidth(16,4000);
              spreadsheet.setColumnWidth(17,4000);
      
      cr_Y = new CellReference("A9");
      row = spreadsheet.getRow(cr_Y.getRow());
      cell2 = row.getCell(cr_Y.getCol()); 
      cell2.setCellValue("");
      drawChart(workbook, spreadsheet, 9, 0, 1, values0);
      drawChart(workbook, spreadsheet, 9, 7, 2, values1);
      drawChart(workbook, spreadsheet, 9, 14, 3, values2);
      
       ////////////////
       ////////////////
       ///////////////
          }
      }
      //Write the workbook in file system
      FileOutputStream out = new FileOutputStream(new File(name)); //+".xlsx"
      workbook.write(out);
      out.close();
        
        FileInputStream fi = new FileInputStream(new File(name));
        XSSFWorkbook workbk = new XSSFWorkbook(fi);
        num_sheets = workbk.getNumberOfSheets();
        fi.close();
      
      System.out.println(name+".xlsx written successfully" );
    }
    
    public Object[] traitement(String ville, Map<String, List<String>> identification, int numero_sheet, FileInputStream fis, String jour, String attr) throws FileNotFoundException, IOException{
      XSSFWorkbook workbook = new XSSFWorkbook(fis);
      XSSFSheet spreadsheet = workbook.getSheetAt(numero_sheet);
      Object[] obj_Y = null; 
      CellReference cr_Y = new CellReference(attr+"5");
      row = spreadsheet.getRow(cr_Y.getRow());
      Cell cell_Y = row.getCell(cr_Y.getCol());
      Double max_Y = cell_Y.getNumericCellValue();
      int indice =5;
      for(int i=1; i<identification.get(ville).size();i++){
           cr_Y = new CellReference(attr+(5+i));
           row = spreadsheet.getRow(cr_Y.getRow());
           cell_Y = row.getCell(cr_Y.getCol());
           if(cell_Y.getNumericCellValue()>max_Y){
               max_Y = cell_Y.getNumericCellValue();
               indice = 5+i;
           }
      }
      long max_YY = Math.round(max_Y);
      cr_Y = new CellReference("A"+indice);
      row = spreadsheet.getRow(cr_Y.getRow());
      cell_Y = row.getCell(cr_Y.getCol());
      String chaine_Y = cell_Y.getStringCellValue();
      //Le total
      cr_Y = new CellReference(attr+(5+identification.get(ville).size()));
      row = spreadsheet.getRow(cr_Y.getRow());
      cell_Y = row.getCell(cr_Y.getCol());
      Long total = Math.round(cell_Y.getNumericCellValue());
      obj_Y = new Object[]{ville, total, chaine_Y, max_YY, "", ""};  ///////////////++++++++++  *_*
      return obj_Y;
    }
    
    public String getStringCellContent(String coor, int sheet_number, FileInputStream fis) throws IOException{
      XSSFWorkbook workbook = new XSSFWorkbook(fis);
      XSSFSheet spreadsheet = workbook.getSheetAt(sheet_number);
      CellReference cr_Y = new CellReference(coor);
      row = spreadsheet.getRow(cr_Y.getRow());
      Cell cell_Y = row.getCell(cr_Y.getCol());
      String total = cell_Y.getStringCellValue();
      return total;
    }
    
    public boolean isCellContentNull(String coor, int sheet_number, String fs) throws IOException{
        FileInputStream fis = new FileInputStream(new File(fs));
      XSSFWorkbook workbook = new XSSFWorkbook(fis);
      XSSFSheet spreadsheet = workbook.getSheetAt(sheet_number);
      CellReference cr_Y = new CellReference(coor);
      row = spreadsheet.getRow(cr_Y.getRow());
      Cell cell_Y = row.getCell(cr_Y.getCol());
     if (cell_Y == null || cell_Y.getCellType() == Cell.CELL_TYPE_BLANK) {
    // This cell is empty
         return true;
        }else{
          return false;
      } 
         
      
    }
    
    public Long getNumericCellContent(String coor, int sheet_number, FileInputStream fis) throws IOException{
      XSSFWorkbook workbook = new XSSFWorkbook(fis);
      XSSFSheet spreadsheet = workbook.getSheetAt(sheet_number);
      CellReference cr_Y = new CellReference(coor);
      row = spreadsheet.getRow(cr_Y.getRow());
      Cell cell_Y = row.getCell(cr_Y.getCol());
      Long total = Math.round(cell_Y.getNumericCellValue());
      return total;
    }
    
    public Map < String, Object[] > readVillesHP(String name, String jour, String attr) throws IOException{
     FileInputStream fis = new FileInputStream(new File(name)); //+".xlsx"

      //Iterator < Row > rowIterator = spreadsheet.iterator();
      Object[] obj_Y = null;
      Object[] obj_K = null;
      Object[] obj_N = null;
      Object[] obj_B = null;
      Object[] obj_G = null;
      Object[] obj_S = null;
      
      Map<String, List<String>> identification = new HashMap<String, List<String>>();
      
      ///////////YAOUNDE
      List<String> list_Y = new ArrayList<String>();
      list_Y.add("CH6");
      list_Y.add("CH5");
      list_Y.add("PET1");
      list_Y.add("CH4");
      list_Y.add("BG");
      list_Y.add("PET8");
      identification.put("YAOUNDE",list_Y); 
      obj_Y = traitement("YAOUNDE", identification, 1, fis, jour, attr); //}}
      
       
      ///////////KOUMASSI
      fis = new FileInputStream(new File(name)); //+".xlsx"
      Map<String, List<String>> identification_K = new HashMap<String, List<String>>();
      List<String> list_K = new ArrayList<String>();
      list_K.add("PET9");
      list_K.add("CH5");
      list_K.add("CH4");
      list_K.add("PET1"); 
      identification_K.put("KOUMASSI",list_K); 
      obj_K = traitement("KOUMASSI", identification_K, 2, fis, jour, attr);
      
      
      ///////////NDOKOTI
      fis = new FileInputStream(new File(name)); //+".xlsx"
      Map<String, List<String>> identification_N = new HashMap<String, List<String>>();
      List<String> list_N = new ArrayList<String>();
      list_N.add("CH8");
      list_N.add("CH7");
      list_N.add("CH6"); 
      identification_N.put("NDOKOTI",list_N); 
      obj_N = traitement("NDOKOTI", identification_N, 3, fis, jour, attr);
      
      
      ///////////BAFOUSSAM
      fis = new FileInputStream(new File(name)); //+".xlsx"
      Map<String, List<String>> identification_B = new HashMap<String, List<String>>();
      List<String> list_B = new ArrayList<String>();
      list_B.add("CH5");
      list_B.add("CH4"); 
      identification_B.put("BAFOUSSAM",list_B); 
      obj_B = traitement("BAFOUSSAM", identification_B, 4, fis, jour, attr);
      
      
      ///////////GAROUA
      fis = new FileInputStream(new File(name)); //+".xlsx"
      Map<String, List<String>> identification_G = new HashMap<String, List<String>>();
      List<String> list_G = new ArrayList<String>();
      list_G.add("CH4");
      list_G.add("CH6");
      list_G.add("PET7"); 
      identification_G.put("GAROUA",list_G); 
      obj_G = traitement("GAROUA", identification_G, 5, fis, jour, attr);
      
      ///////////SEMC
      fis = new FileInputStream(new File(name));  //+".xlsx"
      Map<String, List<String>> identification_S = new HashMap<String, List<String>>();
      List<String> list_S = new ArrayList<String>();
      list_S.add("PET2");
      list_S.add("PET3");
      list_S.add("PET4"); 
      identification_S.put("SEMC",list_S); 
      obj_S = traitement("SEMC", identification_S, 6, fis, jour, attr);
      ///////////YAOUNDE
//      while (rowIterator.hasNext()) 
//      {
//         row = (XSSFRow) rowIterator.next();
//         Iterator < Cell > cellIterator = row.cellIterator();
//         while ( cellIterator.hasNext()) 
//         {
//            Cell cell = cellIterator.next();
//            switch (cell.getCellType()) 
//            {
//               case Cell.CELL_TYPE_NUMERIC:
//               System.out.print( 
//               cell.getNumericCellValue() + " \t\t " );
//               break;
//               case Cell.CELL_TYPE_STRING:
//               System.out.print(
//               cell.getStringCellValue() + " \t\t " );
//               break;
//            }
//         }
//         System.out.println();
//      }
      fis = new FileInputStream(new File(name)); //+".xlsx"
      String jr = "";
      switch(jour){
          case "LUNDI":
              jr = getStringCellContent("E3", 1, fis);
              break;
          case "MARDI":
              jr = getStringCellContent("J3", 1, fis);
              break;
          case "MERCREDI":
              jr = getStringCellContent("O3", 1, fis);
              break;
          case "JEUDI":
              jr = getStringCellContent("T3", 1, fis);
              break;
          case "VENDREDI":
              jr = getStringCellContent("Y3", 1, fis);
              break;
          case "SAMEDI":
              jr = getStringCellContent("AD3", 1, fis);
              break;
          case "DIMANCHE":
              jr = getStringCellContent("AI3", 1, fis);
              break;
      }
        
      Map < String, Object[] > result = new HashMap< String, Object[] >();
      if(jour.equals("HEBDO")){
          System.out.println("OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO");
      result.put("1", new Object[]{"", "Classement des HP Hebdo a date"});
      }else if(jour.equals("MENSUELLE")){
          System.out.println("AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOOOOOOOOOOOOOOOOOOO");
      result.put("1", new Object[]{"", "Classement des HP Mensuelle a date"});
      }else{
      result.put("1", new Object[]{"", "Classement des HP du "+jr});
      }
      result.put("2", new Object[]{"", "HP JOUR", "MOINS BON", "HP MOINS BON", "Cause Majeure", "Statut"}); //*_*
        Vector<Object[]> trop = new  Vector<Object[]>();
        Vector<Object[]> trop1 = new  Vector<Object[]>();
        trop.add(obj_Y);
        trop.add(obj_K);
        trop.add(obj_N);
        trop.add(obj_B);
        trop.add(obj_G);
        trop.add(obj_S);
        System.out.println("LONGUEUR TROP = " + trop.size());
        trop1 = tri(trop);
        int p=0;
        System.out.println("LONGUEUR TROP1 = "+trop1.size());
        for(int l=3; l<=3+trop1.size()-1; l++){
            result.put(Integer.toString(l),trop1.get(p));
            p++;
        }
      result.put("9",new Object[]{jr});
      fis.close();
      return result;
    }

    public Vector<Object[]> tri(Vector<Object[]> p){
        Vector<Object[]> finale = new Vector<Object[]>();
        HashMap<Object[],Long> sorter = new HashMap<>();
        for(int i=0; i< p.size(); i++){
            sorter.put(p.get(i),(Long)p.get(i)[1]);
        }


        Map<Object[], Long> values = sortByComparatorO(sorter, false);


        Set<Object[]> keys = values.keySet(); //////////&&&&
        for (Object[] key : keys) {
            finale.add(key);
            // do something
        }
        return finale;
    }

    public ArrayList<Map < String, Object[] >> readALL(String name) throws IOException{
        ArrayList<Map < String, Object[] >> all = new ArrayList<Map < String, Object[] >>();
        if(testPresence(name, "LUNDI")){
            all.add(readVillesHP(name, "LUNDI", "H"));
        }
        if(testPresence(name, "MARDI")){
            all.add(readVillesHP(name, "MARDI", "M"));
        }
        if(testPresence(name, "MERCREDI")){
            all.add(readVillesHP(name, "MERCREDI", "R"));
        }
        if(testPresence(name, "JEUDI")){
            all.add(readVillesHP(name, "JEUDI", "W"));
        }
        if(testPresence(name, "VENDREDI")){
            all.add(readVillesHP(name, "VENDREDI", "AB"));
        }
        if(testPresence(name, "SAMEDI")){
            all.add(readVillesHP(name, "SAMEDI", "AG"));
        }
        if(testPresence(name, "DIMANCHE")){
            all.add(readVillesHP(name, "DIMANCHE", "AL"));
        }
        all.add(readVillesHP(name, "HEBDO", "AQ"));
        all.add(readVillesHP(name, "MENSUELLE", "AW"));
        return all;
    }
     public static void main(String args[]) throws IOException {
          }
     
     public void colorHeader(String cell_reference, short color, XSSFWorkbook workbook, XSSFSheet spreadsheet){
      XSSFCellStyle style2 = workbook.createCellStyle();
      style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
      style2.setVerticalAlignment( 
      XSSFCellStyle.VERTICAL_CENTER);
      XSSFFont font = workbook.createFont();  
      font.setColor(color);
      style2.setFont(font);

      CellReference cr_Y = new CellReference(cell_reference);
      row = spreadsheet.getRow(cr_Y.getRow());
      Cell cell2 = row.getCell(cr_Y.getCol()); 
      cell2.setCellStyle(style2);
     }
     public void style(int type,String reference, XSSFWorkbook workbook, XSSFSheet spreadsheet){
      
         XSSFCellStyle style = workbook.createCellStyle();
         Font bld = workbook.createFont();
         bld.setBold(true);
      switch(type){
          case 3:
           style.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM);
           style.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);
           style.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);
           style.setFont(bld);  
           break;
          case 4:
           style.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM);
           style.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);
           style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
           style.setFont(bld);  
           break;
          case 5:
           style.setBorderRight(XSSFCellStyle.BORDER_MEDIUM);
            style.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);
            style.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM);
            style.setFont(bld);
            style.setBorderLeft(XSSFCellStyle.BORDER_THIN);  
            break;
          case 6:
            style.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);
            style.setBorderRight(XSSFCellStyle.BORDER_MEDIUM); 
              break;
          case 7:
              style.setBorderRight(XSSFCellStyle.BORDER_MEDIUM); 
      style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
      style.setBorderTop(XSSFCellStyle.BORDER_THIN);
      style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
              break;
          case 8:
              style.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM);  
      style.setBorderTop(XSSFCellStyle.BORDER_THIN);
      style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
      style.setBorderRight(XSSFCellStyle.BORDER_THIN);
      style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
      style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
              break;
          case 9:
              style.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);
              break;
          case 10:
              style.setBorderRight(XSSFCellStyle.BORDER_MEDIUM); 
      style.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM); 
      style.setBorderTop(XSSFCellStyle.BORDER_THIN);
      style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
              break;
          case 11:
              style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
      style.setBorderTop(XSSFCellStyle.BORDER_THIN);
      style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
      style.setBorderRight(XSSFCellStyle.BORDER_THIN);
      style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
      style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
              break;
          case 12:
              style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
      style.setBorderTop(XSSFCellStyle.BORDER_THIN);
      style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
      style.setBorderRight(XSSFCellStyle.BORDER_THIN); 
              break; 
      }
      CellReference cr_Y = new CellReference(reference);
      row = spreadsheet.getRow(cr_Y.getRow());
      Cell cell2 = row.getCell(cr_Y.getCol()); 
      cell2.setCellStyle(style);
     }
     
     public void drawChart(XSSFWorkbook workbook, XSSFSheet spreadsheet, int nrow, int ncol, int temp, Map<String, Integer> value) throws IOException{
         DefaultCategoryDataset my_bar_chart_dataset = new DefaultCategoryDataset(); 
        System.out.println("nrow = "+nrow+" ncol = "+ncol);
//        System.out.println("factories-size = "+factories.size()+" hps-size = "+hps.size());
//         if(factories.size() == hps.size()){
        Map<String, Integer> values = sortByComparator(value, false);
        Set<String> ss = values.keySet();
           for(String s: ss){
               System.out.println("usine = "+s+", hp = "+values.get(s));
//               System.out.println("nrow = "+nrow+"ncol = "+ncol+" -- factory(i) = "+factories.get(j)+"hp(i) = "+hps.get(j));
               my_bar_chart_dataset.addValue(values.get(s),"Heures Perdues",s);
               JFreeChart BarChartObject = null;
               switch(temp){
                   case 1:
                   BarChartObject=ChartFactory.createBarChart("HP Jour","Usine","Nombre D'Heures",my_bar_chart_dataset,PlotOrientation.VERTICAL,true,true,false);  
                       break;
                   case 2:
                   BarChartObject=ChartFactory.createBarChart("HP Hebdo","Usine","Nombre D'Heures",my_bar_chart_dataset,PlotOrientation.VERTICAL,true,true,false);  
                       break;  
                   case 3:
                   BarChartObject=ChartFactory.createBarChart("HP Mensuelle","Usine","Nombre D'Heures",my_bar_chart_dataset,PlotOrientation.VERTICAL,true,true,false);  
                       break;
               }
               BarRenderer renderer = new BarRenderer();
                 renderer.setSeriesItemLabelGenerator(0, new StandardCategoryItemLabelGenerator()); 
               renderer.setSeriesItemLabelsVisible(0, true);
                
                BarChartObject.getCategoryPlot().setBackgroundPaint(Color.white);
               //////////////// ++++

               ((BarRenderer)BarChartObject.getCategoryPlot().getRenderer()).setBarPainter(new StandardBarPainter());

               final CategoryPlot plot = BarChartObject.getCategoryPlot();
               plot.setDomainGridlinePaint(Color.black);
               plot.setRangeGridlinePaint(Color.black);
               /////////////// ++++
                ((BarRenderer) BarChartObject.getCategoryPlot().getRenderer()).setItemMargin(12);
                renderer.setMaximumBarWidth(0.05);
                renderer.setSeriesPaint(0, new Color(91, 155, 213)); 
                //renderer.setMaximumBarWidth(0.1);
                renderer.setBarPainter(new StandardBarPainter());
               BarChartObject.getCategoryPlot().setDomainGridlineStroke(new BasicStroke());
               BarChartObject.getCategoryPlot().setRenderer(renderer);
               int width=660; /* Width of the chart */
               int height=400; /* Height of the chart */
               ByteArrayOutputStream chart_out = new ByteArrayOutputStream();          
               ChartUtilities.writeChartAsPNG(chart_out,BarChartObject,width,height);
               int my_picture_id = workbook.addPicture(chart_out.toByteArray(), Workbook.PICTURE_TYPE_PNG);
                /* we close the output stream as we don't need this anymore */
                chart_out.close();
                /* Create the drawing container */
                XSSFDrawing drawing = spreadsheet.createDrawingPatriarch();
                /* Create an anchor point */
                ClientAnchor my_anchor = new XSSFClientAnchor();
                /* Define top left corner, and we can resize picture suitable from there */
                my_anchor.setCol1(ncol);
                my_anchor.setRow1(nrow);
                /* Invoke createPicture and pass the anchor point and ID */
                XSSFPicture  my_picture = drawing.createPicture(my_anchor, my_picture_id);
                /* Call resize method, which resizes the image */
                my_picture.resize();             
           }
//       }else{
//            System.out.println("Error: Hp size not equal to factory size"); 
//       }
     }

     private static Map<String, Integer> sortByComparator(Map<String, Integer> unsortMap, final boolean order){

        List<Entry<String, Integer>> list = new LinkedList<Entry<String, Integer>>(unsortMap.entrySet());
        Set<String> pp = unsortMap.keySet();
        for(String p : pp){
            System.out.println("---- key is "+p+" value is "+unsortMap.get(p));
        }  
        // Sorting the list based on values
        Collections.sort(list, new Comparator<Entry<String, Integer>>()
        {
            public int compare(Entry<String, Integer> o1,
                    Entry<String, Integer> o2)
            {
                if (order)
                {
                    return o1.getValue().compareTo(o2.getValue());
                }
                else
                {
                    return o2.getValue().compareTo(o1.getValue());

                }
            }
        });

        // Maintaining insertion order with the help of LinkedList
        Map<String, Integer> sortedMap = new LinkedHashMap<String, Integer>();
        for (Entry<String, Integer> entry : list)
        {
            sortedMap.put(entry.getKey(), entry.getValue());
        }

        return sortedMap;
    }

    private static Map<Object[], Long> sortByComparatorO(Map<Object[], Long> unsortMap, final boolean order){

        List<Entry<Object[], Long>> list = new LinkedList<Entry<Object[], Long>>(unsortMap.entrySet());
        Set<Object[]> pp = unsortMap.keySet();
        for(Object[] p : pp){
            System.out.println("---- key is "+p+" value is "+unsortMap.get(p));
        }
        // Sorting the list based on values
        Collections.sort(list, new Comparator<Entry<Object[], Long>>()
        {
            public int compare(Entry<Object[], Long> o1,
                               Entry<Object[], Long> o2)
            {
                if (order)
                {
                    return o1.getValue().compareTo(o2.getValue());
                }
                else
                {
                    return o2.getValue().compareTo(o1.getValue());

                }
            }
        });

        // Maintaining insertion order with the help of LinkedList
        Map<Object[], Long> sortedMap = new LinkedHashMap<Object[], Long>();
        for (Entry<Object[], Long> entry : list)
        {
            sortedMap.put(entry.getKey(), entry.getValue());
        }

        return sortedMap;
    }
}

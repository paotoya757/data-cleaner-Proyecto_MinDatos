/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.mavenproject1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author pa.otoya757
 */
public class MainExecutor {

    private static final String pathToInput = "C:\\Users\\peter\\Documents\\NetBeansProjects\\data-cleaner-Proyecto_MinDatos\\dataIO\\Input\\GrandSlams-2013.xls";
    private static final String pathToModel1 = "C:\\Users\\peter\\Documents\\NetBeansProjects\\data-cleaner-Proyecto_MinDatos\\dataIO\\Output\\GrandSlams-output-model-1y2.xls";

    public static void main(String[] args) throws Exception {
        FileInputStream fio = new FileInputStream(new File(pathToInput));

        //Read
        HSSFWorkbook inputExcel = new HSSFWorkbook(fio);

        // Write 
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Sample sheet");

        LinkedHashSet<String> players = getAllPlayerNames(inputExcel);
        System.out.println(players.size());
        Iterator<String> playersIterator = players.iterator();

        Map<String, Object[]> data = getNewWorkbookData(inputExcel, playersIterator);
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset) {
            System.out.println("com.mycompany.mavenproject1.MainExecutor.main");
            System.out.println(rownum);
            Row row = sheet.createRow(rownum++);
            Object[] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof String) {
                    cell.setCellValue((String) obj);
                } else if (obj instanceof Double) {
                    cell.setCellValue((Double) obj);
                }
            }
        }

        try {
            FileOutputStream out
                    = new FileOutputStream(new File(pathToModel1));
            workbook.write(out);
            out.close();
            System.out.println("Excel written successfully..");

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     *
     * @param inputExcel
     * @return
     */
    static LinkedHashSet<String> getAllPlayerNames(HSSFWorkbook inputExcel) {
        LinkedHashSet<String> playerlist = new LinkedHashSet<String>();
        for (int i = 0; i < inputExcel.getNumberOfSheets(); i++) {
            Iterator<Row> rowIterator = inputExcel.getSheetAt(i).iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                for (int n = 0; n < 2; n++) {
                    Cell cell = cellIterator.next();
                    String possibleNewName = cell.getStringCellValue();
                    playerlist.add(possibleNewName);
                }
            }
        }

        return playerlist;
    }

    //numeric cell value is a double
    static Map<String, Object[]> getNewWorkbookData(HSSFWorkbook inputExcel, Iterator<String> playersIterator) {
        Map<String, Object[]> data = new HashMap<String, Object[]>();
        data.put("1", new Object[]{"Player", "FSP.1", "FSW.1", "SSP.1", "SSW.1", "ACE.1", "DBF.1", "WNR.1", "UFE.1", "BPC.1", "BPW.1", "NPA.1", "NPW.1"});
        int key = 2;
        while (playersIterator.hasNext()) {
            String player = playersIterator.next();
            for (int i = 0; i < inputExcel.getNumberOfSheets(); i++) {
                Iterator<Row> rowIterator = inputExcel.getSheetAt(i).iterator();
                rowIterator.next();
                int row_counter = 0 ;
                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    Object[] rowValues = new Object[13];

                    String player_1_cell = row.getCell(0).getStringCellValue();
                    String player_2_cell = row.getCell(1).getStringCellValue();
                    int dataRangeLow = 0;
                    int dataRangeHigh = 0;

                    if (player.equals(player_1_cell)) {
                        dataRangeLow = 6;
                        dataRangeHigh = 17;
                        rowValues[0] = player;
                    } else if (player.equals(player_2_cell)) {
                        dataRangeLow = 24;
                        dataRangeHigh = 35;
                        rowValues[0] = player;
                    } else {
                        // Go to another row.
                    }
                    int cell_counter = 0;
                    for (int j = dataRangeLow; j < dataRangeHigh ; j++) {
                        Cell cell = row.getCell(j);
                        System.out.println(cell_counter);
                        if( cell != null )
                            rowValues[j - dataRangeLow + 1] = cell.getNumericCellValue();
                        
                        if(cell_counter==5){
                           Object debug = new Object();
                        }
                        cell_counter++;
                        data.put("" + key, rowValues);
                        key++;
                    }
                    
                    System.out.println(row_counter);
                    row_counter++;
                }
            }
        }

        return data;
    }

}

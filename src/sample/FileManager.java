package sample;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;

public class FileManager {

    public static final String resultFile = "result-20-CVO.xls";

    public void combineFile(Optional<String> previousFile, Optional<String> currentFile, Optional<String> pathToSave) throws FileManagerException, IOException, InvalidFormatException {
        previousFile.orElseThrow(() -> new FileManagerException("Отсутсвует файл за прошлые периоды !!!"));
        currentFile.orElseThrow(() -> new FileManagerException("Отсутсвует файл за прошедший месяц !!!"));
        pathToSave.orElseThrow(() -> new FileManagerException("Отсутсвует путь для результата !!!"));
        Workbook previous = WorkbookFactory.create(new FileInputStream(previousFile.get()));
        Workbook current = WorkbookFactory.create(new FileInputStream(currentFile.get()));
        combineData(previous, current);
        writeToFile(previous, pathToSave.get());
    }

    public boolean isNewPeriodValidator(String path) {
        try {
            Workbook file = WorkbookFactory.create(new FileInputStream(path));
            Cell cell = file.getSheetAt(0).getRow(3).getCell(0);
            return cell.getCellType() == Cell.CELL_TYPE_STRING && !cell.getStringCellValue().toLowerCase().contains("січень");
        }  catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return false;
    }

    private void writeToFile(Workbook previous, String pathToSave) throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream(pathToSave+"/"+resultFile);
        previous.write(fileOutputStream);
        fileOutputStream.flush();
        fileOutputStream.close();
    }
    private void combineData(Workbook previous, Workbook current) throws FileManagerException {
        Sheet previousSheet = previous.getSheetAt(0);
        Sheet currentSheet = current.getSheetAt(0);
        FormulaEvaluator previousEvaluator = previous.getCreationHelper().createFormulaEvaluator();
        FormulaEvaluator currentEvaluator = current.getCreationHelper().createFormulaEvaluator();
        combinePeriodTitle(previousSheet, currentSheet);
        combineGeneralTable(previousSheet, currentSheet, previousEvaluator, currentEvaluator);
        combineFirstSubTable(previousSheet, currentSheet, previousEvaluator, currentEvaluator);
        combineSecondSubTable(previousSheet, currentSheet, previousEvaluator, currentEvaluator);
        combineThirdSubTable(previousSheet, currentSheet, previousEvaluator, currentEvaluator);
        createSignature(previousSheet, currentSheet);
    }
    private void combinePeriodTitle(Sheet previousSheet,  Sheet currentSheet) {
        Cell previousCell = previousSheet.getRow(3).getCell(0);
        String currentString = currentSheet.getRow(3).getCell(0).getStringCellValue();
        previousCell.setCellValue("за січень - " + currentString.substring(2));
    }

    private void combineGeneralTable(Sheet previousSheet,  Sheet currentSheet, FormulaEvaluator previousEvaluator, FormulaEvaluator currentEvaluator){
       for (int i = 7; i <= 19; i++) {
           for (int j = 1; j<=24; j ++) {
               cellDataCombiner(previousSheet, currentSheet, previousEvaluator, currentEvaluator, i, j);
           }
       }
    }
    private void combineFirstSubTable(Sheet previousSheet,  Sheet currentSheet, FormulaEvaluator previousEvaluator, FormulaEvaluator currentEvaluator){
        for (int i = 23; i <= 27; i=i+2) {
            for (int j = 1; j<=7; j=j+2) {
                cellDataCombiner(previousSheet, currentSheet, previousEvaluator, currentEvaluator, i, j);
            }
        }
    }
    private void combineSecondSubTable(Sheet previousSheet,  Sheet currentSheet, FormulaEvaluator previousEvaluator, FormulaEvaluator currentEvaluator){
        for (int i = 23; i <= 27; i++) {
            for (int j = 12; j<=16; j=j+2) {
                cellDataCombiner(previousSheet, currentSheet, previousEvaluator, currentEvaluator, i, j);
            }
        }
    }
    private void combineThirdSubTable(Sheet previousSheet,  Sheet currentSheet, FormulaEvaluator previousEvaluator, FormulaEvaluator currentEvaluator){
        List<Integer> iList = Arrays.asList(23,24,26);
        for (Integer i : iList) {
            for (int j = 21; j<=23; j=j+2) {
                cellDataCombiner(previousSheet, currentSheet, previousEvaluator, currentEvaluator, i, j);
            }
        }
    }
    private void createSignature(Sheet previousSheet,  Sheet currentSheet) {
        for (int i=29; i<=31; i=i+2) {
            Cell previousCell = previousSheet.getRow(i).getCell(0);
            Cell currentCell = currentSheet.getRow(i).getCell(0);

            if (previousCell.getCellType() == Cell.CELL_TYPE_STRING && currentCell.getCellType() == Cell.CELL_TYPE_STRING
                    && previousCell.getStringCellValue().length() < currentCell.getStringCellValue().length()) {
                previousCell.setCellValue(currentCell.getStringCellValue());
            }
        }

    }
    private void cellDataCombiner(Sheet previousSheet, Sheet currentSheet, FormulaEvaluator previousEvaluator, FormulaEvaluator currentEvaluator, int i, int j) {
        Cell previousCell = previousEvaluator.evaluateInCell(previousSheet.getRow(i).getCell(j));
        Cell currentCell = currentEvaluator.evaluateInCell(currentSheet.getRow(i).getCell(j));

        if (previousCell.getCellType() == Cell.CELL_TYPE_NUMERIC && currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC){
            previousCell.setCellValue(previousCell.getNumericCellValue()+currentCell.getNumericCellValue());
        } else if (currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            previousCell.setCellValue(currentCell.getNumericCellValue());
        }
    }



}

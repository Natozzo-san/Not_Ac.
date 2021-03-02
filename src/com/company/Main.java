package com.company;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;


import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

class ExcelParser {

    public static String parse(String fileName) {
        List<ArrayList<Double>> TimeImidazol=new ArrayList();
        List<ArrayList<Double>> TimeFon=new ArrayList();
        ArrayList<String> names = new ArrayList<>();
        ArrayList<Integer> indexOfCells = new ArrayList<>();
        int timeIndexCell=0, imidazolIndexCell=0,fonIndexCell=0, presentIndexCell;//нумерация столбцов
        int cellIndex=0;//нумерация строк
        int sheetIndex;//нумерация листов
        int maximidazol=0, maxfon=0, maxImidazolTime=0, maxFonTime=0;
        String finalNameResult;
        //инициализируем потоки
        String result = "";
        InputStream inputStream = null;
        HSSFWorkbook workBook = null;
        try {
            inputStream = new FileInputStream(fileName);
            workBook = new HSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //разбираем первый лист входного файла на объектную модель
        Sheet sheet = workBook.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        //проходим по всему листу
        while (it.hasNext()) {
            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            cellIndex++;
            presentIndexCell=0;
            ArrayList <Double> TimeImidazol1=new ArrayList<>();
            ArrayList <Double> TimeFon1=new ArrayList<>();
            while (cells.hasNext()) {
                Cell cell = cells.next();
                presentIndexCell++;
                int cellType = cell.getCellType();
                //перебираем возможные типы ячеек

                switch (cellType) {

                    case Cell.CELL_TYPE_STRING:
                        result += cell.getStringCellValue() + "=";
                            String s = cell.getStringCellValue(), sub = " ";
                            if (s.indexOf(sub) != -1){
                                finalNameResult = cell.getStringCellValue().replaceAll("[\\s]{2,}", " ");
                            }
                            else finalNameResult = cell.getStringCellValue();
                            if (finalNameResult.equals("Время")){
                                timeIndexCell =presentIndexCell;
                            }
                            else if (finalNameResult.equals("имидазол")){
                                imidazolIndexCell =presentIndexCell;
                            }
                            else if (finalNameResult.equals("фон")){
                                fonIndexCell =presentIndexCell;
                            }
                            else if (finalNameResult.equals("")){
                                System.out.println("Пустое название");
                            }
                            else {System.out.println("Неверно назван столбец. Остановите работу программы. Столбец №"+presentIndexCell);
                            }
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        if (presentIndexCell == timeIndexCell){
                            double pass = Cell.CELL_TYPE_NUMERIC;
                            TimeImidazol1.add(pass);
                            TimeFon1.add(pass);
                        }
                        if (presentIndexCell == imidazolIndexCell){
                            double pass = Cell.CELL_TYPE_NUMERIC;
                            TimeImidazol1.add(pass);
                            TimeImidazol.add(TimeImidazol1);
                        }
                        if (presentIndexCell == fonIndexCell){
                            double pass = Cell.CELL_TYPE_NUMERIC;
                            TimeFon1.add(pass);
                            TimeFon.add(TimeFon1);
                        }
                        result += "[" + cell.getNumericCellValue() + "]";
                        break;

                    case Cell.CELL_TYPE_FORMULA:
                        result += "[" + cell.getNumericCellValue() + "]";
                        break;
                    default:
                        result += "|";
                        break;
                }
            }
            result += "\n";
        }
        for(int i=0; i<TimeImidazol.size(); i++){
            ArrayList <Integer> arr1=new ArrayList<>();
            for(int j=0; j<2;j++) {


            }
        }

        return result;
    }



    public static int getLastIndex(String fileName) {
        int timeIndexCell=0, imidazolIndexCell=0,fonIndexCell=0, presentIndexCell;//нумерация столбцов
        int cellIndex=0, lastIndex=0;//нумерация строк
        int sheetIndex;//нумерация листов
        int maximidazol=0, maxfon=0, maxImidazolTime=0, maxFonTime=0;
        String finalNameResult;
        //инициализируем потоки
        String result = "";
        InputStream inputStream = null;
        HSSFWorkbook workBook = null;
        try {
            inputStream = new FileInputStream(fileName);
            workBook = new HSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //разбираем первый лист входного файла на объектную модель
        Sheet sheet = workBook.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        //проходим по всему листу
        while (it.hasNext()) {
            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            cellIndex++;
            presentIndexCell=0;
            while (cells.hasNext()) {
                Cell cell = cells.next();
                presentIndexCell++;
                int cellType = cell.getCellType();
                //перебираем возможные типы ячеек

                switch (cellType) {

                    case Cell.CELL_TYPE_STRING:
                        String s = cell.getStringCellValue(), sub = " ";
                        if (s.indexOf(sub) != -1){
                            finalNameResult = cell.getStringCellValue().replaceAll("[\\s]{2,}", " ");
                        }
                        else finalNameResult = cell.getStringCellValue();
                        if (finalNameResult.equals("Время")){
                            timeIndexCell =presentIndexCell;
                            if (lastIndex != cellIndex){
                                lastIndex=cellIndex;
                            }
                        }
                        else if (finalNameResult.equals("имидазол")){
                            imidazolIndexCell =presentIndexCell;
                            if (lastIndex != cellIndex){
                                lastIndex=cellIndex;
                            }
                        }
                        else if (finalNameResult.equals("фон")){
                            fonIndexCell =presentIndexCell;
                            if (lastIndex != cellIndex){
                                lastIndex=cellIndex;
                            }
                        }
                        else if (finalNameResult.equals("")){
                            System.out.println("Пустое название");
                            if (lastIndex != cellIndex){
                                lastIndex=cellIndex;
                            }
                        }
                        else {System.out.println("Неверно назван столбец. Остановите работу программы. Столбец №"+presentIndexCell);
                            if (lastIndex != cellIndex){
                                lastIndex=cellIndex;
                            }}
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        if (presentIndexCell == timeIndexCell){
                            if (lastIndex != cellIndex){
                                lastIndex=cellIndex;
                            }
                        }
                        if (presentIndexCell == imidazolIndexCell){
                            if (lastIndex != cellIndex){
                                lastIndex=cellIndex;
                            }
                        }
                        if (presentIndexCell == fonIndexCell){
                            if (lastIndex != cellIndex){
                                lastIndex=cellIndex;
                            }
                        }
                        break;

                    case Cell.CELL_TYPE_FORMULA:
                        if (lastIndex != cellIndex){
                            lastIndex=cellIndex;
                        }
                        break;
                    default:
                        break;
                }
            }
        }

        return lastIndex;
    }





    public static String result(String fileName, String filetwo, int Greatindex, int lastindex) {
        int finalCount=0, notFCount=0;
        LinkedList<Double> Time= new LinkedList<Double>();
        LinkedList<Double> Fon= new LinkedList<Double>();
        LinkedList<Double> Imidazol= new LinkedList<Double>();
        LinkedList<Double> finalResult= new LinkedList<Double>();
        int count = lastindex - Greatindex;//количество полученных строк
        ArrayList<String> names = new ArrayList<>();
        int timeIndexCell=0, imidazolIndexCell=0,fonIndexCell=0, presentIndexCell;//нумерация столбцов
        int cellIndex=0;//нумерация строк
        int sheetIndex;//нумерация листов
        int maximidazol=0, maxfon=0, maxImidazolTime=0, maxFonTime=0;
        String finalNameResult;
        //инициализируем потоки
        String result = "";
        InputStream inputStream = null;
        HSSFWorkbook workBook = null;
        try {
            inputStream = new FileInputStream(fileName);
            workBook = new HSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //разбираем первый лист входного файла на объектную модель
        Sheet sheet = workBook.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        //проходим по всему листу
        while (it.hasNext()) {
            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            cellIndex++;
            presentIndexCell=0;
            while (cells.hasNext()) {
                Cell cell = cells.next();
                presentIndexCell++;
                int cellType = cell.getCellType();
                //перебираем возможные типы ячеек

                switch (cellType) {

                    case Cell.CELL_TYPE_STRING:
                        result += cell.getStringCellValue() + "=";
                        String s = cell.getStringCellValue(), sub = " ";
                        if (s.indexOf(sub) != -1){
                            finalNameResult = cell.getStringCellValue().replaceAll("[\\s]{2,}", " ");
                        }
                        else finalNameResult = cell.getStringCellValue();
                        if (finalNameResult.equals("Время")){
                            timeIndexCell =presentIndexCell;
                            names.add(cell.getStringCellValue());
                        }
                        else if (finalNameResult.equals("имидазол")){
                            imidazolIndexCell =presentIndexCell;
                            names.add(cell.getStringCellValue());
                        }
                        else if (finalNameResult.equals("фон")){
                            fonIndexCell =presentIndexCell;
                            names.add(cell.getStringCellValue());
                        }
                        else if (finalNameResult.equals("")){
                            System.out.println("Пустое название");
                        }
                        else {System.out.println("Неверно назван столбец. Остановите работу программы. Столбец №"+presentIndexCell);
                        }
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        if (presentIndexCell == timeIndexCell && cellIndex<=lastindex){
                            double pass = Cell.CELL_TYPE_NUMERIC;
                            Time.add(cell.getNumericCellValue());
                        }
                        if (presentIndexCell == imidazolIndexCell && cellIndex>=(Greatindex+2)){
                            double pass = Cell.CELL_TYPE_NUMERIC;
                            Imidazol.add(cell.getNumericCellValue());

                        }
                        if (presentIndexCell == fonIndexCell && cellIndex<=lastindex){
                            double pass = Cell.CELL_TYPE_NUMERIC;
                            Fon.add(cell.getNumericCellValue());
                        }
                        break;

                    case Cell.CELL_TYPE_FORMULA:
                        break;
                    default:
                        break;
                }
            }
        }
        //кусок кода- создание нового листа и его заливка данными


        try{Workbook book = new HSSFWorkbook();
            Sheet secondsheet = book.createSheet("SecondSheet");
            Row row = secondsheet.createRow((short)0);
            row.setHeightInPoints(40.0f);
            Cell one = row.createCell(0);
            one.setCellType(CellType._NONE.ordinal());
            one.setCellValue(names.get(0));

            Cell two = row.createCell(1);
            two.setCellValue(CellType._NONE.ordinal());
            two.setCellValue(names.get(1));

            Cell three = row.createCell(2);
            three.setCellValue(CellType._NONE.ordinal());
            three.setCellValue(names.get(2));

            Cell cell = row.createCell(3);
            cell.setCellValue(CellType._NONE.ordinal());
            cell.setCellValue("результат");
            for (int in=1; in<(1002);in++){
                // Нумерация начинается с нуля
                Row roww = secondsheet.createRow((short)in);
                roww.setHeightInPoints(40.0f);
                // Мы запишем имя и дату в два столбца
                // имя будет String, а дата рождения --- Date,
                // формата dd.mm.yyyy
                Cell first = roww.createCell(0);
                first.setCellType(CellType.NUMERIC.ordinal());
                first.setCellValue(Time.get(in));

                Cell second = roww.createCell(1);
                second.setCellType(CellType.NUMERIC.ordinal());
                second.setCellValue(Imidazol.get(in));
                finalCount += Imidazol.get(in);

                Cell last = roww.createCell(2);
                last.setCellType(CellType.NUMERIC.ordinal());
                last.setCellValue(Fon.get(in));
                finalCount -= Fon.get(in);

                Cell cel = roww.createCell(3);
                cel.setCellValue(CellType.NUMERIC.ordinal());
                double finResult = Imidazol.get(in)-Fon.get(in);
                cel.setCellValue(finResult);
                notFCount+=finResult;
                }
            Cell cel = row.createCell(4);
            cel.setCellValue(CellType.NUMERIC.ordinal());
            cel.setCellValue(finalCount);

            // Меняем размер столбца
            sheet.autoSizeColumn(1);

            // Записываем всё в файл
            book.write(new FileOutputStream(filetwo));
            book.close();}catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println(finalCount+"  finC");
        return "Возможно сработало";
    }



    public static Double timeImidazolReturnwithHelpCellNumber(String fileName, int index) {
        double finalTimeIndex=0;
        int timeIndexCell=0, imidazolIndexCell=0,fonIndexCell=0, presentIndexCell;//нумерация столбцов
        int cellIndex=0;//нумерация строк
        int sheetIndex;//нумерация листов
        int maximidazol=0, maxfon=0, maxImidazolTime=0, maxFonTime=0;
        String finalNameResult;
        //инициализируем потоки
        String result = "";
        InputStream inputStream = null;
        HSSFWorkbook workBook = null;
        try {
            inputStream = new FileInputStream(fileName);
            workBook = new HSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //разбираем первый лист входного файла на объектную модель
        Sheet sheet = workBook.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        //проходим по всему листу
        while (it.hasNext()) {
            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            cellIndex++;
            presentIndexCell=0;
            ArrayList <Double> TimeImidazol1=new ArrayList<>();
            ArrayList <Double> TimeFon1=new ArrayList<>();
            while (cells.hasNext()) {
                Cell cell = cells.next();
                presentIndexCell++;
                int cellType = cell.getCellType();
                //перебираем возможные типы ячеек

                switch (cellType) {

                    case Cell.CELL_TYPE_STRING:
                        result += cell.getStringCellValue() + "=";
                        String s = cell.getStringCellValue(), sub = " ";
                        if (s.indexOf(sub) != -1){
                            finalNameResult = cell.getStringCellValue().replaceAll("[\\s]{2,}", " ");
                        }
                        else finalNameResult = cell.getStringCellValue();
                        if (finalNameResult.equals("Время")){
                            timeIndexCell =presentIndexCell;
                        }
                        else if (finalNameResult.equals("имидазол")){
                            imidazolIndexCell =presentIndexCell;
                        }
                        else if (finalNameResult.equals("фон")){
                            fonIndexCell =presentIndexCell;
                        }
                        else if (finalNameResult.equals("")){
                            System.out.println("Пустое название");
                        }
                        else {System.out.println("Неверно назван столбец. Остановите работу программы. Столбец №"+presentIndexCell);
                        }
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        if (presentIndexCell == timeIndexCell  && cellIndex==index){
                            double pass = Cell.CELL_TYPE_NUMERIC;
                            finalTimeIndex = cell.getNumericCellValue();
                            TimeImidazol1.add(pass);
                            TimeFon1.add(pass);
                        }

                        result += "[" + cell.getNumericCellValue() + "]";
                        break;

                    case Cell.CELL_TYPE_FORMULA:
                        result += "[" + cell.getNumericCellValue() + "]";
                        break;
                    default:
                        result += "|";
                        break;
                }
            }
            result += "\n";
        }

        return finalTimeIndex;
    }


    public static Double timeDelphaFonReturnwithHelpCellNumber(String fileName, int index) {
        double finalTimeIndex=0;
        boolean flag = true;
        int timeIndexCell=0, imidazolIndexCell=0,fonIndexCell=0, presentIndexCell;//нумерация столбцов
        int cellIndex=0;//нумерация строк
        int sheetIndex;//нумерация листов
        double maxFonTime=0, primarytimeFon=0;
        String finalNameResult;
        //инициализируем потоки
        String result = "";
        InputStream inputStream = null;
        HSSFWorkbook workBook = null;
        try {
            inputStream = new FileInputStream(fileName);
            workBook = new HSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //разбираем первый лист входного файла на объектную модель
        Sheet sheet = workBook.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        //проходим по всему листу
        while (it.hasNext()) {
            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            cellIndex++;
            presentIndexCell=0;
            while (cells.hasNext()) {
                Cell cell = cells.next();
                presentIndexCell++;
                int cellType = cell.getCellType();
                //перебираем возможные типы ячеек

                switch (cellType) {

                    case Cell.CELL_TYPE_STRING:
                        String s = cell.getStringCellValue(), sub = " ";
                        if (s.indexOf(sub) != -1){
                            finalNameResult = cell.getStringCellValue().replaceAll("[\\s]{2,}", " ");
                        }
                        else finalNameResult = cell.getStringCellValue();
                        if (finalNameResult.equals("Время")){
                            timeIndexCell =presentIndexCell;
                        }
                        else if (finalNameResult.equals("имидазол")){
                            imidazolIndexCell =presentIndexCell;
                        }
                        else if (finalNameResult.equals("фон")){
                            fonIndexCell =presentIndexCell;
                        }
                        else if (finalNameResult.equals("")){
                            System.out.println("Пустое название");
                        }
                        else {System.out.println("Неверно назван столбец. Остановите работу программы. Столбец №"+presentIndexCell);
                        }
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        if (presentIndexCell == timeIndexCell  && flag==true){
                            primarytimeFon = cell.getNumericCellValue();
                            flag = false;
                        }
                        if (presentIndexCell == timeIndexCell  && cellIndex==index){
                            maxFonTime = cell.getNumericCellValue();
                            finalTimeIndex = maxFonTime - primarytimeFon;
                        }

                        break;

                    case Cell.CELL_TYPE_FORMULA:
                        result += "[" + cell.getNumericCellValue() + "]";
                        break;
                    default:
                        result += "|";
                        break;
                }
            }
            result += "\n";
        }

        return finalTimeIndex;
    }



    public static Float DelphaTime(String fileName) {
        float finalTimeIndex=0;
        boolean firstflag = true, secondflag = true;
        int timeIndexCell=0,presentIndexCell;//нумерация столбцов
        int cellIndex=0;//нумерация строк
        int sheetIndex;//нумерация листов
        float firstTime=0, secondTime=0;
        String finalNameResult;
        //инициализируем потоки
        String result = "";
        InputStream inputStream = null;
        HSSFWorkbook workBook = null;
        try {
            inputStream = new FileInputStream(fileName);
            workBook = new HSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //разбираем первый лист входного файла на объектную модель
        Sheet sheet = workBook.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        //проходим по всему листу
        while (it.hasNext()) {
            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            cellIndex++;
            presentIndexCell=0;
            while (cells.hasNext()) {
                Cell cell = cells.next();
                presentIndexCell++;
                int cellType = cell.getCellType();
                //перебираем возможные типы ячеек

                switch (cellType) {

                    case Cell.CELL_TYPE_STRING:
                        String s = cell.getStringCellValue(), sub = " ";
                        if (s.indexOf(sub) != -1){
                            finalNameResult = cell.getStringCellValue().replaceAll("[\\s]{2,}", " ");
                        }
                        else finalNameResult = cell.getStringCellValue();
                        if (finalNameResult.equals("Время")){
                            timeIndexCell =presentIndexCell;
                        }
                        else if (finalNameResult.equals("имидазол")){

                        }
                        else if (finalNameResult.equals("фон")){

                        }
                        else if (finalNameResult.equals("")){
                            System.out.println("Пустое название");
                        }
                        else {System.out.println("Неверно назван столбец. Остановите работу программы. Столбец №"+presentIndexCell);
                        }
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        if (presentIndexCell == timeIndexCell  && firstflag==true  && secondflag==true){
                            firstTime = (float) cell.getNumericCellValue();
                            firstflag = false;
                        }
                        else if (presentIndexCell == timeIndexCell  && firstflag==false  && secondflag==true){
                            secondTime = (float) cell.getNumericCellValue();
                            finalTimeIndex = - firstTime + secondTime;
                            secondflag=false;
                        }

                        break;

                    case Cell.CELL_TYPE_FORMULA:
                        result += "[" + cell.getNumericCellValue() + "]";
                        break;
                    default:
                        result += "|";
                        break;
                }
            }
            result += "\n";
        }

        return finalTimeIndex;
    }



    public static int findMaxFonReturnSell(String fileName) {
        int timeIndexCell=0, imidazolIndexCell=0,fonIndexCell=0, presentIndexCell;//нумерация столбцов
        int cellIndex=0;//нумерация строк
        int sheetIndex;//нумерация листов
        int  maxFonCell=0;
        double maxfon=0;
        String finalNameResult;
        //инициализируем потоки
        String result = "";
        InputStream inputStream = null;
        HSSFWorkbook workBook = null;
        try {
            inputStream = new FileInputStream(fileName);
            workBook = new HSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //разбираем первый лист входного файла на объектную модель
        Sheet sheet = workBook.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        //проходим по всему листу
        while (it.hasNext()) {
            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            cellIndex++;
            presentIndexCell=0;
            while (cells.hasNext()) {
                Cell cell = cells.next();
                presentIndexCell++;
                int cellType = cell.getCellType();
                //перебираем возможные типы ячеек

                switch (cellType) {

                    case Cell.CELL_TYPE_STRING:
                        result += cell.getStringCellValue() + "=";
                        String s = cell.getStringCellValue(), sub = " ";
                        if (s.indexOf(sub) != -1){
                            finalNameResult = cell.getStringCellValue().replaceAll("[\\s]{2,}", " ");
                        }
                        else finalNameResult = cell.getStringCellValue();
                        if (finalNameResult.equals("Время")){
                            timeIndexCell =presentIndexCell;
                        }
                        else if (finalNameResult.equals("имидазол")){
                            imidazolIndexCell =presentIndexCell;
                        }
                        else if (finalNameResult.equals("фон")){
                            fonIndexCell =presentIndexCell;
                        }
                        else if (finalNameResult.equals("")){
                            System.out.println("Пустое название");
                        }
                        else {System.out.println("Неверно назван столбец. Остановите работу программы. Столбец №"+presentIndexCell);
                        }
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        if (presentIndexCell == timeIndexCell){
                        }
                        if (presentIndexCell == imidazolIndexCell){
                        }
                        if (presentIndexCell == fonIndexCell){
                            if (maxfon < cell.getNumericCellValue()){
                                maxfon = cell.getNumericCellValue();
                                maxFonCell= cellIndex;
                            }
                        }
                        result += "[" + cell.getNumericCellValue() + "]";
                        break;

                    case Cell.CELL_TYPE_FORMULA:
                        result += "[" + cell.getNumericCellValue() + "]";
                        break;
                    default:
                        result += "|";
                        break;
                }
            }
            result += "\n";
        }
        return maxFonCell;
    }

    public static int findmaximidazolReturnCell(String fileName) {
        int timeIndexCell=0, imidazolIndexCell=0,fonIndexCell=0, presentIndexCell;//нумерация столбцов
        int cellIndex=0;//нумерация строк
        int sheetIndex;//нумерация листов
        int  maxSell=0;
        double maximidazol=0;
        String finalNameResult;
        //инициализируем потоки
        String result = "";
        InputStream inputStream = null;
        HSSFWorkbook workBook = null;
        try {
            inputStream = new FileInputStream(fileName);
            workBook = new HSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //разбираем первый лист входного файла на объектную модель
        Sheet sheet = workBook.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        //проходим по всему листу
        while (it.hasNext()) {
            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            cellIndex++;
            presentIndexCell=0;
            while (cells.hasNext()) {
                Cell cell = cells.next();
                presentIndexCell++;
                int cellType = cell.getCellType();
                //перебираем возможные типы ячеек

                switch (cellType) {

                    case Cell.CELL_TYPE_STRING:
                        result += cell.getStringCellValue() + "=";
                        String s = cell.getStringCellValue(), sub = " ";
                        if (s.indexOf(sub) != -1){
                            finalNameResult = cell.getStringCellValue().replaceAll("[\\s]{2,}", " ");
                        }
                        else finalNameResult = cell.getStringCellValue();

                        if (finalNameResult.equals("имидазол")){
                            imidazolIndexCell =presentIndexCell;
                        }
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        if (presentIndexCell == imidazolIndexCell){
                            if (maximidazol< cell.getNumericCellValue()){
                                maximidazol=cell.getNumericCellValue();
                                maxSell=cellIndex;
                            }
                            double pass = Cell.CELL_TYPE_NUMERIC;
                        }
                        result += "[" + cell.getNumericCellValue() + "]";
                        break;

                    case Cell.CELL_TYPE_FORMULA:
                        result += "[" + cell.getNumericCellValue() + "]";
                        break;
                    default:
                        result += "|";
                        break;
                }
            }
            result += "\n";
        }

        return maxSell;
    }

}


public class Main {
    public static void main(String[] args) {
        int GreatInd=ExcelParser.findmaximidazolReturnCell("vichitFromSpectraImidazola.xls")-ExcelParser.findMaxFonReturnSell("vichitFromSpectraImidazola.xls");
        System.out.println(ExcelParser.result("vichitFromSpectraImidazola.xls","get.xls",GreatInd,ExcelParser.getLastIndex("vichitFromSpectraImidazola.xls")));
    }
}

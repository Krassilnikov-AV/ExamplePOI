package poi.excel;

import static com.sun.xml.bind.util.CalendarConv.formatter;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.ParsePosition;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Locale;
import javax.ejb.Local;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Класс считывающий данные столбца файла Excel в коллекцию List
 *
 * @author Aleks
 */
public class ReadingExcelData {

    /**
     *
     * @param args
     * @throws IOException
     */
    // выбрать столбец для чтения данных (для проверки/тестировниая)
    final static int code = 0;    // код (строка)
    final static int divID = 1;   // ID подразделения (число)
    final static int gpoupID = 2;   // ID группы  (число)
    final static int codeGroup = 3;   // код группы  (число)
    final static int group = 4;   // название группы (строка)
    final static int dateStart = 5;   // дата начала (дата)
    final static int timeStart = 6;   // время начала (время)
    final static int dateEnd = 7;   // дата завершения (дата)
    final static int timeEnd = 8;   // время завершения (время)
    final static int classID = 9;   // ID аудитории (число)
    final static int clasRum = 10;   // №аудитории или вариант (ОнЛайн) (число/строка)
    final static int typeLearn = 11;   // тип занятия (строка)
    final static int codeDirectionProgramm = 12;   // код-направление-программа (число-строка)
    final static int courseID = 13;   // ID курса (число)
    final static int discipline = 14;   // предмет/дисциплина/программа (число/строка)
    final static int period = 15;   // период (число)
    final static int teacherID = 16;   // ID преподавателя (число)
    final static int teacher = 17;   // преподаватель (строка)
    final static int periodDay = 18;   // период дней(число)
    final static int academHour = 19;   // академических часов (число)
    final static int academRecord = 20;   // академических записей (число)

 //   String fileName = "PrimerRaspisania.xlsx";
    String fileName = "fileToRead";
    LinkedList<String> columndata = null;

// основной метод класса для проверки считывания данных с таблицы

    public static void main(String[] args) throws IOException, ParseException {

        ReadExcelData code = new ReadExcelData();
        code.getDataStringIntegerDate(ReadExcelData.clasRum);
    //    code.getDataTime(ReadExcelData.timeEnd);
    }

    
// метод для получения строчных, целочисленных данных и даты в формате "число.месяц.год"
    /**
     * ! добавить метод, который будет получать путь и номер колонки, который вызывается в сервлете 
     * @param columnIndex
     * @return 
     */
    public List<String> getDataStringIntegerDate(int columnIndex)  {

        try {
            File f = new File(fileName);
            FileInputStream ios = new FileInputStream(f);
            XSSFWorkbook workbook = new XSSFWorkbook(ios);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            columndata = new LinkedList<>();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    if (row.getRowNum() > 0) { //фильтрация заголовков столбцов
                        if (cell.getColumnIndex() == columnIndex) {// соответствие индекса столбца
                            switch (cell.getCellType()) {
                                case Cell.CELL_TYPE_NUMERIC:
                                    String date = "dd.MM.yyyy";
                                    if (DateUtil.isCellDateFormatted(cell)) {  // получение данных даты                                      
                                        SimpleDateFormat sdfDate = new SimpleDateFormat(date);
                                        columndata.add(sdfDate.format(cell.getDateCellValue()));
                                    } else { // получение целочисленных данных
                                        columndata.add((int) cell.getNumericCellValue() + "");
                                    }
                                    break;
                                case Cell.CELL_TYPE_STRING: // получение строчных данных
                                    columndata.add(cell.getStringCellValue());
                                    break;
                            }
                        }
                    }
                }
            }
            ios.close();
            Iterator it = columndata.iterator();
            while (it.hasNext()) {
                System.out.println(it.next());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return columndata;
    }
// метод для получения данных даты в формате "ЧАСы : минуты"

    public List<String> getDataTime(int columnIndex) throws ParseException {
        try {
            File f = new File(fileName);
            FileInputStream ios = new FileInputStream(f);
            XSSFWorkbook workbook = new XSSFWorkbook(ios);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            columndata = new LinkedList<>();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    if (row.getRowNum() > 0) { //фильтрация заголовков столбцов
                        if (cell.getColumnIndex() == columnIndex) {// соответствие индекса столбца
                            switch (cell.getCellType()) {
                                case Cell.CELL_TYPE_NUMERIC:
                                    String time = "HH:mm";
                                    if (DateUtil.isCellDateFormatted(cell)) {// получение данных времени
                                        SimpleDateFormat sdfTime = new SimpleDateFormat(time, Locale.UK);
                                        columndata.add(sdfTime.format(cell.getDateCellValue()));
                                    }
                                    break;
                            }
                        }
                    }
                }
            }
            ios.close();
            Iterator it = columndata.iterator();
            while (it.hasNext()) {
                System.out.println(it.next());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return columndata;
    }
    public List<String> getDataStringIntegerDate(String pathToFile, int columnIndex)  {

        try {
            File f = new File(pathToFile);
            FileInputStream ios = new FileInputStream(f);
            XSSFWorkbook workbook = new XSSFWorkbook(ios);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            columndata = new LinkedList<>();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    if (row.getRowNum() > 0) { //фильтрация заголовков столбцов
                        if (cell.getColumnIndex() == columnIndex) {// соответствие индекса столбца
                            switch (cell.getCellType()) {
                                case Cell.CELL_TYPE_NUMERIC:
                                    String date = "dd.MM.yyyy";
                                    if (DateUtil.isCellDateFormatted(cell)) {  // получение данных даты                                      
                                        SimpleDateFormat sdfDate = new SimpleDateFormat(date);
                                        columndata.add(sdfDate.format(cell.getDateCellValue()));
                                    } else { // получение целочисленных данных
                                        columndata.add((int) cell.getNumericCellValue() + "");
                                    }
                                    break;
                                case Cell.CELL_TYPE_STRING: // получение строчных данных
                                    columndata.add(cell.getStringCellValue());
                                    break;
                            }
                        }
                    }
                }
            }
            ios.close();
            Iterator it = columndata.iterator();
            while (it.hasNext()) {
                System.out.println(it.next());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return columndata;
    }
}

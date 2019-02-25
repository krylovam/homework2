import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.*;

public class Homework {
    public static void main(String[] args) throws Exception {

        int constCount = 30;
        //генерируем количество людей
        Random rand = new Random();
        int count = 1 + rand.nextInt(constCount);
        //генерируем количество мужчин и женщин
        int countMen = rand.nextInt(count + 1);
        int countWomen = count - countMen;

        String dir = System.getProperty("user.dir") + "\\resources\\"; //путь к файлам
        String [] FilenameMen = {"SurnameMen.txt","NameMen.txt","PatronymicMen.txt"}; //имена файлов мужских ФИО
        String [] FilenameWomen = {"SurnameWomen.txt", "NameWomen.txt", "PatronymicWomen.txt"};//имена файлов женских ФИО
        String [] FilenameAddress = {"Country.txt", "City.txt", "Street.txt"}; //имена файлов адресов
        //наборы последовательностей рандомных номеров для списков
        ArrayList<ArrayList<Integer>> NumbersFIOMen = new ArrayList<ArrayList<Integer>>(3);
        ArrayList<ArrayList<Integer>> NumbersFIOWomen = new ArrayList<ArrayList<Integer>>(3);
        ArrayList<ArrayList<Integer>> NumbersAddress = new ArrayList<ArrayList<Integer>>(3);
        //наборы массивов строк, принятых из файлов
        ArrayList<ArrayList<String>> LineFIOMen = new ArrayList<ArrayList<String>>(3);
        ArrayList<ArrayList<String>> LineFIOWomen = new ArrayList<ArrayList<String>>(3);
        ArrayList<ArrayList<String>> LineAddress = new ArrayList<ArrayList<String>>(3);
        //массивы с генерируемыми данными
        ArrayList<String> INN = new ArrayList<String>(getINN(count));//ИНН
        ArrayList<Integer> House = new ArrayList<Integer>(getHouse(count));//Дом
        ArrayList<Integer> Flat = new ArrayList<Integer>(getFlat(count));//Квартира
        ArrayList<String> Index = new ArrayList<String>(getIndex(count));//Индекс
        ArrayList<GregorianCalendar> DateOfBirth = new ArrayList<GregorianCalendar>(getDayOfBirth(count));//Дата рождения
        ArrayList<Integer> Age = new ArrayList<Integer>(getAge(DateOfBirth));//Возраст
        ArrayList<Boolean> Gender = new ArrayList<Boolean>(getGender(countMen,countWomen));//Пол
        //генерируем номера выборки из файлов
        for (int i = 0; i < 3; i ++){
            NumbersFIOMen.add(getNumbers(countMen,constCount));
            NumbersFIOWomen.add(getNumbers(countWomen, constCount));
            NumbersAddress.add(getNumbers(count, constCount));
        }
        //сканируем нужные данные из файлов
        for (int i = 0; i < 3; i ++){
            LineFIOMen.add(getLineByLine(NumbersFIOMen.get(i), dir+FilenameMen[i]));
            LineFIOWomen.add(getLineByLine(NumbersFIOWomen.get(i),dir+FilenameWomen[i]));
            LineAddress.add(getLineByLine(NumbersAddress.get(i),dir+ FilenameAddress[i]));
            LineFIOMen.get(i).add("NULL");
            LineFIOWomen.get(i).add("NULL");
        }
        //Создание Excel файла
        Workbook Excel = new HSSFWorkbook();
        Sheet Sheet = Excel.createSheet("MyWork");
        FileOutputStream fos = new FileOutputStream("my.xls");
        //вывод в файд
        String [] Header = {"Фамилия", "Имя", "Отчество","Возраст","Пол","Дата рождения","ИНН","Индекс","Страна", "Город", "Улица","Дом","Квартира"};//заголовки
        int pointerMen = 0;
        for (int i = 0; i <= count; i ++) {
            Row row = Sheet.createRow(i);
            if (i == 0) {
                for (int j = 0; j < 13; j++) {
                    row.createCell(j).setCellValue(Header[j]);

                }
            } else {
                if (Gender.get(i-1) == false) {
                    row.createCell(4).setCellValue("М");
                }
                else{
                    row.createCell(4).setCellValue("Ж");
                }
                row.createCell(3).setCellValue(Age.get(i-1));
                row.createCell(5).setCellValue(DateOfBirth.get(i-1).get(Calendar.DATE) +"." + DateOfBirth.get(i-1).get(Calendar.MONTH)+"."+ DateOfBirth.get(i-1).get(Calendar.YEAR));
                row.createCell(6).setCellValue(INN.get(i-1));
                row.createCell(7).setCellValue(Index.get(i-1));
                row.createCell(11).setCellValue(House.get(i-1));
                row.createCell(12).setCellValue(Flat.get(i-1));
                for (int j = 0; j < 3; j ++){
                    row.createCell(j+8).setCellValue(LineAddress.get(j).get(i-1));
                }
                if (Gender.get(i-1) == false) {
                    for (int j = 0; j < 3; j++) {
                        row.createCell(j).setCellValue(LineFIOMen.get(j).get(pointerMen));
                    }
                    pointerMen++;
                }
                else{
                    for (int j = 0; j < 3; j++) {
                        row.createCell(j).setCellValue(LineFIOWomen.get(j).get(i - 1 - pointerMen));
                    }
                }
            }
        }

        Excel.write(fos);
        fos.close();
        //окончание работы с файлом
        System.out.println("Файл создан. Путь:" + System.getProperty("user.dir"));
    }

    private static ArrayList getNumbers(int countPart, int count) { //набор номеров для выборки из файла
        ArrayList<Integer> ArrayRandom = new ArrayList();
        int i = 0;
        while (i < countPart) {
            Random rand = new Random();
            int x = 1 +  rand.nextInt(count);
            if (!(ArrayRandom.contains(x))) {
                ArrayRandom.add(x);
                ++i;
            }
        }
        return ArrayRandom;

    }

    private static ArrayList getLineByLine(ArrayList Numbers, String filename) throws Exception{//считывание строк под нужными номерами строк
        BufferedReader bufferedReader = new BufferedReader(beginReading(filename));
        ArrayList<String> Lines = new ArrayList();
        String line;
        Integer pointLine = 1;
        while (bufferedReader.ready()){
            line = bufferedReader.readLine();
            boolean f = false;
            for (int i = 0;i < Numbers.size(); i ++){
                if (pointLine == Numbers.get(i)){
                    f = true;
                }
            }
            if (f) {
                Lines.add(line);
            }
            ++pointLine;
        }
        Collections.shuffle(Lines);
        return Lines;
    }
    private static BufferedReader beginReading(String filename)throws Exception{//правильное считывание из файла
        File file = new File(filename);
        String encoding = System.getProperty("console.encoding", "Cp1251");
        Reader fileReader = new InputStreamReader(new FileInputStream(file), encoding);
        BufferedReader bufferedReader = new BufferedReader(fileReader);
        return bufferedReader;
    }
    private static ArrayList<String> getINN(int count) throws Exception{//генерирование ИНН
        ArrayList<String> Ans = new ArrayList<String>();
        Random rand = new Random();
        for (int k = 0; k < count; k ++) {
            ArrayList<Integer> Inn = new ArrayList<Integer>(12);
            Inn.add(7);
            Inn.add(7);
            for (int i = 2; i < 10; i++) {
                Inn.add(rand.nextInt(9));
            }
            int kontrNumber1 = 0;
            int[] Arr1 = {7, 2, 4, 10, 3, 5, 9, 4, 6, 8};
            for (int i = 0; i < 10; i++) {
                kontrNumber1 += (Arr1[i] * Inn.get(i));
            }

            int x = (kontrNumber1 % 11);
            Inn.add(x % 10);
            int[] Arr2 = {3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8};
            int kontrNumber2 = 0;
            for (int i = 0; i < 11; i++) {
                kontrNumber2 += (Arr2[i] * Inn.get(i));
            }
            x = kontrNumber2 % 11;
            Inn.add(x % 10);
            String str = new String();
            for (int i : Inn) {
                str += i;
            }

            Ans.add(str);

        }
        return Ans;
    }
    private static ArrayList<Boolean> getGender(int countMen,int countWomen){//генерирование пола
        ArrayList Ans = new ArrayList();
        for (int i = 0; i < countMen; i ++){
            Ans.add(false);
        }
        for (int i = 0; i < countWomen; i ++){
            Ans.add(true);
        }
        Collections.shuffle(Ans);
        return Ans;
    }
    private static ArrayList<String> getIndex(int count){//генерирование индекса
        Random rand = new Random();
        ArrayList<Integer> Index = new ArrayList<Integer>();
        for (int i = 0; i < count; i ++){
            Index.add(100000 + rand.nextInt(100000));
        }
        ArrayList<String> Ans = new ArrayList<String>();
        for (int i:Index){
            Ans.add("" + i);
        }
        return Ans;
    }
    private static ArrayList<Integer> getHouse(int count){//генерирование номера дома
        Random rand = new Random();
        ArrayList<Integer> Ans = new ArrayList<Integer>();
        for (int i = 0; i < count; i ++){
            Ans.add(1+ rand.nextInt(100));
        }
        return Ans;
    }
    private static ArrayList<Integer> getFlat(int count){//генерирование номера квартиры
        Random rand = new Random();
        ArrayList<Integer> Ans = new ArrayList<Integer>();
        for (int i = 0; i < count; i ++){
            Ans.add(1+ rand.nextInt(1000));
        }
        return Ans;
    }
    private static ArrayList<GregorianCalendar> getDayOfBirth(int count){//генерирование даты рождения
        ArrayList<GregorianCalendar> Ans = new ArrayList<GregorianCalendar>();
        for (int i = 0; i < count; i ++) {
            GregorianCalendar gc = new GregorianCalendar();
            int year = randBetween(1900, 2018);
            gc.set(gc.YEAR, year);
            int dayOfYear = randBetween(1, gc.getActualMaximum(gc.DAY_OF_YEAR));
            gc.set(gc.DAY_OF_YEAR, dayOfYear);
            Ans.add(gc);
        }
        return Ans;
    }
    public static int randBetween(int start, int end) {
        return start + (int)Math.round(Math.random() * (end - start));
    }
    private static ArrayList<Integer> getAge(ArrayList<GregorianCalendar> DateOfBirth){//подсчет возраста
        ArrayList<Integer> Ans = new ArrayList<Integer>();
        GregorianCalendar today = new GregorianCalendar();
        for (int i = 0; i < DateOfBirth.size();i ++) {
            Calendar dob = DateOfBirth.get(i);
            Calendar now = today;
            int year1 = now.get(Calendar.YEAR);
            int year2 = dob.get(Calendar.YEAR);
            int age = year1 - year2;
            int month1 = now.get(Calendar.MONTH);
            int month2 = dob.get(Calendar.MONTH);
            if (month2 > month1) {
                age--;
            } else if (month1 == month2) {
                int day1 = now.get(Calendar.DAY_OF_MONTH);
                int day2 = dob.get(Calendar.DAY_OF_MONTH);
                if (day2 > day1) {
                    age--;
                }
            }
            Ans.add(age);
        }
        return Ans;
    }
}




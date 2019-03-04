package main.java.Api;

import com.itextpdf.text.Document;
import com.itextpdf.text.Font;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.*;

public class Homework {
    public static void main(String[] args) throws Exception {
        Random rand = new Random();
        final int constCount = 30;
        int count = rand.nextInt(constCount);
        int countMen = rand.nextInt(count + 1);
        int countWomen = count - countMen;
        String dir = System.getProperty("user.dir") + File.separator+"src"+File.separator +"main"+File.separator+"resources"+File.separator;
        ArrayList<Integer> NumbersFIOMen = new ArrayList(getNumbers(countMen,constCount));
        ArrayList<Integer> NumbersFIOWomen = new ArrayList(getNumbers(countWomen,constCount));
        ArrayList<Integer> NumbersAddress = new ArrayList(getNumbers(count, constCount));
        ArrayList<String> SurnameMen = new ArrayList<String>(getLineByLine(NumbersFIOMen,dir + "SurnameMen.txt"));
        ArrayList<String> NameMen = new ArrayList<String>(getLineByLine(NumbersFIOMen,dir + "NameMen.txt"));
        ArrayList<String> PatronymicMen = new ArrayList<String>(getLineByLine(NumbersFIOMen,dir + "PatronymicMen.txt"));
        ArrayList<String> SurnameWomen = new ArrayList<String>(getLineByLine(NumbersFIOWomen,dir + "SurnameWomen.txt"));
        ArrayList<String> NameWomen = new ArrayList<String>(getLineByLine(NumbersFIOWomen,dir + "NameWomen.txt"));
        ArrayList<String> PatronymicWomen = new ArrayList<String>(getLineByLine(NumbersFIOWomen,dir + "PatronymicWomen.txt"));
        ArrayList<User> UserMas = new ArrayList<User>();
        for (int i = 0; i < count; i ++){
            UserMas.add(new User());
            UserMas.get(i).setCountry(User.getAddress(dir + "Country.txt", constCount));
            UserMas.get(i).setRegion(User.getAddress(dir + "Region.txt",constCount));
            UserMas.get(i).setCity(User.getAddress(dir +"City.txt", constCount));
            UserMas.get(i).setStreet(User.getAddress(dir + "Street.txt",constCount));
            UserMas.get(i).setPostcode();
            UserMas.get(i).setHouse();
            UserMas.get(i).setFlat();
            UserMas.get(i).setINN();
            UserMas.get(i).setDayOfBirth();
            UserMas.get(i).setAge(UserMas.get(i).getDob());
        }
        for(int i = 0; i < countMen; i ++){
            UserMas.get(i).setLastName(SurnameMen.get(i));
            UserMas.get(i).setFirstName(NameMen.get(i));
            try {
                UserMas.get(i).setPatronymicName(PatronymicMen.get(i));
            }catch(IndexOutOfBoundsException e){}
            UserMas.get(i).setGender("M");
        }
        for(int i = 0;i < countWomen; i ++){
            UserMas.get(i+ countMen).setLastName(SurnameWomen.get(i));
            UserMas.get(i + countMen).setFirstName(NameWomen.get(i));
            try {
                UserMas.get(i + countMen).setPatronymicName(PatronymicWomen.get(i));
            }catch(IndexOutOfBoundsException e){}
            UserMas.get(i + countMen).setGender("Ж");
        }
        Collections.shuffle(UserMas);
        String [] Header = new String[]{"Фамилия", "Имя", "Отчество", "Возраст", "Пол", "Дата рождения", "ИНН", "Индекс", "Страна","Область", "Город", "Улица", "Дом", "Квартира"};
        writeToExcel(UserMas,count,Header);
        writeToPdf(UserMas,count,Header);
    }

    private static ArrayList getNumbers(int countPart, int count) {
        ArrayList<Integer> ArrayRandom = new ArrayList();
        int i = 0;

        while(i < countPart) {
            Random rand = new Random();
            int x = 1 + rand.nextInt(count);
            if (!ArrayRandom.contains(x)) {
                ArrayRandom.add(x);
                ++i;
            }
        }

        return ArrayRandom;
    }
    private static ArrayList getLineByLine(ArrayList Numbers, String filename) throws Exception {
        BufferedReader bufferedReader = new BufferedReader(beginReading(filename));
        ArrayList<String> Lines = new ArrayList();

        for(Integer pointLine = 1; bufferedReader.ready(); pointLine = pointLine + 1) {
            String line = bufferedReader.readLine();
            boolean f = false;

            for(int i = 0; i < Numbers.size(); ++i) {
                if (pointLine == Numbers.get(i)) {
                    f = true;
                }
            }

            if (f) {
                Lines.add(line);
            }
        }

        Collections.shuffle(Lines);
        return Lines;
    }
    private static BufferedReader beginReading(String filename) throws Exception {
        File file = new File(filename);
        String encoding = System.getProperty("console.encoding", "Cp1251");
        Reader fileReader = new InputStreamReader(new FileInputStream(file), encoding);
        BufferedReader bufferedReader = new BufferedReader(fileReader);
        return bufferedReader;
    }
    private static void getUser(User user, int constCount)throws Exception {
        String dir = System.getProperty("user.dir") + File.separator + "src" + File.separator + "main" + File.separator + "resources" + File.separator;
        user.setCountry(User.getAddress(dir +"Country.txt", constCount));
        user.setRegion(User.getAddress(dir + "Region.txt",constCount));
        user.setCity(User.getAddress(dir + "City.txt", constCount));
        user.setStreet(User.getAddress(dir + "Street.txt",constCount));
        user.setPostcode();
        user.setHouse();
        user.setFlat();
        user.setINN();
        user.setDayOfBirth();
        user.setAge(user.getDob());
    }
    private static void writeToPdf(ArrayList<User> UserMas,int count, String[] Header){
        Document document = new Document();
        try{
            PdfWriter.getInstance(document,new FileOutputStream("my.pdf"));
            document.open();
            BaseFont bf =
                    BaseFont.createFont("C:\\Windows\\Fonts\\arial.ttf",BaseFont.IDENTITY_H,BaseFont.EMBEDDED);
            Font font = new Font(bf);

            PdfPTable table = new PdfPTable(14);
            table.setWidthPercentage(100);
            table.setSpacingBefore(0f);
            table.setSpacingAfter(0f);

            for(int i = 0; i < 14; i ++){
                table.addCell(new Paragraph(Header[i],font));
            }
            for(int i = 0; i < count; i ++){
                table.addCell(new Paragraph(UserMas.get(i).getLastName(),font));
                table.addCell(new Paragraph(UserMas.get(i).getFirstName(),font));
                table.addCell(new Paragraph(UserMas.get(i).getPatronymicName(),font));
                table.addCell(new Paragraph(Long.toString(UserMas.get(i).getAge()),font));
                table.addCell(new Paragraph(UserMas.get(i).getGender(),font));
                GregorianCalendar gc = UserMas.get(i).getDob();
                table.addCell(gc.get(Calendar.DAY_OF_MONTH) + "." + gc.get(Calendar.MONTH) + "." + gc.get(Calendar.YEAR));
                table.addCell(new Paragraph(UserMas.get(i).getINN(),font));
                table.addCell(new Paragraph(UserMas.get(i).getPostcode()));
                table.addCell(new Paragraph(UserMas.get(i).getCountry(),font));
                table.addCell(new Paragraph(UserMas.get(i).getRegion(),font));
                table.addCell(new Paragraph(UserMas.get(i).getCity(),font));
                table.addCell(new Paragraph(UserMas.get(i).getStreet(),font));
                table.addCell(new Paragraph(UserMas.get(i).getHouse(),font));
                table.addCell(new Paragraph(Long.toString(UserMas.get(i).getFlat()),font));

            }

            document.add(table);
            document.close();
            System.out.println("Файл создан. Путь:" + System.getProperty("user.dir") + File.separator + "my.pdf");
        }catch (Exception e){
            e.printStackTrace();
        }
    }
    private static void writeToExcel(ArrayList<User>UserMas,int count,String[] Header)throws Exception{
        Workbook Excel = new HSSFWorkbook();
        Sheet Sheet = Excel.createSheet("MyWork");
        try {
            FileOutputStream fos = new FileOutputStream("my.xls");
            for (int i = 0; i <= count; ++i) {
                Row row = Sheet.createRow(i);
                if (i == 0) {
                    for (int j = 0; j < 14; j++) {
                        row.createCell(j).setCellValue(Header[j]);
                    }
                } else {
                    row.createCell(0).setCellValue(UserMas.get(i - 1).getLastName());
                    row.createCell(1).setCellValue(UserMas.get(i - 1).getFirstName());
                    row.createCell(2).setCellValue(UserMas.get(i - 1).getPatronymicName());
                    row.createCell(3).setCellValue(UserMas.get(i - 1).getAge());
                    row.createCell(4).setCellValue(UserMas.get(i - 1).getGender());
                    GregorianCalendar gc = UserMas.get(i - 1).getDob();
                    row.createCell(5).setCellValue(gc.get(Calendar.DAY_OF_MONTH) + "." + gc.get(Calendar.MONTH) + "." + gc.get(Calendar.YEAR));
                    row.createCell(6).setCellValue(UserMas.get(i - 1).getINN());
                    row.createCell(7).setCellValue(UserMas.get(i - 1).getPostcode());
                    row.createCell(8).setCellValue(UserMas.get(i - 1).getCountry());
                    row.createCell(9).setCellValue(UserMas.get(i - 1).getRegion());
                    row.createCell(10).setCellValue(UserMas.get(i - 1).getCity());
                    row.createCell(11).setCellValue(UserMas.get(i - 1).getStreet());
                    row.createCell(12).setCellValue(UserMas.get(i - 1).getHouse());
                    row.createCell(13).setCellValue(UserMas.get(i - 1).getFlat());
                }

            }
            Excel.write(fos);
            fos.close();
            System.out.println("Файл создан. Путь:" + System.getProperty("user.dir") + File.separator + "my.xls");
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}

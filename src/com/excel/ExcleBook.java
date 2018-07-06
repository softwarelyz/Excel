package com.excel;

import java.io.File;
import java.util.*;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class ExcleBook {
	/** 
     * ���Book����е����Ĳ��� 
     * @param list 
     */  
    public void excleOut(List<Book> list) {  
        WritableWorkbook book = null;  
        try {  
            // ����һ��excle����  
            book = Workbook.createWorkbook(new File("F:/excleTest/book.xls"));  
            // ͨ��excle���󴴽�һ��ѡ�����  
            WritableSheet sheet = book.createSheet("sheet1", 0);  
            // ����һ����Ԫ����� �� �� ֵ  
            // Label label = new Label(0, 2, "test");  
            for (int i = 0; i < list.size(); i++) {  
                Book book2 = list.get(i);  
                Label label1 = new Label(0, i, String.valueOf(book2.getId()));  
                Label label2 = new Label(1, i, book2.getName());  
                Label label3 = new Label(2, i, book2.getAuthor());  
  
                // �������õĵ�Ԫ��������ѡ���  
                sheet.addCell(label1);  
                sheet.addCell(label2);  
                sheet.addCell(label3);  
            }  
            // д��Ŀ��·��  
            book.write();  
        } catch (Exception e) {  
            e.printStackTrace();  
        } finally {  
            try {  
                book.close();  
            } catch (Exception e) {  
                // TODO Auto-generated catch block  
                e.printStackTrace();  
            }  
        }  
    }  
  
    /** 
     * ���Book����е���Ĳ��� 
     * @return 
     */  
    public List<Book> excleIn() {  
        List<Book> list = new ArrayList<>();  
        Workbook workbook = null;  
        try {  
            // ��ȡEcle����  
            workbook = Workbook.getWorkbook(new File("F:/excleTest/book.xls"));  
            // ��ȡѡ����� ��0��ѡ�  
            Sheet sheet = workbook.getSheet(0);  
            // ѭ��ѡ��е�ֵ  
            for (int i = 0; i < sheet.getRows(); i++) {  
                Book book = new Book();  
                // ��ȡ��Ԫ�����  
                Cell cell0 = sheet.getCell(0, i);  
                // ȡ�õ�Ԫ���ֵ,�����õ�������  
                book.setId(Integer.valueOf(cell0.getContents()));  
                // ��ȡ��Ԫ�����Ȼ��ȡ�õ�Ԫ���ֵ,�����õ�������  
                book.setName(sheet.getCell(1, i).getContents());  
                book.setAuthor(sheet.getCell(2, i).getContents());  
                list.add(book);  
            }  
        } catch (Exception e) {  
            e.printStackTrace();  
        } finally {  
            workbook.close();  
        }  
        return list;  
    }  
  
    public static void main(String[] args) {  
        ExcleBook book = new ExcleBook();  
        List<Book> list = new ArrayList<>(); 
        
        //��������
        /*Book book2 = new Book();  
        book2.setId(1);  
        book2.setName("�鱾��1");  
        book2.setAuthor("����");  
        Book book3 = new Book();  
        book3.setId(2);  
        book3.setName("�鱾��2");  
        book3.setAuthor("����");  
        list.add(book2);  
        list.add(book3);  
        book.excleOut(list);  
        List<Book> books = book.excleIn();  
        for (Book bo : books) {  
            System.out.println(bo.getId() + " " + bo.getName() + " " + bo.getAuthor());  
        } */
        
        list = book.excleIn();
        
        for(int i = 0;i<list.size();i++){
        	System.out.println(list.get(i).getName());
        }
    }  
}

package com.SuperCombination.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Vector;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.VerticalAlignment;
import jxl.read.biff.BiffException;
import jxl.write.Alignment;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 * ���ļ��Ĳ���
 * @author SkyFreecss
 *
 */
@SuppressWarnings("deprecation")
public class ExcelDao {
       static Log log = LogFactory.getLog("ExcelDao.class");   
	
       public void readExcel(Vector<String> vecfile) throws IOException
       {
			
			 WritableWorkbook wwb = Workbook.createWorkbook(new File("F://TestFile//Excel//New_Test.xls"));
             WritableSheet sheet_1 = wwb.createSheet("�ܱ�",0);
             
    	   int rowsNum=0;
    	   int columnsNum=0;
    	   int a;
    	   int b;
    	   log.info("���ڻ�ȡ�ļ���������,���Եȣ�");
    	   for(int i=0;i<vecfile.size();i++)
    	   {
    		   String filename = vecfile.elementAt(i);
    		   //System.out.println(filename);  
    		   try {
				InputStream is = new FileInputStream(filename);
				
				//��������������
				Workbook wk = Workbook.getWorkbook(is);
				
				//��ù������ĸ���s
				wk.getNumberOfSheets();
				Sheet oFirstSheet = wk.getSheet(0);
				
				int rows = oFirstSheet.getRows();//��ȡ���������������
				int columns = oFirstSheet.getColumns();//��ȡ���������������
				
				//--------------------------------------------------------
				
				a = rowsNum;
				b = columnsNum;
				System.out.println(a+""+b);
				log.info("�������  "+filename+" �ļ������ݣ�");
				for(int m=0;m<rows;m++)
				{
					for(int n=0;n<columns;n++)
					{
						
						Cell ocell = oFirstSheet.getCell(n,m);
						System.out.println(ocell.getContents());
						System.out.println("columns = "+columns+" "+"rows = "+rows);
						System.out.println("n = "+n+" "+"m = "+m);
						
						/*
						Label label = new Label(columnsNum,rowsNum,ocell.getContents());
						sheet_1.addCell(label);
						wwb.write();
						wwb.close();
						*/
						
					}
					
					rowsNum = a;
					columnsNum = b;
				}
				
				rowsNum = rows;
				columnsNum = columns;
				
			} catch (BiffException | IOException  e) {
				log.error("��ȡ�����������쳣�����������");
				e.printStackTrace();
			}
    	   }
			log.info("����������ļ���д�룡");
    	   
    	   /*
    	   try {
    		   log.info("��ȡ�ļ�����");
    		 //��ȡExcel�ļ�����������
			InputStream is = new FileInputStream(pathfile);
			
			//��������������
			Workbook wk = Workbook.getWorkbook(is);
			
			//��ù������ĸ���
			wk.getNumberOfSheets();
			
			Sheet oFirstSheet = wk.getSheet(0);//ʹ����������ʽ��õ�һ��������Ҳ����ʹ��wk.getSheet(sheetName)
			
			int rows = oFirstSheet.getRows();//��ȡ������
			int columns = oFirstSheet.getColumns();//��ȡ������
			
			  log.info("��������������!");
			for(int i=0;i<rows;i++)
			{
				for(int j=0;j<columns;j++)
				{
			        Cell ocell = oFirstSheet.getCell(j,i);
			        System.out.print(ocell.getContents());
				}
				System.out.println();
			}
			log.info("�����ɣ�");
		} catch (BiffException | IOException e) {
			log.info("���ִ���");
			e.printStackTrace();
		}
		*/
       }
       
       @SuppressWarnings({"unused"})
	public void writeExcel()
       {

    	   try {
    		   String cmd[] = {"C:\\Program Files (x86)\\OpenOffice 4\\program\\scalc.exe","F:\\TestFile\\Excel\\New_Test.xls"};
    		   Process p = Runtime.getRuntime().exec(cmd);
    	   //�����������������ļ����������½���
    		   log.info("���ڴ���������������");
			WritableWorkbook wwb = Workbook.createWorkbook(new File("F://TestFile//Excel/New_Test.xls"));
			   log.info("�����������ɹ���");
			
		   //�½���������󣬲�������Ϊ�ڼ�ҳ��
			WritableSheet sheet_1 = wwb.createSheet("�ܱ�",0);
			WritableFont font1 = new WritableFont(WritableFont.ARIAL,10,WritableFont.BOLD,true);
			
			WritableCellFormat titleformat1 = new WritableCellFormat(font1);
			titleformat1.setVerticalAlignment(VerticalAlignment.CENTRE);//��Ԫ����ж���
			titleformat1.setAlignment(Alignment.CENTRE);
			titleformat1.setBackground(jxl.format.Colour.SKY_BLUE);//��Ԫ�񱳾�ɫ
			titleformat1.setWrap(true);//�Ƿ��Զ�����

			
		    sheet_1.setColumnView(0,10);//ָ����Ԫ����
		    sheet_1.setColumnView(1,10);
		    sheet_1.setColumnView(2,30);
		    sheet_1.setColumnView(3,10);
		    sheet_1.setColumnView(4,80);
		    sheet_1.setColumnView(5,40);
		    sheet_1.setColumnView(6,10);
		    sheet_1.setColumnView(7,10);
		    
			sheet_1.setRowView(0,500);//ָ����Ԫ�񳤶�
		   //������Ԫ�����
		    Label label1 = new Label(0,0,"������ҵ��",titleformat1);
		    Label label2 = new Label(1,0,"�����Ŀ������",titleformat1);
		    Label label3 = new Label(2,0,"ģ�����������",titleformat1);
		    Label label4 = new Label(3,0,"������Ա",titleformat1);
		    Label label5 = new Label(4,0,"���幤������",titleformat1);
		    Label label6 = new Label(5,0,"�ƻ���������",titleformat1);
		    Label label7 = new Label(6,0,"ʵ���������",titleformat1);
		    Label label8 = new Label(7,0,"������",titleformat1);
		   /*
		   Label label3 = new Label(2,0,"ģ�����������");
		   Label label4 = new Label(3,0,"������Ա");
		   Label label5 = new Label(4,0,"���幤������");
		   Label label6 = new Label(5,0,"�ƻ���������");
		   Label label7 = new Label(6,0,"ʵ���������");
		   Label label8 = new Label(7,0,"������");
		   */
		   sheet_1.addCell(label1);
		   sheet_1.addCell(label2);
		   sheet_1.addCell(label3);
		   sheet_1.addCell(label4);
		   sheet_1.addCell(label5);
		   sheet_1.addCell(label6);
		   sheet_1.addCell(label7);
		   sheet_1.addCell(label8);
		   
	
		   
		   log.info("���ڴ�����Ԫ��");
		   wwb.write();
		   wwb.close();
		   log.info("������Ԫ��ɹ���");
		} catch (IOException | WriteException e) {
			log.error("����������ʧ�ܣ�----------δ֪����");
			e.printStackTrace();
		}
       }
       
       public static void main(String args[])
       {
           ExcelDao ed = new ExcelDao();
    	   ed.writeExcel();
       }
}



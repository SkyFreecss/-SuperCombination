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
 * 表文件的操作
 * @author SkyFreecss
 *
 */
@SuppressWarnings("deprecation")
public class ExcelDao {
       static Log log = LogFactory.getLog("ExcelDao.class");   
	
       public void readExcel(Vector<String> vecfile) throws IOException
       {
			
			 WritableWorkbook wwb = Workbook.createWorkbook(new File("F://TestFile//Excel//New_Test.xls"));
             WritableSheet sheet_1 = wwb.createSheet("周报",0);
             
    	   int rowsNum=0;
    	   int columnsNum=0;
    	   int a;
    	   int b;
    	   log.info("正在获取文件的输入流,请稍等！");
    	   for(int i=0;i<vecfile.size();i++)
    	   {
    		   String filename = vecfile.elementAt(i);
    		   //System.out.println(filename);  
    		   try {
				InputStream is = new FileInputStream(filename);
				
				//声明工作薄对象
				Workbook wk = Workbook.getWorkbook(is);
				
				//获得工作薄的个数s
				wk.getNumberOfSheets();
				Sheet oFirstSheet = wk.getSheet(0);
				
				int rows = oFirstSheet.getRows();//获取工作表的总行数。
				int columns = oFirstSheet.getColumns();//获取工作表的总列数。
				
				//--------------------------------------------------------
				
				a = rowsNum;
				b = columnsNum;
				System.out.println(a+""+b);
				log.info("正在输出  "+filename+" 文件的内容！");
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
				log.error("获取输入流出现异常，程序崩溃！");
				e.printStackTrace();
			}
    	   }
			log.info("已完成所有文件的写入！");
    	   
    	   /*
    	   try {
    		   log.info("获取文件对象");
    		 //获取Excel文件输入流对象
			InputStream is = new FileInputStream(pathfile);
			
			//声明工作薄对象
			Workbook wk = Workbook.getWorkbook(is);
			
			//获得工作薄的个数
			wk.getNumberOfSheets();
			
			Sheet oFirstSheet = wk.getSheet(0);//使用索引的形式获得第一个工作表，也可以使用wk.getSheet(sheetName)
			
			int rows = oFirstSheet.getRows();//获取总行数
			int columns = oFirstSheet.getColumns();//获取总列数
			
			  log.info("正在输出表格内容!");
			for(int i=0;i<rows;i++)
			{
				for(int j=0;j<columns;j++)
				{
			        Cell ocell = oFirstSheet.getCell(j,i);
			        System.out.print(ocell.getContents());
				}
				System.out.println();
			}
			log.info("输出完成！");
		} catch (BiffException | IOException e) {
			log.info("出现错误！");
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
    	   //创建工作薄对象，若文件不存在则新建。
    		   log.info("正在创建工作薄。。。");
			WritableWorkbook wwb = Workbook.createWorkbook(new File("F://TestFile//Excel/New_Test.xls"));
			   log.info("创建工作薄成功！");
			
		   //新建工作表对象，并声明其为第几页。
			WritableSheet sheet_1 = wwb.createSheet("周报",0);
			WritableFont font1 = new WritableFont(WritableFont.ARIAL,10,WritableFont.BOLD,true);
			
			WritableCellFormat titleformat1 = new WritableCellFormat(font1);
			titleformat1.setVerticalAlignment(VerticalAlignment.CENTRE);//单元格居中对齐
			titleformat1.setAlignment(Alignment.CENTRE);
			titleformat1.setBackground(jxl.format.Colour.SKY_BLUE);//单元格背景色
			titleformat1.setWrap(true);//是否自动换行

			
		    sheet_1.setColumnView(0,10);//指定单元格宽度
		    sheet_1.setColumnView(1,10);
		    sheet_1.setColumnView(2,30);
		    sheet_1.setColumnView(3,10);
		    sheet_1.setColumnView(4,80);
		    sheet_1.setColumnView(5,40);
		    sheet_1.setColumnView(6,10);
		    sheet_1.setColumnView(7,10);
		    
			sheet_1.setRowView(0,500);//指定单元格长度
		   //创建单元格对象
		    Label label1 = new Label(0,0,"归属事业部",titleformat1);
		    Label label2 = new Label(1,0,"设计项目或需求",titleformat1);
		    Label label3 = new Label(2,0,"模块或需求名称",titleformat1);
		    Label label4 = new Label(3,0,"参与人员",titleformat1);
		    Label label5 = new Label(4,0,"具体工作内容",titleformat1);
		    Label label6 = new Label(5,0,"计划工作周期",titleformat1);
		    Label label7 = new Label(6,0,"实际完成天数",titleformat1);
		    Label label8 = new Label(7,0,"完成情况",titleformat1);
		   /*
		   Label label3 = new Label(2,0,"模块或需求名称");
		   Label label4 = new Label(3,0,"参与人员");
		   Label label5 = new Label(4,0,"具体工作内容");
		   Label label6 = new Label(5,0,"计划工作周期");
		   Label label7 = new Label(6,0,"实际完成天数");
		   Label label8 = new Label(7,0,"完成情况");
		   */
		   sheet_1.addCell(label1);
		   sheet_1.addCell(label2);
		   sheet_1.addCell(label3);
		   sheet_1.addCell(label4);
		   sheet_1.addCell(label5);
		   sheet_1.addCell(label6);
		   sheet_1.addCell(label7);
		   sheet_1.addCell(label8);
		   
	
		   
		   log.info("正在创建单元格！");
		   wwb.write();
		   wwb.close();
		   log.info("创建单元格成功！");
		} catch (IOException | WriteException e) {
			log.error("创建工作薄失败！----------未知错误！");
			e.printStackTrace();
		}
       }
       
       public static void main(String args[])
       {
           ExcelDao ed = new ExcelDao();
    	   ed.writeExcel();
       }
}



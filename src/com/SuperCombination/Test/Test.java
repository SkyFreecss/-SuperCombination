package com.SuperCombination.Test;

import java.io.File;
import java.io.IOException;
import java.util.Vector;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

/**
 * �����õ�
 * @author SkyFreecss
 *
 */
public class Test {
       static Log log = LogFactory.getLog("log.class");       
	   public static void main(String args[])
	   {
		   ListFiles lf = new ListFiles();
		   ExcelDao ed = new ExcelDao();
		   Vector<String> vecfile =  lf.printListFiles(new  File("F:/TestFile/Excel/0103-0106"),".xls");
		   try {
		if(vecfile!=null)
		{
			log.info("vecfile��Ϊ�գ�");
			ed.readExcel(vecfile);
		}
		log.error("vecfileΪ�գ�");
		} catch (IOException e) {
			e.printStackTrace();
		}
	   }
}

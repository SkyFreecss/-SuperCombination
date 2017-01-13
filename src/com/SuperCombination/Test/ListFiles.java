package com.SuperCombination.Test;

import java.io.File;
import java.util.Vector;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

/**
 * �г���ǰĿ¼�µ������ļ���
 * @author SkyFreecss
 *
 */
public class ListFiles {
	
       static Log log = LogFactory.getLog("ListFiles.class");
	   
       public static void main(String args[])
       {
    	   printListFiles(new  File("F:/TestFile/Excel/0103-0106"),".xls");
       }
       
	/**
	 * �г��ļ����洢��Vector��.
	 * @param file
	 */
	   public static Vector<String> printListFiles(File file,String str)
	   {
		   log.info("�빺�����棡");
		   log.info("������ֻ֧��xls�ļ���лл��");
		   Vector<String> vecFile = new Vector<String>();
		   if(file!=null)
		   {
			   log.info("���ڴ򿪸�Ŀ¼��");
			   file.isDirectory();
			   File[] fileArray = file.listFiles();
			   if(fileArray!=null)
			   {
				   log.info("���ڻ�ȡ�����ļ���");
				   for(int i = 0;i<fileArray.length;i++)
				   {
					   
					   if(!fileArray[i].isDirectory())
					   {
						   String tempName = fileArray[i].getName();
						   //�ж��Ƿ�Ϊ.xls��β��
						   if(tempName.trim().toLowerCase().endsWith(str))
						   {
							   System.out.println(fileArray[i]);
							   String fileName = fileArray[i].toString();
							   vecFile.add(fileName);
						   }
					   }
				   }
				   log.info("�ļ���ȡ��ɣ�");
			   }
		   }
		   else
		   {
			   log.error("·������");
		   }
		   return vecFile;
	   }
}

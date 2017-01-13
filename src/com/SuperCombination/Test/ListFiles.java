package com.SuperCombination.Test;

import java.io.File;
import java.util.Vector;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

/**
 * 列出当前目录下的所有文件！
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
	 * 列出文件并存储进Vector中.
	 * @param file
	 */
	   public static Vector<String> printListFiles(File file,String str)
	   {
		   log.info("请购买正版！");
		   log.info("本工具只支持xls文件，谢谢！");
		   Vector<String> vecFile = new Vector<String>();
		   if(file!=null)
		   {
			   log.info("正在打开该目录！");
			   file.isDirectory();
			   File[] fileArray = file.listFiles();
			   if(fileArray!=null)
			   {
				   log.info("正在获取所需文件！");
				   for(int i = 0;i<fileArray.length;i++)
				   {
					   
					   if(!fileArray[i].isDirectory())
					   {
						   String tempName = fileArray[i].getName();
						   //判断是否为.xls结尾！
						   if(tempName.trim().toLowerCase().endsWith(str))
						   {
							   System.out.println(fileArray[i]);
							   String fileName = fileArray[i].toString();
							   vecFile.add(fileName);
						   }
					   }
				   }
				   log.info("文件获取完成！");
			   }
		   }
		   else
		   {
			   log.error("路径有误！");
		   }
		   return vecFile;
	   }
}

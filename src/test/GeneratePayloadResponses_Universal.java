package test;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.PrintWriter;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import javax.swing.JOptionPane;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GeneratePayloadResponses_Universal
{
	public static void main(String[] args) throws Exception
	{
		try
		{
			long start = System.currentTimeMillis();
			String timeStamp = "";
			String current = System.getProperty("user.dir");
			String filePathMain = current+"\\ObjectIDs.xlsx";	
			String textfilePath = current+"\\textfile.txt";
			Timestamp timestamp = new Timestamp(System.currentTimeMillis());
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy_MMM_dd_HH_mm_ss");
			timeStamp = sdf.format(timestamp);				
			ArrayList<String> varNamesInText = new ArrayList<String>();
			HashMap<String,Integer> varMatchInExcel = new HashMap<String,Integer>();
			FileInputStream fis = null;
			try
			{
				fis = new FileInputStream(filePathMain);
			}
			catch(Exception e1)
			{
				throw new Exception("Unable to find 'ObjectIDs.xlsx'!");
			}
			XSSFWorkbook w = new XSSFWorkbook(fis);
			fis.close();
			XSSFSheet s = w.getSheet("ObjectIDs");
			String vNameText="";
			BufferedReader br = null;
			FileReader fr = null;
			int noOfLines = 0, lineNo=1, totalRowsToProcess=0;
			
			
			//Get the actual count of number of lines
			try
			{
				fr = new FileReader(textfilePath);
				br = new BufferedReader(fr);
				while(br.readLine()!=null)
				{
					noOfLines++;
				}
			}
			catch(Exception e){}
			
			
			//Store the text file into this string array
			String[] actualTextFile = new String[noOfLines];		
			try
			{
				int counter=0;
				String cfile="";
				fr = new FileReader(textfilePath);
				br = new BufferedReader(fr);
				while((cfile=br.readLine())!=null)
				{
					actualTextFile[counter++]=cfile;
				}	
			}
			catch(Exception e){
				w.close();
				throw new Exception("Unable to find 'textfile.txt'!");
			}
			br.close();
			fr.close();
			
			
			//Store the excel column name and corresponding column number in varMatchInExcel
			for(int i=0; i<s.getRow(0).getPhysicalNumberOfCells(); i++)
			{
				String ss = "";
				try
				{
					ss = s.getRow(0).getCell(i).getStringCellValue().trim();
				}
				catch(Exception e)
				{
					ss = s.getRow(0).getCell(i).getRawValue().trim();
				}			
				if(!varMatchInExcel.containsValue(ss))
				{
					varMatchInExcel.put(ss,i);
				}
			}
			
			
			
			
			//Add the variable names present in text file into arraylist varNamesInText
			lineNo=1;
			for(String a : actualTextFile)
			{
				if(a.contains("${"))
				{
					try
					{
						vNameText = (((a.split("\\$\\{"))[1]).split("\\}\\$"))[0].trim();
						if(vNameText.equalsIgnoreCase(""))
						{
							throw new Exception();
						}
						if(!varNamesInText.contains(vNameText))
						{
							varNamesInText.add(vNameText);
							if(!varMatchInExcel.containsKey(vNameText))
							{
								throw new Exception("Variable '"+vNameText+"' defined in textfile.txt is not present in ObjectIDs.xlsx\nLine No "+lineNo+" : "+a);
							}
						}
					}
					catch(Exception e)
					{
						w.close();
						throw new Exception("Variable name is NOT defined in textfile.txt\nLine No "+lineNo+" : "+a);
					}
				}
				lineNo++;
			}		
			
			
			
			//Logic to know how many files to generate
			String str;
			int index = varMatchInExcel.get(varNamesInText.get(0));
			for(int i=1; i<s.getPhysicalNumberOfRows();i++)
			{
				str="";
				try
				{
					str = s.getRow(i).getCell(index).getStringCellValue().trim();
				}
				catch(Exception e){
					try
					{
						str = s.getRow(i).getCell(index).getRawValue().trim();
					}
					catch(Exception e1){
						str="";
					}
				}
				if(!str.equalsIgnoreCase(""))
				{
					totalRowsToProcess++;
				}
			}
			
			
			
			
			//Create new folder with timestamp
			String filePath = current+"\\AutomationPayloadResponses\\" + timeStamp+"\\";		
			new File(filePath).mkdirs();
			Thread.sleep(500);
				
			
			
			//Write Actual Data / Create text files
			PrintWriter writer;
			for(int i=1; i<=totalRowsToProcess; i++)
			{
				writer = new PrintWriter(filePath+(i)+".txt", "UTF-8");
				for(String a : actualTextFile)
				{
					try
					{
						if(a.contains("${"))
						{
							vNameText = (((a.split("\\$\\{"))[1]).split("\\}\\$"))[0];
							String replaceString = "";
							try
							{
								replaceString = s.getRow(i).getCell(varMatchInExcel.get(vNameText)).getStringCellValue().trim();
							}
							catch(Exception e1){
								replaceString = s.getRow(i).getCell(varMatchInExcel.get(vNameText)).getRawValue().trim();
							}
							a = a.replace(vNameText, replaceString).replace("${", "").replace("}$", "");
						}
						writer.println(a);
					}
					catch(Exception e)
					{
					}
				}
				writer.close();
			}
			
			
			
			//Close Workbook and print success message
			w.close();
			long end = System.currentTimeMillis();
			float diff = ((float)(end-start)/(float)1000.0);
			JOptionPane.showMessageDialog(null, "Completed Successfully!\nTotal Payloads Generated : " + (totalRowsToProcess) + "\nTotal Time elapsed : " + String.format("%.1f", diff) + " seconds", "Payload Generator ~ Amit_Bhaik", JOptionPane.INFORMATION_MESSAGE);
		
		}
		catch(Exception e)
		{
			e.printStackTrace();
			JOptionPane.showMessageDialog(null,"Unable to Generate Payloads!\n" + e.getMessage(), "Payload Generator ~ Amit_Bhaik", JOptionPane.ERROR_MESSAGE);
		}
		
	}
}

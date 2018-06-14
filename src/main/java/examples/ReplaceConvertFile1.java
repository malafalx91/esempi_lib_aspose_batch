package examples;
import com.aspose.words.Document;
import com.aspose.words.FindReplaceDirection;
import com.aspose.words.FindReplaceOptions;

import java.io.*;
import java.util.Scanner;
import java.util.StringTokenizer;
import com.aspose.words.Document;

//versione di prova 


public class ReplaceConvertFile1 {

	static StringBuilder sbword = new StringBuilder();
	static String dirname = null;
	static File[] filenames = null;
	static Scanner sc = new Scanner(System.in);

	public static void main(String args[]) throws Exception{
		com.aspose.words.License license = new com.aspose.words.License();
		license.setLicense("Aspose.Words.lic");

		boolean fileread = ReadFiles();

		sbword = null;
		System.exit(0);



	}
	private static boolean ReadFiles() throws Exception{

		System.out.println("Enter the location of folder:");

		//input folder
		File file = new File(sc.nextLine()); //sostituire sc.nextLine() con un indirizzo statico se non si vuole input automatico
		//File file = new File("C:\\Users\\francesco.zingaro\\Desktop\\FINCONS\\Windward\\rtf abilitati alla modifica"); //sostituire sc.nextLine() con un indirizzo statico se non si vuole input automatico
		
		
		//File dir = new File("C:\\Users\\francesco.zingaro\\Desktop\\prova sostituzione file\\da pulire\\New folder\\fin\\p\\risultati");
		//dir.mkdir();
		//System.out.println(dirname);
		filenames = file.listFiles();
		String line = null;

		for(File file1 : filenames ){
			System.out.println("File name" + file1.toString());
			sbword.setLength(0); 
			BufferedReader br = new BufferedReader(new FileReader(file1));
			

			line = br.readLine();

			while(line != null){
				System.out.println(line);
				sbword.append(line).append("\r\n");
				line = br.readLine();
			}

			//step 1
			ReplaceLines(file1.toString());
			//Document doc = new Document(file1.toString());
			//doc.getRange().replace("<wr:function select=\"", "<wr:out select=\"=SUM(data(&quot;", new FindReplaceOptions(FindReplaceDirection.FORWARD));
			
			//WriteToFile(file1.toString());
			//step 2
			//ConvertFile(file1.toString());
			System.out.println("Process Completed Successfully");

		}

		return true;    
	}

	
	private static void ReplaceLines(String filename) throws Exception {
		
		
		Document doc = new Document(filename);
		int count;
		String from = "";
		String to = "";
		
		
			from = "<wr:function select=\"";
			to = "<wr:out select=\"=SUM(data(&quot;";
			count = doc.getRange().replace(from, to, new FindReplaceOptions(FindReplaceDirection.FORWARD));
			System.out.println("First replacement count = " + count);
			
			from = "<wr:function select=”";
			to = "<wr:out select=\"=SUM(data(&quot;";
			count = doc.getRange().replace(from, to, new FindReplaceOptions(FindReplaceDirection.FORWARD));
			System.out.println("second replacement count = " + count);
			//
			//stringa originale
			//from = "\" function=\"SUM\" type=\"NUMBER\" pattern=\"#,##0.00\"/>";
			//
			from = "” function=\"SUM\" type=”NUMBER” pattern=”#,##0.00”/>";
			to = "&quot;))\" type='NUMBER' format='category:number;negFormat:0;format:#,##0.00;' nickname='[SUM]'/>";
			//to = "pippo";
			count = doc.getRange().replace(from, to, new FindReplaceOptions(FindReplaceDirection.FORWARD));
			System.out.println("Third replacement count = " + count);
		
		
		
		doc.save(filename + ".docx");
		//ExEnd:ReplaceWithString
		
		/*
		
		
		System.out.println("sbword contains :" + sbword.toString());
		//System.out.println("Enter the 1 word to replace from each of the files:");
		
		
		//sbword.toString().replaceAll("“","\"").replaceAll("”","\"");
		String from3 = Character.toString((char) 8221);
		String To3 = "\"";
		//C:\Users\francesco.zingaro\Desktop\prova sostituzione file\da pulire\New foldersbword = new StringBuilder(sbword.toString().replaceAll(from3, To3));
		//count = ReplaceAll(sbword,from3,To3);
		//System.out.println("Third replacement OK: count = " + count);
		
		
		String from1 = "<wr:function select=\""; //impostare stringhe statiche di sostituzione in alternativa
		String To1 = "<wr:out select=\"=SUM(data(&quot;";
		//StringBuilder sbword = new StringBuilder(stbuff);
		count = ReplaceAll(sbword,from1,To1);
		System.out.println("First replacement OK: count = " + count);
		
				
		//System.out.println("Enter the 2 word to replace from each of the files:");
		String from2 = "\" function=\"SUM\" type=\"NUMBER\" pattern=\"#,##0.00\"/>";
		String To2 = "&quot;))\" type='NUMBER' format='category:number;negFormat:0;format:#,##0.00;' nickname='[SUM]'/>";
		count = ReplaceAll(sbword,from2,To2);
		System.out.println("Second replacement OK: count = " + count);
		
		*/
		
		
	}

	private static void ConvertFile(String filename) throws Exception{

		File file1 = new File(filename);
		//da qui la conversione in docx con aspose
		Document doc = new Document(file1.toString());
		// Save as any Aspose.Words supported format, such as DOCX.
		doc.save(file1.toString() + ".docx");


	}



	private static int ReplaceAll(StringBuilder builder, String from, String to){
		
		
		/*
		try (PrintWriter pw = new PrintWriter("C:\\Users\\francesco.zingaro\\Desktop\\prova sostituzione file\\da pulire\\New folder\\out.txt")) {
		    pw.println(builder.toString());
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		*/
		
		int count = 0;
		int index = builder.indexOf(from);
		while(index != -1){
			builder.replace(index, index + from.length(), to);
			count++;
			index += to.length();
			index = builder.indexOf(from,index);
		}
		return count;
	}
	private static void WriteToFile(String filename) throws IOException{
		try{
			File file1 = new File(filename);
			// The path to the documents directory.
			//String dataDir = Utils.getDataDir(AsposeLoadTxtFile.class);

			BufferedWriter bufwriter = new BufferedWriter(new FileWriter(file1));
			bufwriter.write(sbword.toString());
			bufwriter.close();


		}catch(Exception e){
			System.out.println("Error occured while attempting to write to file: " +     e.getMessage());
		}

		
	}
}
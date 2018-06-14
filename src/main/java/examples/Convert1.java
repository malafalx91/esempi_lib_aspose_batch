package examples;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.Scanner;

import org.apache.commons.io.FileUtils;
import org.omg.CORBA.Environment;

import com.aspose.words.Document;
import com.aspose.words.License;



public class Convert1{

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


		File file = new File(sc.nextLine());
		//File dest = new File("C:\\Users\\francesco.zingaro\\Desktop\\prova sostituzione file\\da pulire\\New folder");
		

		filenames = file.listFiles();
		String line = null;

		for(File file1 : filenames ){
			System.out.println("File name" + file1.toString());

			//int num = countLines(file1);
			//System.out.println("N. lines: " + num);
			//da qui la conversione in docx con aspose
			Document doc = new Document(file1.toString());
			// Save as any Aspose.Words supported format, such as DOCX.
			doc.save(file1.toString() + ".rtf");
			
			System.out.println("Process Completed Successfully");

		}

		return true;    
	}

	public static int countLines(File input) throws IOException {
		try (InputStream file1 = new FileInputStream(input)) {
			int count = 1;
			for (int aChar = 0; aChar != -1;aChar = file1.read())
				count += aChar == '\n' ? 1 : 0;
			return count;

		}



	}
}
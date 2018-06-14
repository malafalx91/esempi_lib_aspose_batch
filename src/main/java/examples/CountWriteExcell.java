package examples;
import com.aspose.words.Document;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Scanner;
import java.util.StringTokenizer;
import com.aspose.words.Document;

/*Piccolo BATHC che tramite libreria aspose:
 * 1) cerca e conta il numero di righe del documento su più documenti in ingresso,
 *  salva i risultati in un doc excell.
 * */

public class CountWriteExcell {

	static StringBuilder sbword = new StringBuilder();
	static String dirname = null;
	static File[] filenames = null;
	static Scanner sc = new Scanner(System.in);

	private static final String EXCEL_FILE_LOCATION = "C:\\Users\\francesco.zingaro\\Desktop\\prova sostituzione file\\New folder\\prova\\Test3.xls";

	public static void main(String args[]) throws Exception{
		com.aspose.words.License license = new com.aspose.words.License();
		license.setLicense("Aspose.Words.lic");



		//boolean fileread = ReadFiles();

		sbword = null;
		//1. Create an Excel file
		WritableWorkbook myFirstWbook = null;
		try {

			myFirstWbook = Workbook.createWorkbook(new File(EXCEL_FILE_LOCATION));

			// create an Excel sheet
			WritableSheet excelSheet = myFirstWbook.createSheet("Conteggio righe", 0);

			// add something into the Excel sheet
			Label label = new Label(0, 0, "Row Num:    ");
			excelSheet.addCell(label);
			label = new Label(1, 0, "File(s) Name:    ");
            excelSheet.addCell(label);

			System.out.println("Enter the location of folder:");


			File file = new File(sc.nextLine());
			filenames = file.listFiles();
			String line = null;
			int i = 1;
			
			for(File file1 : filenames ){
				System.out.println("File name " + file1.toString());
				
				label = new Label(1, i, file1.getName());
				
				excelSheet.addCell(label);

				int num = countLines(file1);
				System.out.println("N. lines: " + num);
				Number number = new Number(0, i, countLines(file1));
				
				excelSheet.addCell(number);
				i++;

			}
			    myFirstWbook.write();//scrive sul file

		} catch (IOException e) {
			e.printStackTrace();
		} catch (WriteException e) {
			e.printStackTrace();
		} finally {

			if (myFirstWbook != null) {
				try {
					myFirstWbook.close();
				} catch (IOException e) {
					e.printStackTrace();
				} catch (WriteException e) {
					e.printStackTrace();
				}
			}


		}
		//System.exit(0);

	}


	private static boolean ReadFiles() throws Exception{

		System.out.println("Enter the location of folder:");


		File file = new File(sc.nextLine());
		filenames = file.listFiles();
		String line = null;

		for(File file1 : filenames ){
			System.out.println("File name " + file1.toString());

			int num = countLines(file1);
			System.out.println("N. lines: " + num);



			//System.out.println("Process Completed Successfully");

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
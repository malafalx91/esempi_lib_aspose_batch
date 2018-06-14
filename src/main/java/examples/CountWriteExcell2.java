package examples;
import com.aspose.words.Document;
import com.aspose.words.FindReplaceDirection;
import com.aspose.words.FindReplaceOptions;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import java.io.*;

import java.util.Scanner;


/*Piccolo BATHC che tramite libreria aspose:
 * 1) cerca e conta delle stringhe su multi documenti, salva  i risultati in un excell. (FROM e TO devono essere uguali)
 * 2) cerca e sostituisce delle stringhe su multi documenti, salva i risultati in un excell. (FROM e TO devono essere diversi)
 * */

public class CountWriteExcell2 {

	static StringBuilder sbword = new StringBuilder();
	static String dirname = null;
	static File[] filenames = null;
	static Scanner sc = new Scanner(System.in);

	private static final String EXCEL_FILE_LOCATION = "C:\\Users\\francesco.zingaro\\Desktop\\new folder\\Test.xls";

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
			WritableSheet excelSheet = myFirstWbook.createSheet("Conteggio marcatori", 0);

			// add something into the Excel sheet
			Label label = new Label(0, 0, "marcatori Num:    ");
			excelSheet.addCell(label);
			label = new Label(1, 0, "File(s) Name:    ");
            excelSheet.addCell(label);

			System.out.println("Enter the location of folder:");


			File file = new File(sc.nextLine());
			filenames = file.listFiles();
		
			int i = 1;
			
			for(File file1 : filenames ){
				System.out.println("File name " + file1.toString());
				
				label = new Label(1, i, file1.getName());
				
				excelSheet.addCell(label);

				int num = ReplaceLines(file1.toString());
				System.out.println("N. lines: " + num);
				Number number = new Number(0, i, ReplaceLines(file1.toString()));
				
				excelSheet.addCell(number);
				i++;

			}
			//scrive sul file  
			myFirstWbook.write();

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
		

		for(File file1 : filenames ){
			System.out.println("File name " + file1.toString());

			int num = ReplaceLines(file1.toString());
			System.out.println("N. lines: " + num);



			//System.out.println("Process Completed Successfully");

		}

		return true;    
	}

		
		private static int ReplaceLines(String filename) throws Exception {
			
			
			Document doc = new Document(filename);
			int count;
			String from = "";
			String to = "";
			
			    //modificare la stringa from con quella da cercare e sostituire, se from e to sono 
			    //identiche viene effetuata una semplice ricerca e conteggio,
				// altrimenta ricerca e sostituzione
				from = "DIGITALE";
				to = "DIGITALE";
				count = doc.getRange().replace(from, to, new FindReplaceOptions(FindReplaceDirection.FORWARD));
				System.out.println("First replacement count = " + count);

				return count;
	}
}
package convert;

import java.io.File;
import java.util.Scanner;

import net.windward.tools.CleanTemplates;
/*
 *Options are:
  -ll link tags(powerpoint
  -ps Substitution parameter mode.
  -pp Parameter Plus parameter mode. This is the default.
  -po Parameter Only parameter mode.
  -ru Remove Unused styles, etc. OFF. The default is ON.
  -sq Smart quote OFF - remove quotes around ${vars}. The default is ON. Must be before -sql/xml.
  -tt Text tags.
  -tf Field tags (macro in Excel).
  -tp Field Plus tags (macro in Excel). This is the default.
  -tc Content Control tags (macro in Excel).
  -v9 Template was created in version 9 or earlier (forces -v14). Setting this wrong will change select types incorrectly.
  -v14 Template was created in version 14 or earlier. Setting this wrong will change formatting patterns incorrectly.
  -fe Attempt to fix evaluable expressions.
Datasources are:
  -sql:name
  -xml:name
Process Completed Successfully
 *
 */
public class CleanTemplate {
	static StringBuilder sbword = new StringBuilder();
	static String dirname = null;
	static File[] filenames = null;
	static Scanner sc = new Scanner(System.in);

	public static void main(String[] args) throws Exception {
 

		//esempio lettura file da folder
		System.out.println("Enter the location of folder:");
		File file = new File ("C:\\Users\\francesco.zingaro\\Desktop\\New folder\\1"); //(sc.nextLine());//con input manuale
		filenames = file.listFiles();

		for(File file1 : filenames ){
			System.out.println("File name" + file1.toString());
            //nella prima parte si inseriscono le opzioni di pulizia template, poi va inserito il file e infine l'output
			//pulizia template con codice semplificato
			String T[] = {"-pp","-sq","-ru","-tp","-v9","-v14","-fe", file1.toString(), "C:\\Users\\francesco.zingaro\\Desktop\\New folder\\2" };
			//con pulizia template lasciando le linee di codice
			//String T[] = {"-pp","-sq","-ru","-tt","-v9","-v14","-fe", file1.toString(), "C:\\Users\\francesco.zingaro\\Desktop\\New folder\\2" };
			CleanTemplates.main(T);
			
			
			
			System.out.println("Process Completed Successfully");

		}

	}

}

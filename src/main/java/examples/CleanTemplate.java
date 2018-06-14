package examples;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Scanner;

import org.apache.commons.io.IOUtils;

import net.windward.datasource.DataSourceProvider;
import net.windward.datasource.dom4j.Dom4jDataSource;
import net.windward.env.DataConnectionException;
import net.windward.env.DataSourceException;
import net.windward.env.OutputLimitationException;
import net.windward.format.TemplateParseException;
import net.windward.tags.TagException;
import net.windward.util.LicenseException;
import net.windward.xmlreport.AlreadyProcessedException;
import net.windward.xmlreport.ProcessPdf;
import net.windward.xmlreport.ProcessReport;
import net.windward.xmlreport.ProcessReportAPI;
import net.windward.xmlreport.SetupException;
import net.windward.tools.CleanTemplates;

public class CleanTemplate {
	static StringBuilder sbword = new StringBuilder();
	static String dirname = null;
	static File[] filenames = null;
	static Scanner sc = new Scanner(System.in);

	public static void main(String[] args) throws Exception {
 

		//esempio lettura file da folder
		System.out.println("Enter the location of folder:");
		File file = new File(sc.nextLine());//con input manuale
		filenames = file.listFiles();

		for(File file1 : filenames ){
			System.out.println("File name" + file1.toString());
            //nella prima parte si inseriscono le opzioni di pulizia template, poi va inserito il file e infine l'output
			String T[] = {"-pp","-sq","-ru","-tp","-v9","-v14","-fe", file1.toString(), "C:\\Users\\francesco.zingaro\\Desktop\\FINCONS\\Windward\\rtf abilitati alla modifica copia\\Cleaned" };
			CleanTemplates.main(T);
			
			
			
			System.out.println("Process Completed Successfully");

		}

	}

}

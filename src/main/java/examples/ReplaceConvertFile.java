package examples;
import com.aspose.words.Document;
import com.aspose.words.FindReplaceDirection;
import com.aspose.words.FindReplaceOptions;

import java.io.*;
import java.util.Scanner;
import java.util.StringTokenizer;
import com.aspose.words.Document;

/* Esegue i seguenti step:
 * 1. Inserire link in input della cartella contenente i template rtf(preferibile se i file hanno i permessi alla modifica)
 * 2. Ricerca nei documenti di tag deprecati in windward
 * 3. Sostituzione dei vecchi tag con quelli nuovi(da versione 8 --> alla 16) 
 * 4. Conversione da RTF a DOCX poichè dalla versione 13 di Windward non sono più supportati gli RTF
 * 5. Effettuare il CleanTemplate tramite il 2o batch
 * */

public class ReplaceConvertFile {

	static StringBuilder sbword = new StringBuilder();
	static String dirname = null;
	static File[] filenames = null;
	static Scanner sc = new Scanner(System.in);

	public static void main(String args[]) throws Exception{
		//Il batch utilizza la libreria aspose, essenziale al funzionamento la licenza, inserire il file Aspose.Words.lic nella dir
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

		//leggere da cartella tutti i file presenti e caricarli in memoria
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

			//Cerca - Sostituisce - Converte in docx
			ReplaceLines(file1.toString());

			System.out.println("Process Completed Successfully");

		}

		return true;    
	}

	@SuppressWarnings("unused")
	private static void ReplaceLines(String filename) throws Exception {


		Document doc = new Document(filename);
		int count;
		String from = "";
		String to = "";

		//Stringa 1
		from = "<wr:function select=\"";
		to = "<wr:out select=\"=SUM(data(&quot;";
		count = doc.getRange().replace(from, to, new FindReplaceOptions(FindReplaceDirection.FORWARD));
		System.out.println("First replacement count = " + count);

		//Stringa 2 che corregge l'errore del right double quote 
		from = "<wr:function select=”";
		to = "<wr:out select=\"=SUM(data(&quot;";
		count = doc.getRange().replace(from, to, new FindReplaceOptions(FindReplaceDirection.FORWARD));
		System.out.println("second replacement count = " + count);

		//Stringa 3 corregge l'errore del right double quote su tutta la funzione
		from = "” function=\"SUM\" type=”NUMBER” pattern=”#,##0.00”/>";
		to = "&quot;))\" type='NUMBER' format='category:number;negFormat:0;format:#,##0.00;' nickname='[SUM]'/>";			
		count = doc.getRange().replace(from, to, new FindReplaceOptions(FindReplaceDirection.FORWARD));
		System.out.println("Third replacement count = " + count);



		doc.save(filename + ".docx");
		//ExEnd:ReplaceWithString


	}

	//funzione di conversione file tramite libreria aspose
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
package convert;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.FilenameFilter;
import java.util.LinkedList;
import java.util.List;
import java.util.Scanner;

import com.aspose.words.Document;
import com.aspose.words.FindReplaceDirection;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.License;
import com.aspose.words.OoxmlCompliance;
import com.aspose.words.OoxmlSaveOptions;

/** Esegue i seguenti step:
 * 1. Inserire link in input della cartella contenente i template rtf(preferibile se i file hanno i permessi alla modifica)
 * 2. Ricerca nei documenti di tag deprecati in windward
 * 3. Sostituzione dei vecchi tag con quelli nuovi(da versione 8 --> alla 16) 
 * 4. Conversione da RTF a DOCX poichè dalla versione 13 di Windward non sono più supportati gli RTF
 * 5. Effettuare il CleanTemplate tramite il 2o batch
 * */
public class ReplaceConvertFile {

	public static void main(String args[]) throws Exception { 
		// Il batch utilizza la libreria aspose, essenziale al funzionamento la licenza,
		// inserire il file Aspose.Words.lic nella dir
		License license = new License();
		license.setLicense("Aspose.Words.lic");
		processFiles();
		System.exit(0);
	}

	private static void processFiles() throws Exception {

		System.out.println("Enter the location of folder:");

		// input folder
		// sostituire sc.nextLine() con un indirizzo statico se non si vuole input automatico
		Scanner sc = new Scanner(System.in);
		File rtfTemplateDir = new File(sc.nextLine());
		sc.close();

		// leggere da cartella tutti i file presenti e caricarli in memoria
		File[] rtfTemplates = rtfTemplateDir.listFiles(new FilenameFilter() {
			public boolean accept(File dir, String name) {
				return name.toLowerCase().endsWith(".rtf");
			}
		});

		//collects template with conversion errors
		List<String> templatesInError = new LinkedList<>();

		StringBuilder sbword = new StringBuilder();
		for (File rtfTemplate : rtfTemplates) {
			try {
				System.out.println("Processing " + rtfTemplate.toString());
				sbword.setLength(0);
				BufferedReader br = new BufferedReader(new FileReader(rtfTemplate));

				String line = br.readLine();
				while (line != null) {
					//System.out.println(line);
					sbword.append(line).append("\r\n");
					line = br.readLine();
				}

				br.close();


				Document document = new Document(rtfTemplate.getPath());

				// Cerca - Sostituisce
				document = replaceLines(document);

				
				
				
				OoxmlSaveOptions options = new OoxmlSaveOptions();
				options.setCompliance(OoxmlCompliance.ISO_29500_2008_STRICT);
				
				
				
				
				
				//salvataggio
				String newFileName = rtfTemplate.getName().replace("rtf", "docx");
				String docxFilePath = rtfTemplate.getParent() + File.separator + newFileName;
				document.save(docxFilePath, options);
				System.out.println(docxFilePath + " saved\n");
			} catch(Exception e) {
				e.printStackTrace();
				templatesInError.add(rtfTemplate.toString());
			}
		}

		if(templatesInError.size() > 0) {
			System.out.println("Process finished. " + templatesInError.size() + " template(s) with conversion errors:");
			for (String template : templatesInError) {
				System.out.println(template);
			}
		}
		else {
			System.out.println("Process finished with no conversion errors.");
		}

	}


	private static Document replaceLines(Document document) throws Exception {

		int count;
		String from = "";
		String to = "";

		// Stringa 1
		from = "<wr:function select=\"";
		to = "<wr:out select=\"=SUM(data(&quot;";
		count = document.getRange().replace(from, to, new FindReplaceOptions(FindReplaceDirection.FORWARD));
		System.out.print("[" + from + " replacement count = " + count + ", ");

		// Stringa 2 che corregge l'errore del right double quote
		from = "<wr:function select=”";
		to = "<wr:out select=\"=SUM(data(&quot;";
		count = document.getRange().replace(from, to, new FindReplaceOptions(FindReplaceDirection.FORWARD));
		System.out.print(from + " replacement count = " + count + ", ");

		// Stringa 3 corregge l'errore del right double quote su tutta la funzione
		from = "” function=\"SUM\" type=”NUMBER” pattern=”#,##0.00”/>";
		to = "&quot;))\" type='NUMBER' format='category:number;negFormat:0;format:#,##0.00;' nickname='[SUM]'/>";
		count = document.getRange().replace(from, to, new FindReplaceOptions(FindReplaceDirection.FORWARD));
		System.out.print(from + " replacement count = " + count + "]\n");

		// Stringa 4 corregge l'errore del right double quote su tutta la funzione
		from = "” function=\"SUM\" pattern=\"#,##0.00\" type=”NUMBER” default=”0,00”/>";
		to = "&quot;))\" type='NUMBER' format='category:number;negFormat:0;format:#,##0.00;' nickname='[SUM]' default=”0,00”/>";
		count = document.getRange().replace(from, to, new FindReplaceOptions(FindReplaceDirection.FORWARD));
		System.out.print(from + " replacement count = " + count + "]\n");
		
		// Stringa 5 corregge l'errore del right double quote su tutta la funzione
				from = "function=\"SUM\" pattern=\"#,##0.00\" type=”NUMBER” default=”0,00”/>";
				to = "&quot;))\" type='NUMBER' format='category:number;negFormat:0;format:#,##0.00;' nickname='[SUM]' default=”0,00”/>";
				count = document.getRange().replace(from, to, new FindReplaceOptions(FindReplaceDirection.FORWARD));
				System.out.print(from + " replacement count = " + count + "]\n");
		
		// Stringa 6 a volte i documenti includono dei file .docx e non .rtf
		from = ".rtf";
		to = ".docx";
		count = document.getRange().replace(from, to, new FindReplaceOptions(FindReplaceDirection.FORWARD));
		System.out.print("[" + from + " replacement count = " + count + ", ");

		return document;
		// ExEnd:ReplaceWithString
	}

	/*

	// funzione che converte file tramite libreria
	private static void convertFile(String filename) throws Exception {
		File file1 = new File(filename);
		// da qui la conversione in docx con aspose doc
		Document doc = new Document(file1.toString());
		// Save as any Aspose.Words supported format, such as DOCX
		doc.save(file1.toString() + ".docx");
	}

	private static int replaceAll(StringBuilder builder, String from, String to) {

		 //try (PrintWriter pw = new
		 //PrintWriter("C:\\Users\\francesco.zingaro\\Desktop\\prova sostituzione file\\da pulire\\New folder\\out.txt"
		 //)) { pw.println(builder.toString()); } catch (FileNotFoundException e) {
		 //e.printStackTrace(); }

		int count = 0;
		int index = builder.indexOf(from);
		while (index != -1) {
			builder.replace(index, index + from.length(), to);
			count++;
			index += to.length();
			index = builder.indexOf(from, index);
		}
		return count;
	}

	private static void writeToFile(String filename) throws IOException {
		try {
			File file1 = new File(filename);
			// The path to the documents directory.
			// String dataDir = Utils.getDataDir(AsposeLoadTxtFile.class);

			BufferedWriter bufwriter = new BufferedWriter(new FileWriter(file1));
			bufwriter.write(sbword.toString());
			bufwriter.close();

		} catch (Exception e) {
			System.out.println("Error occured while attempting to write to file: " + e.getMessage());
		}
	}

	 */


}
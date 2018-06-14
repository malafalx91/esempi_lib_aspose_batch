

import com.aspose.words.Document;


public class AsposeLoadTxtFile
{
    public static void main(String[] args) throws Exception
    {
	    // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeLoadTxtFile.class);
		
        // The encoding of the text file is automatically detected.
        Document doc = new Document(dataDir + "1.rtf");

        // Save as any Aspose.Words supported format, such as DOCX.
        doc.save(dataDir + "1.docx");
        
	System.out.println("Process Completed Successfully");
    }
}

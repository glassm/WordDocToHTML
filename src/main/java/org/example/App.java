package org.example;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import static fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter.getInstance;
import fr.opensagres.poi.xwpf.converter.core.ImageManager;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.w3c.dom.Document;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        App app = new App();
        app.docxToHtml("example/resume.docx", "example/out/");
        app.docToHtml("example/resume.doc", "example/out/resume_doc.html");


    }


    private void docxToHtml(String sourcePath, String targetDir) {
        File sourceFile = new File(sourcePath);

        XWPFDocument document = null;
        try {

            FileInputStream fis = new FileInputStream(sourcePath);
            document = new XWPFDocument(fis);
        } catch (IOException e) {
            e.printStackTrace();
        }

       XHTMLOptions options = XHTMLOptions.create().setImageManager(new ImageManager(new File(targetDir), "images"));

        FileOutputStream out = null;
        try {
            out = new FileOutputStream(targetDir + sourceFile.getName() + ".html");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        try {
            getInstance().convert(document, out, options);  //OpenSagresConverter is Out Of Date; can't do this with current Apache POI
        } catch (IOException e) {
            e.printStackTrace();
        }

        try {
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            document.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void docToHtml(String docfile, String htmlfile) {
        try {
            InputStream input = new FileInputStream(docfile);
            HWPFDocument wordDocument = null;
            try {
                wordDocument = new HWPFDocument(input);
            } catch(IOException e) {
                e.printStackTrace();
            }
            WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
            wordToHtmlConverter.processDocument(wordDocument);
            Document htmlDocument = wordToHtmlConverter.getDocument();
            ByteArrayOutputStream outStream = new ByteArrayOutputStream();
            DOMSource domSource = new DOMSource(htmlDocument);
            StreamResult streamResult = new StreamResult(outStream);

            TransformerFactory tf = TransformerFactory.newInstance();
            Transformer serializer = tf.newTransformer();
            serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
            serializer.setOutputProperty(OutputKeys.INDENT, "yes");
            serializer.setOutputProperty(OutputKeys.METHOD, "html");
            serializer.transform(domSource, streamResult);
            outStream.close();

            String content = new String(outStream.toByteArray(), "UTF-8");
            writeToFile(htmlfile, content);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

        public void writeToFile(String filePath, String content) {
                try {
                    FileWriter myWriter = new FileWriter(filePath);
                    myWriter.write(content);
                    myWriter.close();
                    System.out.println("Successfully wrote to the file.");
                } catch (IOException e) {
                    System.out.println("An error occurred.");
                    e.printStackTrace();
                }
            }

}


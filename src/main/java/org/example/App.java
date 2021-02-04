package org.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import static fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions;
import fr.opensagres.poi.xwpf.converter.core.ImageManager;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        String docPath = "valid.docx";

        String root = "C:\\Users\\mglass\\workspace\\WordDocToHTML\\src\\main\\resources\\";
        String htmlPath = root + "WordDocument.html";

        XWPFDocument document = null;
        try {

            FileInputStream fis = new FileInputStream("C:\\Users\\mglass\\workspace\\WordDocToHTML\\src\\main\\resources\\valid.docx");
            document = new XWPFDocument(fis);
        } catch (IOException e) {
            e.printStackTrace();
        }

        XHTMLOptions options = XHTMLOptions.create().setImageManager(new ImageManager(new File(root), "images"));

        FileOutputStream out = null;
        try {
            out = new FileOutputStream(htmlPath);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        try {
            getInstance().convert(document, out, options);
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
}


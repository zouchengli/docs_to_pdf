package com.yeokhengmeng.docstopdfconverter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;


public class MainClass {

    public static void main(String[] args) {
        Converter converter;

        try {

            String inPath = "C:\\Users\\Administrator\\Documents\\test1.docx";
            String outPath = "C:\\Users\\Administrator\\Documents\\test1.pdf";

            converter = process(inPath, outPath);
        } catch (Exception e) {
            System.out.println("\n\nInput\\Output file not specified properly.");
            return;
        }

        if (converter == null) {
            System.out.println("Unable to determine type of input file.");
        } else {
            try {
                converter.convert();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

    }


    public static Converter process(String inPath, String outPath) {

        Converter converter = null;
        try {

            boolean shouldShowMessages = true;

            String lowerCaseInPath = inPath.toLowerCase();

            InputStream inStream = getInFileStream(inPath);

            OutputStream outStream = getOutFileStream(outPath);

            if (lowerCaseInPath.endsWith("doc")) {

                converter = new DocToPDFConverter(inStream, outStream, shouldShowMessages, true);

            } else if (lowerCaseInPath.endsWith("docx")) {

                converter = new DocxToPDFConverter(inStream, outStream, shouldShowMessages, true);

            } else if (lowerCaseInPath.endsWith("ppt")) {

                converter = new PptToPDFConverter(inStream, outStream, shouldShowMessages, true);

            } else if (lowerCaseInPath.endsWith("pptx")) {

                converter = new PptxToPDFConverter(inStream, outStream, shouldShowMessages, true);

            } else if (lowerCaseInPath.endsWith("odt")) {

                converter = new OdtToPDF(inStream, outStream, shouldShowMessages, true);

            } else {

                converter = null;

            }

        } catch (Exception e) {
            System.err.println(e.getMessage());
        }

        return converter;

    }


    protected static InputStream getInFileStream(String inputFilePath) throws FileNotFoundException {
        File inFile = new File(inputFilePath);
        FileInputStream iStream = new FileInputStream(inFile);
        return iStream;
    }

    protected static OutputStream getOutFileStream(String outputFilePath) throws IOException {
        File outFile = new File(outputFilePath);

        try {
            //Make all directories up to specified
            outFile.getParentFile().mkdirs();
        } catch (NullPointerException e) {
            //Ignore error since it means not parent directories
        }

        outFile.createNewFile();
        FileOutputStream oStream = new FileOutputStream(outFile);
        return oStream;
    }

}

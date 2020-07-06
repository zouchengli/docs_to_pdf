package site.clzblog.docs.to.pdf.converter;

import java.awt.*;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import com.lowagie.text.DocumentException;
import com.lowagie.text.Font;
import com.lowagie.text.pdf.BaseFont;
import fr.opensagres.xdocreport.itext.extension.font.IFontProvider;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class DocxToPDFConverter extends Converter {

    public DocxToPDFConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
    }

    @Override
    public void convert() throws Exception {
        loading();

        XWPFDocument document = new XWPFDocument(inStream);

        // Customer appoint font
        IFontProvider iFontProvider = (s, s1, v, i, color) -> {
            BaseFont baseFont = null;
            try {
                baseFont = BaseFont.createFont("C:/Windows/Fonts/simsun.ttc,1", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            } catch (DocumentException | IOException e) {
                e.printStackTrace();
            }
            return new Font(baseFont, v, i, color);
        };

        // Customer appoint font
        IFontProvider iFontProvider1 = new IFontProvider() {
            public Font getFont(String s, String s1, float v, int i, Color color) {
                BaseFont baseFont = null;
                try {
                    baseFont = BaseFont.createFont("C:/Windows/Fonts/msyh.ttc,1", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                } catch (DocumentException | IOException e) {
                    e.printStackTrace();
                }
                return new Font(baseFont, v, i, color);
            }
        };

        PdfOptions options = PdfOptions.create().fontProvider(iFontProvider).fontProvider(iFontProvider1);

        processing();

        PdfConverter.getInstance().convert(document, outStream, options);

        finished();

    }

}

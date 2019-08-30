package poiTest;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.hwpf.converter.WordToTextConverter;
import org.apache.poi.extractor.POITextExtractor;

public class PoiTest {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		String path = "d:/zsdigger.docx";
        InputStream stream = null;
        try {
            stream = new FileInputStream(new File(path));
            if (path.endsWith(".doc")) {
                HWPFDocument document = new HWPFDocument(stream);
                WordExtractor extractor = new WordExtractor(document);
                String[] contextArray = extractor.getParagraphText();
                System.out.println(extractor.getText());
                for (int i = 0; i < contextArray.length; i++) {
                	 System.out.println(contextArray[i]);
                }
                extractor.close();
                document.close();
            } else if (path.endsWith(".docx")) {
                XWPFDocument document = new XWPFDocument(stream).getXWPFDocument();
                List<XWPFParagraph> paragraphList = document.getParagraphs();
            	for (XWPFParagraph s : paragraphList) {
                   System.out.println(s.getText());
        		}
                document.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println(e);
        } finally {
            if (null != stream) try {
                stream.close();
            } catch (IOException e) {
                e.printStackTrace();
                System.out.println("¶ÁÈ¡wordÎÄ¼þÊ§°Ü");
            }
        }
	}

}

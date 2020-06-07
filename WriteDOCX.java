import java.awt.Color;
import java.io.FileInputStream;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import com.spire.doc.*;
import com.spire.doc.documents.BorderStyle;
import com.spire.doc.documents.HorizontalAlignment;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.CheckBoxFormField;
import com.spire.doc.fields.TextRange;
/**
 * 
 * @author ananddw1
 *
 */
public class WriteDOCX {
	/**
	 * 
	 * @param args
	 */
	public static void main(String[] args) {
		try {
			FileInputStream fis = new FileInputStream("1.docx");
			XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(fis));
			XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(xdoc);

			XWPFHeader header = policy.getDefaultHeader();
			if (header != null) {
				System.out.println(header.getText());
			}

			XWPFFooter footer = policy.getDefaultFooter();
			if (footer != null) {
				System.out.println(footer.getText());
			}

			createDocment();

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	public static void createDocment() {

		Document document = new Document();
		Section section = document.addSection();

		// Add header
		HeadersFooters headersFooters = section.getHeadersFooters();
		HeaderFooter header = headersFooters.getHeader();
		Paragraph headerParagraph = header.addParagraph();
		TextRange hText = headerParagraph.appendText("Sample Document");
		// Set header text format
		hText.getCharacterFormat().setFontName("Calibri");
		hText.getCharacterFormat().setFontSize(15f);
		hText.getCharacterFormat().setTextColor(Color.blue);
		// Set header paragraph format
		headerParagraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
		// border
		headerParagraph.getFormat().getBorders().getBottom().setBorderType(BorderStyle.Thick_Thin_Small_Gap);
		headerParagraph.getFormat().getBorders().getBottom().setSpace(0.05f);
		headerParagraph.getFormat().getBorders().getBottom().setColor(Color.darkGray);

		// Add footer
		com.spire.doc.HeaderFooter footer = section.getHeadersFooters().getFooter();
		Paragraph footerParagraph = footer.addParagraph();
		footerParagraph.appendText(" Ticket Number ");
		footerParagraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);
		footerParagraph.getFormat().getBorders().getTop().setBorderType(BorderStyle.Thick_Thin_Small_Gap);
		footerParagraph.getFormat().getBorders().getTop().setSpace(0.010f);

		
		CheckBoxFormField checkBox = new CheckBoxFormField(document);
        checkBox.setChecked(true);
        
        Paragraph checkbox = section.addParagraph();
        checkbox.appendText("Details Provided");
        CheckBoxFormField checkeBoxField = new CheckBoxFormField(document);
        checkeBoxField.setChecked(true);
        checkeBoxField.setName("Chk1");
        checkbox.getItems().insert(1, checkeBoxField);
        
		// save to file
		document.saveToFile("HeaderAndFooter.docx", FileFormat.Docx);
	}
}

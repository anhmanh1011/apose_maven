import com.aspose.words.*;

public class Main_APOSE {
    public static void main(String[] args) throws Exception {
        // Create a new empty document A
        Document doc = new Document("C:\\Users\\daoma\\IdeaProjects\\apose_maven\\Export\\HO_SO_DIEN_TU.docx");


// Inisialize a DocumentBuilder
        DocumentBuilder builder = new DocumentBuilder(doc);

        Section section = doc.getFirstSection();
        HeaderFooter header = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY);
        Table table = header.getTables().get(0); // Assuming it's the first table in the header

        Row headerRow = table.getRows().get(0); // Assuming it's the first row in the table
        Cell cell = headerRow.getCells().get(0);
        cell.getFirstParagraph().removeAllChildren();
        cell.getFirstParagraph().appendChild(new Run(doc, "New Content")); // Add new content
        FormFieldCollection formFields = doc.getRange().getFormFields();
        for (FormField formField : formFields) {
            if (formField.getName().equals("HO_VA_TEN")) {
                // Found the form field, now replace its content
                formField.setResult("DAO DUC MANH");
            }
            if (formField.getName().equals("NGAY_SINH")) {
                // Found the form field, now replace its content
                formField.setResult("28/10/1997");
            }
            if (formField.getName().equals("CHECK_BOX_NAM")) {
                // Found the form field, now replace its content
                formField.setChecked(true);
            }

            if (formField.getName().equals("DIA_CHI")) {
                // Found the form field, now replace its content
                formField.setResult("SO 2 HOANG CUONG THANH BA PHU THO");
            }

            if (formField.getName().equals("SO_DIEN_THOAI")) {
                // Found the form field, now replace its content
                formField.setResult("0333514807");
            }

            if (formField.getName().equals("EMAIL")) {
                // Found the form field, now replace its content
                formField.setResult("ANHYEUEM@GMAIL>COM");
            }

            if (formField.getName().equals("LY_DO_KHAM")) {
                // Found the form field, now replace its content
                formField.setResult("DAU RANG" +
                        " LAMNH   LAMNH DẤDADASDADADADDSD  LAMNH   LAMNH DẤDADASDADADADDSD  LAMNH   LAMNH DẤDADASDADADADDSD  LAMNH   LAMNH DẤDADASDADADADDSD  LAMNH   LAMNH DẤDADASDADADADDSD  LAMNH   LAMNH DẤDADASDADADADDSD  LAMNH   LAMNH DẤDADASDADADADDSD  LAMNH   LAMNH DẤDADASDADADADDSD LAMNH");
            }

        }
        Table tableTienSuBenhLy = doc.getFirstSection().getBody().getTables().get(2); // Assuming it's the first table

        Row newRow = new Row(doc);
        // Add cells to the new row
        Cell cell1 = new Cell(doc);
        cell1.appendChild(new Paragraph(doc));
        cell1.getFirstParagraph().appendChild(new Run(doc, "Cell 1 Content"));
        newRow.appendChild(cell1);
        tableTienSuBenhLy.appendChild(newRow);
        Shape textBox = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        // Check if the text box exists
        if (textBox != null && textBox.getShapeType() == ShapeType.TEXT_BOX) {
            // Clear the existing text
            textBox.removeAllChildren();

            // Fill the text box with the desired string
            Run run = new Run(doc, "Your text goes here");
            textBox.appendChild(run);
        }



        doc.save("C:\\Users\\daoma\\IdeaProjects\\apose_maven\\Export\\output_AB.pdf");

    }
}

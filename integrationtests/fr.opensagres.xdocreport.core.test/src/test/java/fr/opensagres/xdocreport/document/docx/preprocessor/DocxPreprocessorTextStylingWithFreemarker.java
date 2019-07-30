/**
 * Copyright (C) 2011-2015 The XDocReport Team <xdocreport@googlegroups.com>
 *
 * All rights reserved.
 *
 * Permission is hereby granted, free  of charge, to any person obtaining
 * a  copy  of this  software  and  associated  documentation files  (the
 * "Software"), to  deal in  the Software without  restriction, including
 * without limitation  the rights to  use, copy, modify,  merge, publish,
 * distribute,  sublicense, and/or sell  copies of  the Software,  and to
 * permit persons to whom the Software  is furnished to do so, subject to
 * the following conditions:
 *
 * The  above  copyright  notice  and  this permission  notice  shall  be
 * included in all copies or substantial portions of the Software.
 *
 * THE  SOFTWARE IS  PROVIDED  "AS  IS", WITHOUT  WARRANTY  OF ANY  KIND,
 * EXPRESS OR  IMPLIED, INCLUDING  BUT NOT LIMITED  TO THE  WARRANTIES OF
 * MERCHANTABILITY,    FITNESS    FOR    A   PARTICULAR    PURPOSE    AND
 * NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
 * LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
 * OF CONTRACT, TORT OR OTHERWISE,  ARISING FROM, OUT OF OR IN CONNECTION
 * WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */
package fr.opensagres.xdocreport.document.docx.preprocessor;

import java.io.InputStream;
import java.io.StringWriter;
import java.util.HashMap;

import org.junit.Assert;
import org.junit.Test;

import fr.opensagres.xdocreport.core.document.SyntaxKind;
import fr.opensagres.xdocreport.core.io.IOUtils;
import fr.opensagres.xdocreport.document.docx.preprocessor.sax.DocxPreprocessor;
import fr.opensagres.xdocreport.template.formatter.FieldsMetadata;
import fr.opensagres.xdocreport.template.formatter.IDocumentFormatter;
import fr.opensagres.xdocreport.template.freemarker.FreemarkerDocumentFormatter;

public class DocxPreprocessorTextStylingWithFreemarker
{

    @Test
    public void test2InstrText()
        throws Exception
    {
        DocxPreprocessor preprocessor = new DocxPreprocessor();
        InputStream stream =
                        IOUtils.toInputStream( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + "<w:document "
                                        + "xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
                                        + "xmlns:o=\"urn:schemas-microsoft-com:office:office\" "
                                        + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
                                        + "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" "
                                        + "xmlns:v=\"urn:schemas-microsoft-com:vml\" "
                                        + "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
                                        + "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" "
                                        + "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
                                        + "xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\">"
                                        
                                        + "<w:p w:rsidR=\"008E751F\" w:rsidRPr=\"00C20656\" w:rsidRDefault=\"00E4086C\" w:rsidP=\"00AA3AE5\">"
                                        + "<w:pPr>"
                                            + "<w:spacing w:after=\"0\"/>"
                                            + "<w:rPr>"
                                                + "<w:noProof/>"
                                                + "<w:lang w:val=\"en-US\"/>"
                                            + "</w:rPr>"
                                        + "</w:pPr>"
                                        + "<w:r>"
                                            + "<w:fldChar w:fldCharType=\"begin\"/>"
                                        + "</w:r>"
                                        + "<w:r w:rsidRPr=\"00C20656\">"
                                            + "<w:rPr>"
                                                + "<w:lang w:val=\"en-US\"/>"
                                            + "</w:rPr>"
                                            + "<w:instrText xml:space=\"preserve\"> MERGEFIELD  ${htmlText}  \\* MERGEFORMAT </w:instrText>"
                                        + "</w:r>"
                                        + "<w:r>"
                                            + "<w:fldChar w:fldCharType=\"separate\"/>"
                                        + "</w:r>"
                                        + "<w:r w:rsidR=\"006E20E5\" w:rsidRPr=\"00C20656\">"
                                            + "<w:rPr>"
                                                + "<w:noProof/>"
                                                + "<w:lang w:val=\"en-US\"/>"
                                            + "</w:rPr>"
                                            + "<w:t>«${htmlText}»</w:t>"
                                        + "</w:r>"
                                        + "<w:r>"
                                            + "<w:rPr>"
                                                + "<w:noProof/>"
                                            + "</w:rPr>"
                                            + "<w:fldChar w:fldCharType=\"end\"/>"
                                        + "</w:r>"
                                        + "</w:p>"
                                        
                                        + "</w:document>", "UTF-8"  );

        StringWriter writer = new StringWriter();
        FieldsMetadata metadata = new FieldsMetadata();
        metadata.addFieldAsTextStyling( "htmlText", SyntaxKind.Html );
        IDocumentFormatter formatter = new FreemarkerDocumentFormatter();
 
        preprocessor.preprocess( "word/document.xml", stream, writer, metadata, formatter, new HashMap<String, Object>() );

        Assert.assertEquals( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + "<w:document "
            + "xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
            + "xmlns:o=\"urn:schemas-microsoft-com:office:office\" "
            + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
            + "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" "
            + "xmlns:v=\"urn:schemas-microsoft-com:vml\" "
            + "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
            + "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" "
            + "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
            + "xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\">"
 
            + "[#assign ___NoEscape0=___TextStylingRegistry.transform(htmlText,\"Html\",false,\"DOCX\",\"0_elementId\",___context,\"word/document.xml\")] [#noescape]${___NoEscape0.textBefore}[/#noescape]"
             
             + "<w:p w:rsidR=\"008E751F\" w:rsidRPr=\"00C20656\" w:rsidRDefault=\"00E4086C\" w:rsidP=\"00AA3AE5\">"
             + "<w:pPr>"
                 + "<w:spacing w:after=\"0\"/>"
                 + "<w:rPr>"
                     + "<w:noProof/>"
                     + "<w:lang w:val=\"en-US\"/>"
                 + "</w:rPr>"
             + "</w:pPr>"
             //+ "<w:r>"
             //    + "<w:fldChar w:fldCharType=\"begin\"/>"
             //+ "</w:r>"
             //+ "<w:r w:rsidRPr=\"00C20656\">"
             //    + "<w:rPr>"
             //        + "<w:lang w:val=\"en-US\"/>"
             //    + "</w:rPr>"
             //    + "<w:instrText xml:space=\"preserve\"> MERGEFIELD  ${htmlText}  \\* MERGEFORMAT </w:instrText>"
             //+ "</w:r>"
             //+ "<w:r>"
             //    + "<w:fldChar w:fldCharType=\"separate\"/>"
             //+ "</w:r>"
             + "<w:r w:rsidR=\"006E20E5\" w:rsidRPr=\"00C20656\">"
                 + "<w:rPr>"
                     + "<w:noProof/>"
                     + "<w:lang w:val=\"en-US\"/>"
                 + "</w:rPr>"
             //    + "<w:t>«${htmlText}»</w:t>"
             + "<w:t>[#noescape]${___NoEscape0.textBody}[/#noescape]</w:t>"    
             + "</w:r>"
             //+ "<w:r>"
             //    + "<w:rPr>"
             //        + "<w:noProof/>"
             //    + "</w:rPr>"
             //    + "<w:fldChar w:fldCharType=\"end\"/>"
             //+ "</w:r>"
             + "</w:p>"
             + "[#noescape]${___NoEscape0.textEnd}[/#noescape]"
             + "</w:document>", writer.toString() );
    }

    @Test
    public void test2InstrText2()
            throws Exception
    {
        DocxPreprocessor preprocessor = new DocxPreprocessor();
        InputStream stream =
                IOUtils.toInputStream( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + "<w:document "
                        + "xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
                        + "xmlns:o=\"urn:schemas-microsoft-com:office:office\" "
                        + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
                        + "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" "
                        + "xmlns:v=\"urn:schemas-microsoft-com:vml\" "
                        + "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
                        + "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" "
                        + "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
                        + "xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\">"

                        + "<w:p w:rsidR=\"008E751F\" w:rsidRPr=\"00C20656\" w:rsidRDefault=\"00E4086C\" w:rsidP=\"00AA3AE5\">\n"
                        + "<w:pPr>\n"
                        + "<w:spacing w:after=\"0\"/>\n"
                        + "<w:rPr>\n"
                        + "<w:noProof/>\n"
                        + "<w:lang w:val=\"en-US\"/>\n"
                        + "</w:rPr>\n"
                        + "</w:pPr>\n"
                        + "<w:r>\n"
                        + "<w:fldChar w:fldCharType=\"begin\"/>\n"
                        + "</w:r>\n"
                        + "<w:r w:rsidRPr=\"00C20656\">\n"
                        + "  <w:rPr>\n"
                        + "    <w:lang w:val=\"en-US\"/>\n"
                        + "  </w:rPr>\n"
                        + "  <w:instrText xml:space=\"preserve\"> MERGEFIELD  ${htmlText}  \\* MERGEFORMAT </w:instrText>\n"
                        + "</w:r>\n"
                        + "<w:r>\n"
                        + "<w:fldChar w:fldCharType=\"separate\"/>\n"
                        + "</w:r>\n"
                        + "<w:r w:rsidR=\"006E20E5\" w:rsidRPr=\"00C20656\">\n"
                        + "<w:rPr>\n"
                        + "<w:noProof/>\n"
                        + "<w:lang w:val=\"en-US\"/>\n"
                        + "</w:rPr>\n"
                        + "<w:t>«${htmlText}»</w:t>\n"
                        + "</w:r>\n"
                        + "<w:r>\n"
                        + "<w:rPr>\n"
                        + "<w:noProof/>\n"
                        + "</w:rPr>\n"
                        + "<w:fldChar w:fldCharType=\"end\"/>\n"
                        + "</w:r>\n"
                        + "</w:p>\n"

                        + "</w:document>", "UTF-8"  );

        StringWriter writer = new StringWriter();
        FieldsMetadata metadata = new FieldsMetadata();
        metadata.addFieldAsTextStyling( "htmlText", SyntaxKind.Html );
        IDocumentFormatter formatter = new FreemarkerDocumentFormatter();

        preprocessor.preprocess( "word/document.xml", stream, writer, metadata, formatter, new HashMap<String, Object>() );

        Assert.assertEquals(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + "<w:document "
                + "xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
                + "xmlns:o=\"urn:schemas-microsoft-com:office:office\" "
                + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
                + "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" "
                + "xmlns:v=\"urn:schemas-microsoft-com:vml\" "
                + "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
                + "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" "
                + "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
                + "xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\">"
                + "[#assign ___NoEscape0=___TextStylingRegistry.transform(htmlText,\"Html\",false,\"DOCX\",\"0_elementId\",___context,\"word/document.xml\")] [#noescape]${___NoEscape0.textBefore}[/#noescape]"
                + "<w:p w:rsidR=\"008E751F\" w:rsidRPr=\"00C20656\" w:rsidRDefault=\"00E4086C\" w:rsidP=\"00AA3AE5\">\n"
                + "<w:pPr>\n"
                + "<w:spacing w:after=\"0\"/>\n"
                + "<w:rPr>\n"
                + "<w:noProof/>\n"
                + "<w:lang w:val=\"en-US\"/>\n"
                + "</w:rPr>\n"
                + "</w:pPr>\n\n\n\n"
                + "<w:r w:rsidR=\"006E20E5\" w:rsidRPr=\"00C20656\">\n"
                + "<w:rPr>\n"
                + "<w:noProof/>\n"
                + "<w:lang w:val=\"en-US\"/>\n"
                + "</w:rPr>\n"
                + "<w:t>[#noescape]${___NoEscape0.textBody}[/#noescape]</w:t>\n"
                + "</w:r>\n\n"
                + "</w:p>"
                + "[#noescape]${___NoEscape0.textEnd}[/#noescape]\n"
                + "</w:document>", writer.toString() );
    }

    @Test
    public void test2InstrTextWith2MD()
        throws Exception
    {
        DocxPreprocessor preprocessor = new DocxPreprocessor();
        InputStream stream =
                        IOUtils.toInputStream( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                                "<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:mo=\"http://schemas.microsoft.com/office/mac/office/2008/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:mv=\"urn:schemas-microsoft-com:mac:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 wp14\">\n" +
                                "    <w:body>\n" +
                                "        <w:p w14:paraId=\"1EFB801B\" w14:textId=\"77777777\" w:rsidR=\"00187DBF\" w:rsidRDefault=\"00187DBF\">\n" +
                                "            <w:pPr>\n" +
                                "                <w:rPr>\n" +
                                "                    <w:lang w:val=\"en-US\"/>\n" +
                                "                </w:rPr>\n" +
                                "            </w:pPr>\n" +
                                "            <w:r>\n" +
                                "                <w:rPr>\n" +
                                "                    <w:lang w:val=\"en-US\"/>\n" +
                                "                </w:rPr>\n" +
                                "                <w:fldChar w:fldCharType=\"begin\"/>\n" +
                                "            </w:r>\n" +
                                "            <w:r>\n" +
                                "                <w:rPr>\n" +
                                "                    <w:lang w:val=\"en-US\"/>\n" +
                                "                </w:rPr>\n" +
                                "                <w:instrText xml:space=\"preserve\"> </w:instrText>\n" +
                                "            </w:r>\n" +
                                "            <w:r w:rsidR=\"00DD1465\">\n" +
                                "                <w:rPr>\n" +
                                "                    <w:lang w:val=\"en-US\"/>\n" +
                                "                </w:rPr>\n" +
                                "                <w:instrText>MERGEFIELD ${field1}</w:instrText>\n" +
                                "            </w:r>\n" +
                                "            <w:r>\n" +
                                "                <w:rPr>\n" +
                                "                    <w:lang w:val=\"en-US\"/>\n" +
                                "                </w:rPr>\n" +
                                "                <w:instrText xml:space=\"preserve\"> \\* MERGEFORMAT </w:instrText>\n" +
                                "            </w:r>\n" +
                                "            <w:r>\n" +
                                "                <w:rPr>\n" +
                                "                    <w:lang w:val=\"en-US\"/>\n" +
                                "                </w:rPr>\n" +
                                "                <w:fldChar w:fldCharType=\"separate\"/>\n" +
                                "            </w:r>\n" +
                                "            <w:r w:rsidR=\"00DD1465\">\n" +
                                "                <w:rPr>\n" +
                                "                    <w:noProof/>\n" +
                                "                    <w:lang w:val=\"en-US\"/>\n" +
                                "                </w:rPr>\n" +
                                "                <w:t>«${field1}»</w:t>\n" +
                                "            </w:r>\n" +
                                "            <w:r>\n" +
                                "                <w:rPr>\n" +
                                "                    <w:lang w:val=\"en-US\"/>\n" +
                                "                </w:rPr>\n" +
                                "                <w:fldChar w:fldCharType=\"end\"/>\n" +
                                "            </w:r>\n" +
                                "        </w:p>\n" +
                                "        <w:p w14:paraId=\"316E5D55\" w14:textId=\"77777777\" w:rsidR=\"00DD1465\" w:rsidRDefault=\"00DD1465\">\n" +
                                "            <w:pPr>\n" +
                                "                <w:rPr>\n" +
                                "                    <w:lang w:val=\"en-US\"/>\n" +
                                "                </w:rPr>\n" +
                                "            </w:pPr>\n" +
                                "        </w:p>\n" +
                                "        <w:p w14:paraId=\"71FF5154\" w14:textId=\"77777777\" w:rsidR=\"00DD1465\" w:rsidRDefault=\"00DD1465\" w:rsidP=\"00DD1465\">\n" +
                                "            <w:pPr>\n" +
                                "                <w:rPr>\n" +
                                "                    <w:lang w:val=\"en-US\"/>\n" +
                                "                </w:rPr>\n" +
                                "            </w:pPr>\n" +
                                "            <w:r>\n" +
                                "                <w:rPr>\n" +
                                "                    <w:lang w:val=\"en-US\"/>\n" +
                                "                </w:rPr>\n" +
                                "                <w:fldChar w:fldCharType=\"begin\"/>\n" +
                                "            </w:r>\n" +
                                "            <w:r>\n" +
                                "                <w:rPr>\n" +
                                "                    <w:lang w:val=\"en-US\"/>\n" +
                                "                </w:rPr>\n" +
                                "                <w:instrText xml:space=\"preserve\"> </w:instrText>\n" +
                                "            </w:r>\n" +
                                "            <w:r>\n" +
                                "                <w:rPr>\n" +
                                "                    <w:lang w:val=\"en-US\"/>\n" +
                                "                </w:rPr>\n" +
                                "                <w:instrText>MERGEFIELD ${field2} \\* MERGEFORMAT </w:instrText>\n" +
                                "            </w:r>\n" +
                                "            <w:r>\n" +
                                "                <w:rPr>\n" +
                                "                    <w:lang w:val=\"en-US\"/>\n" +
                                "                </w:rPr>\n" +
                                "                <w:fldChar w:fldCharType=\"separate\"/>\n" +
                                "            </w:r>\n" +
                                "            <w:r>\n" +
                                "                <w:rPr>\n" +
                                "                    <w:noProof/>\n" +
                                "                    <w:lang w:val=\"en-US\"/>\n" +
                                "                </w:rPr>\n" +
                                "                <w:t>«${field2}»</w:t>\n" +
                                "            </w:r>\n" +
                                "            <w:r>\n" +
                                "                <w:rPr>\n" +
                                "                    <w:lang w:val=\"en-US\"/>\n" +
                                "                </w:rPr>\n" +
                                "                <w:fldChar w:fldCharType=\"end\"/>\n" +
                                "            </w:r>\n" +
                                "        </w:p>\n" +
                                "    </w:body>\n" +
                                "</w:document>"  );

        StringWriter writer = new StringWriter();
        FieldsMetadata metadata = new FieldsMetadata();
        metadata.addFieldAsTextStyling( "field1", SyntaxKind.Html );
        metadata.addFieldAsTextStyling( "field2", SyntaxKind.Html );
        IDocumentFormatter formatter = new FreemarkerDocumentFormatter();

        preprocessor.preprocess( "word/document.xml", stream, writer, metadata, formatter, new HashMap<String, Object>() );

        Assert.assertEquals( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:mo=\"http://schemas.microsoft.com/office/mac/office/2008/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:mv=\"urn:schemas-microsoft-com:mac:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 wp14\">\n" +
                "    <w:body>\n" +
                "        [#assign ___NoEscape0=___TextStylingRegistry.transform(field1,\"Html\",false,\"DOCX\",\"0_elementId\",___context,\"word/document.xml\")] [#noescape]${___NoEscape0.textBefore}[/#noescape]<w:p w14:paraId=\"1EFB801B\" w14:textId=\"77777777\" w:rsidR=\"00187DBF\" w:rsidRDefault=\"00187DBF\">\n" +
                "            <w:pPr>\n" +
                "                <w:rPr>\n" +
                "                    <w:lang w:val=\"en-US\"/>\n" +
                "                </w:rPr>\n" +
                "            </w:pPr>\n            \n            \n            \n            \n            \n" +
                "            <w:r w:rsidR=\"00DD1465\">\n" +
                "                <w:rPr>\n" +
                "                    <w:noProof/>\n" +
                "                    <w:lang w:val=\"en-US\"/>\n" +
                "                </w:rPr>\n" +
                "                <w:t>[#noescape]${___NoEscape0.textBody}[/#noescape]</w:t>\n" +
                "            </w:r>\n" +
                "            \n" +
                "        </w:p>[#noescape]${___NoEscape0.textEnd}[/#noescape]\n" +
                "        <w:p w14:paraId=\"316E5D55\" w14:textId=\"77777777\" w:rsidR=\"00DD1465\" w:rsidRDefault=\"00DD1465\">\n" +
                "            <w:pPr>\n" +
                "                <w:rPr>\n" +
                "                    <w:lang w:val=\"en-US\"/>\n" +
                "                </w:rPr>\n" +
                "            </w:pPr>\n" +
                "        </w:p>\n" +
                "        [#assign ___NoEscape1=___TextStylingRegistry.transform(field2,\"Html\",false,\"DOCX\",\"1_elementId\",___context,\"word/document.xml\")] [#noescape]${___NoEscape1.textBefore}[/#noescape]<w:p w14:paraId=\"71FF5154\" w14:textId=\"77777777\" w:rsidR=\"00DD1465\" w:rsidRDefault=\"00DD1465\" w:rsidP=\"00DD1465\">\n" +
                "            <w:pPr>\n" +
                "                <w:rPr>\n" +
                "                    <w:lang w:val=\"en-US\"/>\n" +
                "                </w:rPr>\n" +
                "            </w:pPr>\n            \n            \n            \n            \n" +
                "            <w:r>\n" +
                "                <w:rPr>\n" +
                "                    <w:noProof/>\n" +
                "                    <w:lang w:val=\"en-US\"/>\n" +
                "                </w:rPr>\n" +
                "                <w:t>[#noescape]${___NoEscape1.textBody}[/#noescape]</w:t>\n" +
                "            </w:r>\n" +
                "            \n" +
                "        </w:p>[#noescape]${___NoEscape1.textEnd}[/#noescape]\n" +
                "    </w:body>\n" +
                "</w:document>", writer.toString() );
    }

    @Test
    public void test2InstrTextDocxBreakingStuff()
            throws Exception
    {
        DocxPreprocessor preprocessor = new DocxPreprocessor();
        InputStream stream =
                IOUtils.toInputStream( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                        "<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 w16se w16cid wp14\">\n" +
                        "    <w:body>\n" +
                        "        <w:p w14:paraId=\"472B0AAD\" w14:textId=\"00CF3D2E\" w:rsidR=\"008C0057\" w:rsidRPr=\"00350669\" w:rsidRDefault=\"00350669\">\n" +
                        "            <w:pPr>\n" +
                        "                <w:rPr>\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "            </w:pPr>\n" +
                        "            <w:bookmarkStart w:id=\"0\" w:name=\"_GoBack\" />\n" +
                        "            <w:bookmarkEnd w:id=\"0\" />\n" +
                        "            <w:r>\n" +
                        "                <w:rPr>\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "                <w:t>Hello,</w:t>\n" +
                        "            </w:r>\n" +
                        "            <w:r w:rsidR=\"003F7C24\">\n" +
                        "                <w:rPr>\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "                <w:t xml:space=\"preserve\"> </w:t>\n" +
                        "            </w:r>\n" +
                        "            <w:r w:rsidR=\"00F24665\">\n" +
                        "                <w:rPr>\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "                <w:fldChar w:fldCharType=\"begin\" />\n" +
                        "            </w:r>\n" +
                        "            <w:r w:rsidR=\"00F24665\">\n" +
                        "                <w:rPr>\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "                <w:instrText xml:space=\"preserve\"> MERGEFIELD ${user_name} </w:instrText>\n" +
                        "            </w:r>\n" +
                        "            <w:r w:rsidR=\"00F24665\">\n" +
                        "                <w:rPr>\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "                <w:fldChar w:fldCharType=\"separate\" />\n" +
                        "            </w:r>\n" +
                        "            <w:r w:rsidR=\"00F24665\">\n" +
                        "                <w:rPr>\n" +
                        "                    <w:noProof />\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "                <w:t>«${user_name}»</w:t>\n" +
                        "            </w:r>\n" +
                        "            <w:r w:rsidR=\"00F24665\">\n" +
                        "                <w:rPr>\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "                <w:fldChar w:fldCharType=\"end\" />\n" +
                        "            </w:r>\n" +
                        "            <w:r w:rsidR=\"00F24665\">\n" +
                        "                <w:rPr>\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "                <w:t>!</w:t>\n" +
                        "            </w:r>\n" +
                        "        </w:p>\n" +
                        "        <w:sectPr w:rsidR=\"008C0057\" w:rsidRPr=\"00350669\" w:rsidSect=\"00B23DB1\">\n" +
                        "            <w:pgSz w:w=\"11900\" w:h=\"16840\" />\n" +
                        "            <w:pgMar w:top=\"1134\" w:right=\"850\" w:bottom=\"1134\" w:left=\"1701\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\" />\n" +
                        "            <w:cols w:space=\"708\" />\n" +
                        "            <w:docGrid w:linePitch=\"360\" />\n" +
                        "        </w:sectPr>\n" +
                        "    </w:body>\n" +
                        "</w:document>"  );

        StringWriter writer = new StringWriter();
        FieldsMetadata metadata = new FieldsMetadata();
        IDocumentFormatter formatter = new FreemarkerDocumentFormatter();

        preprocessor.preprocess( "word/document.xml", stream, writer, metadata, formatter, new HashMap<String, Object>() );

        Assert.assertEquals( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 w16se w16cid wp14\">\n" +
                "    <w:body>\n" +
                "        <w:p w14:paraId=\"472B0AAD\" w14:textId=\"00CF3D2E\" w:rsidR=\"008C0057\" w:rsidRPr=\"00350669\" w:rsidRDefault=\"00350669\">\n" +
                "            <w:pPr>\n" +
                "                <w:rPr>\n" +
                "                    <w:lang w:val=\"en-US\"/>\n" +
                "                </w:rPr>\n" +
                "            </w:pPr>\n" +
                "            <w:bookmarkStart w:id=\"0\" w:name=\"_GoBack\"/>\n" +
                "            <w:bookmarkEnd w:id=\"0\"/>\n" +
                "            <w:r>\n" +
                "                <w:rPr>\n" +
                "                    <w:lang w:val=\"en-US\"/>\n" +
                "                </w:rPr>\n" +
                "                <w:t>Hello,</w:t>\n" +
                "            </w:r>\n" +
                "            <w:r w:rsidR=\"003F7C24\">\n" +
                "                <w:rPr>\n" +
                "                    <w:lang w:val=\"en-US\"/>\n" +
                "                </w:rPr>\n" +
                "                <w:t xml:space=\"preserve\"> </w:t>\n" +
                "            </w:r>\n" +
                "            \n" +
                "            \n" +
                "            \n" +
                "            <w:r w:rsidR=\"00F24665\">\n" +
                "                <w:rPr>\n" +
                "                    <w:noProof/>\n" +
                "                    <w:lang w:val=\"en-US\"/>\n" +
                "                </w:rPr>\n" +
                "                <w:t>${user_name}</w:t>\n" +
                "            </w:r>\n" +
                "            \n" +
                "            <w:r w:rsidR=\"00F24665\">\n" +
                "                <w:rPr>\n" +
                "                    <w:lang w:val=\"en-US\"/>\n" +
                "                </w:rPr>\n" +
                "                <w:t>!</w:t>\n" +
                "            </w:r>\n" +
                "        </w:p>\n" +
                "        <w:sectPr w:rsidR=\"008C0057\" w:rsidRPr=\"00350669\" w:rsidSect=\"00B23DB1\">\n" +
                "            <w:pgSz w:w=\"11900\" w:h=\"16840\"/>\n" +
                "            <w:pgMar w:top=\"1134\" w:right=\"850\" w:bottom=\"1134\" w:left=\"1701\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/>\n" +
                "            <w:cols w:space=\"708\"/>\n" +
                "            <w:docGrid w:linePitch=\"360\"/>\n" +
                "        </w:sectPr>\n" +
                "    </w:body>\n" +
                "</w:document>", writer.toString() );
    }

    @Test
    public void test2InstrTextDocxBreakingStuff2()
            throws Exception
    {
        DocxPreprocessor preprocessor = new DocxPreprocessor();
        InputStream stream =
                IOUtils.toInputStream( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                        "<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 w16se w16cid wp14\">\n" +
                        "    <w:body>\n" +
                        "        <w:p w14:paraId=\"472B0AAD\" w14:textId=\"5F7E69FC\" w:rsidR=\"008C0057\" w:rsidRPr=\"00350669\" w:rsidRDefault=\"00350669\">\n" +
                        "            <w:pPr>\n" +
                        "                <w:rPr>\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "            </w:pPr>\n" +
                        "            <w:r>\n" +
                        "                <w:rPr>\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "                <w:t>Hello,</w:t>\n" +
                        "            </w:r>\n" +
                        "            <w:r w:rsidR=\"0020161A\">\n" +
                        "                <w:rPr>\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "                <w:fldChar w:fldCharType=\"begin\" />\n" +
                        "            </w:r>\n" +
                        "            <w:r w:rsidR=\"0020161A\">\n" +
                        "                <w:rPr>\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "                <w:instrText xml:space=\"preserve\"> MERGEFIELD ${</w:instrText>\n" +
                        "            </w:r>\n" +
                        "            <w:bookmarkStart w:id=\"0\" w:name=\"_GoBack\" />\n" +
                        "            <w:bookmarkEnd w:id=\"0\" />\n" +
                        "            <w:r w:rsidR=\"0020161A\" w:rsidRPr=\"0020161A\">\n" +
                        "                <w:rPr>\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "                <w:instrText>user_name</w:instrText>\n" +
                        "            </w:r>\n" +
                        "            <w:r w:rsidR=\"0020161A\">\n" +
                        "                <w:rPr>\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "                <w:instrText xml:space=\"preserve\">} </w:instrText>\n" +
                        "            </w:r>\n" +
                        "            <w:r w:rsidR=\"0020161A\">\n" +
                        "                <w:rPr>\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "                <w:fldChar w:fldCharType=\"separate\" />\n" +
                        "            </w:r>\n" +
                        "            <w:r w:rsidR=\"0020161A\">\n" +
                        "                <w:rPr>\n" +
                        "                    <w:noProof />\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "                <w:t>«${}»</w:t>\n" +
                        "            </w:r>\n" +
                        "            <w:r w:rsidR=\"0020161A\">\n" +
                        "                <w:rPr>\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "                <w:fldChar w:fldCharType=\"end\" />\n" +
                        "            </w:r>\n" +
                        "            <w:r w:rsidR=\"00F24665\">\n" +
                        "                <w:rPr>\n" +
                        "                    <w:lang w:val=\"en-US\" />\n" +
                        "                </w:rPr>\n" +
                        "                <w:t>!</w:t>\n" +
                        "            </w:r>\n" +
                        "        </w:p>\n" +
                        "        <w:sectPr w:rsidR=\"008C0057\" w:rsidRPr=\"00350669\" w:rsidSect=\"00B23DB1\">\n" +
                        "            <w:pgSz w:w=\"11900\" w:h=\"16840\" />\n" +
                        "            <w:pgMar w:top=\"1134\" w:right=\"850\" w:bottom=\"1134\" w:left=\"1701\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\" />\n" +
                        "            <w:cols w:space=\"708\" />\n" +
                        "            <w:docGrid w:linePitch=\"360\" />\n" +
                        "        </w:sectPr>\n" +
                        "    </w:body>\n" +
                        "</w:document>"  );

        StringWriter writer = new StringWriter();
        FieldsMetadata metadata = new FieldsMetadata();
        IDocumentFormatter formatter = new FreemarkerDocumentFormatter();

        preprocessor.preprocess( "word/document.xml", stream, writer, metadata, formatter, new HashMap<String, Object>() );

        Assert.assertEquals( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 w16se w16cid wp14\">\n" +
                "    <w:body>\n" +
                "        <w:p w14:paraId=\"472B0AAD\" w14:textId=\"5F7E69FC\" w:rsidR=\"008C0057\" w:rsidRPr=\"00350669\" w:rsidRDefault=\"00350669\">\n" +
                "            <w:pPr>\n" +
                "                <w:rPr>\n" +
                "                    <w:lang w:val=\"en-US\"/>\n" +
                "                </w:rPr>\n" +
                "            </w:pPr>\n" +
                "            <w:r>\n" +
                "                <w:rPr>\n" +
                "                    <w:lang w:val=\"en-US\"/>\n" +
                "                </w:rPr>\n" +
                "                <w:t>Hello,</w:t>\n" +
                "            </w:r>\n" +
                "            \n" +
                "            \n" +
                "            <w:bookmarkStart w:id=\"0\" w:name=\"_GoBack\"/>\n" +
                "            <w:bookmarkEnd w:id=\"0\"/>\n" +
                "            \n" +
                "            \n" +
                "            \n" +
                "            <w:r w:rsidR=\"0020161A\">\n" +
                "                <w:rPr>\n" +
                "                    <w:noProof/>\n" +
                "                    <w:lang w:val=\"en-US\"/>\n" +
                "                </w:rPr>\n" +
                "                <w:t>${user_name}</w:t>\n" +
                "            </w:r>\n" +
                "            \n" +
                "            <w:r w:rsidR=\"00F24665\">\n" +
                "                <w:rPr>\n" +
                "                    <w:lang w:val=\"en-US\"/>\n" +
                "                </w:rPr>\n" +
                "                <w:t>!</w:t>\n" +
                "            </w:r>\n" +
                "        </w:p>\n" +
                "        <w:sectPr w:rsidR=\"008C0057\" w:rsidRPr=\"00350669\" w:rsidSect=\"00B23DB1\">\n" +
                "            <w:pgSz w:w=\"11900\" w:h=\"16840\"/>\n" +
                "            <w:pgMar w:top=\"1134\" w:right=\"850\" w:bottom=\"1134\" w:left=\"1701\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/>\n" +
                "            <w:cols w:space=\"708\"/>\n" +
                "            <w:docGrid w:linePitch=\"360\"/>\n" +
                "        </w:sectPr>\n" +
                "    </w:body>\n" +
                "</w:document>", writer.toString() );
    }


    @Test
    public void textStylingWithSimpleField()
        throws Exception
    {
        DocxPreprocessor preprocessor = new DocxPreprocessor();
        InputStream stream =
                        IOUtils.toInputStream( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + "<w:document "
                                        + "xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
                                        + "xmlns:o=\"urn:schemas-microsoft-com:office:office\" "
                                        + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
                                        + "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" "
                                        + "xmlns:v=\"urn:schemas-microsoft-com:vml\" "
                                        + "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
                                        + "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" "
                                        + "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
                                        + "xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\">"
                                        
                                        +"<w:tbl>"
                                            +"<w:tblPr>"
                                                +"<w:tblStyle w:val=\"Grilledutableau\"/>"
                                                +"<w:tblW w:w=\"0\" w:type=\"auto\"/>"
                                                +"<w:tblLook w:val=\"04A0\" w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"1\" w:lastColumn=\"0\" w:noHBand=\"0\" w:noVBand=\"1\"/>"
                                            +"</w:tblPr>"
                                            +"<w:tblGrid>"
                                                +"<w:gridCol w:w=\"3070\"/>"
                                            +"</w:tblGrid>"
                                            +"<w:tr w:rsidR=\"00916516\" w:rsidTr=\"005D6D71\">"
                                                +"<w:tc>"
                                                    +"<w:tcPr>"
                                                        +"<w:tcW w:w=\"3070\" w:type=\"dxa\"/>"
                                                    +"</w:tcPr>"
                                                    +"<w:p w:rsidR=\"00916516\" w:rsidRPr=\"005D6D71\" w:rsidRDefault=\"00916516\">"
                                                        +"<w:fldSimple w:instr=\" MERGEFIELD  ${row.html}  \\* MERGEFORMAT \">"
                                                            +"<w:r>"
                                                                +"<w:rPr>"
                                                                    +"<w:noProof/>"
                                                                +"</w:rPr>"
                                                                +"<w:t>«${row.html}»</w:t>"
                                                            +"</w:r>"
                                                        +"</w:fldSimple>"
                                                    +"</w:p>"
                                                +"</w:tc>"
                                            +"</w:tr>"
                                        +"</w:tbl>"
                                        
                                        + "</w:document>", "UTF-8"  );

        StringWriter writer = new StringWriter();

        IDocumentFormatter formatter = new FreemarkerDocumentFormatter();
        FieldsMetadata metadata = new FieldsMetadata();        
        metadata.addFieldAsTextStyling( "row.html", SyntaxKind.Html );
 
        preprocessor.preprocess( "word/document.xml", stream, writer, metadata, formatter, new HashMap<String, Object>() );

        Assert.assertEquals( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + "<w:document "
            + "xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
            + "xmlns:o=\"urn:schemas-microsoft-com:office:office\" "
            + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
            + "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" "
            + "xmlns:v=\"urn:schemas-microsoft-com:vml\" "
            + "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
            + "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" "
            + "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
            + "xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\">"
 
            +"<w:tbl>"
            +"<w:tblPr>"
                +"<w:tblStyle w:val=\"Grilledutableau\"/>"
                +"<w:tblW w:w=\"0\" w:type=\"auto\"/>"
                +"<w:tblLook w:val=\"04A0\" w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"1\" w:lastColumn=\"0\" w:noHBand=\"0\" w:noVBand=\"1\"/>"
            +"</w:tblPr>"
            +"<w:tblGrid>"
                +"<w:gridCol w:w=\"3070\"/>"
            +"</w:tblGrid>"
            +"<w:tr w:rsidR=\"00916516\" w:rsidTr=\"005D6D71\">"
                +"<w:tc>"
                    +"<w:tcPr>"
                        +"<w:tcW w:w=\"3070\" w:type=\"dxa\"/>"
                    +"</w:tcPr>"
                    +"[#assign ___NoEscape0=___TextStylingRegistry.transform(row.html,\"Html\",false,\"DOCX\",\"0_elementId\",___context,\"word/document.xml\")] [#noescape]${___NoEscape0.textBefore}[/#noescape]"
                    +"<w:p w:rsidR=\"00916516\" w:rsidRPr=\"005D6D71\" w:rsidRDefault=\"00916516\">"
                        //+"<w:fldSimple w:instr=\" MERGEFIELD  ${row.html}  \\* MERGEFORMAT \">"
                            +"<w:r>"
                                +"<w:rPr>"
                                    +"<w:noProof/>"
                                +"</w:rPr>"
                                //+"<w:t>«${row.html}»</w:t>"
                                +"<w:t>[#noescape]${___NoEscape0.textBody}[/#noescape]</w:t>"
                            +"</w:r>"
                        //+"</w:fldSimple>"*/
                    +"</w:p>"
                    +"[#noescape]${___NoEscape0.textEnd}[/#noescape]"
                +"</w:tc>"
            +"</w:tr>"
        +"</w:tbl>"
        + "</w:document>", writer.toString() );
    }
    
    @Test
    public void textStylingInsideTableRow()
        throws Exception
    {
        DocxPreprocessor preprocessor = new DocxPreprocessor();
        InputStream stream =
                        IOUtils.toInputStream( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + "<w:document "
                                        + "xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
                                        + "xmlns:o=\"urn:schemas-microsoft-com:office:office\" "
                                        + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
                                        + "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" "
                                        + "xmlns:v=\"urn:schemas-microsoft-com:vml\" "
                                        + "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
                                        + "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" "
                                        + "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
                                        + "xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\">"
                                        
                                        +"<w:tbl>"
                                            +"<w:tblPr>"
                                                +"<w:tblStyle w:val=\"Grilledutableau\"/>"
                                                +"<w:tblW w:w=\"0\" w:type=\"auto\"/>"
                                                +"<w:tblLook w:val=\"04A0\" w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"1\" w:lastColumn=\"0\" w:noHBand=\"0\" w:noVBand=\"1\"/>"
                                            +"</w:tblPr>"
                                            +"<w:tblGrid>"
                                                +"<w:gridCol w:w=\"3070\"/>"
                                            +"</w:tblGrid>"
                                            +"<w:tr w:rsidR=\"00916516\" w:rsidTr=\"005D6D71\">"
                                                +"<w:tc>"
                                                    +"<w:tcPr>"
                                                        +"<w:tcW w:w=\"3070\" w:type=\"dxa\"/>"
                                                    +"</w:tcPr>"
                                                    +"<w:p w:rsidR=\"00916516\" w:rsidRPr=\"005D6D71\" w:rsidRDefault=\"00916516\">"
                                                        +"<w:fldSimple w:instr=\" MERGEFIELD  ${row.html}  \\* MERGEFORMAT \">"
                                                            +"<w:r>"
                                                                +"<w:rPr>"
                                                                    +"<w:noProof/>"
                                                                +"</w:rPr>"
                                                                +"<w:t>«${row.html}»</w:t>"
                                                            +"</w:r>"
                                                        +"</w:fldSimple>"
                                                    +"</w:p>"
                                                +"</w:tc>"
                                            +"</w:tr>"
                                        +"</w:tbl>"
                                        
                                        + "</w:document>", "UTF-8"  );

        StringWriter writer = new StringWriter();

        IDocumentFormatter formatter = new FreemarkerDocumentFormatter();
        FieldsMetadata metadata = new FieldsMetadata();        
        metadata.addFieldAsTextStyling( "row.html", SyntaxKind.Html );
        metadata.addFieldAsList( "row.html" );
 
        preprocessor.preprocess( "word/document.xml", stream, writer, metadata, formatter, new HashMap<String, Object>() );

        Assert.assertEquals( "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + "<w:document "
            + "xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
            + "xmlns:o=\"urn:schemas-microsoft-com:office:office\" "
            + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
            + "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" "
            + "xmlns:v=\"urn:schemas-microsoft-com:vml\" "
            + "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
            + "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" "
            + "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
            + "xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\">"
 
            +"<w:tbl>"
            +"<w:tblPr>"
                +"<w:tblStyle w:val=\"Grilledutableau\"/>"
                +"<w:tblW w:w=\"0\" w:type=\"auto\"/>"
                +"<w:tblLook w:val=\"04A0\" w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"1\" w:lastColumn=\"0\" w:noHBand=\"0\" w:noVBand=\"1\"/>"
            +"</w:tblPr>"
            +"<w:tblGrid>"
                +"<w:gridCol w:w=\"3070\"/>"
            +"</w:tblGrid>"

            +"[#list row as item_row]"
            +"<w:tr w:rsidR=\"00916516\" w:rsidTr=\"005D6D71\">"
                +"<w:tc>"
                    +"<w:tcPr>"
                        +"<w:tcW w:w=\"3070\" w:type=\"dxa\"/>"
                    +"</w:tcPr>"
                    +"[#assign ___NoEscape0=___TextStylingRegistry.transform(item_row.html,\"Html\",false,\"DOCX\",\"0_elementId\",___context,\"word/document.xml\")] [#noescape]${___NoEscape0.textBefore}[/#noescape]"
                    +"<w:p w:rsidR=\"00916516\" w:rsidRPr=\"005D6D71\" w:rsidRDefault=\"00916516\">"
                        //+"<w:fldSimple w:instr=\" MERGEFIELD  ${row.html}  \\* MERGEFORMAT \">"
                            +"<w:r>"
                                +"<w:rPr>"
                                    +"<w:noProof/>"
                                +"</w:rPr>"
                                //+"<w:t>«${row.html}»</w:t>"
                                +"<w:t>[#noescape]${___NoEscape0.textBody}[/#noescape]</w:t>"
                            +"</w:r>"
                        //+"</w:fldSimple>"*/
                    +"</w:p>"
                    +"[#noescape]${___NoEscape0.textEnd}[/#noescape]"
                +"</w:tc>"
            +"</w:tr>"
            +"[/#list]"
            
        +"</w:tbl>"
        + "</w:document>", writer.toString() );
    }  
}

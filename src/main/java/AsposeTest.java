import com.aspose.words.License;
import com.aspose.words.SaveFormat;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

public class AsposeTest {
    public static void main(String args[]) {
        try {
            wordToHtml("/Users/jimzhou/Documents/doc/bigData/课题3项目预申报.docx", "/Users/jimzhou/Documents/doc/bigData");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static {
        try (InputStream is = new FileInputStream(new File("/Users/jimzhou/coding/j_zjw_common/lib/license.xml"))) {
            com.aspose.words.License license = new License();
            license.setLicense(is);
        } catch (Exception e) {
            System.out.println(e);
        }
    }

    public static void wordToHtml(String wordPath, String htmlPath) throws Exception {

        com.aspose.words.Document document = new com.aspose.words.Document(wordPath);
//        com.aspose.words.HtmlSaveOptions options = new com.aspose.words.HtmlSaveOptions();
//        com.aspose.words.PdfSaveOptions options = new com.aspose.words.PdfSaveOptions();
        com.aspose.words.ImageSaveOptions options = new com.aspose.words.ImageSaveOptions(SaveFormat.PNG);
        com.aspose.words.ImageSaveOptions imageSaveOptions =
                new com.aspose.words.ImageSaveOptions(SaveFormat.JPEG);
        imageSaveOptions.setResolution(128);
        int pageindex = 0;
        int pagecount = document.getPageCount();
        for (int i = 1; i < pagecount + 1; i++) {
            try {
                String imgpath = htmlPath + "/" + "tt" + "_" + i + ".jpg";
                imageSaveOptions.setPageIndex(pageindex);
                document.save(imgpath, imageSaveOptions);
                pageindex++;

            } catch (Exception e) {

            }
        }


    }
}

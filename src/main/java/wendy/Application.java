package wendy;

import com.aspose.pdf.Document;
import com.aspose.pdf.ExcelSaveOptions;
import com.aspose.pdf.facades.PdfFileEditor;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;


public class Application {

    public static void main(String[] args) throws IOException {

        List<InputStream> files;
        try (Stream<Path> paths = Files.walk(Paths.get("/Users/camilo_hernandez/Downloads/wendy2"))) {
            files = paths
                    .filter(Files::isRegularFile)
                    .filter(path -> !isHidden(path))
                    .filter(path -> !isMergedFile(path))
                    .map(path -> getInputStream(path))
                    .collect(Collectors.toList());
        }
        InputStream[] filesInputStream = new InputStream[files.size()];
        filesInputStream = files.toArray(filesInputStream);
        PdfFileEditor fileEditor = new PdfFileEditor();
        OutputStream outputStream = Files.newOutputStream(Path.of("/Users/camilo_hernandez/Downloads/wendy2/merged_files.pdf"));
        fileEditor.concatenate(filesInputStream, outputStream);
        System.out.println("PDFs have been merged");

        Document doc = new Document(Files.newInputStream(Path.of("/Users/camilo_hernandez/Downloads/wendy2/merged_files.pdf")));
        ExcelSaveOptions options = new ExcelSaveOptions();
        options.setFormat(ExcelSaveOptions.ExcelFormat.XLSX);
        doc.save("/Users/camilo_hernandez/Downloads/wendy2/merge_excel.xlsx", options);
        System.out.println("XLS has been created");
        Files.delete(Path.of("/Users/camilo_hernandez/Downloads/wendy2/merged_files.pdf"));
        System.out.println("Merged PDF has been deleted");
    }

    private static boolean isMergedFile(Path path) {
        return path.getFileName().equals("merged_files.pdf");
    }


    public static InputStream getInputStream(Path path) {
        try {
            return Files.newInputStream(path);
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException(e);
        }
    }

    public static boolean isHidden(Path path) {
        try {
            return Files.isHidden(path);
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException(e);
        }
    }
}

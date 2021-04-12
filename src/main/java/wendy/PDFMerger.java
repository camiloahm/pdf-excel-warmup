package wendy;

import com.aspose.pdf.Document;
import com.aspose.pdf.ExcelSaveOptions;
import com.aspose.pdf.facades.PdfFileEditor;
import picocli.CommandLine;
import picocli.CommandLine.Command;
import picocli.CommandLine.Option;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.Callable;
import java.util.stream.Collectors;
import java.util.stream.Stream;

@Command(name = "wendy's pdf merger", mixinStandardHelpOptions = true, version = "1.0",
        description = "Merges pdf files and convert them into a xlsx file")
public class PDFMerger implements Callable<Integer> {

    @Option(names = {"-o", "--origin"}, description = "Directory with all the pdfs", interactive = true)
    private File file;

    @Override
    public Integer call() throws Exception { // your business logic goes here...
        List<InputStream> files;
        System.out.println("The origin file is: " + file.getAbsolutePath());
        try (Stream<Path> paths = Files.walk(file.toPath())) {
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
        String mergePDF = file.getAbsolutePath() + "/merged_files.pdf";
        OutputStream outputStream = Files.newOutputStream(Path.of(mergePDF));
        fileEditor.concatenate(filesInputStream, outputStream);
        System.out.println("PDFs have been merged");

        Document doc = new Document(Files.newInputStream(Path.of(mergePDF)));
        ExcelSaveOptions options = new ExcelSaveOptions();
        options.setFormat(ExcelSaveOptions.ExcelFormat.XLSX);
        String xslxFile = file.getAbsolutePath() + "/merged_excel.xlsx";
        doc.save(xslxFile, options);
        System.out.println("XLS has been created");
        Files.delete(Path.of(mergePDF));
        System.out.println("Merged PDF has been deleted");

        return 0;
    }


    private boolean isMergedFile(Path path) {
        return path.getFileName().equals("merged_files.pdf");
    }


    public InputStream getInputStream(Path path) {
        try {
            return Files.newInputStream(path);
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException(e);
        }
    }

    public boolean isHidden(Path path) {
        try {
            return Files.isHidden(path);
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException(e);
        }
    }

    public static void main(String... args) {
        int exitCode = new CommandLine(new PDFMerger()).execute("-o");
        System.exit(exitCode);
    }


}

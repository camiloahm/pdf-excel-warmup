package wendy;

import com.itextpdf.forms.PdfAcroForm;
import com.itextpdf.forms.PdfPageFormCopier;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.kernel.utils.PdfMerger;
import org.apache.commons.collections4.ArrayStack;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.*;
import picocli.CommandLine;
import picocli.CommandLine.Command;
import picocli.CommandLine.Option;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.List;
import java.util.Scanner;
import java.util.concurrent.Callable;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import static wendy.FileUtil.*;
import static wendy.SheetTemplate.*;


@Command(name = "wendy's pdf merger", mixinStandardHelpOptions = true, version = "1.0",
        description = "Merges pdf files and convert them into a xlsx file")
public class PDFMerger implements Callable<Integer> {

    @Option(names = {"-o", "--origin"}, description = "Directory with all the pdfs", interactive = true)
    private File file;


    static final DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("d-MM-yy");

    @Override
    public Integer call() throws Exception { // your business logic goes here...
        List<InputStream> files;
        String mergePDFName = file.getAbsolutePath() + "/merged_files.pdf";
        String excelFileName = file.getAbsolutePath() + "/merged_excel.xlsx";
        mergePDFs(mergePDFName);
        generateExcel(mergePDFName, excelFileName);
        return 0;
    }

    private void generateExcel(String mergePDFName, String excelFileName) {
        try (PDDocument document = PDDocument.load(new File(mergePDFName));
             FileOutputStream fileOut = new FileOutputStream(excelFileName)) {

            String sheetName = "Wendy";
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.createSheet(sheetName);
            XSSFCell cell;
            int rowPosition = 1;

            PDFTextStripperByArea stripper = new PDFTextStripperByArea();
            stripper.setSortByPosition(true);
            PDFTextStripper tStripper = new PDFTextStripper();
            String str = tStripper.getText(document);

            Scanner scnLine = new Scanner(str);
            String line;

            XSSFRow rowHeaders = sheet.createRow(0);

            for (SheetTemplate column : SheetTemplate.values()) {
                rowHeaders.createCell(column.getCell()).setCellValue(column.getName());
            }

            for (int i = 0; i < SheetTemplate.values().length; i++) {
                sheet.autoSizeColumn(i);
                sheet.setColumnWidth(i, sheet.getColumnWidth(i) * 17 / 10);
            }

            while (scnLine.hasNextLine()) {
                line = scnLine.nextLine();
                Scanner scnWord = new Scanner(line);

                XSSFRow row = sheet.createRow(rowPosition);
                if (line.contains(NAAM.getName())) {
//                    scnWord.next();
                    String name = "";
//                    while (scnWord.hasNext()) {
//                        name += scnWord.next() + " ";
//                    }
                    cell = row.createCell(NAAM.getCell());
                    cell.setCellType(CellType.STRING);
                    cell.setCellValue("Camilo");
                }

                if (line.contains(FACTUUR_DATUM.getName())) {
                    line = scnLine.nextLine();
                    scnWord = new Scanner(line);
                    String factuurDatum = scnWord.next();
                    LocalDate localDate = LocalDate.parse(factuurDatum, dateFormatter);
                    cell = row.createCell(FACTUUR_DATUM.getCell());
                    cell.setCellType(CellType.STRING);
                    cell.setCellValue(localDate.toString());

                    String factuurnummer = scnWord.next();
                    cell = row.createCell(FACTUUR_NUMMER.getCell());
                    cell.setCellType(CellType.NUMERIC);
                    cell.setCellValue(factuurnummer);
                    rowPosition++;
                }
            }
            wb.write(fileOut);
            fileOut.flush();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void mergePDFs(String mergePDFName) throws IOException {
        List<File> files;
        try (Stream<Path> paths = Files.walk(file.toPath())) {
            files = paths
                    .filter(Files::isRegularFile)
                    .filter(path -> !isHidden(path))
                    .filter(path -> !isMergedFile(path))
                    .filter(path -> path.toString().endsWith("pdf"))
                    .map(path -> path.toFile())
                    .collect(Collectors.toList());
        }

        PdfDocument mergedDoc = new PdfDocument(new PdfWriter(mergePDFName));
        PdfMerger merger = new PdfMerger(mergedDoc);

        for (File source : files) {
            PdfDocument sourcePdf = new PdfDocument(new PdfReader(source));
            merger.merge(sourcePdf, 1, sourcePdf.getNumberOfPages()).setCloseSourceDocuments(true);
            sourcePdf.close();
        }
        merger.close();
        mergedDoc.close();

        System.out.println("PDFs have been merged");
    }


    public static void main(String... args) {
        int exitCode = new CommandLine(new PDFMerger()).execute("-o");
        System.exit(exitCode);
    }


}

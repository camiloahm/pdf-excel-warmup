package wendy;

import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.kernel.utils.PdfMerger;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import picocli.CommandLine;
import picocli.CommandLine.Command;
import picocli.CommandLine.Option;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Scanner;
import java.util.concurrent.Callable;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import static wendy.FileUtil.isHidden;
import static wendy.FileUtil.isMergedFile;
import static wendy.SheetTemplate.*;


@Command(name = "wendy's pdf merger", mixinStandardHelpOptions = true, version = "1.0",
        description = "Merges pdf files and convert them into a xlsx file")
public class PDFMerger implements Callable<Integer> {

    @Option(names = {"-o", "--origin"}, description = "Directory with all the pdfs", interactive = true)
    private File file;


    static final DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("d-MM-yy");

    @Override
    public Integer call() throws Exception {
        List<InputStream> files;
        String mergePDFName = file.getAbsolutePath() + "/merged_files.pdf";
        String excelFileName = file.getAbsolutePath() + "/merged_excel" + "_" + System.currentTimeMillis() + ".xlsx";
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
            int rowPosition = 1;

            PDFTextStripperByArea stripper = new PDFTextStripperByArea();
            stripper.setSortByPosition(true);
            PDFTextStripper tStripper = new PDFTextStripper();
            String str = tStripper.getText(document);

            Scanner scnLine = new Scanner(str);
            String line;

            XSSFRow rowHeaders = sheet.createRow(0);

            for (SheetTemplate column : SheetTemplate.values()) {
                rowHeaders.createCell(column.getCell()).setCellValue(column.getName().toUpperCase());
            }

            for (int i = 0; i < SheetTemplate.values().length; i++) {
                sheet.autoSizeColumn(i);
                sheet.setColumnWidth(i, sheet.getColumnWidth(i) * 17 / 10);
            }

            XSSFRow row = sheet.createRow(rowPosition);
            while (scnLine.hasNextLine()) {
                line = scnLine.nextLine();
                Scanner scnWord = new Scanner(line);

                if (line.toUpperCase().contains(NAAM.getName().toUpperCase())) {
                    scnWord.next();
                    String name = "";
                    while (scnWord.hasNext()) {
                        name += scnWord.next() + " ";
                    }
                    XSSFCell cell = row.createCell(NAAM.getCell());
                    cell.setCellType(CellType.STRING);
                    cell.setCellValue(name);
                    continue;
                }

                if (line.toUpperCase().contains(TOTAL_INCL_BTW.getName().toUpperCase())) {
                    scnWord.next();
                    scnWord.next();
                    scnWord.next();
                    if (scnWord.hasNext()) {
                        String totalInclBTW = scnWord.next();
                        totalInclBTW = totalInclBTW.contains("€") ? totalInclBTW.replace("€", "") : totalInclBTW;
                        XSSFCell cell = row.createCell(TOTAL_INCL_BTW.getCell());
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue(totalInclBTW);
                    }
                    continue;
                }

                if (line.toUpperCase().contains(BTW_21.getName().toUpperCase())) {
                    scnWord.next();
                    scnWord.next();
                    scnWord.next();
                    if (scnWord.hasNext()) {
                        scnWord.next();
                        String btw = scnWord.next();
                        btw = btw.contains("€") ? btw.replace("€", "") : btw;
                        XSSFCell cell = row.createCell(BTW_21.getCell());
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue(btw);
                    }
                    continue;
                }

                if (line.toUpperCase().contains(BTW_9.getName().toUpperCase())) {
                    scnWord.next();
                    scnWord.next();
                    scnWord.next();
                    if (scnWord.hasNext()) {
                        scnWord.next();
                        String btw = scnWord.next();
                        btw = btw.contains("€") ? btw.replace("€", "") : btw;
                        XSSFCell cell = row.createCell(BTW_9.getCell());
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue(btw);
                    }
                    continue;
                }

                if (line.toUpperCase().contains(TOTAAL_BTW.getName().toUpperCase())) {
                    scnWord.next();
                    scnWord.next();
                    if (scnWord.hasNext()) {
                        scnWord.next();
                        String totalBtw = scnWord.next();
                        totalBtw = totalBtw.contains("€") ? totalBtw.replace("€", "") : totalBtw;
                        XSSFCell cell = row.createCell(TOTAAL_BTW.getCell());
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue(totalBtw);

                    }
                    continue;
                }

                if (line.toUpperCase().contains(VERZENDKOSTEN.getName().toUpperCase())) {
                    scnWord.next();
                    if (scnWord.hasNext()) {
                        String verzendkosten = scnWord.next();
                        verzendkosten = verzendkosten.contains("€") ? verzendkosten.replace("€", "") : verzendkosten;
                        XSSFCell cell = row.createCell(SheetTemplate.VERZENDKOSTEN.getCell());
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue(verzendkosten);
                    }
                    continue;
                }

                if (line.toUpperCase().contains(TE_BETALEN.getName().toUpperCase())) {
                    scnWord.next();
                    scnWord.next();
                    if (scnWord.hasNext()) {
                        String teBetalen = scnWord.next();
                        teBetalen = teBetalen.contains("€") ? teBetalen.replace("€", "") : teBetalen;
                        XSSFCell cell = row.createCell(SheetTemplate.TE_BETALEN.getCell());
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue(teBetalen);
                    }
                    continue;
                }

                if (line.toUpperCase().contains(FACTUUR_DATUM.getName().toUpperCase())) {
                    line = scnLine.nextLine();
                    scnWord = new Scanner(line);
                    String factuurDatum = "";
                    String factuurnummer = "";
                    if (scnWord.hasNext()) {
                        factuurDatum = scnWord.next();
                        try {
                            LocalDate localDate = LocalDate.parse(factuurDatum, dateFormatter);
                            factuurDatum = localDate.toString();
                        } catch (Exception e) {
                            factuurDatum = "";
                            factuurnummer = factuurDatum;
                        }
                    }

                    XSSFCell cellFactuurDatum = row.createCell(FACTUUR_DATUM.getCell());
                    cellFactuurDatum.setCellType(CellType.STRING);
                    cellFactuurDatum.setCellValue(factuurDatum);

                    if (!factuurDatum.trim().equals("")) {
                        if (scnWord.hasNext()) {
                            factuurnummer = scnWord.next();
                        }
                    }

                    XSSFCell cellFacturNummer = row.createCell(FACTUUR_NUMMER.getCell());
                    cellFacturNummer.setCellType(CellType.NUMERIC);
                    cellFacturNummer.setCellValue(factuurnummer);
                    rowPosition++;
                    row = sheet.createRow(rowPosition);
                    continue;
                }
            }
            wb.write(fileOut);
            fileOut.flush();
            System.out.println(excelFileName + " has been created");
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

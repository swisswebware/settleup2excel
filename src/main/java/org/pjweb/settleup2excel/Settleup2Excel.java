/*
 * Small exemple of program with command line specifications to parse a specific
 * text file into an excel sheet.
 *
 * @author Pierre-Jacques Dauvert 
 * @email pj.dauvert@gmail.com
 */
package org.pjweb.settleup2excel;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.Date;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import javafx.util.converter.DateTimeStringConverter;
import org.apache.commons.cli.BasicParser;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.OptionBuilder;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 *
 * @author avoghaii
 */
public class Settleup2Excel {

    private static final String DELIMITER = "a payé";

    public static void main(String[] args) throws ParseException {

        String inputFile;
        String outputFile;
        String delimiter = DELIMITER;
        Options options = generateOptions();
        CommandLineParser parser = new BasicParser();
        try {
            // parse the command line arguments
            CommandLine line = parser.parse(options, args, true);
            inputFile = line.getOptionValue('i');
            System.out.println("Parsing file: " + inputFile);
            if (line.hasOption('o')) {
                outputFile = line.getOptionValue('o');
            } else {
                outputFile = inputFile.substring(0, inputFile.lastIndexOf(".") + 1) + "xls";
                System.out.println("No output specified, writing result to: " + outputFile);
            }
            if (line.hasOption('d')) {
                delimiter = line.getOptionValue('d');
            }
            Set<Payment> payments = parseInput(inputFile, delimiter);
            writeOutput(outputFile, payments);

        } catch (ParseException exp) {
            HelpFormatter formatter = new HelpFormatter();
            formatter.printHelp("settleup2excel", options);
        }

    }
    
    private static Options generateOptions() {
        Options opts = new Options();

        opts.addOption(OptionBuilder.withArgName("file")
                .hasArg()
                .isRequired()
                .withDescription("the input File")
                .create('i'));
        opts.addOption(OptionBuilder.withArgName("file")
                .hasArg()
                .withDescription("the output File")
                .create('o'));
        opts.addOption(OptionBuilder.withArgName("delimiter")
                .hasArg()
                .withDescription("delimiter string")
                .create('d'));

        return opts;
    }

    private static Set<Payment> parseInput(String input, String delimiter) {
        
        //for accessing innerclass initialization
        Settleup2Excel s2e = new Settleup2Excel();
        //container of inner classes instances
        Set<Payment> payments = new HashSet<>();
        
        //file Parsing
        Path path = Paths.get(input);
        try (Stream<String> lines = Files.lines(path)) {
            Iterator<String> it = lines.iterator();
            it.forEachRemaining(line -> {
                    if (line.contains(delimiter)) {
                        String[] elements = line.split(delimiter);
                        String payer = elements[0].trim();
                        String[] amount = elements[1].trim().split(" ");
                        Payment p = s2e.new Payment(payer, Double.parseDouble(amount[1]), amount[0]);
                        payments.add(p);
                        line = it.next();
                        elements = line.split(" ");
                        p.setDate(parseDate(elements[elements.length - 2] + " " + elements[elements.length - 1]));
                        p.setCategory(elements[0]);
                        p.setComment(Arrays.stream(elements, 1, elements.length - 2).filter(c -> !c.matches("-")).collect(Collectors.joining(" ")));
                    }
            });
            //payments.forEach(System.out::println);

        } catch (IOException ioe) {
            System.err.println("Input file could not be read : " + ioe.getMessage());
        }
        return payments;
    }

    private static Date parseDate(String sequence) {
        return new DateTimeStringConverter("dd.MM.yy HH:mm").fromString(sequence);
    }

    private static void writeOutput(String outputFile, Set<Payment> payments) {

        Path out = Paths.get(outputFile);

        try {
            //creation of output file (overwriting)
            Files.deleteIfExists(out);
            Files.createFile(out);
            OutputStream outputStream = Files.newOutputStream(out);

            //generation of workbook
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.createSheet("Accounting");

            //generation of row counter
            AtomicInteger index = new AtomicInteger(0);
            // creation of head row
            HSSFRow headRow = sheet.createRow(index.get());
            headRow.createCell(headRow.getPhysicalNumberOfCells()).setCellValue("Payer");
            headRow.createCell(headRow.getPhysicalNumberOfCells()).setCellValue("Amount");
            headRow.createCell(headRow.getPhysicalNumberOfCells()).setCellValue("Currency");
            headRow.createCell(headRow.getPhysicalNumberOfCells()).setCellValue("Date");
            headRow.createCell(headRow.getPhysicalNumberOfCells()).setCellValue("Category");
            headRow.createCell(headRow.getPhysicalNumberOfCells()).setCellValue("Comment");

            //creation of Date cell format
            final HSSFCellStyle cellStyle = wb.createCellStyle();
            final HSSFDataFormat hssfDataFormat = wb.createDataFormat();
            cellStyle.setDataFormat(hssfDataFormat.getFormat("dd/mm/yyyy hh:mm"));
            cellStyle.setShrinkToFit(true);

            // creation of rows
            payments.forEach(p -> {
                HSSFRow row = sheet.createRow(index.incrementAndGet());
                row.createCell(row.getPhysicalNumberOfCells()).setCellValue(p.getPayer());
                row.createCell(row.getPhysicalNumberOfCells()).setCellValue(p.getAmount());
                row.createCell(row.getPhysicalNumberOfCells()).setCellValue(p.getCurrency());
                HSSFCell dateCell = row.createCell(row.getPhysicalNumberOfCells());
                dateCell.setCellValue(p.getDate());
                dateCell.setCellStyle(cellStyle);
                row.createCell(row.getPhysicalNumberOfCells()).setCellValue(p.getCategory());
                row.createCell(row.getPhysicalNumberOfCells()).setCellValue(p.getComment());
            });

            // writing of worrkbook to output
            wb.write(outputStream);
            outputStream.close();
        } catch (IOException ex) {
            System.err.println("Error writing file: " + ex);
        }
        System.out.println("File has been written here: " + outputFile);
    }

    class Payment {

        private String payer, currency, category, comment;
        private Double amount;
        private Date date;

        public Payment(String payer, Double amount, String currency) {
            this.payer = payer;
            this.amount = amount;
            this.currency = currency;
        }

        public String getComment() {
            return comment;
        }

        public void setComment(String comment) {
            this.comment = comment;
        }

        public String getCategory() {
            return category;
        }

        public void setCategory(String category) {
            this.category = category;
        }

        public Date getDate() {
            return date;
        }

        public void setDate(Date date) {
            this.date = date;
        }

        public String getPayer() {
            return payer;
        }

        public void setPayer(String payer) {
            this.payer = payer;
        }

        public Double getAmount() {
            return amount;
        }

        public void setAmount(Double amount) {
            this.amount = amount;
        }

        public String getCurrency() {
            return currency;
        }

        public void setCurrency(String currency) {
            this.currency = currency;
        }

        @Override
        public String toString() {
            return "Payment{"
                    + "payer=" + payer
                    + ", amount=" + amount
                    + ", currency=" + currency
                    + ", category=" + category
                    + ", comment=" + comment
                    + ", date=" + date + '}';
        }

    }
}

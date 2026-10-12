package com.massmailer.service;

import com.massmailer.model.Contact;
import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@Service
public class ContactReader {

    public List<Contact> readContacts(MultipartFile file) throws IOException {
        String filename = file.getOriginalFilename().toLowerCase();
        if (filename.endsWith(".csv")) {
            return readCsv(file.getInputStream());
        } else if (filename.endsWith(".xlsx") || filename.endsWith(".xls")) {
            return readExcel(file.getInputStream());
        } else {
            throw new IllegalArgumentException("Unsupported file type. Please upload .xlsx or .csv");
        }
    }

    private List<Contact> readCsv(InputStream inputStream) {
        List<Contact> contacts = new ArrayList<>();
        try (CSVReader reader = new CSVReader(new InputStreamReader(inputStream))) {
            List<String[]> rows = reader.readAll();
            System.out.println("Processing CSV. Total rows: " + rows.size());

            int start = 0;
            // Heuristic: If first row has "college" or "email", treat as header.
            // Otherwise, assume it's data.
            if (!rows.isEmpty() && rows.get(0).length >= 2) {
                String c0 = rows.get(0)[0].toLowerCase();
                String c1 = rows.get(0)[1].toLowerCase();
                if (c0.contains("college") || c1.contains("email") || c0.contains("name")) {
                    start = 1;
                    System.out.println("Header row detected.");
                }
            }

            for (int i = start; i < rows.size(); i++) {
                String[] row = rows.get(i);
                if (row.length < 2)
                    continue; // Skip single column rows

                // Try to allow loose column mapping? For now stick to 0 and 1 but trim heavily.
                String name = row[0].trim();
                String email = row[1].trim();

                if (!email.isEmpty() && email.contains("@")) {
                    contacts.add(new Contact(name, email));
                }
            }
        } catch (Exception e) {
            System.err.println("CSV Parse Error: " + e.getMessage());
            throw new RuntimeException("Failed to parse CSV: " + e.getMessage());
        }
        System.out.println("CSV parsed. Contacts found: " + contacts.size());
        return contacts;
    }

    private List<Contact> readExcel(InputStream inputStream) {
        List<Contact> contacts = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);
            System.out.println("Processing Excel. Rows: " + sheet.getPhysicalNumberOfRows());

            Iterator<Row> rowIterator = sheet.iterator();
            boolean headerSkipped = false;

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // Better header check for Excel
                if (!headerSkipped) {
                    Cell c0 = row.getCell(0);
                    Cell c1 = row.getCell(1);
                    String s0 = getCellValueAsString(c0).toLowerCase();
                    String s1 = getCellValueAsString(c1).toLowerCase();

                    if (s0.contains("college") || s1.contains("email") || s0.contains("name")) {
                        headerSkipped = true;
                        continue;
                    }
                    // If it doesn't look like a header, treat as data?
                    // Or just blindly skip first row if user insists?
                    // Let's assume if it DOESN'T look like a header, it might be data, UNLESS it's
                    // the very first row and we want to be safe.
                    // A safe bet is: if it looks like a header, skip. If not, use it.
                    // But usually row 0 IS a header. Let's stick to heuristic.
                    // IF we are at row index 0 (which iterator doesn't tell us easily without
                    // counter), let's just use the boolean.
                    // Actually logic above: if matches header keywords, skip. Else process.
                    // Wait, if first row is "My College, me@test.com", it won't match
                    // "college/email", so we process it. Good.
                    headerSkipped = true;
                    // Fall through to process this row
                }

                String name = getCellValueAsString(row.getCell(0));
                String email = getCellValueAsString(row.getCell(1));

                if (!email.isEmpty() && email.contains("@")) {
                    contacts.add(new Contact(name, email));
                }
            }
        } catch (Exception e) {
            System.err.println("Excel Parse Error: " + e.getMessage());
            throw new RuntimeException("Failed to parse Excel file: " + e.getMessage());
        }
        System.out.println("Excel parsed. Contacts found: " + contacts.size());
        return contacts;
    }

    private String getCellValueAsString(Cell cell) {
        if (cell == null)
            return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            default:
                return "";
        }
    }
}

package com.massmailer.service;

import jakarta.annotation.PostConstruct;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.List;
import java.util.Optional;
import java.util.Set;

/**
 * Local SQLite record of every address we have emailed or found invalid, so a
 * company is never contacted twice.
 */
@Service
public class SentLog {

    public static final String SENT = "SENT";
    public static final String INVALID = "INVALID";
    public static final String FAILED = "FAILED"; // transient, may be retried

    private static final String DB_URL = "jdbc:sqlite:" + System.getProperty("user.dir") + "/mailer.db";

    private static final List<Path> LEGACY_SENT_FILES = List.of(
            Path.of(System.getProperty("user.dir"), "sent_emails.txt"),
            Path.of(System.getProperty("user.dir"), "faculty_contact", "sent_emails.txt"));

    // Shared mail providers: many unrelated people, so the domain says nothing about the company.
    private static final Set<String> FREE_DOMAINS = Set.of(
            "gmail.com", "yahoo.com", "yahoo.in", "yahoo.co.in", "outlook.com", "hotmail.com",
            "rediffmail.com", "live.com", "icloud.com", "protonmail.com", "aol.com");

    @PostConstruct
    public void init() throws SQLException {
        try (Connection c = connect(); Statement s = c.createStatement()) {
            s.execute("CREATE TABLE IF NOT EXISTS contacts ("
                    + "email TEXT PRIMARY KEY, "
                    + "company TEXT, "
                    + "company_key TEXT, "
                    + "domain TEXT, "
                    + "status TEXT NOT NULL, "
                    + "detail TEXT, "
                    + "updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP)");
            s.execute("CREATE INDEX IF NOT EXISTS idx_contacts_domain ON contacts(domain, status)");
            s.execute("CREATE INDEX IF NOT EXISTS idx_contacts_company ON contacts(company_key, status)");
        }
        importLegacyFiles();
    }

    /** Returns why this contact should be skipped, or empty if it is safe to send. */
    public synchronized Optional<String> skipReason(String email, String company) {
        String domain = domainOf(email);
        String key = companyKey(company);
        try (Connection c = connect()) {
            try (PreparedStatement ps = c.prepareStatement("SELECT status FROM contacts WHERE email = ?")) {
                ps.setString(1, email);
                try (ResultSet rs = ps.executeQuery()) {
                    if (rs.next() && !FAILED.equals(rs.getString(1))) {
                        return Optional.of(SENT.equals(rs.getString(1))
                                ? "already emailed this address" : "address previously marked invalid");
                    }
                }
            }
            if (!FREE_DOMAINS.contains(domain)) {
                try (PreparedStatement ps = c.prepareStatement(
                        "SELECT email FROM contacts WHERE domain = ? AND status = 'SENT' LIMIT 1")) {
                    ps.setString(1, domain);
                    try (ResultSet rs = ps.executeQuery()) {
                        if (rs.next()) {
                            return Optional.of("already emailed " + domain + " (" + rs.getString(1) + ")");
                        }
                    }
                }
            }
            if (!key.isEmpty()) {
                try (PreparedStatement ps = c.prepareStatement(
                        "SELECT email FROM contacts WHERE company_key = ? AND status = 'SENT' LIMIT 1")) {
                    ps.setString(1, key);
                    try (ResultSet rs = ps.executeQuery()) {
                        if (rs.next()) {
                            return Optional.of("already emailed company '" + company + "' (" + rs.getString(1) + ")");
                        }
                    }
                }
            }
        } catch (SQLException e) {
            throw new IllegalStateException("Sent-log lookup failed: " + e.getMessage(), e);
        }
        return Optional.empty();
    }

    public synchronized void record(String email, String company, String status, String detail) {
        try (Connection c = connect(); PreparedStatement ps = c.prepareStatement(
                "INSERT INTO contacts (email, company, company_key, domain, status, detail) VALUES (?,?,?,?,?,?) "
                        + "ON CONFLICT(email) DO UPDATE SET company = excluded.company, "
                        + "company_key = excluded.company_key, domain = excluded.domain, "
                        + "status = excluded.status, detail = excluded.detail, updated_at = CURRENT_TIMESTAMP")) {
            ps.setString(1, email);
            ps.setString(2, company);
            ps.setString(3, companyKey(company));
            ps.setString(4, domainOf(email));
            ps.setString(5, status);
            ps.setString(6, detail);
            ps.executeUpdate();
        } catch (SQLException e) {
            throw new IllegalStateException("Sent-log write failed: " + e.getMessage(), e);
        }
    }

    /** Row counts per status, e.g. {SENT=620, INVALID=3}. */
    public synchronized java.util.Map<String, Integer> totals() {
        java.util.Map<String, Integer> result = new java.util.LinkedHashMap<>();
        result.put(SENT, 0);
        result.put(INVALID, 0);
        result.put(FAILED, 0);
        try (Connection c = connect(); Statement s = c.createStatement();
                ResultSet rs = s.executeQuery("SELECT status, COUNT(*) FROM contacts GROUP BY status")) {
            while (rs.next()) {
                result.put(rs.getString(1), rs.getInt(2));
            }
        } catch (SQLException e) {
            throw new IllegalStateException("Sent-log count failed: " + e.getMessage(), e);
        }
        return result;
    }

    /**
     * Everything the dashboard sidebar shows. Rows imported from the old text files are
     * left out of today / last-7-days / recent, because their timestamp is the import time.
     */
    public synchronized java.util.Map<String, Object> dashboard() {
        java.util.Map<String, Object> out = new java.util.LinkedHashMap<>();
        out.put("totals", totals());
        String real = "detail IS NULL OR detail NOT LIKE 'imported%'";
        try (Connection c = connect(); Statement s = c.createStatement()) {
            try (ResultSet rs = s.executeQuery("SELECT COUNT(*) FROM contacts WHERE status='SENT' AND (" + real
                    + ") AND date(updated_at,'localtime') = date('now','localtime')")) {
                out.put("today", rs.next() ? rs.getInt(1) : 0);
            }

            java.util.Map<String, Integer> perDay = new java.util.HashMap<>();
            try (ResultSet rs = s.executeQuery("SELECT date(updated_at,'localtime'), COUNT(*) FROM contacts "
                    + "WHERE status='SENT' AND (" + real + ") "
                    + "AND date(updated_at,'localtime') >= date('now','localtime','-6 days') GROUP BY 1")) {
                while (rs.next()) perDay.put(rs.getString(1), rs.getInt(2));
            }
            java.util.List<java.util.Map<String, Object>> daily = new java.util.ArrayList<>();
            java.time.LocalDate today = java.time.LocalDate.now();
            for (int i = 6; i >= 0; i--) {
                java.time.LocalDate d = today.minusDays(i);
                daily.add(java.util.Map.of("date", d.toString(), "count", perDay.getOrDefault(d.toString(), 0)));
            }
            out.put("daily", daily);

            java.util.List<java.util.Map<String, Object>> recent = new java.util.ArrayList<>();
            try (ResultSet rs = s.executeQuery("SELECT email, company, status, updated_at FROM contacts WHERE ("
                    + real + ") ORDER BY updated_at DESC, rowid DESC LIMIT 8")) {
                while (rs.next()) {
                    java.util.Map<String, Object> row = new java.util.LinkedHashMap<>();
                    row.put("email", rs.getString(1));
                    row.put("company", rs.getString(2));
                    row.put("status", rs.getString(3));
                    row.put("at", rs.getString(4)); // UTC, "yyyy-MM-dd HH:mm:ss"
                    recent.add(row);
                }
            }
            out.put("recent", recent);
        } catch (SQLException e) {
            throw new IllegalStateException("Dashboard query failed: " + e.getMessage(), e);
        }
        return out;
    }

    private void importLegacyFiles() {
        for (Path file : LEGACY_SENT_FILES) {
            if (!Files.exists(file)) continue;
            int added = 0;
            try (Connection c = connect(); PreparedStatement ps = c.prepareStatement(
                    "INSERT OR IGNORE INTO contacts (email, company, company_key, domain, status, detail) "
                            + "VALUES (?,?,?,?,?,?)")) {
                for (String line : Files.readAllLines(file)) {
                    String email = EmailValidator.normalize(line);
                    if (email.isEmpty() || !email.contains("@")) continue;
                    ps.setString(1, email);
                    ps.setString(2, "");
                    ps.setString(3, "");
                    ps.setString(4, domainOf(email));
                    ps.setString(5, SENT);
                    ps.setString(6, "imported from " + file.getFileName());
                    added += ps.executeUpdate();
                }
            } catch (IOException | SQLException e) {
                System.err.println("Could not import " + file + ": " + e.getMessage());
            }
            System.out.println("SentLog: imported " + added + " new addresses from " + file);
        }
    }

    static boolean isSharedDomain(String domain) {
        return FREE_DOMAINS.contains(domain);
    }

    static String domainOf(String email) {
        int at = email.lastIndexOf('@');
        return at < 0 ? "" : email.substring(at + 1).toLowerCase();
    }

    static String companyKey(String company) {
        if (company == null) return "";
        String key = company.toLowerCase().replaceAll("[^a-z0-9 ]", " ").replaceAll("\\s+", " ").trim();
        return key.equals("your team") ? "" : key;
    }

    private Connection connect() throws SQLException {
        return DriverManager.getConnection(DB_URL);
    }
}

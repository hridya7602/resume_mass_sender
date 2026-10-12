package com.massmailer.service;

import com.massmailer.model.Contact;
import jakarta.annotation.PostConstruct;
import jakarta.mail.internet.MimeMessage;
import jakarta.mail.util.ByteArrayDataSource;
import org.eclipse.angus.mail.smtp.SMTPAddressFailedException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.JavaMailSenderImpl;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.Properties;
import java.util.concurrent.atomic.AtomicBoolean;

@Service
public class MailService {

    private JavaMailSenderImpl mailSender;
    private List<String> logs = new ArrayList<>();
    private AtomicBoolean isRunning = new AtomicBoolean(false);
    private int sentCount = 0;
    private int totalCount = 0;
    private int skippedCount = 0;

    @Autowired
    private SentLog sentLog;

    @Autowired
    private EmailValidator emailValidator;

    // Config options
    private int delaySeconds = 3;

    @Value("${app.smtp.host}")
    private String smtpHost;

    @Value("${app.smtp.port}")
    private int smtpPort;

    @Value("${app.smtp.username}")
    private String smtpUser;

    @Value("${app.smtp.password}")
    private String smtpPassword;

    // File paths
    private File resumeFile;
    private List<Contact> currentContacts;

    @PostConstruct
    public void init() {
        mailSender = new JavaMailSenderImpl();
        mailSender.setHost(smtpHost);
        mailSender.setPort(smtpPort);
        mailSender.setUsername(smtpUser);
        mailSender.setPassword(smtpPassword);

        Properties props = mailSender.getJavaMailProperties();
        props.put("mail.transport.protocol", "smtp");
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.debug", "true"); // Enable debug for troubleshooting

        log("MailService initialized for user: " + smtpUser);
    }

    public void setResumeFile(File file) {
        this.resumeFile = file;
        log("Resume file uploaded: " + file.getName());
    }

    public void setContacts(List<Contact> contacts) {
        this.currentContacts = contacts;
        this.totalCount = contacts.size();
        log("Contacts loaded: " + totalCount);
    }

    public void setDelay(int seconds) {
        this.delaySeconds = seconds;
    }

    public void stopSending() {
        if (isRunning.compareAndSet(true, false)) {
            log("Stopping email sending (requested by user)...");
        }
    }

    public void startSending() {
        if (isRunning.get()) {
            throw new RuntimeException("Already sending emails!");
        }
        if (mailSender == null) {
            throw new RuntimeException("SMTP not configured!");
        }
        if (currentContacts == null || currentContacts.isEmpty()) {
            throw new RuntimeException("No contacts loaded!");
        }
        if (resumeFile == null || !resumeFile.exists()) {
            throw new RuntimeException("Resume file not found!");
        }

        isRunning.set(true);
        sentCount = 0;
        skippedCount = 0;
        logs.clear();
        log("Starting email blast to " + totalCount + " contacts...");

        new Thread(() -> {
            try {
                for (Contact contact : currentContacts) {
                    if (!isRunning.get())
                        break;

                    String email = EmailValidator.normalize(contact.getEmailId());
                    String company = contact.getCollegeName();

                    Optional<String> invalid = emailValidator.problem(email);
                    if (invalid.isPresent()) {
                        skippedCount++;
                        sentLog.record(email, company, SentLog.INVALID, invalid.get());
                        log("- Skipped " + email + ": " + invalid.get());
                        continue;
                    }
                    Optional<String> duplicate = sentLog.skipReason(email, company);
                    if (duplicate.isPresent()) {
                        skippedCount++;
                        log("- Skipped " + email + ": " + duplicate.get());
                        continue;
                    }

                    try {
                        contact.setEmailId(email);
                        sendEmail(contact);
                        sentLog.record(email, company, SentLog.SENT, null);
                        sentCount++;
                        log("✓ Successfully sent to: " + email + " (" + company + ")");
                    } catch (Exception e) {
                        boolean permanent = isPermanentRecipientFailure(e);
                        sentLog.record(email, company, permanent ? SentLog.INVALID : SentLog.FAILED, e.getMessage());
                        log("✗ Failed to send to: " + email + " - " + e.getMessage());
                        recordFailure(contact, e.getMessage());
                        if (permanent) continue; // rejected address, no point waiting
                    }

                    try {
                        Thread.sleep(delaySeconds * 1000L);
                    } catch (InterruptedException e) {
                        Thread.currentThread().interrupt();
                        break;
                    }
                }
            } finally {
                isRunning.set(false);
                log("Email sending process completed.");
            }
        }).start();
    }

    private void sendEmail(Contact contact) throws Exception {
        MimeMessage message = mailSender.createMimeMessage();
        MimeMessageHelper helper = new MimeMessageHelper(message, true);

        helper.setFrom(smtpUser);
        helper.setTo(contact.getEmailId());
        helper.setSubject("Software Engineer Opportunities - Diya Dathan");

        String body = getEmailTemplate(contact.getCollegeName());
        helper.setText(body, false); // False for plain text, true for HTML

        helper.addAttachment("Diya_Dathan_Resume.pdf", resumeFile);

        mailSender.send(message);
    }

    private static final java.util.regex.Pattern INSTITUTION_KEYWORDS = java.util.regex.Pattern.compile(
        "college|university|institute|school|academy|campus|polytechnic|vidyalaya|vidyapeeth|mahavidyalaya|deemed",
        java.util.regex.Pattern.CASE_INSENSITIVE
    );
    private static final java.util.regex.Pattern JOB_TITLE_ONLY = java.util.regex.Pattern.compile(
        "^(assistant|associate|professor|lecturer|faculty|hod|principal|teacher|catering|manager|staff|non.teaching|recruitment|job|opening|wanted).*",
        java.util.regex.Pattern.CASE_INSENSITIVE
    );
    private static final java.util.regex.Pattern TRAILING_NOISE = java.util.regex.Pattern.compile(
        "\\s*[\\-–]?\\s*\\d*\\s*facult\\w*.*$|\\s*[\\-–]\\s*\\d{4}.*$|\\s*\\d{4}\\s*facult\\w*.*$",
        java.util.regex.Pattern.CASE_INSENSITIVE
    );

    private String sanitizeCollegeName(String name) {
        if (name == null || name.isBlank()) return "your team";
        String trimmed = TRAILING_NOISE.matcher(name.trim()).replaceAll("").strip();
        if (trimmed.isBlank()) return "your team";
        if (INSTITUTION_KEYWORDS.matcher(trimmed).find()) return trimmed;
        if (JOB_TITLE_ONLY.matcher(trimmed).matches()) return "your team";
        return trimmed;
    }

    private String getEmailTemplate(String companyName) {
        String company = sanitizeCollegeName(companyName);
        return "Dear Hiring Team,\n\n" +
                "I'm a Software Engineer with one year of experience building backend services, RESTful APIs, and full-stack applications. I am writing to apply for a Software Engineer position "
                + (company.equals("your team") ? "at your organization" : "at " + company) + ".\n\n" +
                "Technical background:\n" +
                "• Languages: Python, Java, C++, JavaScript/TypeScript, SQL\n" +
                "• Backend: FastAPI, Node.js, Express, REST API design\n" +
                "• Frontend: React.js, React Native, Expo\n" +
                "• Databases & Cloud: PostgreSQL, MongoDB, AWS (EC2, RDS), CI/CD, Git\n\n" +
                "I've worked on production systems end-to-end, including a full-suite CRM platform, an analytics dashboard for Kochi Metro's ridership and payment data, and a backend for a consumer mobile app with live tracking and ticketing. I also have experience with SQL/schema optimization and AWS deployment management.\n\n"
                +
                "My resume is attached for your consideration. I would appreciate the opportunity to discuss how my skills and experience align with your Software Engineer requirements.\n\n"
                +
                "Best regards,\n" +
                "Diya Dathan\n" +
                "8921920068 | diyaadathan@gmail.com";
    }

    // A 5xx reply for the recipient means the address is bad; connection/auth errors are transient.
    private static boolean isPermanentRecipientFailure(Throwable e) {
        for (Throwable t = e; t != null; t = t.getCause()) {
            if (t instanceof SMTPAddressFailedException a && a.getReturnCode() >= 500) {
                return true;
            }
        }
        return false;
    }

    private void log(String message) {
        System.out.println(message);
        logs.add(message);
    }

    private void recordFailure(Contact contact, String error) {
        try (FileWriter fw = new FileWriter("failed.txt", true)) {
            fw.write(contact.getEmailId() + "|" + contact.getCollegeName() + "|" + error + "\n");
        } catch (IOException e) {
            System.err.println("Could not write to failed.txt");
        }
    }

    public List<String> getLogs() {
        return logs;
    }

    public String getStatus() {
        return isRunning.get() ? "Sending..." : "Idle";
    }

    public int getSentCount() {
        return sentCount;
    }

    /**
     * Labels each uploaded contact READY, DUPLICATE (would be skipped) or INVALID, using the same
     * rules the send loop applies, minus the DNS lookup so the upload stays fast.
     */
    public List<java.util.Map<String, Object>> previewContacts(List<Contact> contacts) {
        List<java.util.Map<String, Object>> rows = new ArrayList<>();
        java.util.Set<String> seenEmails = new java.util.HashSet<>();
        java.util.Set<String> seenDomains = new java.util.HashSet<>();
        for (Contact c : contacts) {
            String email = EmailValidator.normalize(c.getEmailId());
            String status = "READY";
            String reason = "";

            Optional<String> bad = emailValidator.syntaxProblem(email);
            if (bad.isPresent()) {
                status = "INVALID";
                reason = bad.get();
            } else {
                Optional<String> dup = sentLog.skipReason(email, c.getCollegeName());
                String domain = SentLog.domainOf(email);
                if (dup.isPresent()) {
                    status = "DUPLICATE";
                    reason = dup.get();
                } else if (!seenEmails.add(email)) {
                    status = "DUPLICATE";
                    reason = "listed twice in this file";
                } else if (!SentLog.isSharedDomain(domain) && !seenDomains.add(domain)) {
                    status = "DUPLICATE";
                    reason = "same company as an earlier row";
                }
            }

            java.util.Map<String, Object> row = new java.util.LinkedHashMap<>();
            row.put("company", c.getCollegeName());
            row.put("email", email);
            row.put("status", status);
            row.put("reason", reason);
            rows.add(row);
        }
        return rows;
    }

    public java.util.Map<String, Object> getDashboard() {
        return sentLog.dashboard();
    }

    public java.util.Map<String, Integer> getTotals() {
        return sentLog.totals();
    }

    public int getSkippedCount() {
        return skippedCount;
    }

    public int getTotalCount() {
        return totalCount;
    }
}
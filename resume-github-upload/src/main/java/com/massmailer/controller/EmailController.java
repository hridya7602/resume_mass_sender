package com.massmailer.controller;

import com.massmailer.model.Contact;
import com.massmailer.service.ContactReader;
import com.massmailer.service.MailService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api")
@CrossOrigin(origins = "*") // Allow frontend to call
public class EmailController {

    @Autowired
    private ContactReader contactReader;

    @Autowired
    private MailService mailService;

    private static final String UPLOAD_DIR = System.getProperty("user.dir") + "/uploads/";

    public EmailController() {
        new File(UPLOAD_DIR).mkdirs();
    }

    @PostMapping("/upload-contacts")
    public ResponseEntity<?> uploadContacts(@RequestParam("file") MultipartFile file) {
        try {
            List<Contact> contacts = contactReader.readContacts(file);
            mailService.setContacts(contacts);
            List<Map<String, Object>> rows = mailService.previewContacts(contacts);
            long ready = rows.stream().filter(r -> "READY".equals(r.get("status"))).count();
            long invalid = rows.stream().filter(r -> "INVALID".equals(r.get("status"))).count();
            return ResponseEntity.ok(Map.of(
                    "message", "Contacts loaded",
                    "count", contacts.size(),
                    "summary", Map.of("ready", ready, "duplicate", rows.size() - ready - invalid, "invalid", invalid),
                    "contacts", rows));
        } catch (IOException | IllegalArgumentException e) {
            return ResponseEntity.internalServerError().body("Error reading file: " + e.getMessage());
        }
    }

    @PostMapping("/upload-resume")
    public ResponseEntity<?> uploadResume(@RequestParam("file") MultipartFile file) {
        try {
            Path targetPath = Path.of(UPLOAD_DIR + file.getOriginalFilename());
            Files.copy(file.getInputStream(), targetPath, StandardCopyOption.REPLACE_EXISTING);
            mailService.setResumeFile(targetPath.toFile());
            return ResponseEntity.ok("Resume uploaded: " + file.getOriginalFilename());
        } catch (IOException e) {
            return ResponseEntity.internalServerError().body("Error saving file: " + e.getMessage());
        }
    }

    @PostMapping("/send-emails")
    public ResponseEntity<?> sendEmails(@RequestBody Map<String, Integer> payload) {
        try {
            int delay = payload.getOrDefault("delay", 3);
            mailService.setDelay(delay);
            mailService.startSending();
            return ResponseEntity.ok("Sending started");
        } catch (Exception e) {
            return ResponseEntity.badRequest().body(e.getMessage());
        }
    }

    @PostMapping("/stop-sending")
    public ResponseEntity<?> stopSending() {
        mailService.stopSending();
        return ResponseEntity.ok("Stopping");
    }

    @GetMapping("/dashboard")
    public ResponseEntity<?> getDashboard() {
        return ResponseEntity.ok(mailService.getDashboard());
    }

    @GetMapping("/stats")
    public ResponseEntity<?> getStats() {
        return ResponseEntity.ok(mailService.getTotals());
    }

    @GetMapping("/status")
    public ResponseEntity<?> getStatus() {
        Map<String, Object> status = new HashMap<>();
        status.put("status", mailService.getStatus());
        status.put("sent", mailService.getSentCount());
        status.put("skipped", mailService.getSkippedCount());
        status.put("total", mailService.getTotalCount());
        status.put("logs", mailService.getLogs());
        return ResponseEntity.ok(status);
    }
}

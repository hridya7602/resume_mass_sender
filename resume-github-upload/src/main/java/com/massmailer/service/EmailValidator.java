package com.massmailer.service;

import org.springframework.stereotype.Service;

import javax.naming.NameNotFoundException;
import javax.naming.NamingException;
import javax.naming.directory.Attribute;
import javax.naming.directory.Attributes;
import javax.naming.directory.DirContext;
import javax.naming.directory.InitialDirContext;
import java.util.Hashtable;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import java.util.regex.Pattern;

/**
 * Checks that an address is well-formed and that its domain can receive mail.
 * A mailbox itself cannot be verified without sending, so permanent bounces are
 * recorded separately in {@link SentLog}.
 */
@Service
public class EmailValidator {

    private static final Pattern SYNTAX = Pattern.compile(
            "^[a-z0-9._%+'-]+@([a-z0-9-]+\\.)+[a-z]{2,}$");
    private static final Set<String> PLACEHOLDER_DOMAINS = Set.of(
            "example.com", "example.org", "test.com", "domain.com", "email.com");

    private final Map<String, Boolean> domainCache = new ConcurrentHashMap<>();

    /** Strips markup, mailto: and whitespace that leak in from scraped data. */
    public static String normalize(String raw) {
        if (raw == null) return "";
        return raw.replaceAll("<[^>]*>", "")
                .replaceAll("(?i)^mailto:", "")
                .trim()
                .toLowerCase();
    }

    /** Fast offline checks only (no DNS), used for the upload preview. */
    public Optional<String> syntaxProblem(String email) {
        if (!SYNTAX.matcher(email).matches()) {
            return Optional.of("malformed address");
        }
        String domain = SentLog.domainOf(email);
        if (PLACEHOLDER_DOMAINS.contains(domain)) {
            return Optional.of("placeholder domain " + domain);
        }
        return Optional.empty();
    }

    /** Returns the reason the address is unusable, or empty if it looks deliverable. */
    public Optional<String> problem(String email) {
        Optional<String> syntax = syntaxProblem(email);
        if (syntax.isPresent()) {
            return syntax;
        }
        String domain = SentLog.domainOf(email);
        if (!domainCache.computeIfAbsent(domain, this::domainAcceptsMail)) {
            return Optional.of("domain " + domain + " has no mail server");
        }
        return Optional.empty();
    }

    // Only a definite "no such domain / no records" answer rejects; DNS errors let the send proceed.
    private boolean domainAcceptsMail(String domain) {
        Hashtable<String, String> env = new Hashtable<>();
        env.put("java.naming.factory.initial", "com.sun.jndi.dns.DnsContextFactory");
        env.put("com.sun.jndi.dns.timeout.initial", "3000");
        env.put("com.sun.jndi.dns.timeout.retries", "2");
        try {
            DirContext ctx = new InitialDirContext(env);
            try {
                Attributes attrs = ctx.getAttributes(domain, new String[] { "MX", "A" });
                Attribute mx = attrs.get("MX");
                Attribute a = attrs.get("A");
                return (mx != null && mx.size() > 0) || (a != null && a.size() > 0);
            } finally {
                ctx.close();
            }
        } catch (NameNotFoundException e) {
            return false;
        } catch (NamingException e) {
            return true;
        }
    }
}

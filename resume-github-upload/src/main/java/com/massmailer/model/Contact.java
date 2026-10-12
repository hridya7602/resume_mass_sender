package com.massmailer.model;

public class Contact {
    private String collegeName;
    private String emailId;

    public Contact() {}

    public Contact(String collegeName, String emailId) {
        this.collegeName = collegeName;
        this.emailId = emailId;
    }

    public String getCollegeName() {
        return collegeName;
    }

    public void setCollegeName(String collegeName) {
        this.collegeName = collegeName;
    }

    public String getEmailId() {
        return emailId;
    }

    public void setEmailId(String emailId) {
        this.emailId = emailId;
    }

    @Override
    public String toString() {
        return "Contact{collegeName='" + collegeName + "', emailId='" + emailId + "'}";
    }
}

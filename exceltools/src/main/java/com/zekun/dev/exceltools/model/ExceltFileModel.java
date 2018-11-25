package com.zekun.dev.exceltools.model;

import org.springframework.web.multipart.MultipartFile;

import java.io.Serializable;

public class ExceltFileModel implements Serializable {
    private static final long serialVersionUID = 5319387339018978269L;
    private int grant = 10;
    private String countTimePoint = "20:30";
    private MultipartFile namesFile;
    private MultipartFile workHours;

    public int getGrant() {
        return grant;
    }

    public void setGrant(int grant) {
        this.grant = grant;
    }

    public String getCountTimePoint() {
        return countTimePoint;
    }

    public void setCountTimePoint(String countTimePoint) {
        this.countTimePoint = countTimePoint;
    }

    public MultipartFile getNamesFile() {
        return namesFile;
    }

    public void setNamesFile(MultipartFile namesFile) {
        this.namesFile = namesFile;
    }

    public MultipartFile getWorkHours() {
        return workHours;
    }

    public void setWorkHours(MultipartFile workHours) {
        this.workHours = workHours;
    }
}

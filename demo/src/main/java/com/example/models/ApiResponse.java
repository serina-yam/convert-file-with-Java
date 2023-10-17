package com.example.models;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ApiResponse {

    /**
     * 値はenum Statusで管理.
     */
    private int status;
    private String message;
    private String filePath;

    public ApiResponse(int status, String message, String filePath) {
        this.status = status;
        this.message = message;
        this.filePath = filePath;
    }

}
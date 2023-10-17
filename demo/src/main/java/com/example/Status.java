package com.example;

public enum Status {
    OK(0),
    NG(1);

    private final int value;

    private Status(int value) {
        this.value = value;
    }

    public int getValue() {
        return value;
    }
}

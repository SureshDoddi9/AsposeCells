package com.suresh.model;

import lombok.Getter;

public enum Functionality {
    NUMBER_TO_TEXT("NumberToText");

    @Getter
    private final String function;
    Functionality(String s) {
        this.function = s;
    }
}

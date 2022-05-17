package com.suresh.model;

import lombok.Getter;

public enum Functionality {
    NUMBER_TO_TEXT("NumberToText"),
    TEXT_TO_NUMBER("TextToNumber"),
    RENAME_SHEET("RenameSheet"),
    SHEET_FORMAT("FormatSheet");

    @Getter
    private final String function;
    Functionality(String s) {
        this.function = s;
    }
}

package com.example.exceptions;

public class StartingRowExceededException extends RuntimeException {
    public StartingRowExceededException() {
        super("Starting row is greater, than physical amount of sheet rows.");
    }
}

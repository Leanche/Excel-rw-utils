package com.example.exceptions;

// This exception is thrown, if List page exceeds maximum
// amount of sheet rows according to Excel specification
public class ListTooBigException extends RuntimeException {
    private static final int MAX_SIZE = 1048575;

    public ListTooBigException(int listSize) {
        super("List size: " + listSize + " is too big for a sheet. Max size: " + MAX_SIZE);
    }
}

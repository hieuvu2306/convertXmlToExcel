package com.kms.convert.xml.xlsx;

import com.kms.convert.xml.xlsx.service.Converter;

public class Application {
    public static void main(String[] args) {
        Converter converter = new Converter();
        converter.writeToXLSXFile(args[0], args[1]);
    }
}

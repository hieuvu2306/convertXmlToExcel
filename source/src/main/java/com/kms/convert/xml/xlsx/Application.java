package com.kms.convert.xml.xlsx;

import java.io.IOException;
import java.util.Scanner;

import com.kms.convert.xml.xlsx.service.Converter;

public class Application {
    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in);
        System.out.println("xml path : ");
        String xmlFilePath = sc.nextLine();
        System.out.println("report location path : ");
        String reportFilePath = sc.nextLine();
        sc.close();
        Converter converter = new Converter();
        try {
            converter.writeToXLSXFile(xmlFilePath, reportFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

package com.kms.convert.xml.xlsx.temp;

public class Recursive {

    public int factorial(int n) {
        int result = 1;
        System.out.println("debug " + result);
        System.out.println(n);
        if (n == 0) {
            return 1;
        }
        result = n * factorial(n - 1);
        System.out.println(result);
        return result;
    }
}

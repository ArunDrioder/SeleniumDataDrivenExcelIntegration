package tests;

import java.io.File;

public class FileTest
{


    public static void main(String[] args)
    {

       File file = new File(System.getProperty("user.home")+"\\Downloads\\download.xlsx");
        System.out.println(file.exists());
    }
}

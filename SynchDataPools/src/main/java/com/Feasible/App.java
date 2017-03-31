package com.Feasible;

import java.io.IOException;
import java.util.Scanner;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
    	  Scanner sc=new Scanner(System.in);  
        System.out.println( "Please give the Folder path where Data pools are resided " );
        String RootPath = sc.next();  
		//String RootPath = "D:\\testFiles\\";
		SynchDataPools2 datapool = new SynchDataPools2();
		try {
			datapool.copyMaster(RootPath);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }
}

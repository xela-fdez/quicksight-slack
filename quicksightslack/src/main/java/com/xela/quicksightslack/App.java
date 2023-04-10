package com.xela.quicksightslack;

import java.io.IOException;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
    	String awsAccessKeyId = args[0];
		String awsSecretAccessKey = args[1];
		String awsRegion = args[2];					//"US-EAST-1"
		String dashboardId = args[3];				//"x12xx345-67x8-9x12-3xx4-5678x912345x"
		String AWSAccountId = args[4];				//"123456789123"
		String apiToken = args[5];					//xoxb-not-a-real-token-this-will-not-work
		String channel = args[6];					//"C69NRPF44MH"
		
		QuicksightDashboardDownloader quicksight = new QuicksightDashboardDownloader(awsAccessKeyId, awsSecretAccessKey, awsRegion);
		String fileName = quicksight.downloadDashboard(dashboardId, AWSAccountId);
		int sheetNumber = 0;

		String formattedFileName = "";
		try {
			formattedFileName =  ExcelFormatter.BasicFormatter(fileName,sheetNumber);
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		SlackUpload slack = new SlackUpload(apiToken, channel);
		try {
			slack.upload(formattedFileName, sheetNumber);
		} catch (Exception e) {
			e.printStackTrace();
		}
    }
}

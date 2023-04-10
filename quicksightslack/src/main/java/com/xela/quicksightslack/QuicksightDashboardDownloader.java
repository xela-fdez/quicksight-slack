package com.xela.quicksightslack;

import com.amazonaws.auth.AWSCredentials;
import com.amazonaws.auth.AWSCredentialsProvider;
import com.amazonaws.auth.AWSStaticCredentialsProvider;
import com.amazonaws.auth.BasicAWSCredentials;
import com.amazonaws.regions.Regions;
import com.amazonaws.services.quicksight.AmazonQuickSight;
import com.amazonaws.services.quicksight.AmazonQuickSightClientBuilder;
import com.amazonaws.services.quicksight.model.GenerateEmbedUrlForAnonymousUserRequest;
import com.amazonaws.services.quicksight.model.GenerateEmbedUrlForAnonymousUserResult;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;

public class QuicksightDashboardDownloader {
	private AmazonQuickSight quicksightClient;
	
	public QuicksightDashboardDownloader(String awsAccessKeyId, String awsSecretAccessKey, String awsRegion) {
		BasicAWSCredentials awsCreds = new BasicAWSCredentials(awsAccessKeyId, awsSecretAccessKey);
		quicksightClient = (AmazonQuickSight) ((AmazonQuickSightClientBuilder) ((AmazonQuickSightClientBuilder) AmazonQuickSightClientBuilder
				.standard()
				.withCredentials((AWSCredentialsProvider) new AWSStaticCredentialsProvider((AWSCredentials) awsCreds)))
				.withRegion(Regions.fromName(awsRegion))).build();
		
	}
	
	public String downloadDashboard(String dashboardId, String AWSAccountId) {
		String file = "dashboard.xlsx";
		GenerateEmbedUrlForAnonymousUserRequest generateEmbedUrlRequest = (new GenerateEmbedUrlForAnonymousUserRequest())
				.withAwsAccountId(AWSAccountId);
		GenerateEmbedUrlForAnonymousUserResult generateEmbedUrlResult = quicksightClient
				.generateEmbedUrlForAnonymousUser(generateEmbedUrlRequest).withEmbedUrl(dashboardId);
		String dashboardEmbedUrl = generateEmbedUrlResult.getEmbedUrl();
		try {
			URL url = new URL(dashboardEmbedUrl);
			InputStream is = url.openStream();
			FileOutputStream fos = new FileOutputStream(file);
			byte[] buffer = new byte[4096];
			int bytesRead = 0;
			while ((bytesRead = is.read(buffer)) != -1)
				fos.write(buffer, 0, bytesRead);
			fos.close();
			is.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println("Dashboard downloaded successfully.");
		
		return file;
	}
}
